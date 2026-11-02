from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
import json
import logging
import os
from pathlib import Path
import re
import uuid

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .employment_dates import date_text
from .errors import SchemaMismatch
from .infra import safe_name


logger = logging.getLogger(__name__)


BUSINESS_SHEET = "사업장정보"
YEAR_CELL = "K2"
TAX_SHEET_PATTERN = re.compile(r"^원천징수추출\((\d{2})년\)$")
PAYROLL_YEAR_PATTERN = re.compile(r"_(\d{2})년\s*급여대장\.xlsx$", re.IGNORECASE)
OUTPUT_SHEETS = ("전체인원", "대상자(작업)", "대상자", "비대상자")
BASE_HEADERS = ("사원번호", "성명", "주민등록번호", "입사일", "퇴사일", "지급년월", "급여구분")
TOTAL_HEADERS = ("지급합계", "공제합계", "차인지급액")
ANNUAL_TAX_HEADERS = (
    "근로소득",
    "결정세액_소득세",
    "기납부세액_소득세",
    "차감징수세액_소득세",
    "결정세액_지방세",
    "기납부세액_지방세",
    "차감징수세액_지방세",
)
REFERENCE_HEADERS = (
    "정액급여합계",
    "인덱스",
    "직업분류 코드",
    "직업분류 항목",
    "전년도과세소득_총계",
    "과세소득_총계",
    "비과세소득_총계",
    "비과세소득_현계",
    "경정청구 대상 여부",
)
NON_MONEY_REFERENCE_HEADERS = {"인덱스", "직업분류 코드", "직업분류 항목", "경정청구 대상 여부"}
NON_TAX_OVERTIME_HEADER = "야간연장수당(비과세)"
SOCIAL_INSURANCE_HEADERS = ("국민연금", "건강보험")
ANNUAL_NON_TAX_LIMIT = 2_400_000
FIXED_SALARY = "정액급여"
REVIEW_ALLOWANCE = "수당검토"
EXCLUDED = "제외"
UNCLASSIFIED = "미분류"
CONFLICT = "충돌"
VARIABLE_FIXED = "변동·부정기"
BASE_FIXED_ITEMS = {
    "기본급",
    "본봉",
    "기준급",
    "호봉급",
    "월급",
    "급료",
    "직무급",
    "직능급",
    "기본급여",
    "식대",
    "식사대",
    "중식대",
}
BASE_FIXED_KEYWORDS = ("식대", "식사대")
MONEY_FORMAT = "#,##0;-#,##0;"
THIN_GRAY = Side(style="thin", color="FF808080")
CELL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)


@dataclass(frozen=True)
class ItemDecision:
    item_name: str
    category: str
    reason: str


@dataclass(frozen=True)
class TaxRecord:
    employee_code: str
    name: str
    resident_id: str
    index: str
    taxable_total: int
    non_tax_total: int
    non_tax_current: int
    details: dict[str, int]


@dataclass(frozen=True)
class OccupationRecord:
    name: str
    resident_key: str
    code: str
    item: str
    acquired: str
    lost: str


@dataclass
class EmployeeData:
    key: tuple[str, str]
    code: str
    name: str
    resident_id: str
    monthly_rows: list[dict[str, object]] = field(default_factory=list)
    current_tax: TaxRecord | None = None
    prior_taxable: int | None = None
    occupation: OccupationRecord | None = None
    forced_reasons: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class PayrollStatusResult:
    years: tuple[int, ...]
    output_files: tuple[Path, ...]
    report_path: Path
    updated_basic_info_path: Path


def _normalized_name(value: object) -> str:
    text = re.sub(r"\([^)]*\)|\[[^]]*\]|\{[^}]*\}|<[^>]*>", "", str(value or ""))
    return re.sub(r"[^0-9A-Za-z가-힣]", "", text).casefold()


def _normalized_item_name(value: object) -> str:
    return re.sub(r"[^0-9A-Za-z가-힣]", "", str(value or "")).casefold()


_BASE_FIXED_ITEM_KEYS = frozenset(_normalized_item_name(name) for name in BASE_FIXED_ITEMS)
_BASE_FIXED_KEYWORD_KEYS = tuple(_normalized_item_name(keyword) for keyword in BASE_FIXED_KEYWORDS)


def _digits(value: object) -> str:
    return re.sub(r"\D", "", str(value or ""))


def _resident_id(value: object) -> str:
    digits = _digits(value)
    return digits if len(digits) == 13 else ""


def _resident_key(value: object) -> str:
    digits = _digits(value)
    return digits[:7] if len(digits) >= 7 else ""


def _employee_code(value: object) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _money(value: object) -> int:
    if value in (None, ""):
        return 0
    if isinstance(value, (int, float)):
        return int(round(value))
    text = re.sub(r"[₩$,\s]", "", str(value)).replace("(", "-").replace(")", "")
    try:
        return int(round(float(text)))
    except ValueError as exc:
        raise SchemaMismatch(f"금액으로 해석할 수 없습니다: {value!r}") from exc


def _header_map(values: tuple[object, ...] | list[object]) -> dict[str, int]:
    return {str(value).strip(): index for index, value in enumerate(values) if value not in (None, "")}


def _find_header_row(sheet, required: set[str], max_rows: int = 20) -> tuple[int, dict[str, int]]:
    for row_index, values in enumerate(
        sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, max_rows), values_only=True), 1
    ):
        columns = _header_map(values)
        if required <= set(columns):
            return row_index, columns
    raise SchemaMismatch(f"{sheet.title} 시트에서 필수 헤더를 찾지 못했습니다: {', '.join(sorted(required))}")


def _find_header_column(sheet, aliases: tuple[str, ...], max_rows: int = 20) -> int | None:
    wanted = set(aliases)
    for values in sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, max_rows), values_only=True):
        for index, value in enumerate(values):
            if str(value or "").strip() in wanted:
                return index
    return None


class AutomaticItemRules:
    def __init__(
        self,
        direct: dict[str, tuple[str, str]],
        keywords: list[tuple[str, str]],
        exclusions: list[tuple[str, str]],
    ) -> None:
        self.direct = direct
        self.keywords = keywords
        self.exclusions = exclusions

    @classmethod
    def load(cls, path: Path) -> "AutomaticItemRules":
        source = Path(path).resolve(strict=True)
        workbook = load_workbook(source, data_only=True, read_only=True)
        try:
            required = {"칼럼매칭", "자동판단키워드", "제외규칙"}
            if not required <= set(workbook.sheetnames):
                raise SchemaMismatch(f"지급항목 매칭표 시트가 부족합니다: {source}")
            direct: dict[str, tuple[str, str]] = {}
            for raw, standard, category, *_ in workbook["칼럼매칭"].iter_rows(min_row=2, values_only=True):
                key = _normalized_item_name(raw)
                if not key:
                    continue
                value = (str(standard or raw).strip(), str(category or "").strip())
                if key in direct and direct[key] != value:
                    raise SchemaMismatch(f"지급항목 직접 매칭이 충돌합니다: {raw}")
                direct[key] = value
            keywords = [
                (_normalized_item_name(keyword), str(category or "").strip())
                for keyword, category, *_ in workbook["자동판단키워드"].iter_rows(min_row=2, values_only=True)
                if _normalized_item_name(keyword)
            ]
            exclusions = [
                (_normalized_item_name(keyword), str(reason or "").strip())
                for keyword, reason, *_ in workbook["제외규칙"].iter_rows(min_row=2, values_only=True)
                if _normalized_item_name(keyword)
            ]
        finally:
            workbook.close()
        return cls(direct, keywords, exclusions)

    def classify(self, item_name: str) -> ItemDecision:
        key = _normalized_item_name(item_name)
        if key in self.direct:
            standard, category = self.direct[key]
            return ItemDecision(item_name, category or UNCLASSIFIED, f"직접 매칭: {standard}")
        exclusions = [(keyword, reason) for keyword, reason in self.exclusions if keyword in key]
        if exclusions:
            return ItemDecision(item_name, EXCLUDED, exclusions[0][1] or f"제외 키워드: {exclusions[0][0]}")
        if key in _BASE_FIXED_ITEM_KEYS or any(keyword in key for keyword in _BASE_FIXED_KEYWORD_KEYS):
            return ItemDecision(item_name, FIXED_SALARY, "기본 정액급여 항목")
        matches = {(category, keyword) for keyword, category in self.keywords if keyword in key}
        categories = {category for category, _ in matches}
        if len(categories) > 1:
            detail = ", ".join(sorted(f"{keyword}:{category}" for category, keyword in matches))
            return ItemDecision(item_name, CONFLICT, detail)
        if len(categories) == 1:
            category = next(iter(categories))
            keywords = ", ".join(sorted(keyword for matched, keyword in matches if matched == category))
            return ItemDecision(item_name, category, f"키워드: {keywords}")
        return ItemDecision(item_name, UNCLASSIFIED, "일치하는 직접 매칭·제외·키워드 규칙 없음")

    def is_direct(self, item_name: str) -> bool:
        return _normalized_item_name(item_name) in self.direct


def _requires_regularity_check(item_name: str) -> bool:
    key = _normalized_item_name(item_name)
    if any(keyword in key for keyword in _BASE_FIXED_KEYWORD_KEYS):
        return False
    return key not in _BASE_FIXED_ITEM_KEYS


def update_basic_info_item_markers(
    path: Path,
    rules: AutomaticItemRules,
    payroll_directory: Path | None = None,
) -> tuple[int, int]:
    target = Path(path).resolve(strict=True)
    regularity = (
        _load_fixed_allowance_regularity(Path(payroll_directory), rules)
        if payroll_directory is not None
        else {}
    )
    workbook = load_workbook(target)
    temporary = target.with_name(f".{target.stem}.{uuid.uuid4().hex}.tmp.xlsx")
    fixed_count = review_count = 0
    try:
        if BUSINESS_SHEET not in workbook.sheetnames:
            raise SchemaMismatch(f"기초정보에 {BUSINESS_SHEET} 시트가 없습니다.")
        sheet = workbook[BUSINESS_SHEET]
        for item_column, fixed_column, review_column, year_column in (
            (10, 11, 12, None),
            (16, 17, 18, 19),
        ):
            for row in range(5, sheet.max_row + 1):
                item_name = str(sheet.cell(row, item_column).value or "").strip()
                if not item_name:
                    continue
                category = rules.classify(item_name).category
                if (
                    category == FIXED_SALARY
                    and not rules.is_direct(item_name)
                    and _requires_regularity_check(item_name)
                    and regularity
                ):
                    item_key = _normalized_item_name(item_name)
                    requested_years = {
                        2000 + int(value)
                        for value in re.findall(r"\d{2}", str(sheet.cell(row, year_column).value or ""))
                    } if year_column else set()
                    results = [
                        result
                        for (year, key), result in regularity.items()
                        if key == item_key and (not requested_years or year in requested_years)
                    ]
                    if not results or not all(is_regular for is_regular, _ in results):
                        category = VARIABLE_FIXED
                fixed = "O" if category == FIXED_SALARY else None
                review = "O" if category == REVIEW_ALLOWANCE else None
                sheet.cell(row, fixed_column).value = fixed
                sheet.cell(row, review_column).value = review
                fixed_count += fixed is not None
                review_count += review is not None
        workbook.save(temporary)
    finally:
        workbook.close()
    try:
        os.replace(temporary, target)
    finally:
        temporary.unlink(missing_ok=True)
    logger.info(
        "기초정보 지급항목 자동 표시 완료: 정액급여 %d건, 수당검토 %d건",
        fixed_count,
        review_count,
    )
    return fixed_count, review_count


def detect_work_years(basic_info_path: Path, payroll_directory: Path) -> tuple[tuple[int, ...], dict[str, list[int]]]:
    payroll_years = {
        2000 + int(match.group(1))
        for path in Path(payroll_directory).glob("*.xlsx")
        if (match := PAYROLL_YEAR_PATTERN.search(path.name))
    }
    workbook = load_workbook(basic_info_path, read_only=True, data_only=True)
    try:
        withholding_years = {
            2000 + int(match.group(1))
            for name in workbook.sheetnames
            if (match := TAX_SHEET_PATTERN.match(name))
        }
    finally:
        workbook.close()
    years = tuple(sorted(payroll_years & withholding_years))
    if not years:
        raise SchemaMismatch("급여대장과 원천징수추출 시트에 공통으로 존재하는 작업연도가 없습니다.")
    return years, {
        "payroll_years": sorted(payroll_years),
        "withholding_years": sorted(withholding_years),
        "payroll_only_years": sorted(payroll_years - withholding_years),
        "withholding_only_years": sorted(withholding_years - payroll_years),
    }


def update_basic_info_years(path: Path, years: tuple[int, ...]) -> None:
    workbook = load_workbook(path)
    temporary = path.with_name(f".{path.stem}.{uuid.uuid4().hex}.tmp.xlsx")
    try:
        if BUSINESS_SHEET not in workbook.sheetnames:
            raise SchemaMismatch(f"기초정보에 {BUSINESS_SHEET} 시트가 없습니다.")
        workbook[BUSINESS_SHEET][YEAR_CELL] = ",".join(str(year)[-2:] for year in years)
        workbook.save(temporary)
    finally:
        workbook.close()
    try:
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _load_tax_records(
    workbook,
    year: int,
) -> tuple[
    dict[tuple[str, str], list[TaxRecord]],
    dict[str, list[TaxRecord]],
    dict[str, list[TaxRecord]],
    list[str],
]:
    sheet_name = f"원천징수추출({str(year)[-2:]}년)"
    if sheet_name not in workbook.sheetnames:
        return {}, {}, {}, []
    sheet = workbook[sheet_name]
    header_row, columns = _find_header_row(sheet, {"성명", "주민등록번호", "과세소득_총계"})
    by_identity: defaultdict[tuple[str, str], list[TaxRecord]] = defaultdict(list)
    by_employee_code: defaultdict[str, list[TaxRecord]] = defaultdict(list)
    by_name: defaultdict[str, list[TaxRecord]] = defaultdict(list)
    aggregate_headers = {
        "전년도과세소득_총계",
        "과세소득_총계",
        "과세소득_현계",
        "비과세소득_총계",
        "비과세소득_현계",
    }
    detail_headers = [
        *[header for header in ANNUAL_TAX_HEADERS if header in columns],
        *[
            header
            for header, _ in sorted(columns.items(), key=lambda item: item[1])
            if header not in aggregate_headers
            and header not in ANNUAL_TAX_HEADERS
            and header.endswith(("_총계", "_현계"))
        ],
    ]
    for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        name = str(values[columns["성명"]] or "").strip()
        resident = _resident_id(values[columns["주민등록번호"]])
        if not name or not resident:
            continue
        record = TaxRecord(
            employee_code=_employee_code(values[columns["사원번호"]]) if "사원번호" in columns else "",
            name=name,
            resident_id=resident,
            index=_employee_code(values[columns["인덱스"]]) if "인덱스" in columns else "",
            taxable_total=_money(values[columns["과세소득_총계"]]),
            non_tax_total=_money(values[columns.get("비과세소득_총계", -1)]) if "비과세소득_총계" in columns else 0,
            non_tax_current=_money(values[columns.get("비과세소득_현계", -1)]) if "비과세소득_현계" in columns else 0,
            details={header: _money(values[columns[header]]) for header in detail_headers},
        )
        identity = (resident, _normalized_name(name))
        by_identity[identity].append(record)
        if record.employee_code:
            by_employee_code[record.employee_code].append(record)
        by_name[_normalized_name(name)].append(record)
    return dict(by_identity), dict(by_employee_code), dict(by_name), detail_headers


def _load_occupations(workbook) -> tuple[dict[tuple[str, str], OccupationRecord], dict[str, list[OccupationRecord]]]:
    if "직업분류코드" not in workbook.sheetnames:
        return {}, {}
    sheet = workbook["직업분류코드"]
    header_row, columns = _find_header_row(sheet, {"성명", "주민등록번호"})
    code_column = _find_header_column(sheet, ("직업분류 코드", "직업분류코드"))
    item_column = _find_header_column(sheet, ("직업분류 항목", "직업분류항목"))
    acquired_column = _find_header_column(sheet, ("취득일", "입사일"))
    lost_column = _find_header_column(sheet, ("상실일", "퇴사일"))
    by_identity: dict[tuple[str, str], OccupationRecord] = {}
    for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
        name = str(values[columns["성명"]] or "").strip()
        resident_key = _resident_key(values[columns["주민등록번호"]])
        if not name or not resident_key:
            continue
        record = OccupationRecord(
            name=name,
            resident_key=resident_key,
            code=str(values[code_column] or "").strip() if code_column is not None else "",
            item=str(values[item_column] or "").strip() if item_column is not None else "",
            acquired=str(values[acquired_column] or "").strip() if acquired_column is not None else "",
            lost=str(values[lost_column] or "").strip() if lost_column is not None else "",
        )
        identity = (resident_key, _normalized_name(name))
        existing = by_identity.get(identity)
        if existing is None or (not record.lost, record.acquired) > (not existing.lost, existing.acquired):
            by_identity[identity] = record
    by_name: defaultdict[str, list[OccupationRecord]] = defaultdict(list)
    for record in by_identity.values():
        by_name[_normalized_name(record.name)].append(record)
    return by_identity, dict(by_name)


def _match_tax(
    employee: EmployeeData,
    by_identity: dict[tuple[str, str], list[TaxRecord]],
    by_employee_code: dict[str, list[TaxRecord]],
    by_name: dict[str, list[TaxRecord]],
) -> tuple[TaxRecord | None, str | None]:
    normalized_name = _normalized_name(employee.name)
    resident = _resident_id(employee.resident_id) or _resident_id(employee.code)
    if resident:
        identity_candidates = by_identity.get((resident, normalized_name), [])
        if len(identity_candidates) == 1:
            return identity_candidates[0], None
        if len(identity_candidates) > 1:
            return None, "원천징수 자료 동일 주민번호·성명 중복"
        return None, "원천징수 자료 주민번호·성명 미연결"
    employee_code = _employee_code(employee.code)
    code_candidates = by_employee_code.get(employee_code, []) if employee_code else []
    if len(code_candidates) == 1:
        return code_candidates[0], None
    if len(code_candidates) > 1:
        return None, "원천징수 자료 사원번호 중복"
    candidates = by_name.get(normalized_name, [])
    if len(candidates) == 1:
        return candidates[0], None
    if not candidates:
        return None, "원천징수 자료 미연결"
    return None, "원천징수 자료 동명이인·중복"


def _load_payroll(path: Path, year: int) -> tuple[list[str], list[str], list[str], list[EmployeeData]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    employees: dict[tuple[str, str], EmployeeData] = {}
    employee_order: list[tuple[str, str]] = []
    payment_headers: list[str] = []
    deduction_headers: list[str] = []
    try:
        for sheet in workbook.worksheets:
            month_match = re.search(r"(?<!\d)(\d{1,2})\s*월", sheet.title)
            if not month_match:
                continue
            month = int(month_match.group(1))
            header_row, columns = _find_header_row(sheet, {"사원번호", "성명", *TOTAL_HEADERS}, 10)
            ordered_headers = [str(value or "").strip() for value in next(
                sheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True)
            )]
            pay_total_index = columns["지급합계"]
            deduction_total_index = columns["공제합계"]
            first_payment_index = max(columns.get("급여구분", 3) + 1, 4)
            for header in ordered_headers[first_payment_index:pay_total_index]:
                if header and header not in payment_headers:
                    payment_headers.append(header)
            for header in ordered_headers[pay_total_index + 1:deduction_total_index]:
                if header and header not in deduction_headers:
                    deduction_headers.append(header)
            for values in sheet.iter_rows(min_row=header_row + 1, values_only=True):
                code = _employee_code(values[columns["사원번호"]])
                name = str(values[columns["성명"]] or "").strip()
                if not code and not name:
                    continue
                if not name:
                    raise SchemaMismatch(f"{path.name}/{sheet.title}: 성명이 없는 급여 행이 있습니다.")
                row = {header: values[index] if index < len(values) else None for header, index in columns.items()}
                key = (code, _normalized_name(name))
                if key not in employees:
                    employees[key] = EmployeeData(
                        key=key,
                        code=code,
                        name=name,
                        resident_id=_resident_id(row.get("주민등록번호")),
                    )
                    employee_order.append(key)
                pay_type = str(row.get("급여구분") or "급여").strip()
                if pay_type == "급상여":
                    pay_type = "급여"
                canonical = {
                    "사원번호": code,
                    "성명": name,
                    "주민등록번호": _resident_id(row.get("주민등록번호")),
                    "입사일": date_text(row.get("입사일")),
                    "퇴사일": date_text(row.get("퇴사일")),
                    "지급년월": f"{str(year)[-2:]}년{month:02d}월",
                    "급여구분": pay_type,
                }
                for header in (*payment_headers, "지급합계", *deduction_headers, "공제합계", "차인지급액"):
                    value = row.get(header)
                    canonical[header] = None if value in (None, "") else _money(value)
                employees[key].monthly_rows.append(canonical)
    finally:
        workbook.close()
    if not employees:
        raise SchemaMismatch(f"급여 월별 행을 찾지 못했습니다: {path}")
    unresolved_pattern = re.compile(
        r"(?:A\d+\(지급\)|(?:BF|BU)\d+\(공제\)|"
        r"미확정_원본열_.+\((?:지급|공제)\))$",
        flags=re.IGNORECASE,
    )
    inactive_unresolved = {
        header
        for header in (*payment_headers, *deduction_headers)
        if unresolved_pattern.fullmatch(header)
        and not any(
            _money(row.get(header)) != 0
            for employee in employees.values()
            for row in employee.monthly_rows
        )
    }
    if inactive_unresolved:
        logger.info(
            "급상여현황에서 연간 금액이 없는 미확정 열을 제거합니다: %s",
            ", ".join(sorted(inactive_unresolved)),
        )
        payment_headers = [
            header for header in payment_headers
            if header not in inactive_unresolved
        ]
        deduction_headers = [
            header for header in deduction_headers
            if header not in inactive_unresolved
        ]
        for employee in employees.values():
            for row in employee.monthly_rows:
                for header in inactive_unresolved:
                    row.pop(header, None)
    headers = [*BASE_HEADERS, *payment_headers, "지급합계", *deduction_headers, "공제합계", "차인지급액", *REFERENCE_HEADERS]
    return headers, payment_headers, deduction_headers, [employees[key] for key in employee_order]


def _fixed_allowance_regularity(
    employees: list[EmployeeData],
    item_name: str,
) -> tuple[bool, str]:
    recipients = evaluable = 0
    for employee in employees:
        by_month: defaultdict[str, int] = defaultdict(int)
        for row in employee.monthly_rows:
            by_month[str(row["지급년월"])] += _money(row.get(item_name))
        amounts = [by_month[month] for month in sorted(by_month)]
        if not any(amount != 0 for amount in amounts):
            continue
        recipients += 1
        if len(amounts) < 2:
            continue
        evaluable += 1
        if any(amount == 0 for amount in amounts):
            return False, f"{employee.code} {employee.name}: 재직 급여월 중 미지급 월 존재"
        if len(set(amounts)) != 1:
            return False, f"{employee.code} {employee.name}: 월별 지급액 변동"
    if recipients == 0:
        return False, "지급 실적 없음"
    if evaluable == 0:
        return False, "2개월 이상 지급된 근로자가 없어 고정성 확인 불가"
    return True, f"수급자 {recipients}명 모두 재직 급여월 매월 동일액 지급"


def _load_fixed_allowance_regularity(
    payroll_directory: Path,
    rules: AutomaticItemRules,
) -> dict[tuple[int, str], tuple[bool, str]]:
    root = Path(payroll_directory).resolve(strict=True)
    results: dict[tuple[int, str], tuple[bool, str]] = {}
    for path in root.glob("*.xlsx"):
        match = PAYROLL_YEAR_PATTERN.search(path.name)
        if not match:
            continue
        year = 2000 + int(match.group(1))
        _, payment_headers, _, employees = _load_payroll(path, year)
        for item_name in payment_headers:
            decision = rules.classify(item_name)
            if (
                decision.category != FIXED_SALARY
                or rules.is_direct(item_name)
                or not _requires_regularity_check(item_name)
            ):
                continue
            results[(year, _normalized_item_name(item_name))] = _fixed_allowance_regularity(
                employees,
                item_name,
            )
    return results


def _total_row(employee: EmployeeData, headers: list[str], monthly_rows: list[dict[str, object]]) -> dict[str, object]:
    row: dict[str, object] = {header: "" for header in headers}
    detail_headers = set(employee.current_tax.details) if employee.current_tax else set()
    acquired = next((month.get("입사일") for month in monthly_rows if month.get("입사일") not in (None, "")), "")
    lost = next((month.get("퇴사일") for month in monthly_rows if month.get("퇴사일") not in (None, "")), "")
    row.update(
        {
            "사원번호": employee.code,
            "성명": employee.name,
            "주민등록번호": employee.resident_id,
            "입사일": acquired,
            "퇴사일": lost,
            "지급년월": f"{monthly_rows[0]['지급년월'][:2]}년합계",
            "급여구분": "급여",
        }
    )
    for header in headers:
        if (
            header not in BASE_HEADERS
            and header not in REFERENCE_HEADERS
            and header not in ANNUAL_TAX_HEADERS
            and header not in detail_headers
        ):
            row[header] = sum(_money(month.get(header)) for month in monthly_rows)
    row["경정청구 대상 여부"] = monthly_rows[0].get("경정청구 대상 여부", "")
    if employee.occupation:
        row["직업분류 코드"] = (
            int(employee.occupation.code)
            if employee.occupation.code.isdigit()
            else employee.occupation.code
        )
        row["직업분류 항목"] = employee.occupation.item
    row["전년도과세소득_총계"] = employee.prior_taxable if employee.prior_taxable is not None else ""
    if employee.current_tax:
        row["인덱스"] = employee.current_tax.index
        row["과세소득_총계"] = employee.current_tax.taxable_total
        row["비과세소득_총계"] = employee.current_tax.non_tax_total
        row["비과세소득_현계"] = employee.current_tax.non_tax_current
        row.update(employee.current_tax.details)
    return row


def _enrich_employees(
    workbook,
    year: int,
    employees: list[EmployeeData],
    issues: list[dict[str, object]],
) -> list[str]:
    current_identity, current_codes, current_names, detail_headers = _load_tax_records(workbook, year)
    prior_identity, prior_codes, prior_names, _ = _load_tax_records(workbook, year - 1)
    occupations, occupation_names = _load_occupations(workbook)
    employee_keys_by_identity: defaultdict[tuple[str, str], set[tuple[str, str]]] = defaultdict(set)
    for employee in employees:
        resident = _resident_id(employee.resident_id)
        if resident:
            employee_keys_by_identity[(resident, _normalized_name(employee.name))].add(employee.key)
    for employee in employees:
        resident_identity = (_resident_id(employee.resident_id), _normalized_name(employee.name))
        split_employee_keys = employee_keys_by_identity.get(resident_identity, set())
        if len(split_employee_keys) > 1:
            employee.forced_reasons.append("재입사자 복수 사원번호")
            issues.append({
                "year": year,
                "employee_code": employee.code,
                "employee_name": employee.name,
                "type": "대상 제외",
                "detail": "동일 주민번호·성명이 복수 사원번호로 분리됨: "
                + ", ".join(sorted(code for code, _ in split_employee_keys)),
            })
        current, reason = _match_tax(employee, current_identity, current_codes, current_names)
        if current is None:
            issues.append({"year": year, "employee_code": employee.code, "employee_name": employee.name, "type": "미연결", "detail": reason})
        else:
            employee.current_tax = current
            employee.resident_id = current.resident_id
            if not current.index:
                employee.forced_reasons.append("인덱스 없음")
                issues.append({"year": year, "employee_code": employee.code, "employee_name": employee.name, "type": "대상 제외", "detail": "원천징수 자료 인덱스 없음"})
        prior_employee = EmployeeData(employee.key, employee.code, employee.name, employee.resident_id)
        prior, _ = _match_tax(prior_employee, prior_identity, prior_codes, prior_names)
        if prior:
            employee.prior_taxable = prior.taxable_total
        elif year == 2023 and not prior_identity and current is not None:
            employee.prior_taxable = current.taxable_total
            issues.append({"year": year, "employee_code": employee.code, "employee_name": employee.name, "type": "임시설정", "detail": "2022년 자료가 없어 2023년 과세소득을 전년도 과세소득으로 사용"})
        else:
            employee.prior_taxable = 0
        occupation = occupations.get((_resident_key(employee.resident_id), _normalized_name(employee.name)))
        if occupation is None:
            candidates = occupation_names.get(_normalized_name(employee.name), [])
            if len(candidates) == 1 and candidates[0].resident_key == _resident_key(employee.resident_id):
                occupation = candidates[0]
        if occupation is None:
            issues.append({"year": year, "employee_code": employee.code, "employee_name": employee.name, "type": "미연결", "detail": "직업분류 자료 미연결"})
        else:
            employee.occupation = occupation
            if "아님" in occupation.code:
                employee.forced_reasons.append("직업분류 대상 아님")
    return detail_headers


def _classify_and_adjust(
    year: int,
    payment_headers: list[str],
    employees: list[EmployeeData],
    rules: AutomaticItemRules,
    issues: list[dict[str, object]],
) -> tuple[list[ItemDecision], set[tuple[str, str]], set[tuple[tuple[str, str], str]], dict[tuple[str, str], list[dict[str, object]]]]:
    decisions: list[ItemDecision] = []
    for header in payment_headers:
        decision = rules.classify(header)
        if (
            decision.category == FIXED_SALARY
            and not rules.is_direct(header)
            and _requires_regularity_check(header)
        ):
            is_regular, reason = _fixed_allowance_regularity(employees, header)
            if not is_regular:
                decision = ItemDecision(header, VARIABLE_FIXED, reason)
        decisions.append(decision)
    fixed_headers = [decision.item_name for decision in decisions if decision.category == FIXED_SALARY]
    review_headers = [decision.item_name for decision in decisions if decision.category == REVIEW_ALLOWANCE]
    if not fixed_headers:
        raise SchemaMismatch(f"{year}년 급여대장에서 정액급여 항목을 자동 판정하지 못했습니다.")
    for decision in decisions:
        if decision.category in {UNCLASSIFIED, CONFLICT, VARIABLE_FIXED}:
            issues.append({"year": year, "type": decision.category, "item": decision.item_name, "detail": decision.reason})
    fixed_limit = 2_100_000 if year <= 2025 else 2_600_000
    prior_limit = 30_000_000 if year <= 2025 else 37_000_000
    target_employees: set[tuple[str, str]] = set()
    target_months: set[tuple[tuple[str, str], str]] = set()
    adjusted_rows: dict[tuple[str, str], list[dict[str, object]]] = {}
    for employee in employees:
        for row in employee.monthly_rows:
            row["정액급여합계"] = sum(_money(row.get(header)) for header in fixed_headers)
        if (
            employee.forced_reasons
            or employee.prior_taxable is None
            or employee.current_tax is None
            or not employee.current_tax.index
        ):
            continue
        zero_insurance_months: set[str] = set()
        if any(
            all(header in row for header in SOCIAL_INSURANCE_HEADERS)
            for row in employee.monthly_rows
        ):
            insurance_by_month: defaultdict[str, dict[str, int]] = defaultdict(
                lambda: {header: 0 for header in SOCIAL_INSURANCE_HEADERS}
            )
            for row in employee.monthly_rows:
                month = str(row["지급년월"])
                for header in SOCIAL_INSURANCE_HEADERS:
                    insurance_by_month[month][header] += _money(row.get(header))
            zero_insurance_months = {
                month
                for month, amounts in insurance_by_month.items()
                if all(amounts[header] == 0 for header in SOCIAL_INSURANCE_HEADERS)
            }
        for row in employee.monthly_rows:
            month = str(row["지급년월"])
            if month in zero_insurance_months:
                continue
            fixed_total = _money(row.get("정액급여합계"))
            review_total = sum(_money(row.get(header)) for header in review_headers)
            if review_total > 0 and fixed_total <= fixed_limit and employee.prior_taxable <= prior_limit:
                target_employees.add(employee.key)
                target_months.add((employee.key, month))
        if employee.key not in target_employees:
            continue
        adjusted = [dict(row) for row in employee.monthly_rows]
        for row in adjusted:
            row.setdefault(NON_TAX_OVERTIME_HEADER, 0)
        existing_non_tax = sum(_money(row.get(NON_TAX_OVERTIME_HEADER)) for row in adjusted)
        remaining = max(0, ANNUAL_NON_TAX_LIMIT - existing_non_tax)
        for row in adjusted:
            if (employee.key, str(row["지급년월"])) not in target_months:
                continue
            for header in review_headers:
                source_amount = _money(row.get(header))
                moved = min(source_amount, remaining)
                if moved <= 0:
                    continue
                row[header] = source_amount - moved
                row[NON_TAX_OVERTIME_HEADER] = _money(row.get(NON_TAX_OVERTIME_HEADER)) + moved
                remaining -= moved
                if remaining == 0:
                    break
        adjusted_rows[employee.key] = adjusted
    return decisions, target_employees, target_months, adjusted_rows


def _ordered_status_headers(
    headers: list[str],
    payment_headers: list[str],
    decisions: list[ItemDecision],
) -> list[str]:
    category_by_header = {decision.item_name: decision.category for decision in decisions}
    basic_salary_key = _normalized_item_name("기본급")
    basic_headers = [
        header for header in payment_headers
        if _normalized_item_name(header) == basic_salary_key
    ]
    fixed_headers = [
        header for header in payment_headers
        if header not in basic_headers and category_by_header.get(header) == FIXED_SALARY
    ]
    review_headers = [
        header for header in payment_headers
        if category_by_header.get(header) == REVIEW_ALLOWANCE
    ]
    excluded_headers = [
        header for header in payment_headers
        if category_by_header.get(header) == EXCLUDED
    ]
    ordered_headers = [*basic_headers, *fixed_headers, *review_headers, *excluded_headers]
    ordered_headers.extend(header for header in payment_headers if header not in ordered_headers)
    payment_header_set = set(payment_headers)
    pay_total_index = headers.index("지급합계")
    prefix_headers = [header for header in headers[:pay_total_index] if header not in payment_header_set]
    return [*prefix_headers, *ordered_headers, *headers[pay_total_index:]]


def _write_sheet(
    sheet,
    headers: list[str],
    rows: list[dict[str, object]],
    target_months: set[tuple[tuple[str, str], str]],
    category_by_header: dict[str, str],
) -> None:
    header_fill = PatternFill("solid", fgColor="FFD9EAD3")
    review_header_fill = PatternFill("solid", fgColor="FFFFC7CE")
    excluded_header_fill = PatternFill("solid", fgColor="FFD9D9D9")
    total_fill = PatternFill("solid", fgColor="FFA9D08E")
    target_fill = PatternFill("solid", fgColor="FFFFE699")
    header_font = Font(name="맑은 고딕", size=10, bold=True)
    header_red_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFF0000")
    data_font = Font(name="맑은 고딕", size=10)
    data_red_font = Font(name="맑은 고딕", size=10, color="FFFF0000")
    center_alignment = Alignment(horizontal="center", vertical="center")
    right_alignment = Alignment(horizontal="right", vertical="center")
    money_headers = set(headers) - set(BASE_HEADERS) - NON_MONEY_REFERENCE_HEADERS
    sheet.freeze_panes = "F2"
    sheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(rows) + 1}"
    sheet.row_dimensions[1].height = 24
    for column, header in enumerate(headers, 1):
        width = max(10, min(24, len(header) * 2 + 2))
        sheet.column_dimensions[get_column_letter(column)].width = width
    for column, header in enumerate(headers, 1):
        cell = sheet.cell(1, column, header)
        cell.font = header_red_font if header == NON_TAX_OVERTIME_HEADER else header_font
        category = category_by_header.get(header)
        if header == NON_TAX_OVERTIME_HEADER or category == REVIEW_ALLOWANCE:
            cell.fill = review_header_fill
        elif category == EXCLUDED:
            cell.fill = excluded_header_fill
        else:
            cell.fill = header_fill
        cell.border = CELL_BORDER
        cell.alignment = center_alignment
    for row_index, row in enumerate(rows, 2):
        is_total = str(row.get("지급년월") or "").endswith("년합계")
        key = ((_employee_code(row.get("사원번호")), _normalized_name(row.get("성명"))), str(row.get("지급년월") or ""))
        row_fill = total_fill if is_total else target_fill if key in target_months else None
        for column, header in enumerate(headers, 1):
            value = row.get(header, "")
            cell = sheet.cell(row_index, column, value)
            cell.font = data_red_font if header == NON_TAX_OVERTIME_HEADER else data_font
            cell.border = CELL_BORDER
            cell.alignment = right_alignment if header in money_headers else center_alignment
            if header in money_headers and value not in (None, ""):
                cell.number_format = MONEY_FORMAT
            if row_fill is not None:
                cell.fill = row_fill


def _generate_year(
    company_name: str,
    year: int,
    payroll_path: Path,
    basic_workbook,
    rules: AutomaticItemRules,
    output_directory: Path,
    issues: list[dict[str, object]],
) -> tuple[Path, list[ItemDecision], int, int]:
    headers, payment_headers, _, employees = _load_payroll(payroll_path, year)
    detail_headers = _enrich_employees(basic_workbook, year, employees, issues)
    headers = [*headers[:-1], *detail_headers, headers[-1]]
    decisions, targets, target_months, adjusted_rows = _classify_and_adjust(
        year, payment_headers, employees, rules, issues
    )
    headers = _ordered_status_headers(headers, payment_headers, decisions)
    category_by_header = {
        decision.item_name: decision.category
        for decision in decisions
    }
    target_headers = list(headers)
    pay_total_index = target_headers.index("지급합계")
    if NON_TAX_OVERTIME_HEADER not in target_headers:
        target_headers.insert(pay_total_index, NON_TAX_OVERTIME_HEADER)
    all_rows: list[dict[str, object]] = []
    target_rows: list[dict[str, object]] = []
    target_original_rows: list[dict[str, object]] = []
    non_target_rows: list[dict[str, object]] = []
    for employee in employees:
        target_status = "대상" if employee.key in targets else "작업대상이 아님"
        raw_monthly = employee.monthly_rows
        for row in raw_monthly:
            row["경정청구 대상 여부"] = target_status
        all_rows.extend(raw_monthly)
        all_rows.append(_total_row(employee, headers, raw_monthly))
        if employee.key in targets:
            target_original_rows.extend(raw_monthly)
            target_original_rows.append(_total_row(employee, target_headers, raw_monthly))
            monthly = adjusted_rows[employee.key]
            for row in monthly:
                row["경정청구 대상 여부"] = target_status
            target_rows.extend(monthly)
            target_rows.append(_total_row(employee, target_headers, monthly))
        else:
            non_target_rows.extend(raw_monthly)
            non_target_rows.append(_total_row(employee, headers, raw_monthly))
    workbook = Workbook()
    workbook.remove(workbook.active)
    _write_sheet(workbook.create_sheet("전체인원"), headers, all_rows, set(), category_by_header)
    _write_sheet(workbook.create_sheet("대상자(작업)"), target_headers, target_rows, target_months, category_by_header)
    _write_sheet(workbook.create_sheet("대상자"), target_headers, target_original_rows, target_months, category_by_header)
    _write_sheet(workbook.create_sheet("비대상자"), headers, non_target_rows, set(), category_by_header)
    output_directory.mkdir(parents=True, exist_ok=True)
    output_path = output_directory / safe_name(f"{company_name}_{str(year)[-2:]}년 급상여현황.xlsx")
    temporary = output_path.with_name(f".{output_path.stem}.{uuid.uuid4().hex}.tmp.xlsx")
    try:
        workbook.save(temporary)
        check = load_workbook(temporary, read_only=True, data_only=False)
        try:
            if tuple(check.sheetnames) != OUTPUT_SHEETS:
                raise SchemaMismatch(f"급상여현황 시트 구성이 잘못되었습니다: {output_path.name}")
        finally:
            check.close()
        os.replace(temporary, output_path)
    finally:
        workbook.close()
        temporary.unlink(missing_ok=True)
    return output_path, decisions, len(targets), len(employees) - len(targets)


def generate_payroll_status_files(
    *,
    company_name: str,
    basic_info_path: Path,
    payroll_directory: Path,
    item_rules_path: Path,
    output_directory: Path,
) -> PayrollStatusResult:
    basic_info = Path(basic_info_path).resolve(strict=True)
    payroll_root = Path(payroll_directory).resolve(strict=True)
    output_root = Path(output_directory).resolve()
    years, year_inventory = detect_work_years(basic_info, payroll_root)
    update_basic_info_years(basic_info, years)
    rules = AutomaticItemRules.load(item_rules_path)
    update_basic_info_item_markers(basic_info, rules, payroll_root)
    basic_workbook = load_workbook(basic_info, data_only=True, read_only=True)
    outputs: list[Path] = []
    issues: list[dict[str, object]] = []
    year_reports: list[dict[str, object]] = []
    try:
        logger.info("급상여현황 생성 시작 (총 %d개 연도)", len(years))
        for year in years:
            logger.info("%d년 급상여현황 생성 시작", year)
            candidates = [
                path
                for path in payroll_root.glob("*.xlsx")
                if (match := PAYROLL_YEAR_PATTERN.search(path.name)) and 2000 + int(match.group(1)) == year
            ]
            if len(candidates) != 1:
                raise SchemaMismatch(f"{year}년 급여대장 후보가 {len(candidates)}개입니다: {payroll_root}")
            output, decisions, target_count, non_target_count = _generate_year(
                company_name, year, candidates[0], basic_workbook, rules, output_root, issues
            )
            outputs.append(output)
            year_reports.append(
                {
                    "year": year,
                    "source": str(candidates[0]),
                    "output": str(output),
                    "target_employees": target_count,
                    "non_target_employees": non_target_count,
                    "item_decisions": [decision.__dict__ for decision in decisions],
                }
            )
            logger.info("%d년 급상여현황 생성 완료: %s", year, output)
        logger.info("급상여현황 생성 완료: %d개", len(outputs))
    except Exception as exc:
        logger.error("급상여현황 생성 오류: %s", exc)
        raise
    finally:
        basic_workbook.close()
    report = {
        "company_name": company_name,
        "detected_years": list(years),
        "year_inventory": year_inventory,
        "basic_info_year_cell": f"{BUSINESS_SHEET}!{YEAR_CELL}",
        "basic_info_year_value": ",".join(str(year)[-2:] for year in years),
        "temporary_prior_income_rule": "2022년 자료가 없을 때 2023년 과세소득을 2023년 전년도 과세소득으로 사용",
        "years": year_reports,
        "issues": issues,
        "status": "NEEDS_REVIEW" if any(issue.get("type") != "임시설정" for issue in issues) else "SUCCEEDED",
    }
    output_root.mkdir(parents=True, exist_ok=True)
    report_path = output_root / "급상여현황_검토보고서.json"
    temporary = report_path.with_name(f".{report_path.name}.{uuid.uuid4().hex}.tmp")
    try:
        temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8-sig")
        os.replace(temporary, report_path)
    finally:
        temporary.unlink(missing_ok=True)
    return PayrollStatusResult(years, tuple(outputs), report_path, basic_info)


def generate_company_payroll_status(company_directory: Path, resources_directory: Path) -> PayrollStatusResult:
    company_root = Path(company_directory).resolve(strict=True)
    resources_root = Path(resources_directory).resolve(strict=True)
    basic_candidates = [path for path in company_root.glob("*_기초정보.xlsx") if not path.name.startswith("~$")]
    if len(basic_candidates) != 1:
        raise SchemaMismatch(f"회사 폴더의 기초정보 파일 후보가 {len(basic_candidates)}개입니다: {company_root}")
    company_name = basic_candidates[0].name.removesuffix("_기초정보.xlsx")
    return generate_payroll_status_files(
        company_name=company_name,
        basic_info_path=basic_candidates[0],
        payroll_directory=company_root / "2.근로소득 수정신고" / "1.Input",
        item_rules_path=resources_root / "지급항목_매칭표.xlsx",
        output_directory=company_root / "1. 급상여현황 생성" / "급상여현황",
    )
