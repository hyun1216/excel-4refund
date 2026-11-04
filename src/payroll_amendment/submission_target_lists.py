from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime
from collections.abc import Callable
from copy import copy
import json
import os
from pathlib import Path
import re
from uuid import uuid4

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .correction_inputs import _convert_xlsx_to_xls, _match_tool_tax_record, _tool_tax_records
from .employment_dates import parse_date
from .errors import DuplicatePdfIndex, SchemaMismatch
from .infra import format_resident_id, safe_name
from .payroll_status import NON_TAX_OVERTIME_HEADER, _digits, _normalized_name, _resident_key


STATUS_YEAR_PATTERN = re.compile(r"_(\d{2})년\s*급상여현황(?:\(수정\))?\.xlsx$", re.IGNORECASE)
TARGET_REASON = "생산직 야간·연장·휴일근로수당 비과세 적용"
NTS_HEADERS = (
    "연번",
    "귀속연도",
    "사원번호",
    "성명",
    "주민등록번호",
    "PDF 인덱스",
    "경정청구 사유",
    "추가 비과세금액",
    "기존 총급여",
    "수정 총급여",
    "기존 과세표준",
    "기존 결정세액(소득세)",
    "기존 기납부세액(소득세)",
    "기존 차감징수세액(소득세)",
)
INSURANCE_HEADERS = (
    "연번",
    "귀속연도",
    "사원번호",
    "성명",
    "주민등록번호",
    "PDF 인덱스",
    "입사일",
    "퇴사일",
    "추가 비과세금액",
    "기존 보수총액",
    "수정 보수총액",
    "직업분류 코드",
    "직업분류 항목",
    "건강보험 증번호",
    "건강보험 자격취득일",
    "건강보험 자격상실일",
)
MONEY_HEADERS = {
    "추가 비과세금액",
    "기존 총급여",
    "수정 총급여",
    "기존 과세표준",
    "기존 결정세액(소득세)",
    "기존 기납부세액(소득세)",
    "기존 차감징수세액(소득세)",
    "기존 보수총액",
    "수정 보수총액",
}
THIN_GRAY = Side(style="thin", color="FF808080")
CELL_BORDER = Border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY)


@dataclass(frozen=True)
class HealthRecord:
    certificate_number: str
    acquired: object
    lost: object


@dataclass(frozen=True)
class SubmissionTargetListResult:
    years: tuple[int, ...]
    nts_files: tuple[Path, ...]
    health_files: tuple[Path, ...]
    employment_files: tuple[Path, ...]
    general_employment_files: tuple[Path, ...]
    legacy_employment_files: tuple[Path, ...]
    report_path: Path
    issues: tuple[dict[str, object], ...]

    @property
    def insurance_files(self) -> tuple[Path, ...]:
        return tuple(
            path
            for pair in zip(
                self.health_files,
                self.employment_files,
                self.general_employment_files,
                self.legacy_employment_files,
                strict=True,
            )
            for path in pair
        )

    @property
    def terminated_employment_files(self) -> tuple[Path, ...]:
        return self.employment_files


def _year_from_status(path: Path) -> int:
    match = STATUS_YEAR_PATTERN.search(path.name)
    if match is None:
        raise SchemaMismatch(f"급상여현황 파일명에서 연도를 찾지 못했습니다: {path.name}")
    return 2000 + int(match.group(1))


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


def _employee_code(value: object) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _formatted_resident(value: object) -> str:
    return format_resident_id(value)


def _row_as_dict(headers: list[object], values: tuple[object, ...]) -> dict[str, object]:
    return {
        str(header).strip(): values[index] if index < len(values) else None
        for index, header in enumerate(headers)
        if header not in (None, "")
    }


def _identity(row: dict[str, object]) -> tuple[str, str]:
    resident = _digits(row.get("주민등록번호") or row.get("주민번호"))
    name = _normalized_name(row.get("성명") or row.get("이름"))
    if len(resident) != 13 or not name:
        raise SchemaMismatch("대상자 합계행의 성명 또는 주민등록번호가 올바르지 않습니다.")
    return resident, name


def _annual_target_rows(sheet) -> dict[tuple[str, str], dict[str, object]]:
    headers = [cell.value for cell in sheet[1]]
    if "지급년월" not in headers:
        raise SchemaMismatch(f"{sheet.title} 시트에 지급년월 열이 없습니다.")
    period_column = headers.index("지급년월")
    rows: dict[tuple[str, str], dict[str, object]] = {}
    for values in sheet.iter_rows(min_row=2, values_only=True):
        if not str(values[period_column] or "").strip().endswith("년합계"):
            continue
        row = _row_as_dict(headers, values)
        key = _identity(row)
        if key in rows:
            raise SchemaMismatch(f"대상자 합계행이 중복됐습니다: {row.get('성명')}")
        rows[key] = row
    return rows


def _half_year_taxable_totals(sheet) -> dict[tuple[str, str], tuple[int, int]]:
    headers = [cell.value for cell in sheet[1]]
    required = {"지급년월", "지급합계", NON_TAX_OVERTIME_HEADER}
    missing = required - set(headers)
    if missing:
        raise SchemaMismatch(f"{sheet.title} 시트의 반기 보수총액 열이 부족합니다: {sorted(missing)}")
    period_column = headers.index("지급년월")
    totals: dict[tuple[str, str], list[int]] = {}
    for values in sheet.iter_rows(min_row=2, values_only=True):
        period = str(values[period_column] or "").strip()
        match = re.search(r"년(\d{2})월$", period)
        if match is None:
            continue
        row = _row_as_dict(headers, values)
        key = _identity(row)
        taxable_pay = (
            _money(row.get("지급합계"))
            - _money(row.get(NON_TAX_OVERTIME_HEADER))
            - _money(row.get("비과세소득_현계"))
        )
        half = 0 if int(match.group(1)) <= 6 else 1
        totals.setdefault(key, [0, 0])[half] += taxable_pay
    return {key: (value[0], value[1]) for key, value in totals.items()}


def _load_targets(status_path: Path) -> list[dict[str, object]]:
    workbook = load_workbook(status_path, read_only=True, data_only=True)
    try:
        required = {"대상자(작업)", "대상자"}
        missing = required - set(workbook.sheetnames)
        if missing:
            raise SchemaMismatch(f"급상여현황 필수 시트가 없습니다: {sorted(missing)}")
        adjusted = _annual_target_rows(workbook["대상자(작업)"])
        original = _annual_target_rows(workbook["대상자"])
        half_year_totals = _half_year_taxable_totals(workbook["대상자(작업)"])
        if set(adjusted) != set(original):
            raise SchemaMismatch(
                f"수정 전·후 대상자 합계행이 다릅니다: {status_path.name}"
            )
        targets: list[dict[str, object]] = []
        indexes: set[int] = set()
        for key, adjusted_row in adjusted.items():
            original_row = original[key]
            index_text = str(adjusted_row.get("인덱스") or "").strip()
            if index_text.endswith(".0"):
                index_text = index_text[:-2]
            if not index_text.isdigit():
                raise SchemaMismatch(
                    f"대상자 PDF 인덱스가 숫자가 아닙니다: {adjusted_row.get('성명')}, {index_text}"
                )
            pdf_index = int(index_text)
            if pdf_index in indexes:
                raise DuplicatePdfIndex(f"대상자 PDF 인덱스가 중복됐습니다: {pdf_index}")
            indexes.add(pdf_index)
            added_non_tax = max(
                0,
                _money(adjusted_row.get(NON_TAX_OVERTIME_HEADER))
                - _money(original_row.get(NON_TAX_OVERTIME_HEADER)),
            )
            targets.append(
                {
                    **original_row,
                    "PDF 인덱스": pdf_index,
                    "추가 비과세금액": added_non_tax,
                    "수정전 상반기 보수총액": half_year_totals.get(key, (0, 0))[0],
                    "수정전 하반기 보수총액": half_year_totals.get(key, (0, 0))[1],
                }
            )
        return sorted(targets, key=lambda row: int(row["PDF 인덱스"]))
    finally:
        workbook.close()


def _active_in_year(acquired: object, lost: object, year: int) -> bool:
    acquired_date = parse_date(acquired)
    lost_date = parse_date(lost)
    if isinstance(acquired_date, datetime):
        acquired_date = acquired_date.date()
    if isinstance(lost_date, datetime):
        lost_date = lost_date.date()
    if acquired_date and acquired_date > date(year, 12, 31):
        return False
    return not lost_date or lost_date >= date(year, 1, 1)


def _load_health_records(basic_info_path: Path, year: int) -> dict[tuple[str, str], HealthRecord]:
    workbook = load_workbook(basic_info_path, read_only=True, data_only=True)
    try:
        if "건보정보" not in workbook.sheetnames:
            return {}
        sheet = workbook["건보정보"]
        headers = [str(cell.value or "").strip() for cell in sheet[1]]
        required = ("증번호", "주민등록번호", "성명", "자격취득일", "자격상실일")
        if any(header not in headers for header in required):
            raise SchemaMismatch("건보정보 시트의 필수 열이 부족합니다.")
        columns = {header: headers.index(header) for header in required}
        candidates: dict[tuple[str, str], list[HealthRecord]] = {}
        for values in sheet.iter_rows(min_row=2, values_only=True):
            resident_key = _resident_key(values[columns["주민등록번호"]])
            name_key = _normalized_name(values[columns["성명"]])
            if not resident_key or not name_key:
                continue
            acquired = values[columns["자격취득일"]]
            lost = values[columns["자격상실일"]]
            if not _active_in_year(acquired, lost, year):
                continue
            record = HealthRecord(
                _employee_code(values[columns["증번호"]]),
                acquired,
                lost,
            )
            candidates.setdefault((resident_key, name_key), []).append(record)
        result: dict[tuple[str, str], HealthRecord] = {}
        for key, records in candidates.items():
            records.sort(key=lambda item: str(item.acquired or ""), reverse=True)
            result[key] = records[0]
        return result
    finally:
        workbook.close()


def _match_health_record(
    records: dict[tuple[str, str], HealthRecord],
    resident_id: str,
    employee_name: str,
) -> HealthRecord | None:
    resident_key = _resident_key(resident_id)
    exact = records.get((resident_key, _normalized_name(employee_name)))
    if exact is not None:
        return exact
    candidates = [record for (key, _), record in records.items() if key == resident_key]
    return candidates[0] if len(candidates) == 1 else None


def _write_workbook(path: Path, title: str, headers: tuple[str, ...], rows: list[list[object]]) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "제출대상자"
    sheet.freeze_panes = "A3"
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
    title_cell = sheet.cell(1, 1, title)
    title_cell.font = Font(name="맑은 고딕", size=14, bold=True)
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill("solid", fgColor="FFD9EAD3")
    sheet.row_dimensions[1].height = 26
    header_fill = PatternFill("solid", fgColor="FFBDD7EE")
    for column, header in enumerate(headers, 1):
        cell = sheet.cell(2, column, header)
        cell.font = Font(name="맑은 고딕", size=10, bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = header_fill
        cell.border = CELL_BORDER
        width = max(11, min(28, len(header) * 2 + 2))
        if header == "경정청구 사유":
            width = 38
        elif "주민등록번호" in header:
            width = 17
        elif header == "직업분류 항목":
            width = 30
        sheet.column_dimensions[get_column_letter(column)].width = width
    for row_index, values in enumerate(rows, 3):
        for column, value in enumerate(values, 1):
            header = headers[column - 1]
            cell = sheet.cell(row_index, column, value)
            cell.font = Font(name="맑은 고딕", size=10)
            cell.alignment = Alignment(
                horizontal="right" if header in MONEY_HEADERS else "center",
                vertical="center",
            )
            cell.border = CELL_BORDER
            if header in MONEY_HEADERS:
                cell.number_format = "#,##0;-#,##0;-"
            elif "주민등록번호" in header:
                cell.number_format = "@"
    sheet.auto_filter.ref = f"A2:{get_column_letter(len(headers))}{len(rows) + 2}"
    sheet.sheet_view.showGridLines = False
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.stem}.{uuid4().hex}.tmp.xlsx")
    try:
        workbook.save(temporary)
        os.replace(temporary, path)
    finally:
        workbook.close()
        temporary.unlink(missing_ok=True)


def _value(row: dict[str, object], *headers: str) -> object:
    for header in headers:
        if row.get(header) not in (None, ""):
            return row[header]
    return ""


def _as_date(value: object) -> date | None:
    parsed = parse_date(value)
    if isinstance(parsed, datetime):
        return parsed.date()
    return parsed


def _health_date_text(value: object) -> str:
    parsed = _as_date(value)
    return parsed.strftime("%Y.%m.%d") if parsed else ""


def _compact_date(value: object) -> int | str:
    parsed = _as_date(value)
    return int(parsed.strftime("%Y%m%d")) if parsed else ""


def _health_settlement(acquired: object, lost: object, year: int) -> tuple[int | str, int | str]:
    acquired_date = _as_date(acquired)
    lost_date = _as_date(lost)
    if acquired_date is None:
        return "", ""
    start_month = 1 if acquired_date.year < year else acquired_date.month
    if lost_date is None or lost_date.year > year:
        end_month = 12
    else:
        end_month = lost_date.month - (1 if lost_date.day == 1 else 0)
    work_months = max(0, end_month - start_month + 1)
    annual_cutoff = date(year + 1, 5, 1)
    settlement_type = 1 if lost_date is None or lost_date >= annual_cutoff else 2
    return settlement_type, work_months


def _numeric_text(value: object) -> str:
    return _digits(value)


def _write_official_form(
    template_path: Path,
    staging_path: Path,
    rows: list[list[object]],
    *,
    expected_columns: int,
    validation_formula: bool = False,
    text_columns: tuple[int, ...] = (),
) -> None:
    workbook = load_workbook(template_path)
    try:
        sheet = workbook.worksheets[0]
        header_count = max(
            column
            for column in range(1, sheet.max_column + 1)
            if sheet.cell(1, column).value not in (None, "")
        )
        if header_count != expected_columns:
            raise SchemaMismatch(
                f"공단 제출 템플릿 열 수가 다릅니다: {template_path.name}, "
                f"기대={expected_columns}, 실제={header_count}"
            )
        if sheet.max_row > 1:
            sheet.delete_rows(2, sheet.max_row - 1)
        sheet.row_dimensions[1].height = 27
        if validation_formula:
            legacy_widths = (
                7.5703125, 19.7109375, 11.42578125, 11.5703125, 6.85546875,
                12.5703125, 16.7109375, 21.7109375, 12.5703125, 12.5703125,
                18.42578125, 18.42578125, 20.140625, 17.7109375, 10.5703125,
            )
            for column, width in enumerate(legacy_widths, 1):
                sheet.column_dimensions[get_column_letter(column)].width = width
        thin = Side(style="thin", color="FF000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        last_column = expected_columns + (1 if validation_formula else 0)
        for row_number, values in enumerate(rows, 2):
            for column, value in enumerate(values, 1):
                cell = sheet.cell(row_number, column, value)
                if column in text_columns:
                    cell.number_format = "@"
            if validation_formula:
                sheet.cell(row_number, expected_columns + 1, f"=L{row_number}-M{row_number}-N{row_number}")
            for column in range(1, last_column + 1):
                cell = sheet.cell(row_number, column)
                cell.font = Font(name="맑은 고딕", size=11)
                cell.border = border
            sheet.row_dimensions[row_number].height = 13.5
        staging_path.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(staging_path)
    finally:
        workbook.close()


GENERAL_EMPLOYMENT_HEADERS = (
    "성명",
    "주민(외국인)등록번호",
    "건설,벌목업\n근무이력자",
    "(산재)취득일(전입일)",
    "산재보험 수정후\n연간보수총액",
    "(고용)취득일(전입일)",
    "고용보험 수정후\n연간보수총액",
    "오류내역",
)


def _write_general_employment_form(
    template_path: Path,
    output_path: Path,
    rows: list[list[object]],
) -> None:
    workbook = load_workbook(template_path)
    try:
        sheet = workbook.worksheets[0]
        headers = tuple(sheet.cell(1, column).value for column in range(1, 9))
        if headers != GENERAL_EMPLOYMENT_HEADERS:
            raise SchemaMismatch(
                f"고용산재 일반 보수총액신고서 헤더가 다릅니다: {template_path.name}"
            )
        template_styles = [copy(sheet.cell(2, column)._style) for column in range(1, 9)]
        if sheet.max_row > 1:
            sheet.delete_rows(2, sheet.max_row - 1)
        for row_number, values in enumerate(rows, 2):
            for column, value in enumerate(values, 1):
                cell = sheet.cell(row_number, column, value)
                cell._style = copy(template_styles[column - 1])
                if column == 2:
                    cell.value = format_resident_id(value)
                    cell.number_format = "@"
        output_path.parent.mkdir(parents=True, exist_ok=True)
        temporary = output_path.with_name(f".{output_path.stem}.{uuid4().hex}.tmp.xlsx")
        workbook.save(temporary)
        os.replace(temporary, output_path)
    finally:
        workbook.close()
        if "temporary" in locals():
            temporary.unlink(missing_ok=True)


def generate_submission_target_lists(
    *,
    company_name: str,
    basic_info_path: Path,
    payroll_status_files: tuple[Path, ...] | list[Path],
    output_root: Path,
    health_template_path: Path | None = None,
    employment_template_path: Path | None = None,
    general_employment_template_path: Path | None = None,
    legacy_employment_template_path: Path | None = None,
    xls_converter: Callable[[list[tuple[Path, Path]]], None] | None = None,
) -> SubmissionTargetListResult:
    basic_info = Path(basic_info_path).resolve(strict=True)
    template_directory = Path(__file__).resolve().parents[2] / "Resources" / "Template"
    health_template = Path(
        health_template_path
        or template_directory / "TempA_건강보험_퇴직정산_연말정산_재정산신청서.xlsx"
    ).resolve(strict=True)
    employment_template = Path(
        employment_template_path
        or template_directory / "고용종료자_보수총액수정신고서.xlsx"
    ).resolve(strict=True)
    general_employment_template = Path(
        general_employment_template_path
        or template_directory / "일반_보수총액신고서.xlsx"
    ).resolve(strict=True)
    legacy_employment_template = Path(
        legacy_employment_template_path
        or template_directory / "TempB_고용보험_보수총액수정신고.xlsx"
    ).resolve(strict=True)
    status_by_year: dict[int, Path] = {}
    for path in payroll_status_files:
        source = Path(path).resolve(strict=True)
        year = _year_from_status(source)
        if year in status_by_year:
            raise SchemaMismatch(f"{year}년 급상여현황 파일이 중복됐습니다.")
        status_by_year[year] = source
    if not status_by_year:
        raise SchemaMismatch("제출 대상자 명단을 만들 급상여현황 파일이 없습니다.")

    root = Path(output_root).resolve()
    nts_files: list[Path] = []
    health_files: list[Path] = []
    employment_files: list[Path] = []
    general_employment_files: list[Path] = []
    legacy_employment_files: list[Path] = []
    conversions: list[tuple[Path, Path]] = []
    staging_files: list[Path] = []
    issues: list[dict[str, object]] = []
    year_reports: list[dict[str, object]] = []
    for year, status_path in sorted(status_by_year.items()):
        targets = _load_targets(status_path)
        if not targets:
            raise SchemaMismatch(f"{year}년 제출 대상자가 없습니다: {status_path.name}")
        tax_records = _tool_tax_records(basic_info, year)
        health_records = _load_health_records(basic_info, year)
        nts_rows: list[list[object]] = []
        health_rows: list[list[object]] = []
        employment_rows: list[list[object]] = []
        general_employment_rows: list[list[object]] = []
        legacy_employment_rows: list[list[object]] = []
        for serial, target in enumerate(targets, 1):
            resident = _digits(_value(target, "주민등록번호", "주민번호"))
            name = str(_value(target, "성명", "이름"))
            tax = _match_tool_tax_record(target, tax_records)
            if tax is None:
                raise SchemaMismatch(
                    f"기초정보 원천징수 자료에 제출 대상자가 없습니다: {year}, {name}, "
                    f"인덱스={target['PDF 인덱스']}"
                )
            added_non_tax = _money(target.get("추가 비과세금액"))
            original_income = _money(tax.get("과세소득_현계") or tax.get("과세소득_총계"))
            corrected_income = max(0, original_income - added_non_tax)
            original_total_income = _money(tax.get("과세소득_총계") or tax.get("과세소득_현계"))
            corrected_total_income = max(0, original_total_income - added_non_tax)
            nts_rows.append(
                [
                    serial,
                    year,
                    _employee_code(_value(target, "사원번호", "사번")),
                    name,
                    _formatted_resident(resident),
                    target["PDF 인덱스"],
                    TARGET_REASON,
                    added_non_tax,
                    original_income,
                    corrected_income,
                    _money(tax.get("과세표준")),
                    _money(tax.get("결정세액_소득세")),
                    _money(tax.get("기납부세액_소득세")),
                    _money(tax.get("차감징수세액_소득세")),
                ]
            )
            health = _match_health_record(health_records, resident, name)
            if health is None:
                issues.append(
                    {
                        "year": year,
                        "pdf_index": target["PDF 인덱스"],
                        "employee_name": name,
                        "type": "건보정보 미연결",
                    }
                )
            settlement_type, work_months = _health_settlement(
                health.acquired if health else None,
                health.lost if health else None,
                year,
            )
            health_rows.append(
                [
                    name,
                    _formatted_resident(resident),
                    _numeric_text(health.certificate_number) if health else "",
                    settlement_type,
                    year,
                    corrected_total_income,
                    work_months,
                    _health_date_text(health.acquired) if health else "",
                    _health_date_text(health.lost) if health else "",
                ]
            )
            acquired = _compact_date(target.get("입사일"))
            lost = _compact_date(target.get("퇴사일"))
            employment_row = [
                0,
                _numeric_text(resident),
                name,
                "",
                year,
                acquired,
                lost,
                corrected_income,
                acquired,
                lost,
                corrected_income,
                corrected_income,
                _money(target.get("수정전 상반기 보수총액")),
                _money(target.get("수정전 하반기 보수총액")),
            ]
            # 산재 고용종료일(G열)과 고용 상실일(J열)이 모두 없을 때만 일반 신고서로 분리한다.
            # 고용종료자 양식의 주민번호는 하이픈 없는 13자리 텍스트를 사용한다.
            if employment_row[6] not in (None, "") or employment_row[9] not in (None, ""):
                employment_rows.append(employment_row[:12])
                legacy_employment_row = list(employment_row)
                legacy_employment_row[1] = _formatted_resident(resident)
                legacy_employment_rows.append(legacy_employment_row)
            else:
                general_employment_rows.append(
                    [
                        employment_row[2],
                        employment_row[1],
                        "",
                        employment_row[5],
                        employment_row[7],
                        employment_row[8],
                        employment_row[11],
                        "",
                    ]
                )

        nts_directory = root / f"★국세청제출자료_{company_name}" / f"{year}년 {company_name} 국세청 제출 서류"
        insurance_directory = root / f"★공단제출자료_{company_name}" / f"{year}년 {company_name} 공단 제출 서류"
        nts_path = nts_directory / safe_name(f"{company_name}_{year}년_국세청 제출용 대상자 명단.xlsx")
        health_path = insurance_directory / safe_name(
            f"{year}년 건강보험_퇴직정산_연말정산_재정산신청서_{company_name}.xls"
        )
        employment_path = insurance_directory / safe_name(
            f"{year}년 고용종료자_보수총액수정신고서_{company_name}.xlsx"
        )
        general_employment_path = insurance_directory / safe_name(
            f"{year}년 일반_보수총액신고서_{company_name}.xlsx"
        )
        legacy_employment_path = insurance_directory / safe_name(
            f"{year}년 고용보험_보수총액수정신고_{company_name}.xls"
        )
        _write_workbook(
            nts_path,
            f"{company_name} {year}년 국세청 제출용 대상자 명단",
            NTS_HEADERS,
            nts_rows,
        )
        health_staging = insurance_directory / f".{health_path.stem}.{uuid4().hex}.xlsx"
        employment_staging = insurance_directory / f".{employment_path.stem}.{uuid4().hex}.tmp.xlsx"
        legacy_employment_staging = (
            insurance_directory / f".{legacy_employment_path.stem}.{uuid4().hex}.xlsx"
        )
        _write_official_form(
            health_template,
            health_staging,
            health_rows,
            expected_columns=9,
            text_columns=(2, 3),
        )
        _write_official_form(
            employment_template,
            employment_staging,
            employment_rows,
            expected_columns=12,
            text_columns=(2,),
        )
        os.replace(employment_staging, employment_path)
        _write_general_employment_form(
            general_employment_template,
            general_employment_path,
            general_employment_rows,
        )
        _write_official_form(
            legacy_employment_template,
            legacy_employment_staging,
            legacy_employment_rows,
            expected_columns=14,
            validation_formula=True,
            text_columns=(2,),
        )
        staging_files.extend((health_staging, employment_staging, legacy_employment_staging))
        conversions.extend(
            ((health_staging, health_path), (legacy_employment_staging, legacy_employment_path))
        )
        nts_files.append(nts_path)
        health_files.append(health_path)
        employment_files.append(employment_path)
        general_employment_files.append(general_employment_path)
        legacy_employment_files.append(legacy_employment_path)
        year_reports.append(
            {
                "year": year,
                "targets": len(targets),
                "nts_file": str(nts_path),
                "health_file": str(health_path),
                "employment_file": str(employment_path),
                "general_employment_file": str(general_employment_path),
                "legacy_employment_file": str(legacy_employment_path),
                "terminated_employees": len(employment_rows),
                "active_employees": len(general_employment_rows),
                "added_non_tax_total": sum(_money(row.get("추가 비과세금액")) for row in targets),
            }
        )

    try:
        (xls_converter or _convert_xlsx_to_xls)(conversions)
    finally:
        for staging_file in staging_files:
            staging_file.unlink(missing_ok=True)

    for year in status_by_year:
        insurance_directory = (
            root
            / f"★공단제출자료_{company_name}"
            / f"{year}년 {company_name} 공단 제출 서류"
        )
        for obsolete_name in (
            f"{company_name}_{year}년_공단 제출용 대상자 명단.xlsx",
            f"{year}년 고용산재(일반)_보수총액신고서_{company_name}.xlsx",
        ):
            (insurance_directory / safe_name(obsolete_name)).unlink(missing_ok=True)

    report = {
        "company_name": company_name,
        "status": "NEEDS_REVIEW" if issues else "SUCCEEDED",
        "years": year_reports,
        "issues": issues,
    }
    report_path = root / "제출용_대상자명단_생성보고서.json"
    report_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = report_path.with_name(f".{report_path.name}.{uuid4().hex}.tmp")
    try:
        temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8-sig")
        os.replace(temporary, report_path)
    finally:
        temporary.unlink(missing_ok=True)
    return SubmissionTargetListResult(
        tuple(sorted(status_by_year)),
        tuple(nts_files),
        tuple(health_files),
        tuple(employment_files),
        tuple(general_employment_files),
        tuple(legacy_employment_files),
        report_path,
        tuple(issues),
    )


def generate_company_submission_target_lists(company_directory: Path) -> SubmissionTargetListResult:
    company_root = Path(company_directory).resolve(strict=True)
    basic_files = sorted(company_root.glob("*_기초정보.xlsx"))
    if len(basic_files) != 1:
        raise SchemaMismatch(f"기초정보 파일이 {len(basic_files)}개입니다: {company_root}")
    company_name = basic_files[0].name.removesuffix("_기초정보.xlsx")
    status_directory = company_root / "1. 급상여현황 생성" / "급상여현황"
    status_files = sorted(
        path
        for path in status_directory.glob("*_??년 급상여현황.xlsx")
        if "(수정)" not in path.name
    )
    return generate_submission_target_lists(
        company_name=company_name,
        basic_info_path=basic_files[0],
        payroll_status_files=status_files,
        output_root=company_root / "2.근로소득 수정신고" / "2.Output",
    )
