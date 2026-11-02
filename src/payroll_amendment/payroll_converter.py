from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
import argparse
import json
import logging
import os
from pathlib import Path
import re
import shutil
from typing import Callable, Iterable
from uuid import uuid4

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from pypdf import PdfReader, PdfWriter

from .dawon_pipeline import (
    _parse_payroll_name,
    _receipt_range,
    build_dawon_basic_info,
    build_dawon_payroll_workbook,
    extract_dawon_receipts,
)
from .errors import PdfToExcelError, SchemaMismatch
from .employee_info import (
    discover_employee_info_sources,
    populate_basic_info_employee_sheets,
)
from .infra import sha256
from .pdf_split import archive_prepared_withholding_pdfs, inspect_withholding_pdf
from .withholding_basic_info import (
    _extract_all,
    build_basic_info_from_withholding_pdfs,
    extract_ledger_employees,
)
from .withholding_payroll_ledger import build_monthly_payroll_ledger_from_pdfs


logger = logging.getLogger(__name__)


IDENTITY_HEADER_ORDER = (
    "사원번호",
    "성명",
    "주민등록번호",
    "입사일",
    "퇴사일",
    "급여구분",
)

STANDARD_HEADERS = (
    *IDENTITY_HEADER_ORDER,
    "기본급",
    "식대(비과세)",
    "연장수당",
    "야간수당",
    "휴일수당",
    "상여금",
    "기타수당",
    "지급합계",
    "국민연금",
    "건강보험",
    "장기요양",
    "고용보험",
    "소득세",
    "지방소득세",
    "기타공제",
    "공제합계",
    "차인지급액",
)

IDENTITY_HEADERS = set(IDENTITY_HEADER_ORDER)
IGNORED_HEADERS = {"실입사일"}
PAYMENT_HEADERS = (
    "기본급",
    "식대(비과세)",
    "연장수당",
    "야간수당",
    "휴일수당",
    "상여금",
    "기타수당",
)
DEDUCTION_HEADERS = (
    "국민연금",
    "건강보험",
    "장기요양",
    "고용보험",
    "소득세",
    "지방소득세",
    "기타공제",
)
TOTAL_HEADERS = {"지급합계", "공제합계", "차인지급액"}
NUMERIC_HEADERS = set(PAYMENT_HEADERS) | set(DEDUCTION_HEADERS) | TOTAL_HEADERS
HEADER_SCAN_KEYWORDS = {"성명", "사원번호", "기본급", "지급합계", "차인지급액"}

HEADER_SYNONYMS = {
    "사원번호": ["사원번호", "사번", "사원코드", "직원번호"],
    "성명": ["성명", "이름", "사원명", "직원명"],
    "주민등록번호": ["주민등록번호", "주민번호", "주민등록번호앞자리"],
    "입사일": ["입사일", "입사일자"],
    "퇴사일": ["퇴사일", "퇴사일자"],
    "급여구분": ["급여구분", "지급구분", "급상여구분"],
    "기본급": ["기본급", "기본급여", "기본봉", "통상급", "본봉", "월급여"],
    "식대(비과세)": ["식대(비과세)", "식대", "식사대", "중식대", "식비", "비과세식대"],
    "연장수당": [
        "연장수당", "연장근로", "연장근로수당", "시간외수당", "잔업수당", "고정연장수당",
        "약정연장수당", "약정연장수당1",
    ],
    "야간수당": ["야간수당", "야간근로", "야간근로수당", "심야수당", "야간연장수당(비과세)"],
    "휴일수당": [
        "휴일수당", "휴일근로", "휴일근로수당", "특근수당", "주휴수당", "약정휴일수당",
        "휴일야간", "휴일연장야간", "휴일근로수당",
    ],
    "상여금": ["상여금", "상여", "성과급", "인센티브", "특별상여"],
    "지급합계": ["지급합계", "지급총액", "지급액계", "총지급액", "급여총액"],
    "국민연금": ["국민연금", "국민연금보험료"],
    "건강보험": ["건강보험", "건강보험료"],
    "장기요양": ["장기요양", "노인장기요양보험", "장기요양보험", "장기요양보험료"],
    "고용보험": ["고용보험", "고용보험료"],
    "소득세": ["소득세", "갑근세", "근로소득세"],
    "지방소득세": ["지방소득세", "주민세", "지방세"],
    "공제합계": ["공제합계", "공제총액", "공제액계", "총공제액"],
    "차인지급액": ["차인지급액", "실지급액", "실수령액", "통장입금액"],
}

HEADER_FILL = PatternFill("solid", fgColor="FF1F4E78")
HEADER_FONT = Font(color="FFFFFFFF", bold=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="FFD9E2F3"),
    right=Side(style="thin", color="FFD9E2F3"),
    top=Side(style="thin", color="FFD9E2F3"),
    bottom=Side(style="thin", color="FFD9E2F3"),
)
MONEY_FORMAT = "#,##0;[Red]-#,##0"


@dataclass(frozen=True)
class ValidationWarning:
    row_number: int
    employee_name: str
    check: str
    calculated: float
    source: float


@dataclass
class WorkbookConversionResult:
    output_path: Path
    sheet_row_counts: dict[str, int]
    warnings: list[ValidationWarning] = field(default_factory=list)
    unknown_columns: dict[str, str] = field(default_factory=dict)
    payment_headers: list[str] = field(default_factory=list)
    deduction_headers: list[str] = field(default_factory=list)
    active_payment_headers: list[str] = field(default_factory=list)
    active_deduction_headers: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class UniversalConversionResult:
    output_directory: Path
    output_files: tuple[Path, ...]
    report_path: Path


def _normalized_header(value: object) -> str:
    text = str(value or "").replace("\n", "").replace("\r", "")
    text = re.sub(r"\(\s*원\s*\)", "", text)
    return re.sub(r"[^0-9A-Za-z가-힣]", "", text).lower()


def _first_nonblank(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def _append_unique_headers(target: list[str], values: Iterable[object]) -> None:
    known = {_normalized_header(value) for value in target}
    for value in values:
        display = str(value or "").strip()
        key = _normalized_header(display)
        if key and key not in known:
            target.append(display)
            known.add(key)


def _merge_staged_directory(source: Path, destination: Path) -> None:
    destination.mkdir(parents=True, exist_ok=True)
    for child in source.iterdir():
        target = destination / child.name
        if child.is_dir():
            _merge_staged_directory(child, target)
        else:
            os.replace(child, target)
    source.rmdir()


def _source_header(value: object) -> str:
    if value is None:
        return ""
    try:
        if bool(pd.isna(value)):
            return ""
    except (TypeError, ValueError):
        pass
    display = str(value or "").strip()
    if display.lower() in {"nan", "none", "<na>"}:
        return ""
    if re.fullmatch(r"Unnamed:\s*\d+(?:\.\d+)?", display, flags=re.IGNORECASE):
        return ""
    return display


def _numeric_values(series: pd.Series) -> pd.Series:
    return pd.to_numeric(
        series
        .astype(str)
        .str.replace(r"[₩$원,\s]", "", regex=True)
        .str.replace(r"^\((.+)\)$", r"-\1", regex=True),
        errors="coerce",
    ).fillna(0)


def _has_numeric_activity(series: pd.Series) -> bool:
    return bool(_numeric_values(series).ne(0).any())


def _unresolved_source_header(raw_header: object, index: int, kind: str) -> str:
    matched = re.fullmatch(
        r"Unnamed:\s*(\d+)(?:\.\d+)?",
        str(raw_header or "").strip(),
        flags=re.IGNORECASE,
    )
    source_index = int(matched.group(1)) if matched else index
    section = "지급" if kind == "payment" else "공제"
    return f"미확정_원본열_{get_column_letter(source_index + 1)}({section})"


def _coded_source_code(value: object) -> str:
    display = str(value or "").strip().upper()
    return display if re.fullmatch(r"A\d+|(?:BF|BU)\d+", display) else ""


CODED_ITEM_LABEL_ALIASES = {
    "연장근로 수당": "연장근로수당",
    "휴일근로 수당": "휴일근로수당",
    "시간외 수당": "시간외수당",
    "년차수당": "연차수당",
    "전월소급": "전월소급분",
    "전월소급금": "전월소급분",
}


def _canonical_coded_item_label(value: object) -> str:
    display = _source_header(value)
    if not display:
        return ""
    key = _normalized_header(display)
    for source, target in CODED_ITEM_LABEL_ALIASES.items():
        if key == _normalized_header(source):
            return target
    return re.sub(r"\s+", " ", display).strip()


def _target_years(value: object) -> list[int]:
    if value in (None, ""):
        return []
    if isinstance(value, (int, float)):
        return [int(value) % 100]
    years: list[int] = []
    for part in re.split(r"[,;/\s]+", str(value).strip()):
        if not part:
            continue
        try:
            year = int(float(part)) % 100
        except ValueError:
            continue
        if year not in years:
            years.append(year)
    return years


def _append_item_year(target: dict[str, list[int]], item: object, year: int) -> None:
    key = _normalized_header(item)
    if not key:
        return
    years = target.setdefault(key, [])
    normalized_year = int(year) % 100
    if normalized_year not in years:
        years.append(normalized_year)


def _receipt_year(path: Path) -> int:
    matched = re.search(r"(?<!\d)(20\d{2}|\d{2})\s*년", path.name)
    if matched is not None:
        value = int(matched.group(1))
        return value if value >= 2000 else 2000 + value
    try:
        return _receipt_range(path)[0]
    except SchemaMismatch:
        pass
    reader = PdfReader(str(path))
    compact = re.sub(r"\s+", "", reader.pages[0].extract_text() or "")
    matched = re.search(r"(?:근무기간|귀속연도).*?(20\d{2})", compact)
    if matched is None:
        raise SchemaMismatch(f"원천징수영수증에서 귀속연도를 찾지 못했습니다: {path.name}")
    return int(matched.group(1))


def _payroll_workbook_year(path: Path) -> int:
    matched = re.search(r"(?<!\d)(20\d{2}|\d{2})\s*년", path.name)
    if matched is None:
        raise SchemaMismatch(f"Excel 급여대장 파일명에서 연도를 찾지 못했습니다: {path.name}")
    value = int(matched.group(1))
    return value if value >= 2000 else 2000 + value


def _update_basic_info_item_headers(
    basic_info_path: Path,
    payment_headers_by_year: dict[int, Iterable[str]],
    deduction_headers_by_year: dict[int, Iterable[str]],
    all_payment_headers: Iterable[str] | None = None,
    all_deduction_headers: Iterable[str] | None = None,
) -> None:
    try:
        import win32com.client
    except ImportError as exc:
        raise SchemaMismatch("기초정보 지급·공제항목을 입력하려면 pywin32가 필요합니다.") from exc

    destination = Path(basic_info_path).resolve(strict=True)
    requested_payments: list[str] = []
    requested_deductions: list[str] = []
    payment_years: dict[str, list[int]] = {}
    deduction_years: dict[str, list[int]] = {}
    covered_payment_years = {int(year) % 100 for year in payment_headers_by_year}
    covered_deduction_years = {int(year) % 100 for year in deduction_headers_by_year}
    for year in sorted(payment_headers_by_year):
        year_headers: list[str] = []
        _append_unique_headers(year_headers, payment_headers_by_year[year])
        _append_unique_headers(requested_payments, year_headers)
        for header in year_headers:
            _append_item_year(payment_years, header, year)
    for year in sorted(deduction_headers_by_year):
        year_headers = []
        _append_unique_headers(year_headers, deduction_headers_by_year[year])
        _append_unique_headers(requested_deductions, year_headers)
        for header in year_headers:
            _append_item_year(deduction_years, header, year)
    if all_payment_headers is not None:
        requested_payments = []
        _append_unique_headers(requested_payments, all_payment_headers)
    if all_deduction_headers is not None:
        requested_deductions = []
        _append_unique_headers(requested_deductions, all_deduction_headers)
    if not requested_payments:
        raise SchemaMismatch("기초정보에 입력할 지급항목이 없습니다.")
    if not requested_deductions:
        raise SchemaMismatch("기초정보에 입력할 공제항목이 없습니다.")

    temporary = destination.parent / f".{destination.stem}.{uuid4().hex}.tmp.xlsx"
    shutil.copy2(destination, temporary)
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(str(temporary))
        worksheet = workbook.Worksheets("사업장정보")
        last_rows = {
            column: max(5, int(worksheet.Cells(worksheet.Rows.Count, column).End(-4162).Row))
            for column in (10, 14, 16, 21)
        }

        classifications: dict[str, tuple[object, object]] = {}
        existing_payments: list[str] = []
        for row in range(5, last_rows[10] + 1):
            display = str(worksheet.Cells(row, 10).Value or "").strip()
            key = _normalized_header(display)
            if not key:
                continue
            if key not in classifications:
                classifications[key] = (
                    worksheet.Cells(row, 11).Value,
                    worksheet.Cells(row, 12).Value,
                )
                existing_payments.append(display)

        existing_deductions = [
            str(worksheet.Cells(row, 14).Value or "").strip()
            for row in range(5, last_rows[14] + 1)
            if str(worksheet.Cells(row, 14).Value or "").strip()
        ]
        existing_annual_classifications: dict[str, tuple[object, object]] = {}
        existing_annual_payments: list[str] = []
        for row in range(5, last_rows[16] + 1):
            display = str(worksheet.Cells(row, 16).Value or "").strip()
            year_value = worksheet.Cells(row, 19).Value
            if not display or year_value in (None, ""):
                continue
            item_key = _normalized_header(display)
            if item_key not in existing_annual_classifications:
                existing_annual_classifications[item_key] = (
                    worksheet.Cells(row, 17).Value,
                    worksheet.Cells(row, 18).Value,
                )
                existing_annual_payments.append(display)
            for year in _target_years(year_value):
                if year not in covered_payment_years:
                    _append_item_year(payment_years, display, year)
        existing_annual_deductions: list[str] = []
        for row in range(5, last_rows[21] + 1):
            display = str(worksheet.Cells(row, 21).Value or "").strip()
            year_value = worksheet.Cells(row, 22).Value
            if display and year_value not in (None, ""):
                _append_unique_headers(existing_annual_deductions, [display])
                for year in _target_years(year_value):
                    if year not in covered_deduction_years:
                        _append_item_year(deduction_years, display, year)

        payments = list(requested_payments)
        deductions = list(requested_deductions)
        _append_unique_headers(payments, existing_payments)
        _append_unique_headers(payments, existing_annual_payments)
        _append_unique_headers(deductions, existing_deductions)
        _append_unique_headers(deductions, existing_annual_deductions)
        annual_payments = [
            (display, ",".join(str(year) for year in sorted(payment_years[_normalized_header(display)])))
            for display in payments
            if _normalized_header(display) in payment_years
        ]
        annual_deductions = [
            (display, ",".join(str(year) for year in sorted(deduction_years[_normalized_header(display)])))
            for display in deductions
            if _normalized_header(display) in deduction_years
        ]

        counts = {
            (10, 12): len(payments),
            (14, 14): len(deductions),
            (16, 19): len(annual_payments),
            (21, 22): len(annual_deductions),
        }
        for (start_column, end_column), count in counts.items():
            output_last_row = 4 + count
            if output_last_row > worksheet.Rows.Count:
                raise SchemaMismatch(f"사업장정보 항목 입력 행이 부족합니다: {count}개")
            if count:
                source_style = worksheet.Range(
                    worksheet.Cells(5, start_column),
                    worksheet.Cells(5, end_column),
                )
                target_range = worksheet.Range(
                    worksheet.Cells(5, start_column),
                    worksheet.Cells(output_last_row, end_column),
                )
                source_style.Copy(Destination=target_range)
            worksheet.Range(
                worksheet.Cells(5, start_column),
                worksheet.Cells(max(last_rows[start_column], output_last_row), end_column),
            ).ClearContents()

        payment_values = tuple(
            (
                display,
                classifications.get(_normalized_header(display), (None, None))[0],
                classifications.get(_normalized_header(display), (None, None))[1],
            )
            for display in payments
        )
        worksheet.Range(worksheet.Cells(5, 10), worksheet.Cells(4 + len(payments), 12)).Value = payment_values
        worksheet.Range(worksheet.Cells(5, 14), worksheet.Cells(4 + len(deductions), 14)).Value = tuple(
            (display,) for display in deductions
        )

        annual_payment_values = []
        for display, years in annual_payments:
            master = classifications.get(_normalized_header(display))
            previous = existing_annual_classifications.get(_normalized_header(display), (None, None))
            fixed, review = master if master and any(value not in (None, "") for value in master) else previous
            annual_payment_values.append((display, fixed, review, years))
        if annual_payment_values:
            worksheet.Range(
                worksheet.Cells(5, 19), worksheet.Cells(4 + len(annual_payment_values), 19)
            ).NumberFormat = "@"
            worksheet.Range(
                worksheet.Cells(5, 16), worksheet.Cells(4 + len(annual_payment_values), 19)
            ).Value = tuple(annual_payment_values)
        if annual_deductions:
            worksheet.Range(
                worksheet.Cells(5, 22), worksheet.Cells(4 + len(annual_deductions), 22)
            ).NumberFormat = "@"
            worksheet.Range(
                worksheet.Cells(5, 21), worksheet.Cells(4 + len(annual_deductions), 22)
            ).Value = tuple(annual_deductions)

        for (start_column, end_column), count in counts.items():
            formatted = worksheet.Range(
                worksheet.Cells(4, start_column),
                worksheet.Cells(4 + count, end_column),
            )
            formatted.Borders.LineStyle = 1
            formatted.Font.Name = "맑은 고딕"
            formatted.Font.Size = 10
            formatted.Font.Bold = True
            formatted.HorizontalAlignment = -4108
            blank_start_row = 5 + count
            used_last_row = int(worksheet.UsedRange.Row + worksheet.UsedRange.Rows.Count - 1)
            if blank_start_row <= used_last_row:
                worksheet.Range(
                    worksheet.Cells(blank_start_row, start_column),
                    worksheet.Cells(used_last_row, end_column),
                ).Borders.LineStyle = -4142

        actual_payments = [str(worksheet.Cells(row, 10).Value or "").strip() for row in range(5, 5 + len(payments))]
        actual_deductions = [str(worksheet.Cells(row, 14).Value or "").strip() for row in range(5, 5 + len(deductions))]
        actual_annual_payments = [
            (
                str(worksheet.Cells(row, 16).Value or "").strip(),
                ",".join(str(year) for year in _target_years(worksheet.Cells(row, 19).Value)),
            )
            for row in range(5, 5 + len(annual_payments))
        ]
        actual_annual_deductions = [
            (
                str(worksheet.Cells(row, 21).Value or "").strip(),
                ",".join(str(year) for year in _target_years(worksheet.Cells(row, 22).Value)),
            )
            for row in range(5, 5 + len(annual_deductions))
        ]
        if (
            actual_payments != payments
            or actual_deductions != deductions
            or actual_annual_payments != annual_payments
            or actual_annual_deductions != annual_deductions
        ):
            raise SchemaMismatch("사업장정보 지급·공제항목 입력 검증에 실패했습니다.")

        workbook.Save()
        workbook.Close(SaveChanges=True)
        workbook = None
        excel.Quit()
        excel = None
        os.replace(temporary, destination)
    except Exception:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        raise
    finally:
        temporary.unlink(missing_ok=True)


class PayrollConverter:
    """헤더명과 문서 시그니처를 기준으로 급여 자료를 표준 월별 원장으로 변환한다."""

    def __init__(
        self,
        *,
        unknown_item_policy: str = "aggregate",
        tolerance: float = 0,
        llm_header_classifier: Callable[[list[str]], dict[str, str]] | None = None,
    ) -> None:
        if unknown_item_policy not in {"aggregate", "separate"}:
            raise ValueError("unknown_item_policy는 aggregate 또는 separate여야 합니다.")
        self.unknown_item_policy = unknown_item_policy
        self.tolerance = tolerance
        self.llm_header_classifier = llm_header_classifier
        self._synonym_lookup = {
            _normalized_header(alias): standard
            for standard, aliases in HEADER_SYNONYMS.items()
            for alias in aliases
        }
        self._last_column_kinds: dict[str, str] = {}
        self._last_unknown_columns: dict[str, str] = {}

    def _detect_header_row_from_excel(
        self,
        excel_file: pd.ExcelFile,
        sheet_name: str | int,
        source_name: str,
    ) -> int:
        preview = pd.read_excel(
            excel_file,
            sheet_name=sheet_name,
            header=None,
            nrows=15,
            dtype=object,
        )
        best_row = -1
        best_score = 0
        for row_index, row in preview.iterrows():
            mapped = {
                self._synonym_lookup.get(_normalized_header(value))
                for value in row.tolist()
                if _normalized_header(value)
            }
            score = len(HEADER_SCAN_KEYWORDS & mapped)
            if score > best_score:
                best_row, best_score = int(row_index), score
        if best_score < 2:
            raise SchemaMismatch(
                f"상위 15행에서 급여 헤더를 찾지 못했습니다: {source_name}, 시트={sheet_name}"
            )
        return best_row

    def detect_header_row(self, file_path: Path, sheet_name: str | int = 0) -> int:
        source = Path(file_path)
        with pd.ExcelFile(source) as excel_file:
            return self._detect_header_row_from_excel(excel_file, sheet_name, source.name)

    def normalize_and_map_columns(self, raw_columns: Iterable[object]) -> dict[object, str | None]:
        columns = list(raw_columns)
        direct = [self._synonym_lookup.get(_normalized_header(value)) for value in columns]
        pay_total = next((index for index, value in enumerate(direct) if value == "지급합계"), None)
        deduction_total = next((index for index, value in enumerate(direct) if value == "공제합계"), None)
        net_total = next((index for index, value in enumerate(direct) if value == "차인지급액"), None)
        if pay_total is None or deduction_total is None or net_total is None:
            raise SchemaMismatch("지급합계·공제합계·차인지급액 열을 모두 확정할 수 없습니다.")
        if not pay_total < deduction_total < net_total:
            raise SchemaMismatch("지급합계·공제합계·차인지급액의 열 순서가 예상과 다릅니다.")

        identity_positions = [
            index for index, target in enumerate(direct[:pay_total]) if target in IDENTITY_HEADERS
        ]
        payment_start = max(identity_positions, default=3) + 1

        mapping: dict[object, str | None] = {}
        kinds: dict[str, str] = {}
        unknown: dict[str, str] = {}
        for index, raw in enumerate(columns):
            normalized = _normalized_header(raw)
            canonical = direct[index]
            target = canonical if canonical in IDENTITY_HEADERS | TOTAL_HEADERS else None
            kind = "known"
            if normalized in IGNORED_HEADERS:
                kind = "ignored"
                target = None
            elif normalized and payment_start <= index < pay_total:
                kind = "payment"
                target = _source_header(raw) or None
                if canonical is None and target:
                    unknown[target] = kind
            elif normalized and pay_total < index < deduction_total:
                kind = "deduction"
                target = _source_header(raw) or None
                if canonical is None and target:
                    unknown[target] = kind
            elif target is None:
                kind = "ignored"
            mapping[raw] = target
            kinds[str(raw)] = kind
        self._last_column_kinds = kinds
        self._last_unknown_columns = unknown
        return mapping

    def clean_numeric_data(self, df: pd.DataFrame) -> pd.DataFrame:
        cleaned = df.copy()
        for column in cleaned.columns:
            if column in IDENTITY_HEADERS:
                cleaned[column] = _first_nonblank(cleaned[column])
                continue
            values = (
                cleaned[column]
                .astype(str)
                .str.replace(r"[₩$원,\s]", "", regex=True)
                .str.replace(r"^\((.+)\)$", r"-\1", regex=True)
                .replace({"": None, "None": None, "nan": None})
            )
            cleaned[column] = pd.to_numeric(values, errors="coerce").fillna(0)
        return cleaned

    def validate_totals(self, df: pd.DataFrame) -> list[ValidationWarning]:
        warnings: list[ValidationWarning] = []
        headers = list(df.columns)
        payment_total_index = headers.index("지급합계")
        deduction_total_index = headers.index("공제합계")
        payment_start = max(
            (headers.index(header) for header in IDENTITY_HEADERS if header in headers),
            default=3,
        ) + 1
        payment_headers = headers[payment_start:payment_total_index]
        deduction_headers = headers[payment_total_index + 1:deduction_total_index]
        for position, (_, row) in enumerate(df.iterrows(), 2):
            employee = str(row.get("성명", "") or "")
            payment = float(sum(row.get(header, 0) or 0 for header in payment_headers))
            deduction = float(sum(row.get(header, 0) or 0 for header in deduction_headers))
            source_payment = float(row.get("지급합계", 0) or 0)
            source_deduction = float(row.get("공제합계", 0) or 0)
            source_net = float(row.get("차인지급액", 0) or 0)
            checks = (
                ("지급항목 합계", payment, source_payment),
                ("공제항목 합계", deduction, source_deduction),
                ("차인지급액", source_payment - source_deduction, source_net),
            )
            for check, calculated, source in checks:
                if abs(calculated - source) > self.tolerance:
                    warnings.append(ValidationWarning(position, employee, check, calculated, source))
        return warnings

    def classify_unknown_headers_with_llm(self, headers: Iterable[str]) -> dict[str, str]:
        names = sorted({_normalized_header(header) for header in headers if _normalized_header(header)})
        if self.llm_header_classifier is None:
            return {name: "UNCLASSIFIED" for name in names}
        return self.llm_header_classifier(names)

    def _standardize_dataframe(self, raw: pd.DataFrame) -> pd.DataFrame:
        mapping = self.normalize_and_map_columns(raw.columns)
        output = pd.DataFrame(index=raw.index)
        payment_headers: list[str] = []
        deduction_headers: list[str] = []
        unknown = dict(self._last_unknown_columns)
        for source_index, (source_column, target) in enumerate(mapping.items()):
            kind = self._last_column_kinds.get(str(source_column))
            numeric_values = None
            if target is None and kind in {"payment", "deduction"}:
                numeric_values = _numeric_values(raw[source_column])
                if not numeric_values.ne(0).any():
                    continue
                target = _unresolved_source_header(source_column, source_index, kind)
                unknown[target] = kind
            if target is None:
                continue
            if kind == "payment" and target not in payment_headers:
                payment_headers.append(target)
            elif kind == "deduction" and target not in deduction_headers:
                deduction_headers.append(target)
            if target in IDENTITY_HEADERS:
                values = _first_nonblank(raw[source_column])
                if target not in output:
                    output[target] = values
                else:
                    output[target] = output[target].where(output[target].astype(str).str.strip() != "", values)
            else:
                values = numeric_values if numeric_values is not None else _numeric_values(raw[source_column])
                output[target] = output.get(target, 0) + values

        headers = [
            *IDENTITY_HEADER_ORDER,
            *payment_headers,
            "지급합계",
            *deduction_headers,
            "공제합계",
            "차인지급액",
        ]
        for header in headers:
            if header not in output:
                output[header] = "" if header in IDENTITY_HEADERS else 0
        if not output["급여구분"].astype(str).str.strip().any():
            output["급여구분"] = "급상여"
        output = output[headers]
        output = self.clean_numeric_data(output)
        meaningful = output["성명"].astype(str).str.strip().ne("")
        self._last_unknown_columns = unknown
        return output.loc[meaningful].reset_index(drop=True)

    def _collect_coded_item_labels(
        self,
        file_path: Path,
        sheet_names: Iterable[str],
        *,
        excel_file: pd.ExcelFile | None = None,
        header_rows: dict[str, int] | None = None,
    ) -> dict[str, str]:
        labels: dict[str, str] = {}
        ambiguous_codes: set[str] = set()
        reader: Path | pd.ExcelFile = excel_file or file_path
        for sheet_name in sheet_names:
            try:
                header_row = (
                    header_rows[sheet_name]
                    if header_rows is not None and sheet_name in header_rows
                    else self.detect_header_row(file_path, sheet_name)
                )
            except SchemaMismatch:
                continue
            preview = pd.read_excel(
                reader,
                sheet_name=sheet_name,
                header=header_row,
                nrows=1,
                dtype=object,
            )
            if preview.empty:
                continue
            item_labels = preview.iloc[0]
            for index, column in enumerate(preview.columns):
                code = _coded_source_code(column)
                label = _source_header(item_labels.iloc[index])
                if not code or not label:
                    continue
                if code in ambiguous_codes:
                    continue
                existing = labels.get(code)
                if existing and _normalized_header(existing) != _normalized_header(label):
                    logger.warning(
                        "동일 급여 코드가 서로 다른 항목명으로 사용되어 "
                        "전역 헤더 보완에서 제외합니다: "
                        "%s=%r/%r, 시트=%s",
                        code,
                        existing,
                        label,
                        sheet_name,
                    )
                    labels.pop(code, None)
                    ambiguous_codes.add(code)
                    continue
                labels.setdefault(code, label)
        return labels

    def _load_excel_sheet(
        self,
        file_path: Path,
        sheet_name: str,
        coded_labels: dict[str, str] | None = None,
        *,
        excel_file: pd.ExcelFile | None = None,
        header_row: int | None = None,
    ) -> tuple[pd.DataFrame, list[str], list[str], list[str], list[str]]:
        if header_row is None:
            header_row = self.detect_header_row(file_path, sheet_name)
        source_raw = pd.read_excel(
            excel_file or file_path,
            sheet_name=sheet_name,
            header=header_row,
            dtype=object,
        )
        source_columns = list(source_raw.columns)
        coded_result = self._load_coded_excel_sheet(source_raw, coded_labels)
        if coded_result is not None:
            return coded_result
        raw = source_raw.dropna(how="all").dropna(axis=1, how="all")
        mapped = [self._synonym_lookup.get(_normalized_header(value)) for value in source_columns]
        try:
            payment_total = mapped.index("지급합계")
            deduction_total = mapped.index("공제합계")
        except ValueError as exc:
            raise SchemaMismatch(
                f"지급합계·공제합계 열을 찾지 못했습니다: {Path(file_path).name}, 시트={sheet_name}"
            ) from exc
        if payment_total >= deduction_total:
            raise SchemaMismatch(f"지급·공제 열 순서가 잘못되었습니다: {Path(file_path).name}, 시트={sheet_name}")
        identity_positions = [
            index for index, target in enumerate(mapped[:payment_total]) if target in IDENTITY_HEADERS
        ]
        payment_start = max(identity_positions, default=3) + 1
        standardized = self._standardize_dataframe(raw)
        standardized_headers = list(standardized.columns)
        standardized_payment_total = standardized_headers.index("지급합계")
        standardized_deduction_total = standardized_headers.index("공제합계")
        payment_headers = standardized_headers[len(IDENTITY_HEADER_ORDER):standardized_payment_total]
        deduction_headers = standardized_headers[
            standardized_payment_total + 1:standardized_deduction_total
        ]
        active_payment_headers = [
            header for header in payment_headers
            if _has_numeric_activity(standardized[header])
        ]
        active_deduction_headers = [
            header for header in deduction_headers
            if _has_numeric_activity(standardized[header])
        ]
        return (
            standardized,
            payment_headers,
            deduction_headers,
            active_payment_headers,
            active_deduction_headers,
        )

    def _load_coded_excel_sheet(
        self,
        source_raw: pd.DataFrame,
        coded_labels: dict[str, str] | None = None,
    ) -> tuple[pd.DataFrame, list[str], list[str], list[str], list[str]] | None:
        """A/BF/BU 코드와 다음 행의 항목명을 사용하는 급여대장을 읽는다."""
        source_columns = list(source_raw.columns)
        payment_positions = [
            index
            for index, column in enumerate(source_columns)
            if re.fullmatch(r"A\d+", str(column).strip(), re.IGNORECASE)
        ]
        deduction_positions = [
            index
            for index, column in enumerate(source_columns)
            if re.fullmatch(r"(?:BF|BU)\d+", str(column).strip(), re.IGNORECASE)
        ]
        if not payment_positions or not deduction_positions or source_raw.empty:
            return None

        mapped = [self._synonym_lookup.get(_normalized_header(value)) for value in source_columns]
        required_totals = {"지급합계", "공제합계", "차인지급액"}
        if not required_totals.issubset(set(mapped)):
            return None

        known_labels = coded_labels or {}
        identity_positions = [
            index for index, target in enumerate(mapped) if target in IDENTITY_HEADERS
        ]

        def is_item_header(row: pd.Series) -> bool:
            if any(_source_header(row.iloc[index]) for index in identity_positions):
                return False
            payment_labels = [
                _source_header(row.iloc[index]) for index in payment_positions
            ]
            deduction_labels = [
                _source_header(row.iloc[index]) for index in deduction_positions
            ]
            payment_text_count = sum(
                bool(label) and pd.isna(pd.to_numeric(label, errors="coerce"))
                for label in payment_labels
            )
            deduction_text_count = sum(
                bool(label) and pd.isna(pd.to_numeric(label, errors="coerce"))
                for label in deduction_labels
            )
            return payment_text_count >= 1 and deduction_text_count >= 1 and (
                payment_text_count + deduction_text_count >= 4
            )

        header_rows = [
            index for index, row in source_raw.iterrows() if is_item_header(row)
        ]
        if 0 not in header_rows:
            header_rows.insert(0, 0)
        repeated_headers = len(header_rows) > 1

        def collect_specs(
            header: pd.Series,
            data: pd.DataFrame,
            positions: list[int],
        ) -> list[tuple[int, str, str, bool, bool]]:
            specs: list[tuple[int, str, str, bool, bool]] = []
            for index in positions:
                code = _coded_source_code(source_columns[index])
                label = _canonical_coded_item_label(header.iloc[index])
                if not label and not repeated_headers:
                    label = _canonical_coded_item_label(known_labels.get(code, ""))
                active = _has_numeric_activity(data.iloc[:, index])
                unresolved = not label
                if unresolved:
                    if not active:
                        continue
                    label = code
                specs.append((index, code, label, active, unresolved))
            return specs
        segment_specs: list[
            tuple[
                pd.DataFrame,
                list[tuple[int, str, str, bool, bool]],
                list[tuple[int, str, str, bool, bool]],
            ]
        ] = []
        for position, header_index in enumerate(header_rows):
            next_header = (
                header_rows[position + 1]
                if position + 1 < len(header_rows)
                else len(source_raw)
            )
            data = source_raw.iloc[header_index + 1:next_header].dropna(how="all").copy()
            if data.empty:
                continue
            header = source_raw.iloc[header_index]
            payment_sources = collect_specs(header, data, payment_positions)
            deduction_sources = collect_specs(header, data, deduction_positions)
            if payment_sources and deduction_sources:
                segment_specs.append((data, payment_sources, deduction_sources))
        if not segment_specs:
            return None

        label_registry: dict[str, str] = {}
        for _, payment_sources, deduction_sources in segment_specs:
            for _, _, label, _, unresolved in (*payment_sources, *deduction_sources):
                if not unresolved:
                    label_registry.setdefault(_normalized_header(label), label)
        deduction_names = {
            _normalized_header(label)
            for _, _, deduction_sources in segment_specs
            for _, _, label, _, unresolved in deduction_sources
            if not unresolved
        }
        unknown: dict[str, str] = {}
        outputs: list[pd.DataFrame] = []
        payment_headers: list[str] = []
        deduction_headers: list[str] = []
        active_payment_headers: list[str] = []
        active_deduction_headers: list[str] = []

        for data, payment_sources, deduction_sources in segment_specs:
            payment_counts = Counter(
                _normalized_header(label) for _, _, label, _, _ in payment_sources
            )
            deduction_counts = Counter(
                _normalized_header(label) for _, _, label, _, _ in deduction_sources
            )
            payment_specs: list[tuple[int, str, bool]] = []
            deduction_specs: list[tuple[int, str, bool]] = []
            for index, code, label, active, unresolved in payment_sources:
                if unresolved:
                    target = f"{code}(지급)"
                    unknown[target] = "payment"
                else:
                    label = label_registry.get(_normalized_header(label), label)
                    target = (
                        f"{label} [{code}]"
                        if payment_counts[_normalized_header(label)] > 1
                        else label
                    )
                    if _normalized_header(label) in deduction_names:
                        target = f"{target}(지급)"
                payment_specs.append((index, target, active))
            for index, code, label, active, unresolved in deduction_sources:
                if unresolved:
                    target = f"{code}(공제)"
                    unknown[target] = "deduction"
                else:
                    label = label_registry.get(_normalized_header(label), label)
                    target = (
                        f"{label} [{code}]"
                        if deduction_counts[_normalized_header(label)] > 1
                        else label
                    )
                deduction_specs.append((index, target, active))

            output = pd.DataFrame(index=data.index)
            for index, target in enumerate(mapped):
                if target not in IDENTITY_HEADERS:
                    continue
                values = _first_nonblank(data.iloc[:, index])
                if target not in output:
                    output[target] = values
                else:
                    output[target] = output[target].where(
                        output[target].astype(str).str.strip() != "",
                        values,
                    )

            def add_numeric_column(source_index: int, target: str) -> None:
                values = _numeric_values(data.iloc[:, source_index])
                output[target] = output.get(target, 0) + values

            for index, target, _ in payment_specs:
                add_numeric_column(index, target)
            for total in ("지급합계", "공제합계", "차인지급액"):
                add_numeric_column(mapped.index(total), total)
            for index, target, _ in deduction_specs:
                add_numeric_column(index, target)

            _append_unique_headers(payment_headers, (target for _, target, _ in payment_specs))
            _append_unique_headers(deduction_headers, (target for _, target, _ in deduction_specs))
            _append_unique_headers(
                active_payment_headers,
                (target for _, target, active in payment_specs if active),
            )
            _append_unique_headers(
                active_deduction_headers,
                (target for _, target, active in deduction_specs if active),
            )
            if "급여구분" not in output:
                output["급여구분"] = "급상여"
            elif not output["급여구분"].astype(str).str.strip().any():
                output["급여구분"] = "급상여"
            employee_number = output["사원번호"].astype(str).str.strip()
            employee_name = output["성명"].astype(str).str.strip()
            summary_keys = {"조회건수", "합계", "총계", "계"}
            meaningful = employee_name.ne("") & ~employee_number.map(
                lambda value: _normalized_header(value) in summary_keys
            )
            outputs.append(output.loc[meaningful])

        headers = [
            *IDENTITY_HEADER_ORDER,
            *payment_headers,
            "지급합계",
            *deduction_headers,
            "공제합계",
            "차인지급액",
        ]
        output = pd.concat(outputs, ignore_index=True).reindex(columns=headers)
        for header in IDENTITY_HEADERS:
            output[header] = output[header].fillna("")
        for header in (set(headers) - IDENTITY_HEADERS):
            output[header] = output[header].fillna(0)
        output = self.clean_numeric_data(output[headers])
        self._last_column_kinds = {
            str(source_columns[index]): "payment" for index in payment_positions
        }
        self._last_column_kinds.update(
            {str(source_columns[index]): "deduction" for index in deduction_positions}
        )
        self._last_unknown_columns = unknown
        return (
            output,
            payment_headers,
            deduction_headers,
            active_payment_headers,
            active_deduction_headers,
        )

    @staticmethod
    def _month_from_sheet(sheet_name: str, fallback: int) -> int:
        matched = re.search(r"(?<!\d)(0?[1-9]|1[0-2])\s*월", sheet_name)
        return int(matched.group(1)) if matched else fallback

    def _write_months(
        self,
        months: dict[int, pd.DataFrame],
        output_path: Path,
        template_path: Path | None = None,
        headers: list[str] | None = None,
    ) -> None:
        if template_path is not None and Path(template_path).is_file():
            workbook = load_workbook(template_path)
            for worksheet in list(workbook.worksheets):
                workbook.remove(worksheet)
        else:
            workbook = Workbook()
            workbook.remove(workbook.active)
        if headers is None:
            headers = []
            for frame in months.values():
                _append_unique_headers(headers, frame.columns)
        if not headers:
            headers = list(STANDARD_HEADERS)
        header_alignment = Alignment(horizontal="center", vertical="center")
        left_alignment = Alignment(horizontal="left")
        right_alignment = Alignment(horizontal="right")
        for month in range(1, 13):
            worksheet = workbook.create_sheet(f"{month:02d}월")
            worksheet.sheet_view.showGridLines = False
            frame = months.get(month, pd.DataFrame(columns=headers))
            worksheet.append(headers)
            ordered_frame = frame.reindex(columns=headers)
            for record in ordered_frame.itertuples(index=False, name=None):
                values = []
                for header, value in zip(headers, record):
                    if pd.isna(value) or (header not in IDENTITY_HEADERS and value == 0):
                        value = None
                    values.append(value)
                worksheet.append(values)
            worksheet.freeze_panes = "E2"
            worksheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, worksheet.max_row)}"
            worksheet.row_dimensions[1].height = 22
            for cell in worksheet[1]:
                cell.fill = HEADER_FILL
                cell.font = HEADER_FONT
                cell.alignment = header_alignment
                cell.border = THIN_BORDER
            for column in range(1, len(headers) + 1):
                worksheet.column_dimensions[get_column_letter(column)].width = 16 if column > 4 else (16 if column == 3 else 13)
            for row in range(2, worksheet.max_row + 1):
                worksheet.cell(row, 1).number_format = "@"
                worksheet.cell(row, 3).number_format = "@"
                for column in range(1, len(headers) + 1):
                    cell = worksheet.cell(row, column)
                    cell.border = THIN_BORDER
                    cell.alignment = right_alignment if column > 4 else left_alignment
                    if column > 4:
                        cell.number_format = MONEY_FORMAT
            worksheet.sheet_properties.pageSetUpPr.fitToPage = True
            worksheet.page_setup.orientation = "landscape"
            worksheet.page_setup.fitToWidth = 1
            worksheet.page_setup.fitToHeight = 0
            worksheet.print_title_rows = "1:1"
        workbook.calculation.calcMode = "auto"
        destination = Path(output_path).resolve()
        destination.parent.mkdir(parents=True, exist_ok=True)
        temporary = destination.parent / f".{destination.stem}.{uuid4().hex}.tmp.xlsx"
        try:
            workbook.save(temporary)
            checked = load_workbook(temporary, read_only=True, data_only=False)
            try:
                if checked.sheetnames != [f"{month:02d}월" for month in range(1, 13)]:
                    raise SchemaMismatch("표준 급여대장의 월별 시트 검증에 실패했습니다.")
                for worksheet in checked.worksheets:
                    if [cell.value for cell in worksheet[1]] != headers:
                        raise SchemaMismatch(f"표준 급여대장 헤더 검증에 실패했습니다: {worksheet.title}")
            finally:
                checked.close()
            os.replace(temporary, destination)
        finally:
            temporary.unlink(missing_ok=True)

    def write_to_template(
        self,
        df: pd.DataFrame,
        template_path: Path | None,
        output_path: Path,
    ) -> None:
        self._write_months({1: df}, output_path, template_path)

    def convert_excel_workbook(
        self,
        source_path: Path,
        output_path: Path,
        *,
        template_path: Path | None = None,
        strict: bool = True,
    ) -> WorkbookConversionResult:
        source = Path(source_path).resolve(strict=True)
        months: dict[int, pd.DataFrame] = {}
        warnings: list[ValidationWarning] = []
        unknown: dict[str, str] = {}
        payment_headers: list[str] = []
        deduction_headers: list[str] = []
        active_payment_headers: list[str] = []
        active_deduction_headers: list[str] = []
        fallback = 1
        with pd.ExcelFile(source) as excel_file:
            sheet_names = list(excel_file.sheet_names)
            header_rows: dict[str, int] = {}
            for sheet_name in sheet_names:
                try:
                    header_rows[sheet_name] = self._detect_header_row_from_excel(
                        excel_file,
                        sheet_name,
                        source.name,
                    )
                except SchemaMismatch:
                    continue
            coded_labels = self._collect_coded_item_labels(
                source,
                sheet_names,
                excel_file=excel_file,
                header_rows=header_rows,
            )
            for sheet_name in sheet_names:
                if sheet_name not in header_rows:
                    continue
                try:
                    (
                        frame,
                        sheet_payment_headers,
                        sheet_deduction_headers,
                        sheet_active_payment_headers,
                        sheet_active_deduction_headers,
                    ) = self._load_excel_sheet(
                        source,
                        sheet_name,
                        coded_labels,
                        excel_file=excel_file,
                        header_row=header_rows[sheet_name],
                    )
                except SchemaMismatch:
                    continue
                _append_unique_headers(payment_headers, sheet_payment_headers)
                _append_unique_headers(deduction_headers, sheet_deduction_headers)
                _append_unique_headers(active_payment_headers, sheet_active_payment_headers)
                _append_unique_headers(active_deduction_headers, sheet_active_deduction_headers)
                month = self._month_from_sheet(sheet_name, fallback)
                fallback = min(12, fallback + 1)
                months[month] = pd.concat(
                    [months.get(month, pd.DataFrame()), frame],
                    ignore_index=True,
                )
                warnings.extend(self.validate_totals(frame))
                unknown.update(self._last_unknown_columns)
        if not months:
            raise SchemaMismatch(f"급여 시트를 찾지 못했습니다: {source.name}")
        if strict and warnings:
            first = warnings[0]
            raise SchemaMismatch(
                f"급여 합계 검증에 실패했습니다: {source.name}, 행 {first.row_number}, "
                f"{first.employee_name}, {first.check}, 계산={first.calculated:g}, 원본={first.source:g}"
            )
        inactive_unknown_headers = {
            header
            for header in unknown
            if not any(
                header in frame and _has_numeric_activity(frame[header])
                for frame in months.values()
            )
        }
        if inactive_unknown_headers:
            logger.info(
                "최종 직원 데이터에 금액이 없는 미확정 열을 제거합니다: %s",
                ", ".join(sorted(inactive_unknown_headers)),
            )
            payment_headers = [
                header for header in payment_headers
                if header not in inactive_unknown_headers
            ]
            deduction_headers = [
                header for header in deduction_headers
                if header not in inactive_unknown_headers
            ]
            active_payment_headers = [
                header for header in active_payment_headers
                if header not in inactive_unknown_headers
            ]
            active_deduction_headers = [
                header for header in active_deduction_headers
                if header not in inactive_unknown_headers
            ]
            unknown = {
                header: kind
                for header, kind in unknown.items()
                if header not in inactive_unknown_headers
            }
        output_headers = [
            *IDENTITY_HEADER_ORDER,
            *payment_headers,
            "지급합계",
            *deduction_headers,
            "공제합계",
            "차인지급액",
        ]
        invalid_headers = [header for header in output_headers if not _source_header(header)]
        if invalid_headers:
            raise SchemaMismatch(f"유효하지 않은 출력 헤더가 있습니다: {invalid_headers}")
        self._write_months(months, output_path, template_path, output_headers)
        return WorkbookConversionResult(
            Path(output_path).resolve(),
            {f"{month:02d}월": len(frame) for month, frame in sorted(months.items())},
            warnings,
            unknown,
            payment_headers,
            deduction_headers,
            active_payment_headers,
            active_deduction_headers,
        )

    @staticmethod
    def _document_signature(path: Path) -> str:
        try:
            _parse_payroll_name(path)
            return "tabular_payroll"
        except SchemaMismatch:
            pass
        reader = PdfReader(str(path))
        text = re.sub(r"\s+", "", reader.pages[0].extract_text() or "")
        if "소득자별근로소득원천징수부" in text:
            return "withholding_ledger"
        if "별지제24호서식(1)" in text or "근로소득원천징수영수증" in text:
            return "withholding_receipt"
        if all(keyword in text for keyword in ("지급총액", "공제총액", "실지급액")):
            return "tabular_payroll"
        return "unknown"

    @staticmethod
    def _business_id(path: Path) -> str:
        reader = PdfReader(str(path))
        compact = re.sub(r"\s+", "", reader.pages[0].extract_text() or "")
        matched = re.search(r"사업자등록번호([0-9-]{10,14})", compact)
        if matched is None:
            raise SchemaMismatch(f"문서에서 사업자번호를 찾지 못했습니다: {path.name}")
        digits = re.sub(r"\D", "", matched.group(1))[:10]
        if len(digits) != 10:
            raise SchemaMismatch(f"사업자번호가 10자리가 아닙니다: {path.name}")
        return f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"

    @staticmethod
    def _company_metadata(receipt_path: Path) -> tuple[str, str]:
        reader = PdfReader(str(receipt_path))
        text = reader.pages[0].extract_text() or ""
        compact = re.sub(r"\s+", "", text)
        name_match = re.search(
            r"(?:①|1\.)?법인명\(상호\)(.+?)(?:②대표자|2\.대표자|징수의무자|②사업자등록번호)",
            compact,
        )
        if name_match is None:
            name_match = re.search(
                r"([가-힣()]{2,50})1\.법인명\(상호\)2\.대표자",
                compact,
            )
        id_match = re.search(r"(?:③|3\.)?사업자등록번호([0-9-]{10,14})", compact)
        if id_match is None:
            id_match = re.search(r"(\d{3}-\d{2}-\d{5})", compact)
        if name_match is None or id_match is None:
            raise SchemaMismatch(f"원천징수영수증에서 회사명·사업자번호를 찾지 못했습니다: {receipt_path.name}")
        company_name = re.sub(r"\s+", " ", name_match.group(1)).strip()
        if company_name.startswith("주식회사") and not company_name.startswith("주식회사 "):
            company_name = f"주식회사 {company_name[len('주식회사'):]}"
        digits = re.sub(r"\D", "", id_match.group(1))[:10]
        if len(digits) != 10:
            raise SchemaMismatch(f"사업자번호가 10자리가 아닙니다: {receipt_path.name}")
        return company_name, f"{digits[:3]}-{digits[3:5]}-{digits[5:]}"

    @staticmethod
    def _split_bundle(source: Path, output_directory: Path) -> list[Path]:
        inspection = inspect_withholding_pdf(source)
        reader = PdfReader(str(source))
        output_directory.mkdir(parents=True, exist_ok=False)
        outputs: list[Path] = []
        total = len(inspection.first_pages)
        logger.info("원천징수 PDF 분할 시작: %s (총 %d개)", source.name, total)
        try:
            for index, start in enumerate(inspection.first_pages, 1):
                end = inspection.first_pages[index] if index < len(inspection.first_pages) else inspection.page_count
                writer = PdfWriter()
                for page in range(start, end):
                    writer.add_page(reader.pages[page])
                target = output_directory / f"{index}_receipt.pdf"
                with target.open("wb") as stream:
                    writer.write(stream)
                outputs.append(target)
        except Exception as exc:
            logger.error("원천징수 PDF 분할 오류: %s, %s", source.name, exc)
            raise
        logger.info("원천징수 PDF 분할 성공: %d개", len(outputs))
        return outputs

    @staticmethod
    def _split_ledger(source: Path, output_directory: Path) -> list[Path]:
        employees = extract_ledger_employees(source)
        reader = PdfReader(str(source))
        starts = [employee.source_page - 1 for employee in employees]
        output_directory.mkdir(parents=True, exist_ok=False)
        outputs: list[Path] = []
        for index, start in enumerate(starts, 1):
            end = starts[index] if index < len(starts) else len(reader.pages)
            if end - start != 3:
                raise SchemaMismatch(f"소득자별 원천징수부 직원별 페이지가 3쪽이 아닙니다: 인덱스 {index}")
            writer = PdfWriter()
            for page in range(start, end):
                writer.add_page(reader.pages[page])
            target = output_directory / f"{index}_ledger.pdf"
            with target.open("wb") as stream:
                writer.write(stream)
            outputs.append(target)
        return outputs

    @staticmethod
    def _year_from_text(path: Path) -> int:
        reader = PdfReader(str(path))
        text = reader.pages[0].extract_text() or ""
        compact = re.sub(r"\s+", "", text)
        matched = re.search(r"(?:귀속연도|근무기간)(20\d{2})", compact)
        if matched is None:
            matched = re.search(r"(20\d{2})", path.name)
        if matched is None:
            raise SchemaMismatch(f"문서에서 귀속연도를 찾지 못했습니다: {path.name}")
        return int(matched.group(1))

    def _convert_tabular_group(
        self,
        source_root: Path,
        staging: Path,
        basic_template: Path,
        existing_output_directory: Path,
    ) -> dict[str, object]:
        payroll_files = []
        payroll_workbooks = sorted(
            path
            for path in source_root.rglob("*")
            if path.is_file()
            and path.suffix.lower() in {".xlsx", ".xlsm"}
            and "급여대장" in path.parent.name
            and not path.name.startswith("~$")
        )
        receipt_files = []
        for path in source_root.rglob("*.pdf"):
            try:
                _parse_payroll_name(path)
                payroll_files.append(path)
                continue
            except SchemaMismatch:
                pass
            if self._document_signature(path) == "withholding_receipt":
                receipt_files.append(path)
        if not (payroll_files or payroll_workbooks) or not receipt_files:
            raise SchemaMismatch(f"표형 급여대장 그룹의 입력이 부족합니다: {source_root}")
        logger.info(
            "변환 입력 확인 완료: 급여대장 PDF %d개, Excel %d개, 원천징수영수증 PDF %d개",
            len(payroll_files),
            len(payroll_workbooks),
            len(receipt_files),
        )
        company_name, company_id = self._company_metadata(receipt_files[0])
        split_root = staging / f".{uuid4().hex}_receipts"
        split_root.mkdir(parents=True, exist_ok=False)
        receipt_paths_by_year: dict[int, list[Path]] = {}
        receipt_total = len(receipt_files)
        logger.info("원천징수 통합본 처리 시작 (총 %d개)", receipt_total)
        for source_index, receipt_path in enumerate(sorted(receipt_files), 1):
            year = _receipt_year(receipt_path)
            source_parts = self._split_bundle(receipt_path, split_root / f"source_{source_index}")
            year_directory = split_root / str(year)
            year_directory.mkdir(parents=True, exist_ok=True)
            collected = receipt_paths_by_year.setdefault(year, [])
            for part in source_parts:
                target = year_directory / f"{len(collected) + 1}_receipt.pdf"
                os.replace(part, target)
                collected.append(target)
        logger.info(
            "원천징수 통합본 처리 완료: %d개",
            sum(len(paths) for paths in receipt_paths_by_year.values()),
        )
        receipt_data = {}
        receipt_years = sorted(receipt_paths_by_year.items())
        for year, paths in receipt_years:
            receipt_data[year] = _extract_all(paths, year)
        basic_output = staging / f"{company_name}_기초정보.xlsx"
        existing_basic = existing_output_directory / basic_output.name
        build_dawon_basic_info(
            template_path=existing_basic if existing_basic.is_file() else basic_template,
            receipt_data=receipt_data,
            output_path=basic_output,
            company_name=company_name,
            company_id=company_id,
        )
        receipt_pdf_archive = {}
        archive_root = staging / "1. 급상여현황 생성" / "원천징수(PDF)"
        archive_years = sorted(receipt_data.items())
        logger.info("개인별 원천징수 PDF 보관 시작 (총 %d개 연도)", len(archive_years))
        for year, rows in archive_years:
            batch = archive_prepared_withholding_pdfs(
                pdf_files=receipt_paths_by_year[year],
                identities=rows,
                year=year,
                output_dir=archive_root / f"{str(year)[-2:]}년",
            )
            receipt_pdf_archive[str(year)] = {
                "count": len(batch.output_files),
                "folder": f"{str(year)[-2:]}년",
                "files": [path.name for path in batch.output_files],
            }
        logger.info(
            "개인별 원천징수 PDF 보관 완료: %d개",
            sum(item["count"] for item in receipt_pdf_archive.values()),
        )
        shutil.rmtree(split_root, ignore_errors=True)
        outputs = [basic_output]
        payroll = {}
        payment_headers: list[str] = []
        deduction_headers: list[str] = []
        payment_headers_by_year: dict[int, list[str]] = {}
        deduction_headers_by_year: dict[int, list[str]] = {}
        pdf_files_by_year: dict[int, list[Path]] = {}
        for path in payroll_files:
            pdf_files_by_year.setdefault(_parse_payroll_name(path)[0], []).append(path)
        workbooks_by_year: dict[int, list[Path]] = {}
        for path in payroll_workbooks:
            workbooks_by_year.setdefault(_payroll_workbook_year(path), []).append(path)
        duplicate_years = sorted(set(pdf_files_by_year) & set(workbooks_by_year))
        if duplicate_years:
            raise SchemaMismatch(f"같은 연도의 PDF·Excel 급여대장이 함께 있습니다: {duplicate_years}")
        multiple_workbook_years = sorted(
            year for year, paths in workbooks_by_year.items() if len(paths) > 1
        )
        if multiple_workbook_years:
            raise SchemaMismatch(
                f"같은 연도의 Excel 급여대장이 두 개 이상입니다: {multiple_workbook_years}"
            )
        payroll_years = sorted(set(pdf_files_by_year) | set(workbooks_by_year))
        logger.info("연도별 급여대장 생성 시작 (총 %d개 연도)", len(payroll_years))
        for year in payroll_years:
            target = staging / f"{company_name}_{str(year)[-2:]}년 급여대장.xlsx"
            year_files = pdf_files_by_year.get(year, [])
            if year_files:
                intermediate = staging / f".{uuid4().hex}_{year}.xlsx"
                _, controls, _, _ = build_dawon_payroll_workbook(
                    year=year,
                    pdf_files=year_files,
                    output_path=intermediate,
                )
                conversion = self.convert_excel_workbook(intermediate, target, strict=True)
                intermediate.unlink(missing_ok=True)
            else:
                controls = []
                conversion = self.convert_excel_workbook(
                    workbooks_by_year[year][0],
                    target,
                    strict=True,
                )
            _append_unique_headers(payment_headers, conversion.payment_headers)
            _append_unique_headers(deduction_headers, conversion.deduction_headers)
            payment_headers_by_year[year] = conversion.active_payment_headers
            deduction_headers_by_year[year] = conversion.active_deduction_headers
            outputs.append(target)
            payroll[str(year)] = {
                "rows": sum(conversion.sheet_row_counts.values()),
                "source_controls": len(controls),
                "payment_headers": conversion.payment_headers,
                "deduction_headers": conversion.deduction_headers,
                "active_payment_headers": conversion.active_payment_headers,
                "active_deduction_headers": conversion.active_deduction_headers,
                "unknown_columns": conversion.unknown_columns,
            }
        logger.info("연도별 급여대장 생성 완료: %d개", len(payroll))
        _update_basic_info_item_headers(
            basic_output,
            payment_headers_by_year,
            deduction_headers_by_year,
            payment_headers,
            deduction_headers,
        )
        employment_info, health_roster = discover_employee_info_sources(source_root)
        employee_info = populate_basic_info_employee_sheets(
            basic_info_path=basic_output,
            employment_info_path=employment_info,
            health_roster_path=health_roster,
            occupation_linkage_path=basic_template.parent.parent / "직업분류_연계표.xlsx",
        )
        return {
            "profile": "tabular_payroll_pdf",
            "company_name": company_name,
            "company_id": company_id,
            "receipt_counts": {str(year): len(rows) for year, rows in receipt_data.items()},
            "receipt_pdf_archive": receipt_pdf_archive,
            "payment_headers": payment_headers,
            "deduction_headers": deduction_headers,
            "employee_info": employee_info.to_report(),
            "payroll": payroll,
            "outputs": outputs,
        }

    def _convert_ledger_group(
        self,
        receipt: Path,
        ledger: Path,
        staging: Path,
        basic_template: Path,
        existing_output_directory: Path,
    ) -> dict[str, object]:
        company_name, company_id = self._company_metadata(receipt)
        year = self._year_from_text(ledger)
        receipt_dir = staging / f".{uuid4().hex}_receipts"
        ledger_dir = staging / f".{uuid4().hex}_ledgers"
        receipt_files = self._split_bundle(receipt, receipt_dir)
        ledger_files = self._split_ledger(ledger, ledger_dir)
        if len(receipt_files) != len(ledger_files):
            raise SchemaMismatch(
                f"영수증과 소득자별 원천징수부 인원수가 다릅니다: {len(receipt_files)}, {len(ledger_files)}"
            )
        basic_output = staging / f"{company_name}_기초정보.xlsx"
        existing_basic = existing_output_directory / basic_output.name
        basic = build_basic_info_from_withholding_pdfs(
            basic_info_path=existing_basic if existing_basic.is_file() else basic_template,
            pdf_files=receipt_files,
            year=year,
            output_path=basic_output,
            ledger_pdf=ledger,
            company_name=company_name,
            company_id=company_id,
        )
        intermediate = staging / f".{uuid4().hex}_{year}.xlsx"
        ledger_result = build_monthly_payroll_ledger_from_pdfs(
            ledger_pdf_directory=ledger_dir,
            basic_info_path=basic_output,
            output_path=intermediate,
            year=year,
        )
        target = staging / f"{company_name}_{str(year)[-2:]}년 급여대장.xlsx"
        conversion = self.convert_excel_workbook(intermediate, target, strict=True)
        _update_basic_info_item_headers(
            basic_output,
            {year: conversion.active_payment_headers},
            {year: conversion.active_deduction_headers},
            conversion.payment_headers,
            conversion.deduction_headers,
        )
        employment_info, health_roster = discover_employee_info_sources(receipt.parent)
        employee_info = populate_basic_info_employee_sheets(
            basic_info_path=basic_output,
            employment_info_path=employment_info,
            health_roster_path=health_roster,
            occupation_linkage_path=basic_template.parent.parent / "직업분류_연계표.xlsx",
        )
        intermediate.unlink(missing_ok=True)
        basic.manifest_path.unlink(missing_ok=True)
        ledger_result.manifest_path.unlink(missing_ok=True)
        shutil.rmtree(receipt_dir, ignore_errors=True)
        shutil.rmtree(ledger_dir, ignore_errors=True)
        return {
            "profile": "withholding_ledger_pdf",
            "company_name": company_name,
            "company_id": company_id,
            "year": year,
            "employee_count": basic.employee_count,
            "payroll_rows": sum(conversion.sheet_row_counts.values()),
            "payment_headers": conversion.payment_headers,
            "deduction_headers": conversion.deduction_headers,
            "active_payment_headers": conversion.active_payment_headers,
            "active_deduction_headers": conversion.active_deduction_headers,
            "unknown_columns": conversion.unknown_columns,
            "employee_info": employee_info.to_report(),
            "outputs": [basic_output, target],
        }

    def convert_company_input(
        self,
        *,
        input_root: Path,
        output_directory: Path,
        basic_info_template: Path,
    ) -> UniversalConversionResult:
        """한 거래처의 분류된 1.Input을 기초정보·연도별 급여대장으로 변환합니다."""
        source_root = Path(input_root).resolve(strict=True)
        destination = Path(output_directory).resolve()
        template = Path(basic_info_template).resolve(strict=True)
        staging = destination.parent / f".{destination.name}.{uuid4().hex}.tmp"
        staging.mkdir(parents=True, exist_ok=False)
        try:
            group = self._convert_tabular_group(source_root, staging, template, destination)
            outputs = group.pop("outputs")
            report = {
                "input_root": str(source_root),
                "mapping_mode": "header_signature_and_synonym",
                "unknown_item_policy": self.unknown_item_policy,
                "groups": [group],
                "outputs": [
                    {"name": path.name, "size": path.stat().st_size, "sha256": sha256(path)}
                    for path in outputs
                ],
            }
            report_path = staging / "trans_검증보고서.json"
            report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            destination.mkdir(parents=True, exist_ok=True)
            final_outputs = []
            for path in list(staging.iterdir()):
                if path.is_dir():
                    _merge_staged_directory(path, destination / path.name)
                    continue
                target = destination / path.name
                os.replace(path, target)
                if target.suffix.lower() == ".xlsx":
                    final_outputs.append(target)
            shutil.rmtree(staging, ignore_errors=True)
            return UniversalConversionResult(
                destination,
                tuple(sorted(final_outputs)),
                destination / report_path.name,
            )
        except Exception:
            shutil.rmtree(staging, ignore_errors=True)
            raise

    def convert_all(
        self,
        *,
        input_root: Path,
        output_directory: Path,
        basic_info_template: Path,
    ) -> UniversalConversionResult:
        source_root = Path(input_root).resolve(strict=True)
        destination = Path(output_directory).resolve()
        template = Path(basic_info_template).resolve(strict=True)
        staging = destination.parent / f".{destination.name}.{uuid4().hex}.tmp"
        staging.mkdir(parents=True, exist_ok=False)
        groups: list[dict[str, object]] = []
        outputs: list[Path] = []
        try:
            claimed: set[Path] = set()
            for child in source_root.iterdir():
                if not child.is_dir() or child.name.lower() == "output":
                    continue
                signatures = {self._document_signature(path) for path in child.rglob("*.pdf")}
                if "tabular_payroll" in signatures and "withholding_receipt" in signatures:
                    group = self._convert_tabular_group(child, staging, template, destination)
                    groups.append(group)
                    outputs.extend(group.pop("outputs"))
                    claimed.update(path.resolve() for path in child.rglob("*.pdf"))

            top_level_pdfs = [path for path in source_root.glob("*.pdf") if path.resolve() not in claimed]
            receipts = [
                path
                for path in top_level_pdfs
                if self._document_signature(path) == "withholding_receipt"
                and len(inspect_withholding_pdf(path).first_pages) > 1
            ]
            ledgers = [path for path in top_level_pdfs if self._document_signature(path) == "withholding_ledger"]
            ledger_by_id = {self._business_id(path): path for path in ledgers}
            for receipt in receipts:
                try:
                    company_id = self._company_metadata(receipt)[1]
                except SchemaMismatch:
                    continue
                ledger = ledger_by_id.get(company_id)
                if ledger is None:
                    continue
                group = self._convert_ledger_group(receipt, ledger, staging, template, destination)
                groups.append(group)
                outputs.extend(group.pop("outputs"))

            if not groups:
                raise SchemaMismatch(f"변환 가능한 급여 문서 그룹을 찾지 못했습니다: {source_root}")
            report = {
                "input_root": str(source_root),
                "mapping_mode": "header_signature_and_synonym",
                "unknown_item_policy": self.unknown_item_policy,
                "groups": groups,
                "outputs": [
                    {"name": path.name, "size": path.stat().st_size, "sha256": sha256(path)} for path in outputs
                ],
            }
            report_path = staging / "trans_검증보고서.json"
            report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
            destination.mkdir(parents=True, exist_ok=True)
            final_outputs = []
            for path in list(staging.iterdir()):
                if path.is_dir():
                    _merge_staged_directory(path, destination / path.name)
                    continue
                target = destination / path.name
                os.replace(path, target)
                if target.suffix.lower() == ".xlsx":
                    final_outputs.append(target)
            shutil.rmtree(staging, ignore_errors=True)
            return UniversalConversionResult(
                destination,
                tuple(sorted(final_outputs)),
                destination / report_path.name,
            )
        except Exception:
            shutil.rmtree(staging, ignore_errors=True)
            raise


def main() -> None:
    parser = argparse.ArgumentParser(description="이기종 급여 문서를 표준 월별 급여대장으로 변환합니다.")
    parser.add_argument("input_root", type=Path)
    parser.add_argument("basic_info_template", type=Path)
    parser.add_argument("output_directory", type=Path)
    parser.add_argument("--unknown-item-policy", choices=("aggregate", "separate"), default="aggregate")
    args = parser.parse_args()
    result = PayrollConverter(unknown_item_policy=args.unknown_item_policy).convert_all(
        input_root=args.input_root,
        output_directory=args.output_directory,
        basic_info_template=args.basic_info_template,
    )
    print(result.report_path)


if __name__ == "__main__":
    main()
