from __future__ import annotations

from collections import defaultdict
from copy import copy
from dataclasses import dataclass
import logging
import os
from pathlib import Path
import re
import shutil
from typing import Callable
from uuid import uuid4

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font
from openpyxl.formula.translate import Translator
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

from .employment_dates import parse_date
from .errors import SchemaMismatch
from .infra import safe_name, snapshot_file
from .payroll_status import NON_TAX_OVERTIME_HEADER, _digits, _normalized_name, _resident_key


logger = logging.getLogger(__name__)


STATUS_YEAR_PATTERN = re.compile(r"_(\d{2})년\s*급상여현황(?:\(수정\))?\.xlsx$", re.IGNORECASE)
PAYROLL_YEAR_PATTERN = re.compile(r"_(\d{2})년\s*급여대장\.xlsx$", re.IGNORECASE)
PERIOD_PATTERN = re.compile(r"^(?P<year>\d{2})년(?P<month>0[1-9]|1[0-2])월$")
MONTH_SHEET_PATTERN = re.compile(r"^(?P<month>0?[1-9]|1[0-2])월$")
IDENTITY_ALIASES = {
    "employee_code": ("사번", "사원번호"),
    "name": ("성명", "이름"),
    "resident_id": ("주민번호", "주민등록번호"),
    "acquired": ("입사일",),
    "lost": ("퇴사일",),
    "period": ("지급년월",),
    "pay_type": ("급여구분",),
    "payment_total": ("지급총액", "지급합계"),
    "deduction_total": ("공제총액", "공제합계"),
    "net_pay": ("실지급액", "차인지급액"),
}
MONTHLY_BASE_HEADERS = ("사번", "성명", "주민번호", "입사일", "퇴사일")
MONEY_FORMAT = "#,###"
DATE_FORMAT = "yyyy-mm-dd"
LEGACY_PAYMENT_ORDER = {
    2023: ("기본급", "식대", "직책수당", "차량유지비", "연차수당", "기타수당", "전월소급분", "상여금", "반장수당", "지각조퇴", "주휴수당", "교통비", "격려금", "기타수당1", "기타수당2", "경영성과급", "휴일근로", "휴일연장야간", "약정야간수당", "약정휴일수당", "약정연장수당", "연장근로", "야간근로", "휴일야간", "약정연장수당1", "휴일근로수당", "연장근로수당", "야간근로수당"),
    2024: ("기본급", "식대", "직책수당", "차량유지비", "연차수당", "기타수당", "기타수당1", "상여금", "전월소급분", "지각조퇴", "반장수당", "주휴수당", "교통비", "경영성과급", "격려금", "기타수당2", "휴일근로", "휴일연장야간", "약정야간수당", "약정휴일수당", "약정연장수당", "연장근로", "야간근로", "휴일야간", "약정연장수당1", "휴일근로수당", "연장근로수당", "야간근로수당"),
    2025: ("기본급", "식대", "직책수당", "차량유지비", "연차수당", "기타수당", "전월소급분", "상여금", "반장수당", "지각조퇴", "주휴수당", "교통비", "기타수당1", "경영성과급", "격려금", "기타수당2", "휴일근로", "휴일야간", "약정야간수당", "약정휴일수당", "약정연장수당", "연장근로", "야간근로", "휴일연장야간", "약정연장수당1", "휴일근로수당", "연장근로수당", "야간근로수당"),
}
LEGACY_DEDUCTION_ORDER = {
    2023: ("소득세", "장기요양보험", "연말정산지방소득세", "지방소득세", "장기요양보험정산", "선지급금", "국민연금", "고용보험", "기타공제", "건강보험", "고용보험정산", "기타공제1", "건강보험정산", "연말정산소득세", "사우회비", "학자금", "가불상환"),
    2024: ("소득세", "장기요양보험", "연말정산지방소득세", "지방소득세", "장기요양보험정산", "선지급금", "국민연금", "고용보험", "기타공제", "건강보험", "고용보험정산", "기타공제1", "건강보험정산", "연말정산소득세", "가불상환", "사우회비", "학자금", "신원보증보험"),
    2025: ("소득세", "장기요양보험", "연말정산지방소득세", "지방소득세", "장기요양보험정산", "선지급금", "국민연금", "고용보험", "기타공제", "건강보험", "고용보험정산", "건강보험정산", "연말정산소득세", "기타공제1", "가불상환", "신원보증보험", "사우회비", "학자금"),
}
TOOL_REQUIRED_SHEETS = {"prevData", "Data", "Form", "TaxTable"}
TOOL_NON_TAX_DETAILS = (
    ("조특법30(T13)", "조특법30", "T13"),
    ("조특법30(T12)", "조특법30", "T12"),
    ("출산보육수당", "출산보육수당", "Q01"),
    ("야간근로수당", "야간근로수당", "O01"),
    ("비과세식대(P01)", "비과세식대", "P01"),
)


@dataclass(frozen=True)
class CorrectionInputResult:
    years: tuple[int, ...]
    modified_status_files: tuple[Path, ...]
    withholding_tool_files: tuple[Path, ...]
    monthly_payroll_files: tuple[Path, ...]
    issues: tuple[dict[str, object], ...]


def _find_header(headers: list[object], aliases: tuple[str, ...]) -> int:
    columns = {str(value or "").strip(): index for index, value in enumerate(headers)}
    for alias in aliases:
        if alias in columns:
            return columns[alias]
    raise SchemaMismatch(f"필수 헤더가 없습니다: {'/'.join(aliases)}")


def _formatted_resident(value: object) -> str:
    digits = _digits(value)
    if len(digits) == 13:
        return f"{digits[:6]}-{digits[6:]}"
    return str(value or "").strip()


def _year_from_status(path: Path) -> int:
    match = STATUS_YEAR_PATTERN.search(path.name)
    if match is None:
        raise SchemaMismatch(f"급상여현황 파일명에서 연도를 찾지 못했습니다: {path.name}")
    return 2000 + int(match.group(1))


def _year_from_payroll(path: Path) -> int:
    match = PAYROLL_YEAR_PATTERN.search(path.name)
    if match is None:
        raise SchemaMismatch(f"급여대장 파일명에서 연도를 찾지 못했습니다: {path.name}")
    return 2000 + int(match.group(1))


def _monthly_rows_from_status(
    status_path: Path,
    issues: list[dict[str, object]],
) -> tuple[list[str], dict[int, list[list[object]]]]:
    workbook = load_workbook(status_path, data_only=True, read_only=True)
    try:
        if "대상자(작업)" not in workbook.sheetnames or "비대상자" not in workbook.sheetnames:
            raise SchemaMismatch(f"수정 급상여현황의 필수 시트가 없습니다: {status_path.name}")
        target_headers = [cell.value for cell in workbook["대상자(작업)"][1]]
        target_columns = {
            key: _find_header(target_headers, aliases)
            for key, aliases in IDENTITY_ALIASES.items()
        }
        year = _year_from_status(status_path)
        available_payment_headers = [
            str(header).strip()
            for header in target_headers[target_columns["pay_type"] + 1:target_columns["payment_total"]]
            if header not in (None, "") and str(header).strip() != NON_TAX_OVERTIME_HEADER
        ]
        available_deduction_headers = [
            str(header).strip()
            for header in target_headers[target_columns["payment_total"] + 1:target_columns["deduction_total"]]
            if header not in (None, "")
        ]
        payment_headers = [
            header for header in LEGACY_PAYMENT_ORDER.get(year, tuple(available_payment_headers))
            if header in available_payment_headers
        ]
        payment_headers.extend(header for header in available_payment_headers if header not in payment_headers)
        deduction_headers = [
            header for header in LEGACY_DEDUCTION_ORDER.get(year, tuple(available_deduction_headers))
            if header in available_deduction_headers
        ]
        deduction_headers.extend(
            header for header in available_deduction_headers
            if header not in deduction_headers
        )
        output_headers = [
            *MONTHLY_BASE_HEADERS,
            *payment_headers,
            NON_TAX_OVERTIME_HEADER,
            "지급총액",
            *deduction_headers,
            "공제총액",
            "실지급액",
        ]
        monthly: defaultdict[int, list[list[object]]] = defaultdict(list)
        for sheet_name in ("대상자(작업)", "비대상자"):
            sheet = workbook[sheet_name]
            headers = [cell.value for cell in sheet[1]]
            columns = {
                key: _find_header(headers, aliases)
                for key, aliases in IDENTITY_ALIASES.items()
            }
            header_columns = {
                str(header).strip(): index
                for index, header in enumerate(headers)
                if header not in (None, "")
            }
            for values in sheet.iter_rows(min_row=2, values_only=True):
                period = str(values[columns["period"]] or "").strip()
                match = PERIOD_PATTERN.fullmatch(period)
                if match is None:
                    continue
                month = int(match.group("month"))
                name = values[columns["name"]]
                if name in (None, ""):
                    continue
                employee_code = values[columns["employee_code"]]
                resident_id = values[columns["resident_id"]]
                acquired = parse_date(values[columns["acquired"]])
                lost = parse_date(values[columns["lost"]])
                row: list[object] = [
                    _formatted_resident(employee_code),
                    name,
                    _formatted_resident(resident_id or employee_code),
                    acquired,
                    lost,
                ]
                for header in payment_headers:
                    index = header_columns.get(header)
                    value = values[index] if index is not None else None
                    row.append(None if value == 0 else value)
                overtime_column = header_columns.get(NON_TAX_OVERTIME_HEADER)
                overtime = values[overtime_column] if overtime_column is not None else None
                row.append(
                    0 if sheet_name == "대상자(작업)" and overtime in (None, "")
                    else None if overtime in (None, "", 0)
                    else overtime
                )
                row.append(values[columns["payment_total"]])
                for header in deduction_headers:
                    index = header_columns.get(header)
                    value = values[index] if index is not None else None
                    row.append(None if value == 0 else value)
                row.extend(
                    [
                        values[columns["deduction_total"]],
                        values[columns["net_pay"]],
                    ]
                )
                monthly[month].append(row)
        return output_headers, dict(monthly)
    finally:
        workbook.close()


def _row_as_dict(headers: list[object], values: tuple[object, ...]) -> dict[str, object]:
    return {
        str(header).strip(): values[index] if index < len(values) else None
        for index, header in enumerate(headers)
        if header not in (None, "")
    }


def _tool_identity_key(record: dict[str, object]) -> str:
    resident = _digits(record.get("주민등록번호") or record.get("주민번호"))
    employee_code = _digits(record.get("사원번호") or record.get("사번"))
    if len(resident) == 13:
        return resident
    if len(employee_code) == 13:
        return employee_code
    return resident or employee_code


def _tool_index_key(value: object) -> str:
    text = str(value or "").strip()
    return text[:-2] if text.endswith(".0") and text[:-2].isdigit() else text


def _tool_tax_records(basic_info_path: Path, year: int) -> dict[str, list[dict[str, object]]]:
    workbook = load_workbook(basic_info_path, data_only=True, read_only=True)
    try:
        sheet_name = f"원천징수추출({str(year)[-2:]}년)"
        if sheet_name not in workbook.sheetnames:
            raise SchemaMismatch(f"기초정보에 {sheet_name} 시트가 없습니다.")
        sheet = workbook[sheet_name]
        headers = [cell.value for cell in sheet[1]]
        records: defaultdict[str, list[dict[str, object]]] = defaultdict(list)
        for values in sheet.iter_rows(min_row=2, values_only=True):
            record = _row_as_dict(headers, values)
            identity = _tool_identity_key(record)
            if not identity:
                continue
            records[identity].append(record)
        return dict(records)
    finally:
        workbook.close()


def _tool_target_rows(status_path: Path) -> list[dict[str, object]]:
    workbook = load_workbook(status_path, data_only=True, read_only=True)
    try:
        if "대상자(작업)" not in workbook.sheetnames:
            raise SchemaMismatch(f"수정 급상여현황에 대상자(작업) 시트가 없습니다: {status_path.name}")
        sheet = workbook["대상자(작업)"]
        headers = [cell.value for cell in sheet[1]]
        period_column = _find_header(headers, ("지급년월",))
        rows: list[dict[str, object]] = []
        for values in sheet.iter_rows(min_row=2, values_only=True):
            period = str(values[period_column] or "").strip()
            if not period.endswith("년합계"):
                continue
            row = _row_as_dict(headers, values)
            identity = _tool_identity_key(row)
            if not identity:
                raise SchemaMismatch(f"대상자 합계행의 식별정보가 없습니다: {status_path.name}")
            rows.append(row)
        if not rows:
            raise SchemaMismatch(f"Tool에 적재할 대상자 합계행이 없습니다: {status_path.name}")
        return rows
    finally:
        workbook.close()


def _formula_text(value: object) -> str | None:
    if isinstance(value, str) and value.startswith("="):
        return value
    text = getattr(value, "text", None)
    return text if isinstance(text, str) and text.startswith("=") else None


def _match_tool_tax_record(
    target: dict[str, object],
    tax_records: dict[str, list[dict[str, object]]],
) -> dict[str, object] | None:
    candidates = tax_records.get(_tool_identity_key(target), [])
    if not candidates:
        return None
    target_index = _tool_index_key(target.get("인덱스"))
    index_matches = [
        record for record in candidates
        if _tool_index_key(record.get("인덱스")) == target_index
    ]
    if len(index_matches) == 1:
        return index_matches[0]
    if len(candidates) == 1:
        return candidates[0]
    identity = _tool_identity_key(target)
    raise SchemaMismatch(
        f"원천징수 자료를 인덱스로 구분하지 못했습니다: {identity[:6]}-*******, "
        f"인덱스={target_index or '(없음)'}"
    )


def _populate_withholding_tool_data(
    *,
    tool_path: Path,
    status_path: Path,
    basic_info_path: Path,
    year: int,
    issues: list[dict[str, object]],
) -> int:
    tax_records = _tool_tax_records(basic_info_path, year)
    target_rows = _tool_target_rows(status_path)
    workbook = load_workbook(tool_path)
    temporary = tool_path.with_name(f".{tool_path.stem}.{uuid4().hex}.tmp.xlsx")
    try:
        missing = TOOL_REQUIRED_SHEETS - set(workbook.sheetnames)
        if missing:
            raise SchemaMismatch(f"수정 원천징수 Tool 시트가 부족합니다: {sorted(missing)}")
        previous = workbook["prevData"]
        data = workbook["Data"]
        previous_headers = [cell.value for cell in previous[1]]
        data_headers = [
            previous_headers[index]
            if _formula_text(cell.value) and index < len(previous_headers)
            else cell.value
            for index, cell in enumerate(data[1])
        ]
        previous_columns = {
            str(header).strip(): index + 1
            for index, header in enumerate(previous_headers)
            if header not in (None, "")
        }
        data_columns = {
            str(header).strip(): index + 1
            for index, header in enumerate(data_headers)
            if header not in (None, "")
        }
        for required in ("기간", "연번", "사번", "이름", "주민등록번호", "인덱스"):
            if required not in previous_columns or required not in data_columns:
                raise SchemaMismatch(f"수정 원천징수 Tool 필수 헤더가 없습니다: {required}")

        template_formulas = {
            column: (
                _formula_text(data.cell(2, column).value),
                isinstance(data.cell(2, column).value, ArrayFormula),
            )
            for column in range(1, data.max_column + 1)
        }
        template_styles = {
            column: copy(data.cell(2, column)._style)
            for column in range(1, data.max_column + 1)
        }
        for sheet in (previous, data):
            for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row):
                for cell in row:
                    cell.value = None

        populated_rows: list[dict[str, object]] = []
        for serial, target in enumerate(target_rows, 1):
            tax = _match_tool_tax_record(target, tax_records)
            if tax is None:
                issues.append(
                    {
                        "year": year,
                        "employee_code": str(target.get("사원번호") or ""),
                        "employee_name": str(target.get("성명") or ""),
                        "type": "Tool 원천징수 미연결",
                    }
                )
                tax = {}
            row = {
                header: target.get(header, tax.get(header))
                for header in previous_columns
            }
            row.update(
                {
                    "기간": f"{year}년",
                    "연번": serial,
                    "사번": target.get("사원번호") or tax.get("사원번호"),
                    "이름": target.get("성명") or tax.get("성명"),
                    "주민등록번호": target.get("주민등록번호") or tax.get("주민등록번호"),
                    "인덱스": target.get("인덱스") or tax.get("인덱스"),
                    "18-1.생산직등야간근로수당": target.get(NON_TAX_OVERTIME_HEADER) or 0,
                }
            )
            total_details: list[tuple[str, str, object]] = []
            current_details: list[tuple[str, str, object]] = []
            for source_name, display_name, code in TOOL_NON_TAX_DETAILS:
                total = tax.get(f"{source_name}_총계")
                current = tax.get(f"{source_name}_현계")
                if total not in (None, "", 0):
                    total_details.append((display_name, code, total))
                if current not in (None, "", 0):
                    current_details.append((display_name, code, current))
            if len(total_details) > 6 or len(current_details) > 6:
                raise SchemaMismatch(f"Tool 비과세 상세항목이 6개를 초과합니다: {row['이름']}")
            for index, (name, code, amount) in enumerate(total_details, 1):
                row[f"비과세명{index}_총"] = name
                row[f"코드{index}_총"] = code
                row[f"금액{index}_총"] = amount
            for index, (name, code, amount) in enumerate(current_details, 1):
                row[f"비과세명{index}_현"] = name
                row[f"코드{index}_현"] = code
                row[f"금액{index}_현"] = amount
            populated_rows.append(row)

        for row_number, row in enumerate(populated_rows, 2):
            for header, column in previous_columns.items():
                previous.cell(row_number, column, row.get(header))
            for header, column in data_columns.items():
                cell = data.cell(row_number, column)
                cell._style = copy(template_styles[column])
                source_formula, is_array_formula = template_formulas[column]
                if source_formula:
                    translated = Translator(
                        source_formula,
                        origin=data.cell(2, column).coordinate,
                    ).translate_formula(cell.coordinate)
                    cell.value = (
                        ArrayFormula(ref=cell.coordinate, text=translated)
                        if is_array_formula
                        else translated
                    )
                elif header in row:
                    cell.value = row.get(header)

        previous.auto_filter.ref = f"A1:{get_column_letter(previous.max_column)}{len(populated_rows) + 1}"
        data.auto_filter.ref = f"A1:{get_column_letter(data.max_column)}{len(populated_rows) + 1}"
        workbook.calculation.calcMode = "auto"
        workbook.calculation.fullCalcOnLoad = True
        workbook.calculation.forceFullCalc = True
        workbook.save(temporary)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise
    finally:
        workbook.close()

    try:
        verification = load_workbook(temporary, data_only=False, read_only=True)
        try:
            previous = verification["prevData"]
            data = verification["Data"]
            previous_count = sum(
                1 for (period,) in previous.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True)
                if period not in (None, "")
            )
            data_count = sum(
                1 for (period,) in data.iter_rows(min_row=2, min_col=1, max_col=1, values_only=True)
                if period not in (None, "")
            )
            if previous_count != len(populated_rows) or data_count != len(populated_rows):
                raise SchemaMismatch(
                    f"Tool 저장 행 수가 다릅니다: 대상 {len(populated_rows)}, "
                    f"prevData {previous_count}, Data {data_count}"
                )
        finally:
            verification.close()
        os.replace(temporary, tool_path)
    finally:
        temporary.unlink(missing_ok=True)
    return len(populated_rows)


def _sample_withholding_pdf(root: Path, year: int) -> Path:
    year_directory = root / f"{str(year)[-2:]}년"
    candidates = sorted(
        year_directory.glob("*.pdf"),
        key=lambda path: (
            0 if path.stem.partition("_")[0] == "10" else 1,
            int(path.stem.partition("_")[0]) if path.stem.partition("_")[0].isdigit() else 10**9,
            path.name,
        ),
    )
    if not candidates:
        raise SchemaMismatch(f"Tool에 삽입할 {year}년 분할 원천징수 PDF가 없습니다: {year_directory}")
    return candidates[0]


def _populate_tools_with_pdf_images(jobs: list[tuple[Path, Path]]) -> None:
    try:
        import pymupdf
        import win32com.client
    except ImportError as exc:
        raise SchemaMismatch("Tool 이미지 삽입에는 PyMuPDF, pywin32와 Microsoft Excel이 필요합니다.") from exc

    excel = None
    workbook = None
    staging_roots: list[Path] = []
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        box_width = excel.CentimetersToPoints(49)
        box_height = excel.CentimetersToPoints(70)
        for tool_path, pdf_path in jobs:
            staging = tool_path.parent / f".{uuid4().hex}_tool_images"
            staging.mkdir(parents=True, exist_ok=False)
            staging_roots.append(staging)
            image_paths: list[Path] = []
            document = pymupdf.open(pdf_path)
            try:
                if document.page_count < 3:
                    raise SchemaMismatch(f"Tool 샘플 PDF가 3페이지 미만입니다: {pdf_path.name}")
                matrix = pymupdf.Matrix(300 / 72, 300 / 72)
                for page_index in range(3):
                    image_path = staging / f"page_{page_index + 1}.png"
                    document[page_index].get_pixmap(matrix=matrix, alpha=False).save(image_path)
                    image_paths.append(image_path)
            finally:
                document.close()

            workbook = excel.Workbooks.Open(str(tool_path), UpdateLinks=0, ReadOnly=False)
            sheet = workbook.Worksheets("Form")
            for shape_index in range(sheet.Shapes.Count, 0, -1):
                shape = sheet.Shapes.Item(shape_index)
                if str(shape.Name).startswith("InsertedImage_"):
                    shape.Delete()
            for column, image_path in enumerate(image_paths, 1):
                area = sheet.Cells(1, column).MergeArea
                shape = sheet.Shapes.AddPicture(
                    str(image_path), False, True, area.Left, area.Top, -1, -1
                )
                shape.Name = f"InsertedImage_{column}"
                shape.Placement = 1
                original_width = shape.Width
                original_height = shape.Height
                scale = min(box_width / original_width, box_height / original_height)
                shape.LockAspectRatio = -1
                shape.Width = original_width * scale
                shape.Height = original_height * scale
                shape.Left = area.Left + (box_width - shape.Width) / 2
                shape.Top = area.Top
                shape.ZOrder(1)
            excel.CalculateFullRebuild()
            workbook.Save()
            workbook.Close(SaveChanges=False)
            workbook = None
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        for staging in staging_roots:
            shutil.rmtree(staging, ignore_errors=True)


def _write_monthly_workbook(path: Path, headers: list[str], rows: list[list[object]]) -> None:
    workbook = Workbook(write_only=True)
    sheet = workbook.create_sheet("Sheet1")
    sheet.sheet_format.defaultRowHeight = 16.5
    header_alignment = Alignment(vertical="center")
    data_alignment = Alignment(vertical="center")
    header_font = Font(name="맑은 고딕", size=11, color="FF000000")
    header_red_font = Font(name="맑은 고딕", size=11, color="FFFF0000")
    data_font = Font(name="맑은 고딕", size=11)
    for column, header in enumerate(headers, 1):
        if column in (1, 3):
            width = 15
        elif column in (4, 5):
            width = 12
        elif header == NON_TAX_OVERTIME_HEADER:
            width = 20
        elif column > 5:
            width = max(12, min(18.5, len(header) * 1.25 + 2))
        else:
            width = max(6.5, min(18.5, len(header) * 1.25 + 2))
        sheet.column_dimensions[get_column_letter(column)].width = width
    header_cells = []
    for header in headers:
        cell = WriteOnlyCell(sheet, value=header)
        cell.font = header_red_font if header == NON_TAX_OVERTIME_HEADER else header_font
        cell.alignment = header_alignment
        header_cells.append(cell)
    sheet.append(header_cells)
    for values in rows:
        cells = []
        for column, value in enumerate(values, 1):
            cell = WriteOnlyCell(sheet, value=value)
            cell.font = data_font
            cell.alignment = data_alignment
            if column in (1, 3):
                cell.number_format = "@"
            elif column in (4, 5):
                cell.number_format = DATE_FORMAT
            elif column > 5 and cell.value not in (None, ""):
                cell.number_format = MONEY_FORMAT
            cells.append(cell)
        sheet.append(cells)
    workbook.save(path)
    workbook.close()


def _convert_xlsx_to_xls(pairs: list[tuple[Path, Path]]) -> None:
    try:
        import win32com.client
    except ImportError as exc:
        raise SchemaMismatch("월별 XLS 급여대장을 생성하려면 pywin32와 Microsoft Excel이 필요합니다.") from exc

    excel = None
    workbook = None
    temporary_xls: Path | None = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        for source, destination in pairs:
            destination.parent.mkdir(parents=True, exist_ok=True)
            temporary_xls = destination.parent / f".{destination.stem}.{uuid4().hex}.tmp.xls"
            workbook = excel.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=True)
            workbook.SaveAs(str(temporary_xls), FileFormat=56)
            workbook.Close(SaveChanges=False)
            workbook = None
            os.replace(temporary_xls, destination)
            temporary_xls = None
    finally:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        if temporary_xls is not None:
            temporary_xls.unlink(missing_ok=True)


def prepare_correction_inputs(
    *,
    company_name: str,
    basic_info_path: Path,
    payroll_files: tuple[Path, ...],
    payroll_status_files: tuple[Path, ...],
    tool_template_path: Path,
    modified_input_directory: Path,
    correction_input_directory: Path,
    withholding_pdf_root: Path | None = None,
    xls_converter: Callable[[list[tuple[Path, Path]]], None] | None = None,
    tool_populator: Callable[[list[tuple[Path, Path]]], None] | None = None,
) -> CorrectionInputResult:
    template = Path(tool_template_path).resolve(strict=True)
    basic_info = Path(basic_info_path).resolve(strict=True)
    modified_root = Path(modified_input_directory).resolve()
    correction_root = Path(correction_input_directory).resolve()
    modified_root.mkdir(parents=True, exist_ok=True)
    correction_root.mkdir(parents=True, exist_ok=True)
    issues: list[dict[str, object]] = []
    modified_files: list[Path] = []
    tool_files: list[Path] = []
    monthly_files: list[Path] = []
    years: list[int] = []
    converter = xls_converter or _convert_xlsx_to_xls
    staging = correction_root / f".{uuid4().hex}_monthly_xls"
    staging.mkdir(parents=True, exist_ok=False)
    conversions: list[tuple[Path, Path]] = []
    tool_jobs: list[tuple[Path, Path]] = []
    try:
        payroll_by_year = {
            _year_from_payroll(Path(path)): Path(path)
            for path in payroll_files
        }
        for status_path in sorted(payroll_status_files, key=_year_from_status):
            source = Path(status_path).resolve(strict=True)
            year = _year_from_status(source)
            logger.info("%d년 급상여현황(수정) 생성 시작", year)
            years.append(year)
            modified = modified_root / safe_name(
                f"{company_name}_{str(year)[-2:]}년 급상여현황(수정).xlsx"
            )
            snapshot_file(source, modified)
            modified_files.append(modified)
            logger.info("%d년 급상여현황(수정) 생성 완료: %s", year, modified)

            tool = modified_root / safe_name(
                f"{company_name}_{str(year)[-2:]}년 수정 원천징수영수증 생성 Tool.xlsx"
            )
            snapshot_file(template, tool)
            _populate_withholding_tool_data(
                tool_path=tool,
                status_path=modified,
                basic_info_path=basic_info,
                year=year,
                issues=issues,
            )
            tool_files.append(tool)
            if withholding_pdf_root is not None:
                tool_jobs.append((tool, _sample_withholding_pdf(Path(withholding_pdf_root), year)))

            payroll = payroll_by_year.get(year)
            if payroll is None:
                raise SchemaMismatch(f"{year}년 연도별 급여대장이 없습니다.")
            payroll.resolve(strict=True)
            headers, monthly = _monthly_rows_from_status(modified, issues)
            for month in range(1, 13):
                temporary = staging / f"{year}_{month:02d}.xlsx"
                _write_monthly_workbook(temporary, headers, monthly.get(month, []))
                destination = correction_root / safe_name(
                    f"{company_name}_{str(year)[-2:]}년{month:02d}월_급여대장.xls"
                )
                conversions.append((temporary, destination))
                monthly_files.append(destination)
        if tool_jobs:
            (tool_populator or _populate_tools_with_pdf_images)(tool_jobs)
        converter(conversions)
    except Exception:
        for path in (*modified_files, *tool_files, *monthly_files):
            path.unlink(missing_ok=True)
        raise
    finally:
        shutil.rmtree(staging, ignore_errors=True)
    return CorrectionInputResult(
        tuple(years),
        tuple(modified_files),
        tuple(tool_files),
        tuple(monthly_files),
        tuple(issues),
    )
