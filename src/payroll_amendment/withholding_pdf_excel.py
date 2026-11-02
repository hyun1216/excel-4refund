from __future__ import annotations



from dataclasses import dataclass

import json

import os

from pathlib import Path

import re

import shutil

from typing import Any

from uuid import uuid4



from openpyxl import Workbook, load_workbook

from openpyxl.styles import Alignment, Font, PatternFill

import pdfplumber



from .errors import PdfToExcelError

from .infra import safe_name, sha256

from .models import TaxExemption





Table = list[list[str | None]]

PageTables = list[Table]



FORM_TEMPLATE = (

    Path(__file__).resolve().parents[2]

    / "Resources"

    / "원천징수영수증_공용양식_자동연동.xlsx"

)



FIXED_FIELDS = (

    "기간",

    "사번",

    "성명",

    "주민등록번호",

    "인덱스",

    "과세소득_총계",

    "비과세소득_총계",

    "과세소득_현계",

    "비과세소득_현계",

    "결정세액_소득세",

    "기납부세액_소득세",

    "차감징수세액_소득세",

    "결정세액_지방세",

    "기납부세액_지방세",

    "차감징수세액_지방세",

    "근로소득금액",

    "차감소득금액",

    "그밖의소득공제계",

    "과세표준",

    "세액감면 계",

    "근로소득",

    "세액공제 계",

)



_RESIDENT_NUMBER = re.compile(r"(?<!\d)(\d{6})\s*-?\s*(\d{7})(?!\d)")

_INTEGER = re.compile(r"-?[\d,]+")





@dataclass(frozen=True)

class TaxfreeIncome:

    item_name: str

    code: str

    current_amount: int

    total_amount: int





@dataclass(frozen=True)

class WithholdingPdfData:

    employee_name: str

    resident_id: str

    pdf_index: int

    values: dict[str, str | int]

    taxfree_incomes: tuple[TaxfreeIncome, ...]

    page_tables: tuple[tuple[Table, ...], ...]

    page_count: int





@dataclass(frozen=True)

class PreparedWithholdingExcelBatch:

    output_files: list[Path]

    manifest_path: Path





def _compact(value: object) -> str:

    return re.sub(r"\s+", "", str(value or ""))





def _excel_value(value: object) -> object:

    if isinstance(value, str) and value.startswith(("=", "+", "-", "@")):

        return f"'{value}"

    return value





def _money(value: object) -> int | None:

    text = _compact(value).replace("(", "").replace(")", "").removesuffix("원")

    if not text or not _INTEGER.fullmatch(text):

        return None

    return int(text.replace(",", ""))





def _find_cell(table: Table, keyword: str, *, start_row: int = 0) -> tuple[int, int]:

    target = _compact(keyword)

    for row_index, row in enumerate(table[start_row:], start_row):

        for column_index, value in enumerate(row):

            if target in _compact(value):

                return row_index, column_index

    raise PdfToExcelError(f"원천징수 필수 항목을 찾지 못했습니다: {keyword}")





def _find_exact_cell(table: Table, keyword: str, *, start_row: int = 0) -> tuple[int, int]:

    target = _compact(keyword)

    for row_index, row in enumerate(table[start_row:], start_row):

        for column_index, value in enumerate(row):

            compacted = _compact(value)

            if target == compacted or re.fullmatch(rf"\d+(?:-\d+)?\.{re.escape(target)}", compacted):

                return row_index, column_index

    raise PdfToExcelError(f"원천징수 필수 항목을 찾지 못했습니다: {keyword}")





def _amount_at(row: list[str | None], column: int) -> int:

    if column >= len(row):

        return 0

    return _money(row[column]) or 0





def _amount_after(row: list[str | None], column: int, stop: int | None = None) -> int:

    for value in row[column + 1 : stop]:

        amount = _money(value)

        if amount is not None:

            return amount

    return 0





def _largest_table(page: PageTables, keyword: str) -> Table:
    candidates = [table for table in page if any(keyword in _compact(cell) for row in table for cell in row)]
    if not candidates and keyword == "근로소득원천징수영수증":
        candidates = [
            table
            for table in page
            if any("소득자" in _compact(cell) for row in table for cell in row)
            and any("주(현)" == _compact(cell) for row in table for cell in row)
        ]
    if not candidates:
        raise PdfToExcelError(f"원천징수 표를 찾지 못했습니다: {keyword}")
    return max(candidates, key=len)

def _pdf_index(path: Path) -> int:

    match = re.search(r"_(\d+)\.pdf$", path.name, re.IGNORECASE)

    if match is None:

        match = re.match(r"(\d+)_", path.name)

    if not match:

        raise PdfToExcelError(f"PDF 파일명에서 인덱스를 찾지 못했습니다: {path.name}")

    return int(match.group(1))





def _identity(table: Table) -> tuple[str, str]:
    for row in table:
        compacted = [_compact(value) for value in row]
        if not any("소득자" in value for value in compacted):
            continue
        name_cell = next((value for value in compacted if "성명" in value), "")
        resident_cell = next((value for value in compacted if "주민(외국인)등록번호" in value), "")
        match = _RESIDENT_NUMBER.search(resident_cell)
        name = name_cell.split("성명", 1)[-1]
        if name and match:
            return name, "".join(match.groups())

        name_label = next((index for index, value in enumerate(compacted) if "성명" in value), -1)
        resident_label = next(
            (index for index, value in enumerate(compacted) if "주민" in value and "등록번호" in value),
            -1,
        )
        if name_label < 0 or resident_label < 0:
            continue
        name = next((value for value in compacted[name_label + 1 : resident_label] if value), "")
        resident_text = "".join(value for value in compacted[resident_label + 1 :] if value)
        match = _RESIDENT_NUMBER.search(resident_text)
        if name and match:
            return name, "".join(match.groups())
    raise PdfToExcelError("소득자의 성명 또는 주민등록번호를 찾지 못했습니다.")

def _income_columns(table: Table) -> tuple[int, int, int]:

    header_row, current_column = _find_exact_cell(table, "주(현)")

    same_row = table[header_row]

    total_column = next(

        (index for index, value in enumerate(same_row) if _compact(value) == "합계"),

        -1,

    )

    if total_column < 0:

        raise PdfToExcelError("근무처별 소득명세의 합계 열을 찾지 못했습니다.")

    return header_row, current_column, total_column





def _taxfree_incomes(

    table: Table,

    *,

    income_total_row: int,

    taxfree_total_row: int,

    current_column: int,

    total_column: int,

) -> tuple[TaxfreeIncome, ...]:

    result: list[TaxfreeIncome] = []

    for row in table[income_total_row + 1 : taxfree_total_row]:

        current_amount = _amount_at(row, current_column)

        total_amount = _amount_at(row, total_column)

        if current_amount == 0 and total_amount == 0:

            continue

        labels = [_compact(value) for value in row[:current_column]]

        item_name = next((value for value in labels[1:] if value and _money(value) is None), "")

        if not item_name:

            continue

        item_name = re.sub(r"^-?\d+(?:-\d+)?[.]?", "", item_name)

        code = next((value for value in reversed(labels) if re.fullmatch(r"[A-Z]\d{2}", value)), "")

        result.append(TaxfreeIncome(item_name, code, current_amount, total_amount))

    return tuple(result)





def _page_one_values(table: Table) -> tuple[str, str, dict[str, int], tuple[TaxfreeIncome, ...]]:

    employee_name, resident_id = _identity(table)

    header_row, current_column, total_column = _income_columns(table)

    taxfree_total_row, _ = _find_cell(table, "비과세소득계", start_row=header_row)

    income_total_row = next(

        (

            row_index

            for row_index in range(header_row + 1, taxfree_total_row)

            if any(
                re.fullmatch(r"(?:16\.)?계", _compact(value))
                for value in table[row_index]
            )

        ),

        -1,

    )

    if income_total_row < 0:

        raise PdfToExcelError("16. 계 과세소득 행을 찾지 못했습니다.")



    tax_header_row, income_tax_column = _find_cell(table, "소득세", start_row=taxfree_total_row)

    _, local_tax_column = _find_cell(table, "지방소득세", start_row=tax_header_row)

    determined_row, _ = _find_cell(table, "결정세액", start_row=tax_header_row)

    paid_row, _ = _find_cell(table, "기납부세액", start_row=determined_row)

    deducted_row, _ = _find_cell(table, "차감징수세액", start_row=paid_row)



    values = {

        "과세소득_총계": _amount_at(table[income_total_row], total_column),

        "비과세소득_총계": _amount_at(table[taxfree_total_row], total_column),

        "과세소득_현계": _amount_at(table[income_total_row], current_column),

        "비과세소득_현계": _amount_at(table[taxfree_total_row], current_column),

        "결정세액_소득세": _amount_at(table[determined_row], income_tax_column),

        "결정세액_지방세": _amount_at(table[determined_row], local_tax_column),

        "기납부세액_소득세": sum(

            _amount_at(row, income_tax_column) for row in table[paid_row + 1 : deducted_row]

        ),

        "기납부세액_지방세": sum(

            _amount_at(row, local_tax_column) for row in table[paid_row + 1 : deducted_row]

        ),

        "차감징수세액_소득세": _amount_at(table[deducted_row], income_tax_column),

        "차감징수세액_지방세": _amount_at(table[deducted_row], local_tax_column),

    }

    details = _taxfree_incomes(

        table,

        income_total_row=income_total_row,

        taxfree_total_row=taxfree_total_row,

        current_column=current_column,

        total_column=total_column,

    )

    return employee_name, resident_id, values, details





def _page_two_values(table: Table) -> dict[str, int]:

    _, right_start = _find_cell(table, "종합소득과세표준")



    def left_value(keyword: str) -> int:

        row_index, column = _find_cell(table, keyword)

        return _amount_after(table[row_index], column, right_start)



    def right_value(keyword: str, *, exact: bool = False) -> int:

        finder = _find_exact_cell if exact else _find_cell

        row_index, column = finder(table, keyword)

        if column < right_start:

            for candidate in range(column + 1, len(table[row_index])):

                if keyword in _compact(table[row_index][candidate]):

                    column = candidate

                    break

        return _amount_after(table[row_index], column)



    return {

        "근로소득금액": left_value("근로소득금액"),

        "차감소득금액": left_value("차감소득금액"),

        "그밖의소득공제계": left_value("그밖의소득공제계"),

        "과세표준": right_value("종합소득과세표준"),

        "세액감면 계": right_value("세액감면계"),

        "근로소득": right_value("근로소득", exact=True),

        "세액공제 계": right_value("세액공제계"),

    }





def parse_withholding_tables(

    *,

    path: Path,

    year: int,

    page_tables: list[PageTables],

    page_count: int,

    exemption: TaxExemption | None = None,

) -> WithholdingPdfData:

    if len(page_tables) < 2:

        raise PdfToExcelError(f"원천징수 PDF가 2쪽 미만입니다: {path.name}")

    page_one = _largest_table(page_tables[0], "근로소득원천징수영수증")

    page_two = _largest_table(page_tables[1], "정산명세")

    employee_name, resident_id, values, taxfree_incomes = _page_one_values(page_one)

    values.update(_page_two_values(page_two))

    index = _pdf_index(path)



    employee_id = ""

    if exemption is not None:

        expected_resident = re.sub(r"\D", "", exemption.resident_id)

        if expected_resident and resident_id != expected_resident:

            raise PdfToExcelError(f"PDF와 기초정보의 주민번호가 다릅니다: 인덱스 {index}")

        if exemption.name and _compact(exemption.name) != _compact(employee_name):

            raise PdfToExcelError(f"PDF와 기초정보의 성명이 다릅니다: 인덱스 {index}")

        employee_id = exemption.employee_id



    record: dict[str, str | int] = {

        "기간": year,

        "사번": employee_id,

        "성명": employee_name,

        "주민등록번호": resident_id,

        "인덱스": index,

        **values,

    }

    for detail in taxfree_incomes:

        key = detail.item_name if not detail.code else f"{detail.item_name}({detail.code})"

        record[f"{key}_총계"] = detail.total_amount

        record[f"{key}_현계"] = detail.current_amount



    return WithholdingPdfData(

        employee_name,

        resident_id,

        index,

        record,

        taxfree_incomes,

        tuple(tuple(tables) for tables in page_tables),

        page_count,

    )





def extract_withholding_pdf(

    path: Path,

    *,

    year: int,

    exemption: TaxExemption | None = None,

) -> WithholdingPdfData:

    source = Path(path).resolve(strict=True)

    try:

        with pdfplumber.open(source) as document:

            page_tables = [page.extract_tables() for page in document.pages]

            return parse_withholding_tables(

                path=source,

                year=year,

                page_tables=page_tables,

                page_count=len(document.pages),

                exemption=exemption,

            )

    except PdfToExcelError:

        raise

    except Exception as exc:

        raise PdfToExcelError(f"원천징수 PDF 표를 읽을 수 없습니다: {source.name}: {exc}") from exc





def _write_template_inputs(sheet, data: WithholdingPdfData) -> None:

    labels = {

        _compact(sheet.cell(row, 3).value): row

        for row in range(2, sheet.max_row + 1)

        if sheet.cell(row, 3).value not in (None, "")

    }

    resident = data.resident_id

    if len(resident) == 13:

        resident = f"{resident[:6]}-{resident[6:]}"

    sheet["D8"] = _excel_value(data.employee_name)

    sheet["D9"] = _excel_value(resident)

    sheet["D11"] = data.values.get("기간")

    for key, value in data.values.items():

        normalized = _compact(key)

        row = next(

            (candidate for label, candidate in labels.items() if label.startswith(normalized)),

            None,

        )

        if row is not None:

            sheet.cell(row, 4, _excel_value(value))

    for detail in data.taxfree_incomes:

        if detail.code == "P01":

            sheet["D27"] = detail.current_amount

            sheet["D29"] = detail.total_amount

            continue

        sheet["D30"] = _excel_value(detail.item_name)

        sheet["D31"] = _excel_value(detail.code)

        sheet["D32"] = detail.current_amount

        sheet["D33"] = detail.total_amount

        break



def _write_workbook(data: WithholdingPdfData, source: Path, target: Path) -> None:

    if not FORM_TEMPLATE.is_file():

        raise PdfToExcelError(f"원천징수 Excel 양식 템플릿이 없습니다: {FORM_TEMPLATE}")

    workbook = load_workbook(FORM_TEMPLATE)

    if {"Table 1", "데이터입력"} - set(workbook.sheetnames):

        raise PdfToExcelError("공용양식에 Table 1 또는 데이터입력 시트가 없습니다.")

    _write_template_inputs(workbook["데이터입력"], data)

    workbook.calculation.calcMode = "auto"

    workbook.calculation.fullCalcOnLoad = True

    workbook.calculation.forceFullCalc = True



    target.parent.mkdir(parents=True, exist_ok=True)

    workbook.save(target)

    check = load_workbook(target, read_only=True, data_only=False)

    try:

        if {"Table 1", "데이터입력"} - set(check.sheetnames):

            raise PdfToExcelError(f"생성 Excel 시트 검증에 실패했습니다: {target.name}")

        if check["데이터입력"]["D8"].value != data.employee_name:

            raise PdfToExcelError(f"생성 Excel 성명 검증에 실패했습니다: {target.name}")

        if check["Table 1"]["D15"].data_type != "f":

            raise PdfToExcelError(f"생성 Excel 자동연동 수식 검증에 실패했습니다: {target.name}")

    finally:

        check.close()





def _write_form_sheet(workbook: Workbook, source: Path, data: WithholdingPdfData) -> None:

    sheet = workbook["Table 1"]

    sheet.sheet_view.showGridLines = False

    sheet.sheet_view.zoomScale = 80

    sheet._images = []



    for row in sheet.iter_rows():

        for cell in row:

            if not isinstance(cell, MergedCell):

                cell.value = None



    try:

        with pymupdf.open(source) as document:

            for page_number, page in enumerate(document):

                _write_pdf_page_to_cells(sheet, page, page_number)

    except Exception as exc:

        raise PdfToExcelError(f"PDF 양식을 Excel 셀로 변환할 수 없습니다: {source.name}: {exc}") from exc



    _write_vba_compatibility_cells(sheet, data)

    last_row = TEMPLATE_PAGE_ROWS[-1][1]

    if data.page_count > len(TEMPLATE_PAGE_ROWS):

        last_row += (data.page_count - len(TEMPLATE_PAGE_ROWS)) * 64

    sheet.print_area = f"A1:AX{last_row}"

    sheet.sheet_properties.pageSetUpPr.fitToPage = True

    sheet.page_setup.paperSize = sheet.PAPERSIZE_A4

    sheet.page_setup.orientation = sheet.ORIENTATION_PORTRAIT

    sheet.page_setup.fitToWidth = 1

    sheet.page_setup.fitToHeight = 0

    sheet.page_margins.left = 0.1

    sheet.page_margins.right = 0.1

    sheet.page_margins.top = 0.1

    sheet.page_margins.bottom = 0.1

    sheet.page_margins.header = 0

    sheet.page_margins.footer = 0

    sheet.row_breaks.brk = []

    for page_start, _ in TEMPLATE_PAGE_ROWS[1:data.page_count]:

        sheet.row_breaks.append(Break(id=page_start - 1))

    for page_number in range(len(TEMPLATE_PAGE_ROWS), data.page_count):

        sheet.row_breaks.append(Break(id=166 + (page_number - len(TEMPLATE_PAGE_ROWS)) * 64))





def _write_pdf_page_to_cells(sheet, page: pymupdf.Page, page_number: int) -> None:

    start_row, end_row = _page_rows(page_number)

    column_edges = _normalized_edges(

        [sheet.column_dimensions[get_column_letter(column)].width or 13.0 for column in range(1, 51)]

    )

    row_edges = _normalized_edges(

        [sheet.row_dimensions[row].height or 15.0 for row in range(start_row, end_row + 1)]

    )

    merged_anchors = _merged_anchors(sheet)



    page_dict = page.get_text("dict")

    for block in page_dict.get("blocks", []):

        if block.get("type") != 0:

            continue

        for line in block.get("lines", []):

            for span in line.get("spans", []):

                text = str(span.get("text", "")).strip()

                if not text:

                    continue

                x0, y0, x1, y1 = span["bbox"]

                column = min(50, bisect_left(column_edges, ((x0 + x1) / 2) / page.rect.width))

                row_offset = min(end_row - start_row + 1, bisect_left(row_edges, ((y0 + y1) / 2) / page.rect.height))

                cell = sheet.cell(start_row + max(row_offset, 1) - 1, max(column, 1))

                anchor = merged_anchors.get(cell.coordinate, cell.coordinate)

                _append_cell_value(sheet[anchor], text)





def _page_rows(page_number: int) -> tuple[int, int]:

    if page_number < len(TEMPLATE_PAGE_ROWS):

        return TEMPLATE_PAGE_ROWS[page_number]

    start = 167 + (page_number - len(TEMPLATE_PAGE_ROWS)) * 64

    return start, start + 63





def _normalized_edges(sizes: list[float]) -> list[float]:

    total = sum(sizes)

    running = 0.0

    edges: list[float] = []

    for size in sizes:

        running += size

        edges.append(running / total)

    return edges





def _merged_anchors(sheet) -> dict[str, str]:

    anchors: dict[str, str] = {}

    for merged in sheet.merged_cells.ranges:

        anchor = sheet.cell(merged.min_row, merged.min_col).coordinate

        for row in range(merged.min_row, merged.max_row + 1):

            for column in range(merged.min_col, merged.max_col + 1):

                anchors[sheet.cell(row, column).coordinate] = anchor

    return anchors





def _append_cell_value(cell, text: str) -> None:

    value: object = text

    compact = text.replace(",", "")

    if re.fullmatch(r"-?\d+", compact):

        value = int(compact)

    value = _excel_value(value)

    if cell.value in (None, ""):

        cell.value = value

        return

    if str(value) not in str(cell.value):

        cell.value = f"{cell.value} {value}"





def _write_vba_compatibility_cells(sheet, data: WithholdingPdfData) -> None:

    def put(cell: str, value: object, merge: str | None = None) -> None:

        target = sheet[cell]

        target.value = _excel_value(value)



    resident = data.resident_id

    if len(resident) == 13:

        resident = f"{resident[:6]}-{resident[6:]}"



    put("D15", f"⑥성    명        {data.employee_name}", "D15:AC15")

    put("AD15", f"⑦주민등록번호  {resident}", "AD15:AU15")

    put("B17", "구분", "B17:N17")

    put("O17", "주(현)", "O17:T17")

    put("AO17", "합계", "AO17:AU17")

    put("B29", "16. 계", "B29:N29")

    put("O29", data.values["과세소득_현계"], "O29:T29")

    put("AO29", data.values["과세소득_총계"], "AO29:AU29")



    for row, detail in enumerate(data.taxfree_incomes[:19], 30):

        put(f"B{row}", detail.item_name, f"B{row}:K{row}")

        put(f"L{row}", detail.code, f"L{row}:N{row}")

        put(f"O{row}", detail.current_amount, f"O{row}:T{row}")

        put(f"AO{row}", detail.total_amount, f"AO{row}:AU{row}")



    put("B49", "20. 비과세소득계", "B49:N49")

    put("O49", data.values["비과세소득_현계"], "O49:T49")

    put("AO49", data.values["비과세소득_총계"], "AO49:AU49")



    put("U51", "79. 소득세", "U51:AC51")

    put("AD51", "80. 지방소득세", "AD51:AK51")

    put("AL51", "81. 농어촌특별세", "AL51:AU51")

    put("B52", "결정세액", "B52:T52")

    put("U52", data.values["결정세액_소득세"], "U52:AC52")

    put("AD52", data.values["결정세액_지방세"], "AD52:AK52")

    put("B53", "기납부세액", "B53:T55")

    put("E56", "주(현) 근무지", "E56:T56")

    put("U56", data.values["기납부세액_소득세"], "U56:AC56")

    put("AD56", data.values["기납부세액_지방세"], "AD56:AK56")

    put("B58", "차감징수세액", "B58:T58")

    put("U58", data.values["차감징수세액_소득세"], "U58:AC58")

    put("AD58", data.values["차감징수세액_지방세"], "AD58:AK58")



    put("B65", "근로소득금액")

    put("U65", data.values["근로소득금액"])

    put("B106", "차감소득금액")

    put("U106", data.values["차감소득금액"])

    put("C122", "그밖의소득공제계")

    put("U122", data.values["그밖의소득공제계"])

    put("AB62", "종합소득과세표준")

    put("AS62", data.values["과세표준"])

    put("AD70", "세액감면계")

    put("AS70", data.values["세액감면 계"])

    put("AD71", "근로소득")

    put("AS71", data.values["근로소득"])

    put("AD121", "세액공제계")

    put("AS121", data.values["세액공제 계"])





def _quarantine(source: Path, quarantine_dir: Path, error: Exception) -> None:

    quarantine_dir.mkdir(parents=True, exist_ok=True)

    copied = quarantine_dir / safe_name(source.name)

    shutil.copy2(source, copied)

    report = {

        "status": "QUARANTINED",

        "error_code": getattr(error, "code", "E_PDF_TO_EXCEL"),

        "message": str(error),

        "file": copied.name,

    }

    (quarantine_dir / "failure.json").write_text(

        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"

    )





def _exemptions_by_index(exemptions: list[TaxExemption]) -> dict[int, TaxExemption]:

    result: dict[int, TaxExemption] = {}

    for exemption in exemptions:

        try:

            index = int(exemption.pdf_index)

        except (TypeError, ValueError) as exc:

            raise PdfToExcelError(

                f"기초정보의 PDF 인덱스가 숫자가 아닙니다: {exemption.pdf_index}"

            ) from exc

        if index in result:

            raise PdfToExcelError(f"기초정보의 PDF 인덱스가 중복되었습니다: {index}")

        result[index] = exemption

    return result





def convert_withholding_pdfs_to_excel(

    *,

    pdf_files: list[Path],

    exemptions: list[TaxExemption],

    year: int,

    output_dir: Path,

    quarantine_dir: Path,

) -> PreparedWithholdingExcelBatch:

    by_index = _exemptions_by_index(exemptions)

    staging = output_dir.parent / f".{output_dir.name}.{uuid4().hex}.tmp"

    records: list[dict[str, Any]] = []

    outputs: list[Path] = []

    current_source: Path | None = None

    try:

        staging.mkdir(parents=True, exist_ok=False)

        for source in sorted((Path(path).resolve(strict=True) for path in pdf_files), key=lambda path: path.name):

            current_source = source

            index = _pdf_index(source)

            exemption = by_index.get(index)

            if exemption is None:

                raise PdfToExcelError(f"기초정보에 없는 PDF 인덱스입니다: {index}")

            data = extract_withholding_pdf(source, year=year, exemption=exemption)

            target = staging / f"{source.stem}.xlsx"

            if target.exists():

                raise PdfToExcelError(f"변환 Excel 파일명이 중복됩니다: {target.name}")

            _write_workbook(data, source, target)

            outputs.append(target)

            records.append(

                {

                    "pdf_index": index,

                    "source": source.name,

                    "source_sha256": sha256(source),

                    "output": target.name,

                    "output_sha256": sha256(target),

                    "page_count": data.page_count,

                    "field_count": len(data.values),

                    "taxfree_item_count": len(data.taxfree_incomes),

                }

            )



        expected = set(by_index)

        produced = {record["pdf_index"] for record in records}

        if expected != produced:

            missing = sorted(expected - produced)

            raise PdfToExcelError(f"원천징수 Excel 변환 대상이 누락됐습니다: 인덱스 {missing}")



        manifest = staging / "pdf_to_excel_manifest.json"

        manifest.write_text(

            json.dumps({"year": year, "outputs": records}, ensure_ascii=False, indent=2),

            encoding="utf-8",

        )

        if output_dir.exists():

            raise PdfToExcelError(f"원천징수 Excel 출력 폴더가 이미 존재합니다: {output_dir}")

        os.replace(staging, output_dir)

        return PreparedWithholdingExcelBatch(

            [output_dir / path.name for path in outputs],

            output_dir / manifest.name,

        )

    except Exception as exc:

        shutil.rmtree(staging, ignore_errors=True)

        if current_source is not None and current_source.is_file():

            _quarantine(current_source, quarantine_dir, exc)

        if isinstance(exc, PdfToExcelError):

            raise

        raise PdfToExcelError(str(exc)) from exc

