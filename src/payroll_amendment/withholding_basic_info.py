from __future__ import annotations

from concurrent.futures import as_completed, ThreadPoolExecutor
from dataclasses import dataclass
import json
import logging
import os
from pathlib import Path
import re
from uuid import uuid4

from openpyxl import load_workbook
from pypdf import PdfReader

from .errors import PdfToExcelError, SchemaMismatch
from .infra import format_resident_id, sha256
from .withholding_pdf_excel import WithholdingPdfData, extract_withholding_pdf


logger = logging.getLogger(__name__)


WITHHOLDING_HEADERS = (
    "기간",
    "사원번호",
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
LEDGER_FIRST_PAGE_MARKER = "소득세법시행규칙[별지제25호서식(1)]"
_LEDGER_EMPLOYEE = re.compile(r"사원코드:([^()]*)\(성명:(.*?)주민등록번호:")


@dataclass(frozen=True)
class LedgerEmployee:
    pdf_index: int
    employee_id: str
    name: str
    source_page: int


@dataclass(frozen=True)
class BasicInfoExtractionResult:
    output_path: Path
    manifest_path: Path
    sheet_name: str
    employee_count: int
    taxfree_headers: tuple[str, ...]


def _compact(value: object) -> str:
    return re.sub(r"\s+", "", str(value or ""))


def extract_ledger_employees(path: Path) -> list[LedgerEmployee]:
    source = Path(path).resolve(strict=True)
    reader = PdfReader(str(source))
    employees: list[LedgerEmployee] = []
    for page_number, page in enumerate(reader.pages, 1):
        text = _compact(page.extract_text() or "")
        if LEDGER_FIRST_PAGE_MARKER not in text:
            continue
        match = _LEDGER_EMPLOYEE.search(text)
        if match is None:
            raise PdfToExcelError(
                f"소득자별 원천징수부에서 사원코드·성명을 찾지 못했습니다: {page_number}쪽"
            )
        employees.append(
            LedgerEmployee(
                pdf_index=len(employees) + 1,
                employee_id=match.group(1).strip(),
                name=match.group(2).strip(),
                source_page=page_number,
            )
        )
    if not employees:
        raise PdfToExcelError(
            f"소득자별 원천징수부 첫 페이지 표식을 찾지 못했습니다: {source.name}"
        )
    duplicate_ids = sorted(
        employee_id
        for employee_id in {item.employee_id for item in employees}
        if sum(item.employee_id == employee_id for item in employees) > 1
    )
    if duplicate_ids:
        raise PdfToExcelError(f"소득자별 원천징수부 사원코드가 중복됩니다: {duplicate_ids}")
    return employees


def _extract_all(pdf_files: list[Path], year: int) -> list[WithholdingPdfData]:
    sources = [Path(path).resolve(strict=True) for path in pdf_files]
    if not sources:
        raise PdfToExcelError("기초정보에 적재할 원천징수영수증 PDF가 없습니다.")
    rows: list[WithholdingPdfData] = []
    total = len(sources)
    logger.info("원천징수 값 추출 시작: %d년 (총 %d개)", year, total)
    try:
        with ThreadPoolExecutor(max_workers=min(4, total)) as executor:
            futures = {
                executor.submit(extract_withholding_pdf, path, year=year): path
                for path in sources
            }
            for future in as_completed(futures):
                rows.append(future.result())
    except Exception as exc:
        logger.error("원천징수 값 추출 오류: %d년, %s", year, exc)
        raise
    rows.sort(key=lambda item: item.pdf_index)
    indexes = [item.pdf_index for item in rows]
    if indexes != list(range(1, len(rows) + 1)):
        raise PdfToExcelError(f"원천징수영수증 PDF 인덱스가 연속되지 않습니다: {indexes}")
    logger.info("원천징수 값 추출 완료: %d년, %d개", year, len(rows))
    return rows


def _record(data: WithholdingPdfData, employee_id: str) -> dict[str, object]:
    values = dict(data.values)
    values["기간"] = data.values.get("기간")
    values["사원번호"] = employee_id
    values["성명"] = data.employee_name
    values["주민등록번호"] = format_resident_id(data.resident_id)
    values["인덱스"] = data.pdf_index
    return values


def _taxfree_headers(records: list[dict[str, object]]) -> tuple[str, ...]:
    fixed = set(WITHHOLDING_HEADERS) | {"사번"}
    found: list[str] = []
    for record in records:
        for key in record:
            if key not in fixed and key not in found:
                found.append(key)
    return tuple(found)


def _write_with_excel(
    *,
    source: Path,
    destination: Path,
    sheet_name: str,
    headers: tuple[str, ...],
    records: list[dict[str, object]],
    taxfree_headers: tuple[str, ...],
    year: int,
    company_name: str | None,
    company_id: str | None,
) -> Path:
    try:
        import win32com.client
    except ImportError as exc:
        raise SchemaMismatch("Excel ?? ??? ?? pywin32? ???? ?? ????.") from exc

    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.parent / f".{destination.stem}.{uuid4().hex}.tmp.xlsx"
    import shutil

    shutil.copy2(source, temporary)
    excel = None
    workbook = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        workbook = excel.Workbooks.Open(str(temporary.resolve()))
        business = workbook.Worksheets("\uc0ac\uc5c5\uc7a5\uc815\ubcf4")
        business.Range("K2").Value = str(year)[-2:]
        business.Range("N2").Value = "\uc5d1\uc140"
        if company_name:
            business.Range("B5").Value = company_name
        if company_id:
            business.Range("B6").Value = company_id

        for worksheet in workbook.Worksheets:
            if worksheet.Name == sheet_name:
                worksheet.Delete()
                break
        sheet = workbook.Worksheets.Add(After=workbook.Worksheets(workbook.Worksheets.Count))
        sheet.Name = sheet_name
        header_values = (tuple(headers),)
        sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, len(headers))).Value = header_values
        sheet.Columns(2).NumberFormat = "@"
        sheet.Columns(4).NumberFormat = "@"
        if records:
            body = tuple(tuple(record.get(header) for header in headers) for record in records)
            sheet.Range(sheet.Cells(2, 1), sheet.Cells(len(records) + 1, len(headers))).Value = body
        used = sheet.Range(sheet.Cells(1, 1), sheet.Cells(len(records) + 1, len(headers)))
        used.Borders.LineStyle = 1
        used.Font.Name = "\ub9d1\uc740 \uace0\ub515"
        used.Font.Size = 9
        header_range = sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, len(headers)))
        header_range.Font.Bold = True
        header_range.Font.Color = 0xFFFFFF
        header_range.Interior.Color = 0xD59B5B
        header_range.HorizontalAlignment = -4108
        if len(headers) >= 6:
            sheet.Range(sheet.Cells(2, 6), sheet.Cells(len(records) + 1, len(headers))).NumberFormat = "#,##0"
        used.Columns.AutoFit()
        for column in range(1, len(headers) + 1):
            width = sheet.Columns(column).ColumnWidth
            sheet.Columns(column).ColumnWidth = min(max(width, 11), 34)
        header_range.AutoFilter()
        sheet.Activate()
        excel.ActiveWindow.SplitRow = 1
        excel.ActiveWindow.FreezePanes = True
        excel.ActiveWindow.DisplayGridlines = False

        allowance = workbook.Worksheets("\uc218\ub2f9\ub4f1\ub85d")
        last_row = max(allowance.UsedRange.Rows.Count, 4 + len(taxfree_headers))
        allowance.Range(allowance.Cells(5, 1), allowance.Cells(last_row, 4)).ClearContents()
        if taxfree_headers:
            values = tuple((header, "\uae30\uc874\ube44\uacfc\uc138") for header in taxfree_headers)
            allowance.Range(allowance.Cells(5, 1), allowance.Cells(4 + len(values), 2)).Value = values

        workbook.Save()
        workbook.Close(SaveChanges=True)
        workbook = None
        excel.Quit()
        excel = None
        return temporary
    except Exception:
        if workbook is not None:
            workbook.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        temporary.unlink(missing_ok=True)
        raise


def build_basic_info_from_withholding_pdfs(
    *,
    basic_info_path: Path,
    pdf_files: list[Path],
    year: int,
    output_path: Path,
    ledger_pdf: Path | None = None,
    company_name: str | None = None,
    company_id: str | None = None,
) -> BasicInfoExtractionResult:
    source = Path(basic_info_path).resolve(strict=True)
    destination = Path(output_path).resolve()
    data_rows = _extract_all(pdf_files, year)
    ledger = extract_ledger_employees(ledger_pdf) if ledger_pdf is not None else []
    if ledger and len(ledger) != len(data_rows):
        raise PdfToExcelError(
            f"영수증과 소득자별 원천징수부의 인원수가 다릅니다: 영수증={len(data_rows)}, 원천징수부={len(ledger)}"
        )
    records: list[dict[str, object]] = []
    employee_sources: list[dict[str, object]] = []
    for position, data in enumerate(data_rows):
        employee_id = data.resident_id
        if ledger:
            employee = ledger[position]
            if employee.pdf_index != data.pdf_index or _compact(employee.name) != _compact(data.employee_name):
                raise PdfToExcelError(
                    f"영수증과 소득자별 원천징수부의 사원 순서가 다릅니다: 인덱스 {data.pdf_index}"
                )
            employee_id = employee.employee_id
            employee_sources.append(
                {"pdf_index": data.pdf_index, "employee_id": employee_id, "name": data.employee_name, "ledger_page": employee.source_page}
            )
        records.append(_record(data, employee_id))
    dynamic_headers = _taxfree_headers(records)
    headers = WITHHOLDING_HEADERS + dynamic_headers
    sheet_name = f"원천징수추출({str(year)[-2:]}년)"

    temporary = _write_with_excel(
        source=source,
        destination=destination,
        sheet_name=sheet_name,
        headers=headers,
        records=records,
        taxfree_headers=dynamic_headers,
        year=year,
        company_name=company_name,
        company_id=company_id,
    )
    try:
        checked = load_workbook(temporary, read_only=True, data_only=False)
        try:
            if sheet_name not in checked.sheetnames:
                raise SchemaMismatch(f"\uc0dd\uc131\ub41c \uae30\ucd08\uc815\ubcf4\uc5d0 \uc2dc\ud2b8\uac00 \uc5c6\uc2b5\ub2c8\ub2e4: {sheet_name}")
            sheet = checked[sheet_name]
            if sheet.max_row != len(records) + 1:
                raise SchemaMismatch(
                    f"\uc6d0\ucc9c\uc9d5\uc218 \ucd94\ucd9c \ud589 \uc218 \uac80\uc99d \uc2e4\ud328: \uc608\uc0c1={len(records)}, \uc2e4\uc81c={sheet.max_row - 1}"
                )
            checked_names = [sheet.cell(row, 3).value for row in range(2, sheet.max_row + 1)]
            expected_names = [record["\uc131\uba85"] for record in records]
            if checked_names != expected_names:
                raise SchemaMismatch("\uc6d0\ucc9c\uc9d5\uc218 \ucd94\ucd9c \uc131\uba85 \uc21c\uc11c \uac80\uc99d\uc5d0 \uc2e4\ud328\ud588\uc2b5\ub2c8\ub2e4.")
        finally:
            checked.close()
        os.replace(temporary, destination)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise

    manifest_path = destination.with_suffix(".manifest.json")
    manifest = {
        "year": year,
        "sheet_name": sheet_name,
        "employee_count": len(records),
        "taxfree_headers": list(dynamic_headers),
        "basic_info_source": source.name,
        "basic_info_output": destination.name,
        "ledger_source": Path(ledger_pdf).name if ledger_pdf else None,
        "employees": employee_sources,
        "pdfs": [
            {"pdf_index": data.pdf_index, "name": data.employee_name, "resident_id": data.resident_id, "sha256": sha256(path)}
            for data, path in zip(data_rows, sorted((Path(path).resolve() for path in pdf_files), key=lambda item: int(item.name.split("_", 1)[0])))
        ],
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return BasicInfoExtractionResult(destination, manifest_path, sheet_name, len(records), dynamic_headers)

