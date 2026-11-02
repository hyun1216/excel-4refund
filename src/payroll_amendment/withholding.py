from __future__ import annotations

import csv
from pathlib import Path
import re

from openpyxl import load_workbook

from .errors import PdfMatchError, SchemaMismatch
from .infra import safe_name
from .models import StageResult, TaxExemption
from .pdf_adjust import index_from_name, update_pdf


def _digits(value: object) -> str:
    return re.sub(r"\D", "", str(value or ""))


def _tool_rows(path: Path) -> dict[str, dict[str, object]]:
    book = load_workbook(path, data_only=True, read_only=True)
    try:
        if "Data" not in book.sheetnames:
            raise SchemaMismatch("Tool 파일에 Data 시트가 없습니다.")
        sheet = book["Data"]
        headers = [cell.value for cell in sheet[1]]
        if "인덱스" not in headers:
            raise SchemaMismatch("Data 시트에 인덱스 열이 없습니다.")
        result = {}
        for values in sheet.iter_rows(min_row=2, values_only=True):
            row = dict(zip(headers, values))
            if not row.get("주민번호") and row.get("주민등록번호"):
                row["주민번호"] = row["주민등록번호"]
            if row.get("인덱스") not in (None, ""):
                result[str(int(row["인덱스"]))] = row
        return result
    finally:
        book.close()


def generate_amended_withholding(*, company_id: str, year: int, exemptions: list[TaxExemption], tool_workbook: Path, pdf_files: list[Path], output_dir: Path) -> StageResult:
    rows = _tool_rows(tool_workbook)
    pdf_by_index = {str(index_from_name(path)): path for path in pdf_files if index_from_name(path) is not None}
    artifacts: list[Path] = []
    reconciliation = output_dir / "amendment_reconciliation.csv"
    output_dir.mkdir(parents=True, exist_ok=True)
    lines = []
    for exemption in exemptions:
        key = str(int(exemption.pdf_index))
        row = rows.get(key)
        source = pdf_by_index.get(key)
        if row is None:
            raise SchemaMismatch(f"Tool Data 시트에 인덱스 {key}가 없습니다.")
        if source is None:
            raise PdfMatchError(f"원천징수 PDF 인덱스 {key}가 없습니다.")
        if exemption.resident_id and _digits(row.get("주민번호")) != _digits(exemption.resident_id):
            raise SchemaMismatch(f"기초정보와 Tool의 주민번호가 다릅니다: 인덱스 {key}")
        serial = row.get("연번") or exemption.serial_number or key
        name = row.get("성명") or exemption.name
        period = row.get("기간") or f"{year}년"
        target = output_dir / safe_name(f"{int(serial)}_{name}_{period}_원천징수영수증(수정후).pdf")
        update_pdf(source, target, row)
        artifacts.append(target)
        lines.append({"인덱스": key, "사원번호": exemption.employee_id, "성명": name, "비과세반영액": str(sum(line.amount for line in exemption.lines)), "결과파일": target.name})
    with reconciliation.open("w", encoding="utf-8-sig", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=["인덱스", "사원번호", "성명", "비과세반영액", "결과파일"])
        writer.writeheader(); writer.writerows(lines)
    artifacts.append(reconciliation)
    return StageResult("S40", company_id, year, "SUCCEEDED", artifacts, {"employees": len(lines), "pdfs": len(lines)})
