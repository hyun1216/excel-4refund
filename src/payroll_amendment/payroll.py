from __future__ import annotations

import json
from pathlib import Path
import re

from openpyxl import load_workbook, Workbook

from .errors import SchemaMismatch
from .infra import safe_name
from .models import StageResult, TaxExemption


def _sheet_values(sheet) -> tuple[list[str], list[dict[str, object]]]:
    headers = [str(cell.value).strip() if cell.value is not None else "" for cell in sheet[1]]
    rows = [dict(zip(headers, values)) for values in sheet.iter_rows(min_row=2, values_only=True)]
    return headers, rows


def convert_monthly_payroll(*, company_id: str, company_name: str, year: int, payroll_files: list[Path], exemptions: list[TaxExemption], output_dir: Path) -> StageResult:
    candidates = []
    for path in payroll_files:
        if path.suffix.lower() not in {".xlsx", ".xlsm"}:
            continue
        workbook = load_workbook(path, read_only=True, data_only=True)
        if {"대상자(작업)", "비대상자"} <= set(workbook.sheetnames):
            candidates.append(path)
        workbook.close()
    if len(candidates) != 1:
        raise SchemaMismatch(f"{year}년 수정 급상여현황 후보가 {len(candidates)}개입니다.")
    source = candidates[0]
    workbook = load_workbook(source, read_only=True, data_only=True)
    rows: list[dict[str, object]] = []
    try:
        for name in ("대상자(작업)", "비대상자"):
            _, values = _sheet_values(workbook[name])
            rows.extend(values)
    finally:
        workbook.close()
    month_key = "지급년월"
    type_key = "급여구분"
    if not rows or month_key not in rows[0] or type_key not in rows[0]:
        raise SchemaMismatch("수정 급상여현황에 지급년월/급여구분 열이 없습니다.")
    grouped: dict[tuple[str, str], list[dict[str, object]]] = {}
    for row in rows:
        month, pay_type = str(row.get(month_key) or "").strip(), str(row.get(type_key) or "").strip()
        if not month or "합계" in month or not pay_type:
            continue
        grouped.setdefault((month, pay_type), []).append(row)
    output_dir.mkdir(parents=True, exist_ok=True)
    artifacts = []
    summary = []
    for (month, pay_type), items in sorted(grouped.items()):
        headers = [key for key in items[0] if key not in {month_key, type_key, "합계검증"}]
        book = Workbook(); sheet = book.active; sheet.title = safe_name(f"{month}_{pay_type}")[:31]
        sheet.append(headers)
        for item in items:
            sheet.append([item.get(header) for header in headers])
        target = output_dir / safe_name(f"{company_name}_{month}_{pay_type}대장.xlsx")
        book.save(target); book.close(); artifacts.append(target)
        summary.append({"month": month, "pay_type": pay_type, "rows": len(items), "file": target.name})
    summary_path = output_dir / "payroll_summary.json"
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    artifacts.append(summary_path)
    return StageResult("S20", company_id, year, "SUCCEEDED", artifacts, {"ledgers": len(summary), "rows": sum(item["rows"] for item in summary)})

