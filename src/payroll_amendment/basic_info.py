from __future__ import annotations

from collections import defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
import re

from openpyxl import load_workbook

from .errors import DuplicateEmployee, DuplicatePdfIndex, SchemaMismatch
from .infra import safe_name
from .models import CompanyInfo, TaxExemption, TaxExemptionLine


REQUIRED_SHEETS = {"사업장정보", "수당등록", "수당조정(비과세)"}


def _text(value: object) -> str:
    return str(value).strip() if value not in (None, "") else ""


def _money(value: object, *, cell: str) -> Decimal:
    if value in (None, ""):
        return Decimal("0")
    try:
        return Decimal(str(value).replace(",", ""))
    except InvalidOperation as exc:
        raise SchemaMismatch(f"금액 형식 오류: {cell}") from exc


def _year(value: object, *, cell: str) -> int:
    match = re.fullmatch(r"\s*(20\d{2})년?\s*", str(value or ""))
    if not match:
        raise SchemaMismatch(f"귀속연도 형식 오류: {cell}")
    return int(match.group(1))


def parse_basic_info(path: Path, requested_years: set[int], company_name: str | None = None) -> tuple[CompanyInfo, dict[int, list[TaxExemption]]]:
    workbook = load_workbook(path, data_only=True, read_only=False)
    try:
        missing = REQUIRED_SHEETS - set(workbook.sheetnames)
        if missing:
            raise SchemaMismatch(f"필수 시트 누락: {', '.join(sorted(missing))}")
        info_sheet = workbook["사업장정보"]
        fields: dict[str, str] = {}
        blank_run = 0
        for row in range(5, min(info_sheet.max_row, 5000) + 1):
            key, value = _text(info_sheet.cell(row, 1).value), _text(info_sheet.cell(row, 2).value)
            if not key:
                blank_run += 1
                if blank_run >= 100:
                    break
                continue
            blank_run = 0
            if key in fields and fields[key] != value:
                raise SchemaMismatch(f"사업장정보 중복 키: A{row}")
            fields[key] = value
        display_name = company_name or next((fields[k] for k in fields if "사업장명" in k or "회사명" in k or "상호" in k), path.stem.split("_기초정보")[0])
        company_id = safe_name(next((fields[k] for k in fields if "사업자등록번호" in k), display_name)).replace(" ", "")

        allowance_sheet = workbook["수당등록"]
        allowance = "비과세수당"
        for col in range(1, min(allowance_sheet.max_column, 20) + 1):
            if _text(allowance_sheet.cell(3, col).value) == "근로소득세부":
                allowance = _text(allowance_sheet.cell(4, col).value) or allowance
                break

        sheet = workbook["수당조정(비과세)"]
        headers = [_text(sheet.cell(4, col).value) for col in range(1, sheet.max_column + 1)]
        by_header = {name: idx + 1 for idx, name in enumerate(headers) if name}
        required = ["기간", "사원번호", "이름", "주민번호", "인덱스", "과세차감총액"]
        missing_headers = [name for name in required if name not in by_header]
        if missing_headers:
            # 레거시 A:H 고정 규격 보정
            if sheet.max_column < 8:
                raise SchemaMismatch(f"수당조정 헤더 누락: {', '.join(missing_headers)}")
            by_header.update({"기간": 1, "연번": 2, "사원번호": 3, "이름": 4, "주민번호": 5, "인덱스": 6, "과세차감총액": 7})
        allowance_col = by_header.get(allowance, 8)
        records: dict[int, list[TaxExemption]] = defaultdict(list)
        seen_employees: set[tuple[int, str]] = set()
        seen_indexes: set[tuple[int, str]] = set()
        blank_run = 0
        for row in range(5, min(sheet.max_row, 100000) + 1):
            if all(sheet.cell(row, col).value in (None, "") for col in range(1, min(sheet.max_column, 8) + 1)):
                blank_run += 1
                if blank_run >= 100:
                    break
                continue
            blank_run = 0
            year = _year(sheet.cell(row, by_header["기간"]).value, cell=f"A{row}")
            if year not in requested_years:
                continue
            employee_id = _text(sheet.cell(row, by_header["사원번호"]).value)
            pdf_index = _text(sheet.cell(row, by_header["인덱스"]).value)
            if not employee_id:
                raise SchemaMismatch(f"사원번호 누락: C{row}")
            if not pdf_index:
                raise SchemaMismatch(f"PDF 인덱스 누락: F{row}")
            if (year, employee_id) in seen_employees:
                raise DuplicateEmployee(f"동일 연도 사원번호 중복: {year}, C{row}")
            if (year, pdf_index) in seen_indexes:
                raise DuplicatePdfIndex(f"동일 연도 PDF 인덱스 중복: {year}, F{row}")
            seen_employees.add((year, employee_id)); seen_indexes.add((year, pdf_index))
            amount = _money(sheet.cell(row, allowance_col).value, cell=f"{sheet.cell(4, allowance_col).column_letter}{row}")
            records[year].append(TaxExemption(
                company_id=company_id, year=year, employee_id=employee_id,
                name=_text(sheet.cell(row, by_header["이름"]).value),
                resident_id=_text(sheet.cell(row, by_header["주민번호"]).value),
                pdf_index=str(int(float(pdf_index))) if re.fullmatch(r"\d+(?:\.0+)?", pdf_index) else pdf_index,
                serial_number=_text(sheet.cell(row, by_header.get("연번", 2)).value),
                taxable_deduction_total=_money(sheet.cell(row, by_header["과세차감총액"]).value, cell=f"G{row}"),
                lines=(TaxExemptionLine(allowance, amount),), source_sheet=sheet.title, source_row=row,
            ))
        if not any(records.values()):
            raise SchemaMismatch("요청 연도의 수당조정 대상자가 없습니다.")
        return CompanyInfo(company_id, display_name, fields), dict(records)
    finally:
        workbook.close()
