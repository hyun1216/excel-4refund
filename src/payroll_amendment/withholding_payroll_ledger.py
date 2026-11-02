from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
import re
from typing import Any
from uuid import uuid4

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pdfplumber

from .errors import PdfToExcelError, SchemaMismatch
from .infra import sha256


BASE_HEADERS = (
    "사원번호",
    "성명",
    "주민등록번호",
    "급여구분",
    "기본급",
    "상여",
    "식대",
    "보육수당",
    "교통비",
    "차량유지비",
    "연장근로",
    "시간외수당",
    "야간연장수당(비과세)",
    "지급합계",
    "국민연금",
    "건강보험",
    "고용보험",
    "장기요양보험료",
    "학자금상환액",
    "국민연금정산",
    "건강보험료정산",
    "장기요양보험료정산",
    "건강보험료정산(해외근무전액)",
    "중도퇴사자 소득세",
    "중도퇴사자 주민세",
    "소득세",
    "지방소득세",
    "농특세",
    "연말정산소득세",
    "연말정산지방소득세",
    "연말정산농특세",
    "기타공제액",
    "공제합계",
    "차인지급액",
)
HEADER_FILL = PatternFill("solid", fgColor="FF1F4E78")
HEADER_FONT = Font(color="FFFFFFFF", bold=True)
THIN_BORDER = Border(
    left=Side(style="thin", color="FFD9E2F3"),
    right=Side(style="thin", color="FFD9E2F3"),
    top=Side(style="thin", color="FFD9E2F3"),
    bottom=Side(style="thin", color="FFD9E2F3"),
)
MONEY_FORMAT = "#,##0;[Red]-#,##0"
_MONTH_LABEL = re.compile(r"^(?P<month>[1-9]|1[0-2])월$")
_EMPLOYEE_CODE = re.compile(r"사원코드\s*:\s*([^)]*)\)")
_EMPLOYEE_NAME = re.compile(r"성명:\s*(.*?)\s+주민등록번호:")


@dataclass(frozen=True)
class PayrollLedgerBuildResult:
    output_path: Path
    manifest_path: Path
    employee_count: int
    month_row_counts: dict[str, int]


def _compact(value: object) -> str:
    return re.sub(r"\s+", "", str(value or ""))


def _money(value: object) -> int:
    text = _compact(value).replace(",", "")
    if not text:
        return 0
    if not re.fullmatch(r"-?\d+", text):
        raise PdfToExcelError(f"금액을 읽을 수 없습니다: {value!r}")
    return int(text)


def _largest_table(page: Any) -> list[list[str | None]]:
    tables = page.extract_tables()
    if not tables:
        raise PdfToExcelError("소득자별 원천징수부 표를 찾지 못했습니다.")
    return max(tables, key=len)


def _identity(page: Any, source: Path) -> tuple[str, str]:
    text = page.extract_text() or ""
    code = _EMPLOYEE_CODE.search(text)
    name = _EMPLOYEE_NAME.search(text)
    if code is None or name is None:
        raise PdfToExcelError(f"사원코드 또는 성명을 찾지 못했습니다: {source.name}")
    return code.group(1).strip(), name.group(1).strip()


def _month_rows(table: list[list[str | None]], year: int) -> dict[int, list[str | None]]:
    result: dict[int, list[str | None]] = {}
    pattern = re.compile(rf"^{year}/(?P<month>0[1-9]|1[0-2])$")
    for row in table:
        value = str(row[1] or "").strip() if len(row) > 1 else ""
        matched = pattern.fullmatch(value)
        if matched:
            result[int(matched.group("month"))] = row
    return result


def _tax_rows(table: list[list[str | None]]) -> tuple[dict[int, list[str | None]], list[str | None] | None]:
    months: dict[int, list[str | None]] = {}
    settlement = None
    for row in table:
        label = str(row[0] or "").strip() if row else ""
        matched = _MONTH_LABEL.fullmatch(label)
        if matched:
            months[int(matched.group("month"))] = row
        elif label == "연말":
            settlement = row
    return months, settlement


def _read_basic_people(path: Path) -> dict[str, tuple[str, str]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    try:
        candidates = [name for name in workbook.sheetnames if name.startswith("원천징수추출(")]
        if len(candidates) != 1:
            raise SchemaMismatch("기초정보에서 원천징수추출 시트를 하나로 확정할 수 없습니다.")
        sheet = workbook[candidates[0]]
        headers = {str(cell.value or "").strip(): index for index, cell in enumerate(sheet[1], 1)}
        required = {"사원번호", "성명", "주민등록번호"}
        missing = sorted(required - set(headers))
        if missing:
            raise SchemaMismatch(f"기초정보 원천징수추출 시트에 필수 열이 없습니다: {', '.join(missing)}")
        people: dict[str, tuple[str, str]] = {}
        for row in range(2, sheet.max_row + 1):
            employee_id = str(sheet.cell(row, headers["사원번호"]).value or "").strip()
            if not employee_id:
                continue
            if employee_id in people:
                raise SchemaMismatch(f"기초정보 사원번호가 중복되었습니다: {employee_id}")
            name = str(sheet.cell(row, headers["성명"]).value or "").strip()
            resident = str(sheet.cell(row, headers["주민등록번호"]).value or "").strip()
            digits = re.sub(r"\D", "", resident)
            if len(digits) != 13:
                raise SchemaMismatch(f"기초정보 주민등록번호가 13자리가 아닙니다: 사원번호 {employee_id}")
            people[employee_id] = (name, f"{digits[:6]}-{digits[6:]}")
        return people
    finally:
        workbook.close()


def _employee_rows(
    source: Path,
    *,
    year: int,
    people: dict[str, tuple[str, str]],
) -> tuple[str, str, dict[int, dict[str, int | str]]]:
    try:
        with pdfplumber.open(source) as document:
            if len(document.pages) != 3:
                raise PdfToExcelError(f"소득자별 원천징수부는 직원별 3쪽이어야 합니다: {source.name}")
            employee_id, pdf_name = _identity(document.pages[0], source)
            if employee_id not in people:
                raise SchemaMismatch(f"기초정보에 없는 사원번호입니다: {employee_id}")
            basic_name, resident = people[employee_id]
            if _compact(pdf_name) != _compact(basic_name):
                raise SchemaMismatch(
                    f"기초정보와 원천징수부의 성명이 다릅니다: 사원번호 {employee_id}, {basic_name!r}, {pdf_name!r}"
                )
            income = _month_rows(_largest_table(document.pages[0]), year)
            taxfree = _month_rows(_largest_table(document.pages[1]), year)
            tax, settlement = _tax_rows(_largest_table(document.pages[2]))
    except (PdfToExcelError, SchemaMismatch):
        raise
    except Exception as exc:
        raise PdfToExcelError(f"소득자별 원천징수부를 읽을 수 없습니다: {source.name}: {exc}") from exc

    rows: dict[int, dict[str, int | str]] = {}
    for month in range(1, 13):
        income_row = income.get(month, [])
        taxfree_row = taxfree.get(month, [])
        tax_row = tax.get(month, [])
        salary = _money(income_row[2]) if len(income_row) > 2 else 0
        bonus = sum(_money(income_row[index]) for index in range(3, 9) if len(income_row) > index)
        night = _money(taxfree_row[5]) if len(taxfree_row) > 5 else 0
        meal = _money(taxfree_row[6]) if len(taxfree_row) > 6 else 0
        childcare = _money(taxfree_row[7]) if len(taxfree_row) > 7 else 0
        other_reported = _money(taxfree_row[9]) if len(taxfree_row) > 9 else 0
        car = _money(taxfree_row[11]) if len(taxfree_row) > 11 else 0
        other_unreported = _money(taxfree_row[12]) if len(taxfree_row) > 12 else 0
        if other_reported or other_unreported:
            raise SchemaMismatch(
                f"급여대장 열에 없는 기타 비과세소득이 있습니다: {source.name}, {month}월"
            )
        pension = _money(tax_row[8]) if len(tax_row) > 8 else 0
        health = _money(tax_row[9]) if len(tax_row) > 9 else 0
        employment = _money(tax_row[10]) if len(tax_row) > 10 else 0
        donation = _money(tax_row[11]) if len(tax_row) > 11 else 0
        income_tax = _money(tax_row[4]) if len(tax_row) > 4 else 0
        local_tax = _money(tax_row[5]) if len(tax_row) > 5 else 0
        rural_tax = _money(tax_row[6]) if len(tax_row) > 6 else 0
        payment_total = salary + bonus + meal + childcare + car + night
        deduction_total = pension + health + employment + donation + income_tax + local_tax + rural_tax
        if not any((payment_total, deduction_total)):
            continue
        rows[month] = {
            "사원번호": employee_id.zfill(3) if employee_id.isdigit() else employee_id,
            "성명": basic_name,
            "주민등록번호": resident,
            "급여구분": "급상여",
            "기본급": salary,
            "상여": bonus,
            "식대": meal,
            "보육수당": childcare,
            "교통비": 0,
            "차량유지비": car,
            "연장근로": 0,
            "시간외수당": 0,
            "야간연장수당(비과세)": night,
            "지급합계": payment_total,
            "국민연금": pension,
            "건강보험": health,
            "고용보험": employment,
            "장기요양보험료": 0,
            "학자금상환액": 0,
            "국민연금정산": 0,
            "건강보험료정산": 0,
            "장기요양보험료정산": 0,
            "건강보험료정산(해외근무전액)": 0,
            "중도퇴사자 소득세": 0,
            "중도퇴사자 주민세": 0,
            "소득세": income_tax,
            "지방소득세": local_tax,
            "농특세": rural_tax,
            "연말정산소득세": 0,
            "연말정산지방소득세": 0,
            "연말정산농특세": 0,
            "기타공제액": donation,
            "공제합계": deduction_total,
            "차인지급액": payment_total - deduction_total,
        }

    if settlement is not None and rows:
        settlement_month = max(rows)
        row = rows[settlement_month]
        values = {
            "연말정산소득세": _money(settlement[4]) if len(settlement) > 4 else 0,
            "연말정산지방소득세": _money(settlement[5]) if len(settlement) > 5 else 0,
            "연말정산농특세": _money(settlement[6]) if len(settlement) > 6 else 0,
        }
        for header, value in values.items():
            row[header] = value
        adjustment = sum(values.values())
        row["공제합계"] = int(row["공제합계"]) + adjustment
        row["차인지급액"] = int(row["지급합계"]) - int(row["공제합계"])
    return employee_id, basic_name, rows


def _write_month_sheet(worksheet: Any, rows: list[dict[str, int | str]]) -> None:
    worksheet.sheet_view.showGridLines = False
    worksheet.append(BASE_HEADERS)
    for record in rows:
        worksheet.append([None if record.get(header) == 0 else record.get(header) for header in BASE_HEADERS])
    worksheet.freeze_panes = "E2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(BASE_HEADERS))}{max(1, len(rows) + 1)}"
    for cell in worksheet[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = THIN_BORDER
    worksheet.row_dimensions[1].height = 22
    for column in range(1, len(BASE_HEADERS) + 1):
        letter = get_column_letter(column)
        worksheet.column_dimensions[letter].width = 16 if column > 4 else (16 if column == 3 else 13)
    for row_number in range(2, len(rows) + 2):
        worksheet.cell(row_number, 1).number_format = "@"
        worksheet.cell(row_number, 3).number_format = "@"
        for column in range(1, len(BASE_HEADERS) + 1):
            cell = worksheet.cell(row_number, column)
            cell.border = THIN_BORDER
            if column > 4:
                cell.number_format = MONEY_FORMAT
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="center" if column in {1, 4} else "left")


def build_monthly_payroll_ledger_from_pdfs(
    *,
    ledger_pdf_directory: Path,
    basic_info_path: Path,
    output_path: Path,
    year: int,
) -> PayrollLedgerBuildResult:
    source_directory = Path(ledger_pdf_directory).resolve(strict=True)
    basic_info = Path(basic_info_path).resolve(strict=True)
    destination = Path(output_path).resolve()
    pdf_files = sorted(
        source_directory.glob("*.pdf"),
        key=lambda path: int(path.name.split("_", 1)[0]),
    )
    if not pdf_files:
        raise PdfToExcelError(f"소득자별 원천징수부 PDF가 없습니다: {source_directory}")
    people = _read_basic_people(basic_info)
    rows_by_month: dict[int, list[dict[str, int | str]]] = {month: [] for month in range(1, 13)}
    employee_records = []
    seen: set[str] = set()
    for source in pdf_files:
        employee_id, name, rows = _employee_rows(source, year=year, people=people)
        if employee_id in seen:
            raise SchemaMismatch(f"소득자별 원천징수부 사원번호가 중복되었습니다: {employee_id}")
        seen.add(employee_id)
        for month, row in rows.items():
            rows_by_month[month].append(row)
        employee_records.append(
            {
                "employee_id": employee_id,
                "name": name,
                "source": source.name,
                "sha256": sha256(source),
                "active_months": sorted(rows),
            }
        )
    if seen != set(people):
        missing = sorted(set(people) - seen)
        raise SchemaMismatch(f"원천징수부에 없는 기초정보 사원번호가 있습니다: {missing}")

    workbook = Workbook()
    workbook.remove(workbook.active)
    for month in range(1, 13):
        worksheet = workbook.create_sheet(f"{month:02d}월")
        _write_month_sheet(worksheet, rows_by_month[month])
    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True

    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.parent / f".{destination.stem}.{uuid4().hex}.tmp.xlsx"
    try:
        workbook.save(temporary)
        check = load_workbook(temporary, read_only=True, data_only=False)
        try:
            if check.sheetnames != [f"{month:02d}월" for month in range(1, 13)]:
                raise SchemaMismatch("생성 급여대장의 월별 시트 검증에 실패했습니다.")
            for month, rows in rows_by_month.items():
                sheet = check[f"{month:02d}월"]
                if sheet.max_row != len(rows) + 1 or sheet.max_column != len(BASE_HEADERS):
                    raise SchemaMismatch(f"생성 급여대장 크기 검증에 실패했습니다: {month:02d}월")
                if [cell.value for cell in sheet[1]] != list(BASE_HEADERS):
                    raise SchemaMismatch(f"생성 급여대장 헤더 검증에 실패했습니다: {month:02d}월")
        finally:
            check.close()
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)

    manifest_path = destination.with_suffix(".manifest.json")
    manifest = {
        "year": year,
        "output": destination.name,
        "basic_info": basic_info.name,
        "employee_count": len(employee_records),
        "month_row_counts": {f"{month:02d}월": len(rows) for month, rows in rows_by_month.items()},
        "headers": list(BASE_HEADERS),
        "source_limitations": [
            "원천징수부의 건강보험 금액은 장기요양보험료와 분리되지 않아 건강보험 열에 전액 기록",
            "원천징수부에 없는 장기요양보험료와 보험 정산 항목은 빈칸",
            "연말 행의 세액은 해당 직원의 마지막 지급월 연말정산 열에 기록",
        ],
        "employees": employee_records,
    }
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return PayrollLedgerBuildResult(
        destination,
        manifest_path,
        len(employee_records),
        {f"{month:02d}월": len(rows) for month, rows in rows_by_month.items()},
    )


