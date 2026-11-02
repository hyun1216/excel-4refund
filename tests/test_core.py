from decimal import Decimal
from pathlib import Path

from openpyxl import Workbook
import pytest

from payroll_amendment.basic_info import parse_basic_info
from payroll_amendment.errors import DuplicatePdfIndex
from payroll_amendment.year_parser import parse_years


def make_basic(path: Path, duplicate_index: bool = False) -> None:
    book = Workbook()
    info = book.active; info.title = "사업장정보"
    info["A5"] = "사업장명"; info["B5"] = "테스트회사"
    info["A6"] = "사업자등록번호"; info["B6"] = "123-45-67890"
    register = book.create_sheet("수당등록")
    register["A3"] = "수당명"; register["A4"] = "야간연장수당"
    register["D3"] = "근로소득세부"; register["D4"] = "18-1.생산직등야간근로수당"
    adjust = book.create_sheet("수당조정(비과세)")
    adjust.append([]); adjust.append([]); adjust.append([])
    adjust.append(["기간", "연번", "사원번호", "이름", "주민번호", "인덱스", "과세차감총액", "18-1.생산직등야간근로수당"])
    adjust.append(["2025년", 1, "00123", "가명", "900101-1234567", 7, -100000, 100000])
    if duplicate_index:
        adjust.append(["2025년", 2, "00456", "가명2", "900101-2234567", 7, -200000, 200000])
    book.save(path)


def test_parse_years():
    assert parse_years("23년,2024년,25년", max_year=2026) == [2023, 2024, 2025]
    with pytest.raises(ValueError):
        parse_years("2023.5", max_year=2026)


def test_basic_info_preserves_employee_id_and_zero_rules(tmp_path):
    path = tmp_path / "회사_기초정보(3차).xlsx"; make_basic(path)
    company, records = parse_basic_info(path, {2025})
    assert company.display_name == "테스트회사"
    assert records[2025][0].employee_id == "00123"
    assert records[2025][0].lines[0].amount == Decimal("100000")


def test_duplicate_pdf_index_is_blocked(tmp_path):
    path = tmp_path / "회사_기초정보(3차).xlsx"; make_basic(path, True)
    with pytest.raises(DuplicatePdfIndex):
        parse_basic_info(path, {2025})

