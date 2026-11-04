from pathlib import Path

from openpyxl import Workbook, load_workbook

from payroll_amendment.resident_excel import normalize_resident_ids_in_workbook


def test_normalize_resident_ids_in_workbook(tmp_path: Path) -> None:
    path = tmp_path / "기초정보.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.append(["성명", "주민등록번호"])
    sheet.append(["가명", 9001011234567])
    health = workbook.create_sheet("건보정보")
    health.append(["성명", "주민번호\n(숫자만 13자리)"])
    health.append(["가명", "010101-2234567"])
    workbook.save(path)
    workbook.close()

    assert normalize_resident_ids_in_workbook(path) == 2

    checked = load_workbook(path, read_only=True, data_only=True)
    try:
        assert checked.active["B2"].value == "900101-1234567"
        assert checked.active["B2"].number_format == "@"
        assert checked["건보정보"]["B2"].value == "010101-2234567"
        assert checked["건보정보"]["B2"].number_format == "@"
    finally:
        checked.close()
