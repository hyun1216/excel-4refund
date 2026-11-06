from pathlib import Path

from openpyxl import Workbook, load_workbook
import pytest

from payroll_amendment.errors import SchemaMismatch
from payroll_amendment.payroll_converter import (
    PayrollConverter,
    _group_payroll_workbooks,
    _merge_monthly_payroll_workbooks,
    _payroll_workbook_period,
    _payroll_workbook_year,
    _receipt_year,
)


def _make_payroll(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "06월"
    sheet.append([
        "사원번호",
        "성명",
        "주민등록번호",
        "급여구분",
        "기본급",
        "주휴수당",
        "휴일연장야간",
        "지급총액",
        "소득세",
        "공제총액",
        "실지급액",
    ])
    sheet.append([
        "001",
        "윤현길",
        "000107-3114719",
        "급여",
        1_000_000,
        300_000,
        70_000,
        1_370_000,
        10_000,
        10_000,
        1_360_000,
    ])
    workbook.save(path)


def test_preserves_distinct_source_payment_items(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    _make_payroll(source)

    PayrollConverter().convert_excel_workbook(source, output)

    workbook = load_workbook(output, read_only=True, data_only=True)
    try:
        sheet = workbook["06월"]
        headers = [cell.value for cell in sheet[1]]
        values = [cell.value for cell in sheet[2]]
        row = dict(zip(headers, values))
        assert "주휴수당" in headers
        assert "휴일연장야간" in headers
        assert "휴일수당" not in headers
        assert row["주휴수당"] == 300_000
        assert row["휴일연장야간"] == 70_000
        assert row["지급합계"] == 1_370_000
    finally:
        workbook.close()


@pytest.mark.parametrize(
    ("raw_name", "expected_name"),
    [
        ("#$김길동(eirwe23023)1!@$", "김길동"),
        ("123김길동(erwer김딜)", "김길동"),
    ],
)
def test_cleans_employee_name_during_payroll_conversion(tmp_path, raw_name, expected_name):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    _make_payroll(source)
    workbook = load_workbook(source)
    workbook.active["B2"] = raw_name
    workbook.save(source)

    PayrollConverter().convert_excel_workbook(source, output)

    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        assert converted["06월"]["B2"].value == expected_name
    finally:
        converted.close()


def test_scans_each_sheet_header_once_during_conversion(tmp_path, monkeypatch):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    _make_payroll(source)
    converter = PayrollConverter()
    original = converter._detect_header_row_from_excel
    scanned_sheets = []

    def tracked(excel_file, sheet_name, source_name):
        scanned_sheets.append(sheet_name)
        return original(excel_file, sheet_name, source_name)

    monkeypatch.setattr(converter, "_detect_header_row_from_excel", tracked)
    converter.convert_excel_workbook(source, output)

    assert scanned_sheets == ["06월"]


def test_payroll_workbook_year_reads_korean_filename() -> None:
    assert _payroll_workbook_year(Path("(SCK) 2025년 월별 급여대장.xlsx")) == 2025
    assert _payroll_workbook_year(Path("급여대장 202301.xlsx")) == 2023


def test_receipt_year_uses_work_period_start_date(monkeypatch) -> None:
    class Page:
        @staticmethod
        def extract_tables():
            return [[[None, "⑪ 근 무 기 간", None, "2022.10.11 ~ 2022.12.31"]]]

    class Document:
        pages = [Page()]

        def __enter__(self):
            return self

        def __exit__(self, *_):
            return None

    monkeypatch.setattr(
        "payroll_amendment.payroll_converter.pdfplumber.open",
        lambda _: Document(),
    )

    assert _receipt_year(Path("21_이용휘_소득자보관용.pdf")) == 2022


def test_company_metadata_uses_withholding_agent_id_before_prior_employer(monkeypatch) -> None:
    class Page:
        @staticmethod
        def extract_text():
            return """
            (주)다원에이치알
            231-88-02177
            1. 법인명(상호) 주식회사 다원로지스틱스 2. 대표자(성명) 김대표
            110111 - 6455326 803 - 81 - 00806 3. 사업자등록번호 4. 주민(법인)등록번호
            """

    class Reader:
        pages = [Page()]

    monkeypatch.setattr("payroll_amendment.payroll_converter.PdfReader", lambda _: Reader())

    assert PayrollConverter._company_metadata(Path("원천징수영수증.pdf")) == (
        "주식회사 다원로지스틱스",
        "803-81-00806",
    )


def test_receipt_year_falls_back_to_explicit_filename_year(monkeypatch) -> None:
    class Page:
        @staticmethod
        def extract_tables():
            return []

    class Document:
        pages = [Page()]

        def __enter__(self):
            return self

        def __exit__(self, *_):
            return None

    monkeypatch.setattr(
        "payroll_amendment.payroll_converter.pdfplumber.open",
        lambda _: Document(),
    )

    assert _receipt_year(Path("홍길동_2024_근로소득원천징수영수증.pdf")) == 2024


def test_receipt_year_does_not_use_unrelated_numbers(monkeypatch) -> None:
    class Page:
        @staticmethod
        def extract_tables():
            return [[["⑪ 근 무 기 간", None, "사원번호 94"]]]

    class Document:
        pages = [Page()]

        def __enter__(self):
            return self

        def __exit__(self, *_):
            return None

    monkeypatch.setattr(
        "payroll_amendment.payroll_converter.pdfplumber.open",
        lambda _: Document(),
    )

    with pytest.raises(SchemaMismatch, match="귀속연도"):
        _receipt_year(Path("94_정시환_소득자보관용.pdf"))


def test_payroll_workbook_period_distinguishes_annual_and_monthly_files() -> None:
    assert _payroll_workbook_period(Path("(SCK) 2025년 월별 급여대장.xlsx")) == (2025, None)
    assert _payroll_workbook_period(Path("급여대장_2025년 01월.xlsx")) == (2025, 1)
    assert _payroll_workbook_period(Path("급여대장_25년12월.xlsx")) == (2025, 12)
    assert _payroll_workbook_period(Path("급여대장 202301.xlsx")) == (2023, 1)
    assert _payroll_workbook_period(Path("급여대장 2023-02.xlsx")) == (2023, 2)
    assert _payroll_workbook_period(Path("급여대장 2023_03.xlsx")) == (2023, 3)


def test_groups_monthly_workbooks_by_year_and_month() -> None:
    paths = [Path(f"급여대장_2025년 {month:02d}월.xlsx") for month in range(1, 13)]

    annual, monthly = _group_payroll_workbooks(paths)

    assert annual == {}
    assert monthly[2025] == {month: paths[month - 1] for month in range(1, 13)}


def test_rejects_annual_and_monthly_workbooks_for_same_year() -> None:
    with pytest.raises(SchemaMismatch, match="통합·월별"):
        _group_payroll_workbooks([
            Path("2025년 월별 급여대장.xlsx"),
            Path("급여대장_2025년 01월.xlsx"),
        ])


def test_rejects_duplicate_monthly_workbooks() -> None:
    with pytest.raises(SchemaMismatch, match=r"2025년\(01월\)"):
        _group_payroll_workbooks([
            Path("급여대장_2025년 01월.xlsx"),
            Path("2025년01월_급여대장.xlsx"),
        ])


def test_merges_monthly_workbooks_into_month_sheets(tmp_path) -> None:
    january = tmp_path / "급여대장_2025년 01월.xlsx"
    february = tmp_path / "급여대장_2025년 02월.xlsx"
    intermediate = tmp_path / "merged.xlsx"
    output = tmp_path / "output.xlsx"
    _make_payroll(january)
    _make_payroll(february)

    _merge_monthly_payroll_workbooks({1: january, 2: february}, intermediate)
    result = PayrollConverter().convert_excel_workbook(intermediate, output)

    assert result.sheet_row_counts == {"01월": 1, "02월": 1}
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        assert converted.sheetnames == [f"{month:02d}월" for month in range(1, 13)]
        assert converted["01월"].max_row == 2
        assert converted["02월"].max_row == 2
    finally:
        converted.close()


def test_allows_one_won_total_difference(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    _make_payroll(source)
    workbook = load_workbook(source)
    workbook.active["K2"] = 1_359_999
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert result.warnings == []
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        row = dict(zip(
            (cell.value for cell in converted["06월"][1]),
            (cell.value for cell in converted["06월"][2]),
        ))
        assert row["차인지급액"] == 1_359_999
    finally:
        converted.close()


def test_zeroes_amounts_but_keeps_identity_for_invalid_total(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    _make_payroll(source)
    workbook = load_workbook(source)
    workbook.active["K2"] = 1
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert len(result.warnings) == 1
    assert result.warnings[0].check == "차인지급액"
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        row = dict(zip(
            (cell.value for cell in converted["06월"][1]),
            (cell.value for cell in converted["06월"][2]),
        ))
        assert row["성명"] == "윤현길"
        assert row["주민등록번호"] == "000107-3114719"
        for header in [
            *result.payment_headers,
            "지급합계",
            *result.deduction_headers,
            "공제합계",
            "차인지급액",
        ]:
            assert row[header] == 0
    finally:
        converted.close()


def test_zeroes_all_months_for_employee_with_any_invalid_total(tmp_path):
    source = tmp_path / "source.xlsx"
    output = tmp_path / "output.xlsx"
    _make_payroll(source)
    workbook = load_workbook(source)
    january = workbook.active
    january.title = "01월"
    february = workbook.copy_worksheet(january)
    february.title = "02월"
    february["K2"] = 1
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert len(result.warnings) == 1
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        for sheet_name in ("01월", "02월"):
            sheet = converted[sheet_name]
            row = dict(zip(
                (cell.value for cell in sheet[1]),
                (cell.value for cell in sheet[2]),
            ))
            assert row["성명"] == "윤현길"
            assert row["지급합계"] == 0
            assert row["공제합계"] == 0
            assert row["차인지급액"] == 0
    finally:
        converted.close()


def test_converts_coded_payroll_workbook(tmp_path):
    source = tmp_path / "(SCK) 2025년 월별 급여대장.xlsx"
    output = tmp_path / "output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "202501"
    sheet.append([
        "사번",
        "성명",
        "주민번호",
        "실입사일",
        "입사일",
        "퇴사일",
        "A01",
        "A12",
        "BF01",
        "BF13",
        "BU01",
        "지급총액",
        "공제총액",
        "실지급액",
    ])
    sheet.append([
        None,
        None,
        None,
        None,
        None,
        None,
        "기본급",
        "학자금",
        "소득세",
        "학자금",
        "기타공제1",
        None,
        None,
        None,
    ])
    sheet.append([
        "001",
        "홍원구",
        "800101-1234567",
        "2019-01-01",
        "2020-01-01",
        "2023-12-31",
        4_000_000,
        191_667,
        100_000,
        600_000,
        51_330,
        4_191_667,
        751_330,
        3_440_337,
    ])
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert result.payment_headers == ["기본급", "학자금(지급)"]
    assert result.deduction_headers == ["소득세", "학자금", "기타공제1"]
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        sheet = converted["01월"]
        headers = [cell.value for cell in sheet[1]]
        row = dict(zip(headers, [cell.value for cell in sheet[2]]))
        assert "실입사일" not in headers
        assert row["입사일"] == "2020-01-01"
        assert row["퇴사일"] == "2023-12-31"
        assert row["기본급"] == 4_000_000
        assert row["학자금(지급)"] == 191_667
        assert row["학자금"] == 600_000
        assert row["지급합계"] == 4_191_667
        assert row["공제합계"] == 751_330
        assert row["차인지급액"] == 3_440_337
    finally:
        converted.close()


def test_coded_insurance_labels_map_to_standard_headers(tmp_path):
    source = tmp_path / "coded_insurance.xlsx"
    output = tmp_path / "output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "202501"
    sheet.append([
        "사번", "성명", "주민번호", "A01", "BF01", "BF02", "BF03", "BF04",
        "지급총액", "공제총액", "실지급액",
    ])
    sheet.append([
        None, None, None, "기본급", "건강보험료(보수월액*3.545%)",
        "장기요양 보험료(건강보험료*12.81%)", "고용보험료",
        "국민연금(보수월액*4.5%)", None, None, None,
    ])
    sheet.append([
        "001", "홍길동", "900101-1234567", 2_000_000,
        80_000, 10_000, 20_000, 90_000, 2_000_000, 200_000, 1_800_000,
    ])
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert result.deduction_headers == ["건강보험", "장기요양", "고용보험", "국민연금"]
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        row = dict(zip(
            (cell.value for cell in converted["01월"][1]),
            (cell.value for cell in converted["01월"][2]),
        ))
        assert row["건강보험"] == 80_000
        assert row["장기요양"] == 10_000
        assert row["고용보험"] == 20_000
        assert row["국민연금"] == 90_000
    finally:
        converted.close()


def test_duplicate_insurance_aliases_with_same_amount_are_not_doubled(tmp_path):
    source = tmp_path / "duplicate_insurance.xlsx"
    output = tmp_path / "output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "01월"
    sheet.append([
        "사번", "성명", "주민번호", "기본급", "지급총액",
        "건강보험", "건강보험료(보수월액*3.545%)", "공제총액", "실지급액",
    ])
    sheet.append([
        "001", "홍길동", "900101-1234567", 2_000_000, 2_000_000,
        80_000, 80_000, 80_000, 1_920_000,
    ])
    workbook.save(source)

    PayrollConverter().convert_excel_workbook(source, output)

    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        headers = [cell.value for cell in converted["01월"][1]]
        row = dict(zip(headers, [cell.value for cell in converted["01월"][2]]))
        assert headers.count("건강보험") == 1
        assert row["건강보험"] == 80_000
    finally:
        converted.close()


def test_duplicate_insurance_aliases_with_different_amounts_require_review(tmp_path):
    source = tmp_path / "conflicting_insurance.xlsx"
    output = tmp_path / "output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "01월"
    sheet.append([
        "사번", "성명", "주민번호", "기본급", "지급총액",
        "건강보험", "건강보험료(보수월액*3.545%)", "공제총액", "실지급액",
    ])
    sheet.append([
        "001", "홍길동", "900101-1234567", 2_000_000, 2_000_000,
        80_000, 90_000, 170_000, 1_830_000,
    ])
    workbook.save(source)

    with pytest.raises(SchemaMismatch, match="동일 보험료 후보 열의 금액이 서로 다릅니다: 건강보험"):
        PayrollConverter().convert_excel_workbook(source, output)


def test_excludes_actual_join_date_from_generic_mapping() -> None:
    mapping = PayrollConverter().normalize_and_map_columns([
        "사번", "성명", "주민번호", "실입사일", "기본급",
        "지급총액", "소득세", "공제총액", "실지급액",
    ])

    assert mapping["실입사일"] is None


def test_coded_blank_active_items_use_distinct_source_codes(tmp_path):
    source = tmp_path / "coded_blank_headers.xlsx"
    output = tmp_path / "output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "202501"
    sheet.append([
        "사번",
        "성명",
        "주민번호",
        "A18",
        "A19",
        "BF01",
        "BU03",
        "지급총액",
        "공제총액",
        "실지급액",
    ])
    sheet.append([None, None, None, None, None, "소득세", None, None, None, None])
    sheet.append(["001", "홍길동", "900101-1234567", 100_000, 200_000, 10_000, 5_000, 300_000, 15_000, 285_000])
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert result.payment_headers == ["A18(지급)", "A19(지급)"]
    assert result.deduction_headers == ["소득세", "BU03(공제)"]
    assert result.unknown_columns == {
        "A18(지급)": "payment",
        "A19(지급)": "payment",
        "BU03(공제)": "deduction",
    }
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        target = converted["01월"]
        headers = [cell.value for cell in target[1]]
        row = dict(zip(headers, [cell.value for cell in target[2]]))
        assert "nan(지급)" not in headers
        assert "nan" not in headers
        assert row["A18(지급)"] == 100_000
        assert row["A19(지급)"] == 200_000
        assert row["BU03(공제)"] == 5_000
    finally:
        converted.close()


def test_coded_blank_inactive_items_are_removed_after_employee_filter(tmp_path):
    source = tmp_path / "coded_blank_inactive_headers.xlsx"
    output = tmp_path / "output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "202501"
    sheet.append([
        "사번",
        "성명",
        "주민번호",
        "A12",
        "A13",
        "BF01",
        "지급총액",
        "공제총액",
        "실지급액",
    ])
    sheet.append([None, None, None, None, "기술수당", "소득세", None, None, None])
    sheet.append(["001", "홍길동", "900101-1234567", 0, 50_000, 5_000, 50_000, 5_000, 45_000])
    sheet.append(["조회건수:", None, None, 1, None, None, None, None, None])
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert result.payment_headers == ["기술수당"]
    assert result.deduction_headers == ["소득세"]
    assert result.unknown_columns == {}
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        headers = [cell.value for cell in converted["01월"][1]]
        assert "A12(지급)" not in headers
        assert "기술수당" in headers
    finally:
        converted.close()


def test_coded_header_registry_reuses_label_from_another_sheet(tmp_path):
    source = tmp_path / "coded_cross_sheet_headers.xlsx"
    output = tmp_path / "output.xlsx"
    workbook = Workbook()
    for index, (title, label, amount) in enumerate(
        (("202501", None, 100_000), ("202502", "현장수당", 200_000))
    ):
        sheet = workbook.active if index == 0 else workbook.create_sheet()
        sheet.title = title
        sheet.append(["사번", "성명", "주민번호", "A18", "BF01", "지급총액", "공제총액", "실지급액"])
        sheet.append([None, None, None, label, "소득세", None, None, None])
        sheet.append(["001", "홍길동", "900101-1234567", amount, 10_000, amount, 10_000, amount - 10_000])
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert result.payment_headers == ["현장수당"]
    assert result.unknown_columns == {}
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        for sheet_name, amount in (("01월", 100_000), ("02월", 200_000)):
            sheet = converted[sheet_name]
            row = dict(zip(
                (cell.value for cell in sheet[1]),
                (cell.value for cell in sheet[2]),
            ))
            assert row["현장수당"] == amount
    finally:
        converted.close()


def test_coded_header_registry_keeps_different_labels_separate_across_sheets(tmp_path):
    source = tmp_path / "coded_cross_sheet_conflicting_headers.xlsx"
    output = tmp_path / "output.xlsx"
    workbook = Workbook()
    for index, (title, label, amount) in enumerate(
        (
            ("202501", "능력가급", 100_000),
            ("202502", "장기근속수당", 200_000),
            ("202503", None, 300_000),
        )
    ):
        sheet = workbook.active if index == 0 else workbook.create_sheet()
        sheet.title = title
        sheet.append(["사번", "성명", "주민번호", "A11", "BF01", "지급총액", "공제총액", "실지급액"])
        sheet.append([None, None, None, label, "소득세", None, None, None])
        sheet.append(["001", "홍길동", "900101-1234567", amount, 10_000, amount, 10_000, amount - 10_000])
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert result.payment_headers == ["능력가급", "장기근속수당", "A11(지급)"]
    assert result.unknown_columns == {"A11(지급)": "payment"}
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        for sheet_name, header, amount in (
            ("01월", "능력가급", 100_000),
            ("02월", "장기근속수당", 200_000),
            ("03월", "A11(지급)", 300_000),
        ):
            sheet = converted[sheet_name]
            row = dict(zip(
                (cell.value for cell in sheet[1]),
                (cell.value for cell in sheet[2]),
            ))
            assert row[header] == amount
    finally:
        converted.close()


def test_coded_duplicate_labels_keep_distinct_source_codes(tmp_path):
    source = tmp_path / "coded_duplicate_headers.xlsx"
    output = tmp_path / "output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "202501"
    sheet.append(["사번", "성명", "주민번호", "A01", "A02", "BF01", "지급총액", "공제총액", "실지급액"])
    sheet.append([None, None, None, "현장수당", "현장수당", "소득세", None, None, None])
    sheet.append(["001", "홍길동", "900101-1234567", 100_000, 200_000, 10_000, 300_000, 10_000, 290_000])
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert result.payment_headers == ["현장수당 [A01]", "현장수당 [A02]"]
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        target = converted["01월"]
        headers = [cell.value for cell in target[1]]
        row = dict(zip(headers, [cell.value for cell in target[2]]))
        assert row["현장수당 [A01]"] == 100_000
        assert row["현장수당 [A02]"] == 200_000
    finally:
        converted.close()


def test_coded_repeated_headers_follow_each_section_and_map_label_variants(tmp_path):
    source = tmp_path / "coded_repeated_headers.xlsx"
    output = tmp_path / "output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "12월"
    sheet.append([
        "사번",
        "성명",
        "주민번호",
        "A01",
        "A02",
        "A03",
        "A18",
        "BF01",
        "지급총액",
        "공제총액",
        "실지급액",
    ])
    sheet.append([
        None,
        None,
        None,
        "기본급",
        "연장근로수당",
        None,
        None,
        "소득세",
        None,
        None,
        None,
    ])
    sheet.append([
        "001",
        "첫직원",
        "900101-1234567",
        1_000_000,
        100_000,
        0,
        0,
        10_000,
        1_100_000,
        10_000,
        1_090_000,
    ])
    sheet.append([
        None,
        None,
        None,
        "기본근무수당",
        "평일초과근무수당",
        "연장근로 수당",
        "성과급",
        "소득세",
        None,
        None,
        None,
    ])
    sheet.append([
        "002",
        "둘째직원",
        "900102-1234567",
        1_200_000,
        200_000,
        50_000,
        300_000,
        20_000,
        1_750_000,
        20_000,
        1_730_000,
    ])
    sheet.append(["조회건수:", "2건", None, None, None, None, None, None, None, None, None])
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert result.sheet_row_counts == {"12월": 2}
    assert "기본급" in result.payment_headers
    assert "기본근무수당" in result.payment_headers
    assert "평일초과근무수당" in result.payment_headers
    assert "연장근로수당" in result.payment_headers
    assert "성과급" in result.payment_headers
    assert "A18(지급)" not in result.payment_headers
    converted = load_workbook(output, read_only=True, data_only=True)
    try:
        target = converted["12월"]
        headers = [cell.value for cell in target[1]]
        rows = [dict(zip(headers, values)) for values in target.iter_rows(min_row=2, values_only=True)]
        assert len(rows) == 2
        assert rows[0]["기본급"] == 1_000_000
        assert rows[0]["연장근로수당"] == 100_000
        assert rows[1]["기본근무수당"] == 1_200_000
        assert rows[1]["평일초과근무수당"] == 200_000
        assert rows[1]["연장근로수당"] == 50_000
        assert rows[1]["성과급"] == 300_000
    finally:
        converted.close()


def test_active_unnamed_columns_use_stable_source_locations(tmp_path):
    source = tmp_path / "unnamed_headers.xlsx"
    output = tmp_path / "output.xlsx"
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "202501"
    sheet.append(["사번", "성명", "주민번호", None, "지급총액", None, "공제총액", "실지급액"])
    sheet.append(["001", "홍길동", "900101-1234567", 100_000, 100_000, 5_000, 5_000, 95_000])
    workbook.save(source)

    result = PayrollConverter().convert_excel_workbook(source, output)

    assert result.payment_headers == ["미확정_원본열_D(지급)"]
    assert result.deduction_headers == ["미확정_원본열_F(공제)"]
    assert result.unknown_columns == {
        "미확정_원본열_D(지급)": "payment",
        "미확정_원본열_F(공제)": "deduction",
    }
