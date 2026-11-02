from pathlib import Path

from openpyxl import Workbook, load_workbook

from payroll_amendment.payroll_converter import PayrollConverter, _payroll_workbook_year


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
