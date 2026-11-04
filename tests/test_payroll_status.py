from pathlib import Path

from openpyxl import Workbook, load_workbook

from payroll_amendment.payroll_status import (
    AutomaticItemRules,
    ItemDecision,
    _ordered_status_headers,
    _requires_regularity_check,
    _write_sheet,
    generate_payroll_status_files,
    update_basic_info_item_markers,
)


def _make_rules(path: Path) -> None:
    book = Workbook()
    direct = book.active
    direct.title = "칼럼매칭"
    direct.append(["WEHAGO 항목명", "표준 항목명", "판정구분"])
    direct.append(["기본급", "기본급", "정액급여"])
    direct.append(["연장수당", "연장수당", "수당검토"])
    keywords = book.create_sheet("자동판단키워드")
    keywords.append(["키워드", "판정구분"])
    keywords.append(["연장", "수당검토"])
    exclusions = book.create_sheet("제외규칙")
    exclusions.append(["제외 키워드", "제외 사유"])
    exclusions.append(["상여", "상여 제외"])
    book.save(path)


def _make_basic(
    path: Path,
    *,
    include_2022: bool = False,
    include_second_employee: bool = False,
    occupation_code: str = "9309",
    index: object = 1,
) -> None:
    book = Workbook()
    business = book.active
    business.title = "사업장정보"
    business["K2"] = "25"
    business["J5"] = "기본급"
    business["P5"] = "연장수당"
    occupation = book.create_sheet("직업분류코드")
    occupation.append(["주민등록번호", "성명", "직업분류 코드", "직업분류 항목", "입사일", "퇴사일"])
    occupation.append(["900101-1******", "홍길동", occupation_code, "단순 종사자", "2020-01-02", ""])
    health = book.create_sheet("건보정보")
    health.append(["주민등록번호", "성명", "자격취득일", "자격상실일"])
    health.append(["900101-1******", "홍길동", "2019-12-15", "2024-01-01"])
    years = [2023, 2024]
    if include_2022:
        years.insert(0, 2022)
    for year in years:
        sheet = book.create_sheet(f"원천징수추출({str(year)[-2:]}년)")
        sheet.append([
            "기간", "사원번호", "성명", "주민등록번호", "인덱스",
            "과세소득_총계", "비과세소득_총계", "과세소득_현계", "비과세소득_현계",
            "근로소득", "결정세액_소득세", "기납부세액_소득세", "차감징수세액_소득세",
            "결정세액_지방세", "기납부세액_지방세", "차감징수세액_지방세",
            "18-1. 야간근로수당_총계", "18-1. 야간근로수당_현계",
        ])
        sheet.append([
            str(year), "9001011234567", "홍길동", "9001011234567", index,
            20_000_000, 1_000_000, 20_000_000, 1_000_000,
            550_000, 100_000, 120_000, -20_000, 10_000, 12_000, -2_000,
            240_000, 120_000,
        ])
        if include_second_employee:
            sheet.append([
                str(year), "9101011234567", "김연결", "9101011234567", 2,
                20_000_000, 0, 20_000_000, 0,
                650_000, 200_000, 230_000, -30_000, 20_000, 23_000, -3_000,
                0, 0,
            ])
    book.save(path)


def _make_payroll(
    path: Path,
    *,
    include_unknown_employee: bool = False,
    unknown_amount: int = 0,
    employment_dates: tuple[str, str] | None = None,
) -> None:
    book = Workbook()
    for month in range(1, 13):
        sheet = book.active if month == 1 else book.create_sheet()
        sheet.title = f"{month:02d}월"
        date_headers = ["입사일", "퇴사일"] if employment_dates else []
        date_values = list(employment_dates) if employment_dates else []
        sheet.append(["사원번호", "성명", "주민등록번호", *date_headers, "급여구분", "기본급", "연장수당", "기타수당", "지급합계", "소득세", "공제합계", "차인지급액"])
        sheet.append(["001", "홍길동", "", *date_values, "급여", 2_000_000, 100_000, unknown_amount, 2_100_000 + unknown_amount, 10_000, 10_000, 2_090_000 + unknown_amount])
        if include_unknown_employee:
            sheet.append(["002", "미연결", "", "급여", 2_000_000, 0, 0, 2_000_000, 0, 0, 2_000_000])
    book.save(path)


def test_item_rules_precedence(tmp_path):
    rules_path = tmp_path / "rules.xlsx"
    _make_rules(rules_path)
    rules = AutomaticItemRules.load(rules_path)
    assert rules.classify("기본급").category == "정액급여"
    assert rules.classify("연장수당").category == "수당검토"
    assert rules.classify("특별상여").category == "제외"
    assert rules.classify("식대보조금").category == "정액급여"
    assert rules.classify("현장수당").category == "미분류"


def test_status_payment_headers_are_ordered_by_category():
    headers = [
        "사원번호", "성명", "주민등록번호", "입사일", "퇴사일", "지급년월", "급여구분",
        "기타수당", "상여금", "연장수당", "식대", "기본급", "지급합계", "공제합계", "차인지급액",
    ]
    payment_headers = ["기타수당", "상여금", "연장수당", "식대", "기본급"]
    decisions = [
        ItemDecision("기타수당", "미분류", ""),
        ItemDecision("상여금", "제외", ""),
        ItemDecision("연장수당", "수당검토", ""),
        ItemDecision("식대", "정액급여", ""),
        ItemDecision("기본급", "정액급여", ""),
    ]

    ordered = _ordered_status_headers(headers, payment_headers, decisions)

    assert ordered[7:12] == ["기본급", "식대", "연장수당", "상여금", "기타수당"]


def test_status_payment_header_colors_follow_category():
    book = Workbook()
    sheet = book.active
    headers = ["사원번호", "성명", "기본급", "연장수당", "상여금"]

    _write_sheet(
        sheet,
        headers,
        [],
        set(),
        {"기본급": "정액급여", "연장수당": "수당검토", "상여금": "제외"},
    )

    assert sheet["C1"].fill.fgColor.rgb == "FFD9EAD3"
    assert sheet["D1"].fill.fgColor.rgb == "FFFFC7CE"
    assert sheet["E1"].fill.fgColor.rgb == "FFD9D9D9"
    book.close()


def test_direct_mapping_overrides_exclusion_rule(tmp_path):
    rules_path = tmp_path / "rules.xlsx"
    _make_rules(rules_path)
    workbook = load_workbook(rules_path)
    workbook["칼럼매칭"].append(["고정연장수당", "고정연장수당", "수당검토"])
    workbook["제외규칙"].append(["고정연장", "고정·포괄 시간외수당은 자동 조정 제외"])
    workbook.save(rules_path)
    workbook.close()

    decision = AutomaticItemRules.load(rules_path).classify("고정연장수당")

    assert decision.category == "수당검토"
    assert decision.reason == "직접 매칭: 고정연장수당"


def test_weekly_holiday_allowance_is_not_treated_as_holiday_work(tmp_path):
    rules_path = tmp_path / "rules.xlsx"
    _make_rules(rules_path)
    workbook = load_workbook(rules_path)
    workbook["자동판단키워드"].append(["휴일", "수당검토"])
    workbook.save(rules_path)
    workbook.close()

    rules = AutomaticItemRules.load(rules_path)

    assert rules.classify("주휴수당").category == "미분류"
    assert rules.classify("휴일연장야간").category == "수당검토"


def test_item_rules_distinguish_parenthetical_non_tax_name(tmp_path):
    rules_path = tmp_path / "rules.xlsx"
    _make_rules(rules_path)
    workbook = load_workbook(rules_path)
    workbook["칼럼매칭"].append(["야간연장수당(비과세)", "야간연장수당(비과세)", "제외"])
    workbook["칼럼매칭"].append(["야간연장수당", "야간연장수당", "수당검토"])
    workbook.save(rules_path)
    workbook.close()

    rules = AutomaticItemRules.load(rules_path)

    assert rules.classify("야간연장수당").category == "수당검토"
    assert rules.classify("야간연장수당(비과세)").category == "제외"


def test_updates_basic_info_item_markers_from_rules(tmp_path):
    rules_path = tmp_path / "rules.xlsx"
    basic_path = tmp_path / "기초정보.xlsx"
    _make_rules(rules_path)
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "사업장정보"
    sheet.cell(4, 10, "지급항목")
    sheet.cell(4, 11, "정액급여")
    sheet.cell(4, 12, "수당검토")
    sheet.cell(4, 16, "연도별 지급항목")
    sheet.cell(4, 17, "정액급여")
    sheet.cell(4, 18, "수당검토")
    sheet.cell(5, 10, "기본급")
    sheet.cell(6, 10, "연장수당")
    sheet.cell(7, 10, "특별상여")
    sheet.cell(5, 16, "기본급")
    sheet.cell(6, 16, "연장수당")
    workbook.save(basic_path)
    workbook.close()

    counts = update_basic_info_item_markers(basic_path, AutomaticItemRules.load(rules_path))

    assert counts == (2, 2)
    workbook = load_workbook(basic_path, read_only=True, data_only=True)
    try:
        sheet = workbook["사업장정보"]
        assert (sheet["K5"].value, sheet["L5"].value) == ("O", None)
        assert (sheet["K6"].value, sheet["L6"].value) == (None, "O")
        assert (sheet["K7"].value, sheet["L7"].value) == (None, None)
        assert (sheet["Q5"].value, sheet["R5"].value) == ("O", None)
        assert (sheet["Q6"].value, sheet["R6"].value) == (None, "O")
    finally:
        workbook.close()


def test_direct_fixed_allowance_marker_ignores_monthly_regularity(tmp_path):
    rules_path = tmp_path / "rules.xlsx"
    basic_path = tmp_path / "기초정보.xlsx"
    payroll = tmp_path / "payroll"
    payroll.mkdir()
    _make_rules(rules_path)
    rules_book = load_workbook(rules_path)
    rules_book["칼럼매칭"].append(["직책수당", "직책수당", "정액급여"])
    rules_book.save(rules_path)
    rules_book.close()

    basic = Workbook()
    sheet = basic.active
    sheet.title = "사업장정보"
    sheet["J5"] = "직책수당"
    basic.save(basic_path)
    basic.close()

    payroll_path = payroll / "테스트_23년 급여대장.xlsx"
    _make_payroll(payroll_path)
    payroll_book = load_workbook(payroll_path)
    for sheet in payroll_book.worksheets:
        sheet.insert_cols(8)
        sheet.cell(1, 8, "직책수당")
        sheet.cell(2, 8, 100_000)
    payroll_book.save(payroll_path)
    payroll_book.close()

    rules = AutomaticItemRules.load(rules_path)
    update_basic_info_item_markers(basic_path, rules, payroll)
    result = load_workbook(basic_path, read_only=True, data_only=True)
    assert result["사업장정보"]["K5"].value == "O"
    result.close()

    payroll_book = load_workbook(payroll_path)
    payroll_book["06월"].cell(2, 8, 200_000)
    payroll_book.save(payroll_path)
    payroll_book.close()

    update_basic_info_item_markers(basic_path, rules, payroll)
    result = load_workbook(basic_path, read_only=True, data_only=True)
    assert result["사업장정보"]["K5"].value == "O"
    result.close()


def test_meal_named_allowance_marker_does_not_require_monthly_regularity(tmp_path):
    rules_path = tmp_path / "rules.xlsx"
    basic_path = tmp_path / "기초정보.xlsx"
    payroll = tmp_path / "payroll"
    payroll.mkdir()
    _make_rules(rules_path)
    rules_book = load_workbook(rules_path)
    rules_book["자동판단키워드"].append(["식대", "정액급여"])
    rules_book["자동판단키워드"].append(["식사대", "정액급여"])
    rules_book.save(rules_path)
    rules_book.close()

    basic = Workbook()
    sheet = basic.active
    sheet.title = "사업장정보"
    sheet["J5"] = "식대보조금"
    basic.save(basic_path)
    basic.close()

    payroll_path = payroll / "테스트_23년 급여대장.xlsx"
    _make_payroll(payroll_path)
    payroll_book = load_workbook(payroll_path)
    for month, sheet in enumerate(payroll_book.worksheets, 1):
        sheet.insert_cols(8)
        sheet.cell(1, 8, "식대보조금")
        sheet.cell(2, 8, 0 if month == 6 else month * 10_000)
    payroll_book.save(payroll_path)
    payroll_book.close()

    update_basic_info_item_markers(
        basic_path,
        AutomaticItemRules.load(rules_path),
        payroll,
    )

    result = load_workbook(basic_path, read_only=True, data_only=True)
    assert result["사업장정보"]["K5"].value == "O"
    result.close()
    assert all(
        not _requires_regularity_check(item_name)
        for item_name in ("식대", "식대보조", "식대금", "식대보조금", "식사대", "식사대지원")
    )


def test_variable_keyword_fixed_allowance_is_excluded_from_monthly_fixed_total(tmp_path):
    rules_path = tmp_path / "rules.xlsx"
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    payroll.mkdir()
    _make_rules(rules_path)
    rules_book = load_workbook(rules_path)
    rules_book["자동판단키워드"].append(["근무조정", "정액급여"])
    rules_book.save(rules_path)
    rules_book.close()
    _make_basic(basic)
    payroll_path = payroll / "테스트_23년 급여대장.xlsx"
    _make_payroll(payroll_path)
    payroll_book = load_workbook(payroll_path)
    for sheet in payroll_book.worksheets:
        sheet.insert_cols(8)
        sheet.cell(1, 8, "근무조정수당")
        sheet.cell(2, 8, 100_000)
    payroll_book["06월"].cell(2, 8, 200_000)
    payroll_book.save(payroll_path)
    payroll_book.close()

    result = generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules_path,
        output_directory=output,
    )

    status = load_workbook(output / "테스트_23년 급상여현황.xlsx", read_only=True, data_only=True)
    try:
        sheet = status["대상자"]
        headers = [cell.value for cell in sheet[1]]
        first = dict(zip(headers, next(sheet.iter_rows(min_row=2, values_only=True))))
        assert first["정액급여합계"] == 2_000_000
    finally:
        status.close()
    report = result.report_path.read_text(encoding="utf-8-sig")
    assert "변동·부정기" in report
    assert "근무조정수당" in report


def test_direct_fixed_allowance_is_included_despite_monthly_variation(tmp_path):
    rules_path = tmp_path / "rules.xlsx"
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    payroll.mkdir()
    _make_rules(rules_path)
    rules_book = load_workbook(rules_path)
    rules_book["칼럼매칭"].append(["직책수당", "직책수당", "정액급여"])
    rules_book.save(rules_path)
    rules_book.close()
    _make_basic(basic)
    payroll_path = payroll / "테스트_23년 급여대장.xlsx"
    _make_payroll(payroll_path)
    payroll_book = load_workbook(payroll_path)
    for sheet in payroll_book.worksheets:
        sheet.insert_cols(8)
        sheet.cell(1, 8, "직책수당")
        sheet.cell(2, 8, 100_000)
    payroll_book["06월"].cell(2, 8, 200_000)
    payroll_book.save(payroll_path)
    payroll_book.close()

    result = generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules_path,
        output_directory=output,
    )

    status = load_workbook(output / "테스트_23년 급상여현황.xlsx", read_only=True, data_only=True)
    try:
        sheet = status["전체인원"]
        headers = [cell.value for cell in sheet[1]]
        rows = [dict(zip(headers, row)) for row in sheet.iter_rows(min_row=2, values_only=True)]
        assert rows[0]["정액급여합계"] == 2_100_000
        assert rows[5]["정액급여합계"] == 2_200_000
    finally:
        status.close()
    report = result.report_path.read_text(encoding="utf-8-sig")
    assert '"item_name": "직책수당", "category": "변동·부정기"' not in report


def test_regular_fixed_allowances_are_included_in_monthly_fixed_total(tmp_path):
    rules_path = tmp_path / "rules.xlsx"
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    payroll.mkdir()
    _make_rules(rules_path)
    rules_book = load_workbook(rules_path)
    rules_book["칼럼매칭"].append(["직책수당", "직책수당", "정액급여"])
    rules_book["칼럼매칭"].append(["기본근무수당", "기본근무수당", "정액급여"])
    rules_book.save(rules_path)
    rules_book.close()
    _make_basic(basic)
    payroll_path = payroll / "테스트_23년 급여대장.xlsx"
    _make_payroll(payroll_path)
    payroll_book = load_workbook(payroll_path)
    for sheet in payroll_book.worksheets:
        sheet.insert_cols(8, 3)
        sheet.cell(1, 8, "식대보조금")
        sheet.cell(1, 9, "직책수당")
        sheet.cell(1, 10, "기본근무수당")
        sheet.cell(2, 8, 200_000)
        sheet.cell(2, 9, 100_000)
        sheet.cell(2, 10, 150_000)
        headers = [cell.value for cell in sheet[1]]
        sheet.cell(2, headers.index("지급합계") + 1, 2_550_000)
        sheet.cell(2, headers.index("차인지급액") + 1, 2_540_000)
    payroll_book.save(payroll_path)
    payroll_book.close()

    generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules_path,
        output_directory=output,
    )

    status = load_workbook(output / "테스트_23년 급상여현황.xlsx", data_only=True)
    try:
        for sheet in status.worksheets:
            headers = [cell.value for cell in sheet[1]]
            assert headers[7:13] == [
                "기본급",
                "식대보조금",
                "직책수당",
                "기본근무수당",
                "연장수당",
                "기타수당",
            ]
            review_cell = sheet.cell(1, headers.index("연장수당") + 1)
            assert review_cell.fill.fgColor.rgb == "FFFFC7CE"
        sheet = status["전체인원"]
        headers = [cell.value for cell in sheet[1]]
        first = dict(zip(headers, next(sheet.iter_rows(min_row=2, values_only=True))))
        assert first["정액급여합계"] == 2_450_000
    finally:
        status.close()


def test_generates_status_and_updates_detected_years(tmp_path, caplog):
    caplog.set_level("INFO")
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    payroll.mkdir()
    rules = tmp_path / "rules.xlsx"
    _make_basic(basic)
    _make_rules(rules)
    _make_payroll(
        payroll / "테스트_23년 급여대장.xlsx",
        employment_dates=("2018-03-04", "2023-05-06"),
    )
    _make_payroll(payroll / "테스트_24년 급여대장.xlsx")

    result = generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules,
        output_directory=output,
    )

    assert result.years == (2023, 2024)
    for year in result.years:
        assert f"{year}년 급상여현황 생성 시작" in caplog.messages
        assert any(message.startswith(f"{year}년 급상여현황 생성 완료:") for message in caplog.messages)
    basic_book = load_workbook(basic, data_only=True, read_only=True)
    assert basic_book["사업장정보"]["K2"].value == "23,24"
    assert basic_book["사업장정보"]["K5"].value == "O"
    assert basic_book["사업장정보"]["R5"].value == "O"
    basic_book.close()
    status_book = load_workbook(output / "테스트_23년 급상여현황.xlsx", data_only=True, read_only=True)
    assert status_book.sheetnames == ["전체인원", "대상자(작업)", "대상자", "비대상자"]
    headers = [cell.value for cell in status_book["대상자(작업)"][1]]
    assert headers[:7] == ["사원번호", "성명", "주민등록번호", "입사일", "퇴사일", "지급년월", "급여구분"]
    assert headers[headers.index("차인지급액") + 1:headers.index("차인지급액") + 5] == ["정액급여합계", "인덱스", "직업분류 코드", "직업분류 항목"]
    assert "18-1. 야간근로수당_총계" in headers
    assert headers[-1] == "경정청구 대상 여부"
    non_tax_column = headers.index("야간연장수당(비과세)") + 1
    assert non_tax_column == headers.index("지급합계")
    assert status_book["대상자(작업)"].cell(1, non_tax_column).font.color.rgb == "FFFF0000"
    assert status_book["대상자(작업)"].cell(2, non_tax_column).font.color.rgb == "FFFF0000"
    assert status_book["대상자(작업" if False else "대상자(작업)"].max_row == 14
    rows = list(status_book["대상자"].iter_rows(min_row=2, values_only=True))
    first = dict(zip(headers, rows[0]))
    annual = dict(zip(headers, rows[-1]))
    resident_column = headers.index("주민등록번호") + 1
    assert status_book["대상자"].cell(2, resident_column).number_format == "@"
    assert first["입사일"] == "2018-03-04"
    assert first["퇴사일"] == "2023-05-06"
    assert annual["입사일"] == "2018-03-04"
    assert annual["퇴사일"] == "2023-05-06"
    assert first["정액급여합계"] == 2_000_000
    assert first["경정청구 대상 여부"] == "대상"
    assert annual["인덱스"] == "1"
    assert annual["직업분류 코드"] == 9309
    assert annual["18-1. 야간근로수당_총계"] == 240_000
    assert annual["근로소득"] == 550_000
    assert annual["결정세액_소득세"] == 100_000
    assert annual["기납부세액_소득세"] == 120_000
    assert annual["차감징수세액_소득세"] == -20_000
    assert annual["결정세액_지방세"] == 10_000
    assert annual["기납부세액_지방세"] == 12_000
    assert annual["차감징수세액_지방세"] == -2_000
    assert annual["경정청구 대상 여부"] == "대상"
    status_book.close()
    status_book = load_workbook(output / "테스트_24년 급상여현황.xlsx", data_only=True, read_only=True)
    headers = [cell.value for cell in status_book["대상자"][1]]
    first = dict(zip(headers, next(status_book["대상자"].iter_rows(min_row=2, values_only=True))))
    assert first["입사일"] is None
    assert first["퇴사일"] is None
    status_book.close()


def test_same_name_tax_records_match_annual_rows_by_employee_code(tmp_path):
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    rules = tmp_path / "rules.xlsx"
    payroll.mkdir()
    _make_basic(basic, include_second_employee=True)
    _make_rules(rules)
    basic_book = load_workbook(basic)
    for year in (2023, 2024):
        basic_book[f"원천징수추출({str(year)[-2:]}년)"].cell(3, 3, "홍길동")
    basic_book.save(basic)
    basic_book.close()

    payroll_book = Workbook()
    for month in range(1, 13):
        sheet = payroll_book.active if month == 1 else payroll_book.create_sheet()
        sheet.title = f"{month:02d}월"
        sheet.append([
            "사원번호", "성명", "주민등록번호", "급여구분", "기본급", "연장수당",
            "지급합계", "소득세", "공제합계", "차인지급액",
        ])
        sheet.append(["9001011234567", "홍길동", "", "급여", 2_000_000, 100_000, 2_100_000, 10_000, 10_000, 2_090_000])
        sheet.append(["9101011234567", "홍길동", "", "급여", 3_000_000, 100_000, 3_100_000, 20_000, 20_000, 3_080_000])
    payroll_book.save(payroll / "테스트_23년 급여대장.xlsx")
    payroll_book.close()

    generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules,
        output_directory=output,
    )

    status = load_workbook(output / "테스트_23년 급상여현황.xlsx", read_only=True, data_only=True)
    try:
        sheet = status["전체인원"]
        headers = [cell.value for cell in sheet[1]]
        annual_rows = {
            row[headers.index("사원번호")]: dict(zip(headers, row))
            for row in sheet.iter_rows(min_row=2, values_only=True)
            if str(row[headers.index("지급년월")] or "").endswith("년합계")
        }
        assert annual_rows["9001011234567"]["주민등록번호"] == "900101-1234567"
        assert annual_rows["9101011234567"]["주민등록번호"] == "910101-1234567"
        assert annual_rows["9001011234567"]["인덱스"] == "1"
        assert annual_rows["9101011234567"]["인덱스"] == "2"
        assert annual_rows["9001011234567"]["근로소득"] == 550_000
        assert annual_rows["9101011234567"]["근로소득"] == 650_000
        assert annual_rows["9001011234567"]["결정세액_소득세"] == 100_000
        assert annual_rows["9101011234567"]["결정세액_소득세"] == 200_000
    finally:
        status.close()


def test_duplicate_tax_records_for_same_identity_are_excluded(tmp_path):
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    rules = tmp_path / "rules.xlsx"
    payroll.mkdir()
    _make_basic(basic)
    _make_rules(rules)
    payroll_path = payroll / "테스트_23년 급여대장.xlsx"
    _make_payroll(payroll_path)
    payroll_book = load_workbook(payroll_path)
    try:
        for sheet in payroll_book.worksheets:
            sheet.cell(2, 3, "9001011234567")
        payroll_book.save(payroll_path)
    finally:
        payroll_book.close()
    basic_book = load_workbook(basic)
    try:
        sheet = basic_book["원천징수추출(23년)"]
        duplicate = [cell.value for cell in sheet[2]]
        duplicate[4] = 99
        sheet.append(duplicate)
        basic_book.save(basic)
    finally:
        basic_book.close()

    result = generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules,
        output_directory=output,
    )

    status = load_workbook(output / "테스트_23년 급상여현황.xlsx", read_only=True, data_only=True)
    try:
        assert status["대상자"].max_row == 1
        assert {row[1] for row in status["비대상자"].iter_rows(min_row=2, values_only=True)} == {"홍길동"}
    finally:
        status.close()
    assert "원천징수 자료 동일 주민번호·성명 중복" in result.report_path.read_text(encoding="utf-8-sig")


def test_resident_mismatch_does_not_fall_back_to_same_name(tmp_path):
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    rules = tmp_path / "rules.xlsx"
    payroll.mkdir()
    _make_basic(basic)
    _make_rules(rules)
    payroll_path = payroll / "테스트_23년 급여대장.xlsx"
    _make_payroll(payroll_path)
    payroll_book = load_workbook(payroll_path)
    try:
        for sheet in payroll_book.worksheets:
            sheet.cell(2, 3, "9901011234567")
        payroll_book.save(payroll_path)
    finally:
        payroll_book.close()

    result = generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules,
        output_directory=output,
    )

    status = load_workbook(output / "테스트_23년 급상여현황.xlsx", read_only=True, data_only=True)
    try:
        assert status["대상자"].max_row == 1
        annual = next(
            row
            for row in status["비대상자"].iter_rows(min_row=2, values_only=True)
            if str(row[5]).endswith("년합계")
        )
        headers = [cell.value for cell in status["비대상자"][1]]
        values = dict(zip(headers, annual))
        assert values["주민등록번호"] == "990101-1234567"
        assert values["인덱스"] is None
    finally:
        status.close()
    assert "원천징수 자료 주민번호·성명 미연결" in result.report_path.read_text(encoding="utf-8-sig")


def test_rehire_split_across_employee_codes_is_excluded(tmp_path):
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    rules = tmp_path / "rules.xlsx"
    payroll.mkdir()
    _make_basic(basic)
    _make_rules(rules)
    payroll_path = payroll / "테스트_23년 급여대장.xlsx"
    _make_payroll(payroll_path)
    payroll_book = load_workbook(payroll_path)
    try:
        for sheet in payroll_book.worksheets:
            sheet.cell(2, 3, "9001011234567")
            second = [cell.value for cell in sheet[2]]
            second[0] = "002"
            sheet.append(second)
        payroll_book.save(payroll_path)
    finally:
        payroll_book.close()

    result = generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules,
        output_directory=output,
    )

    status = load_workbook(output / "테스트_23년 급상여현황.xlsx", read_only=True, data_only=True)
    try:
        assert status["대상자"].max_row == 1
        non_target_codes = {row[0] for row in status["비대상자"].iter_rows(min_row=2, values_only=True)}
        assert non_target_codes == {"001", "002"}
    finally:
        status.close()
    assert "동일 주민번호·성명이 복수 사원번호로 분리됨" in result.report_path.read_text(encoding="utf-8-sig")


def test_explicit_non_target_occupation_is_excluded_and_blank_is_provisionally_eligible(tmp_path):
    rules = tmp_path / "rules.xlsx"
    _make_rules(rules)
    for occupation_code, expected_sheet, expected_status in (
        ("경정청구대상아님", "비대상자", "작업대상이 아님"),
        ("", "대상자", "대상"),
    ):
        case = tmp_path / ("blank" if not occupation_code else "excluded")
        case.mkdir()
        basic = case / "테스트_기초정보.xlsx"
        payroll = case / "payroll"
        output = case / "output"
        payroll.mkdir()
        _make_basic(basic, occupation_code=occupation_code)
        _make_payroll(payroll / "테스트_23년 급여대장.xlsx")

        generate_payroll_status_files(
            company_name="테스트",
            basic_info_path=basic,
            payroll_directory=payroll,
            item_rules_path=rules,
            output_directory=output,
        )

        book = load_workbook(output / "테스트_23년 급상여현황.xlsx", data_only=True, read_only=True)
        try:
            sheet = book[expected_sheet]
            headers = [cell.value for cell in sheet[1]]
            values = dict(zip(headers, next(sheet.iter_rows(min_row=2, values_only=True))))
            assert values["경정청구 대상 여부"] == expected_status
        finally:
            book.close()


def test_employee_without_index_is_always_excluded_from_targets(tmp_path):
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    rules = tmp_path / "rules.xlsx"
    payroll.mkdir()
    _make_basic(basic, index="")
    _make_rules(rules)
    _make_payroll(payroll / "테스트_23년 급여대장.xlsx")

    result = generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules,
        output_directory=output,
    )

    book = load_workbook(output / "테스트_23년 급상여현황.xlsx", data_only=True, read_only=True)
    try:
        assert book["대상자"].max_row == 1
        non_target_names = {row[1] for row in book["비대상자"].iter_rows(min_row=2, values_only=True)}
        assert non_target_names == {"홍길동"}
    finally:
        book.close()
    assert "원천징수 자료 인덱스 없음" in result.report_path.read_text(encoding="utf-8-sig")


def test_employee_without_current_tax_record_is_excluded_even_when_allowance_qualifies(tmp_path):
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    rules = tmp_path / "rules.xlsx"
    payroll.mkdir()
    _make_basic(basic)
    _make_rules(rules)
    payroll_path = payroll / "테스트_23년 급여대장.xlsx"
    _make_payroll(payroll_path, include_unknown_employee=True)
    payroll_book = load_workbook(payroll_path)
    try:
        for sheet in payroll_book.worksheets:
            sheet.cell(3, 6, 100_000)
            sheet.cell(3, 8, 2_100_000)
            sheet.cell(3, 11, 2_100_000)
        payroll_book.save(payroll_path)
    finally:
        payroll_book.close()

    generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules,
        output_directory=output,
    )

    book = load_workbook(output / "테스트_23년 급상여현황.xlsx", data_only=True, read_only=True)
    try:
        target_names = {row[1] for row in book["대상자"].iter_rows(min_row=2, values_only=True)}
        non_target_names = {row[1] for row in book["비대상자"].iter_rows(min_row=2, values_only=True)}
        assert "미연결" not in target_names
        assert "미연결" in non_target_names
    finally:
        book.close()


def test_zero_pension_and_health_insurance_excludes_only_that_month(tmp_path):
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    rules = tmp_path / "rules.xlsx"
    payroll.mkdir()
    _make_basic(basic)
    _make_rules(rules)
    payroll_path = payroll / "테스트_23년 급여대장.xlsx"
    _make_payroll(payroll_path)

    payroll_book = load_workbook(payroll_path)
    try:
        for month, sheet in enumerate(payroll_book.worksheets, 1):
            sheet.insert_cols(9, 2)
            sheet.cell(1, 9, "국민연금")
            sheet.cell(1, 10, "건강보험")
            insurance = 0 if month <= 3 else 100_000
            sheet.cell(2, 9, insurance)
            sheet.cell(2, 10, insurance)
            sheet.cell(2, 12, 10_000 + insurance * 2)
            sheet.cell(2, 13, 2_090_000 - insurance * 2)
        payroll_book.save(payroll_path)
    finally:
        payroll_book.close()

    generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules,
        output_directory=output,
    )

    book = load_workbook(output / "테스트_23년 급상여현황.xlsx", data_only=True)
    try:
        sheet = book["대상자(작업)"]
        headers = [cell.value for cell in sheet[1]]
        rows = {
            row[headers.index("지급년월")]: dict(zip(headers, row))
            for row in sheet.iter_rows(min_row=2, values_only=True)
            if str(row[headers.index("지급년월")]).endswith("월")
        }
        assert rows["23년03월"]["연장수당"] == 100_000
        assert rows["23년03월"]["야간연장수당(비과세)"] == 0
        assert rows["23년04월"]["연장수당"] == 0
        assert rows["23년04월"]["야간연장수당(비과세)"] == 100_000
    finally:
        book.close()


def test_unmatched_employee_is_non_target_and_reported(tmp_path):
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    payroll.mkdir()
    rules = tmp_path / "rules.xlsx"
    _make_basic(basic)
    _make_rules(rules)
    _make_payroll(payroll / "테스트_23년 급여대장.xlsx", include_unknown_employee=True)

    result = generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules,
        output_directory=output,
    )

    report = result.report_path.read_text(encoding="utf-8-sig")
    assert "원천징수 자료 미연결" in report
    book = load_workbook(output / "테스트_23년 급상여현황.xlsx", data_only=True, read_only=True)
    non_target_names = {row[1] for row in book["비대상자"].iter_rows(min_row=2, values_only=True)}
    assert "미연결" in non_target_names
    book.close()


def test_unclassified_item_is_reported_without_forcing_non_target(tmp_path):
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    payroll.mkdir()
    rules = tmp_path / "rules.xlsx"
    _make_basic(basic)
    _make_rules(rules)
    _make_payroll(payroll / "테스트_23년 급여대장.xlsx", unknown_amount=50_000)

    result = generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules,
        output_directory=output,
    )

    report = result.report_path.read_text(encoding="utf-8-sig")
    assert "기타수당" in report
    book = load_workbook(output / "테스트_23년 급상여현황.xlsx", data_only=True, read_only=True)
    target_names = {row[1] for row in book["대상자(작업)"].iter_rows(min_row=2, values_only=True)}
    assert "홍길동" in target_names
    book.close()


def test_status_removes_only_inactive_unresolved_columns(tmp_path):
    basic = tmp_path / "테스트_기초정보.xlsx"
    payroll = tmp_path / "payroll"
    output = tmp_path / "output"
    payroll.mkdir()
    rules = tmp_path / "rules.xlsx"
    _make_basic(basic)
    _make_rules(rules)
    payroll_path = payroll / "테스트_23년 급여대장.xlsx"
    _make_payroll(payroll_path)

    workbook = load_workbook(payroll_path)
    try:
        for sheet in workbook.worksheets:
            sheet.insert_cols(8, 3)
            sheet.cell(1, 8, "기술수당")
            sheet.cell(1, 9, "A12(지급)")
            sheet.cell(1, 10, "A13(지급)")
            sheet.cell(2, 8, 50_000)
            sheet.cell(2, 9, 0)
            sheet.cell(2, 10, 25_000)
            sheet.cell(2, 11, sheet.cell(2, 11).value + 75_000)
            sheet.cell(2, 14, sheet.cell(2, 14).value + 75_000)
        workbook.save(payroll_path)
    finally:
        workbook.close()

    generate_payroll_status_files(
        company_name="테스트",
        basic_info_path=basic,
        payroll_directory=payroll,
        item_rules_path=rules,
        output_directory=output,
    )

    status = load_workbook(
        output / "테스트_23년 급상여현황.xlsx",
        read_only=True,
        data_only=True,
    )
    try:
        for sheet_name in status.sheetnames:
            headers = [cell.value for cell in status[sheet_name][1]]
            assert "A12(지급)" not in headers
            assert "기술수당" in headers
            assert "A13(지급)" in headers
    finally:
        status.close()
