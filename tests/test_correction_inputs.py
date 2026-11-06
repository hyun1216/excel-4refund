from io import BytesIO
from pathlib import Path
import shutil

from openpyxl import Workbook, load_workbook

from payroll_amendment.correction_inputs import (
    _append_tax_reduction_review_issue,
    prepare_correction_inputs,
)


def _make_basic(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "직업분류코드"
    sheet.append(["주민등록번호", "성명", "취득일", "상실일"])
    sheet.append(["900101-1******", "홍길동", "2023-01-15", "2023-02-20"])
    health = workbook.create_sheet("건보정보")
    health.append(["주민등록번호", "성명", "자격취득일", "자격상실일"])
    health.append(["910101-1******", "김철수", "2023.02.01", ""])
    tax = workbook.create_sheet("원천징수추출(23년)")
    tax.append([
        "기간", "사원번호", "성명", "주민등록번호", "인덱스",
        "과세소득_총계", "비과세소득_총계", "과세소득_현계", "비과세소득_현계",
        "결정세액_소득세", "야간근로수당_총계", "야간근로수당_현계",
        "조특법 감면율", "산출세액", "세액감면 계", "근로소득", "세액공제 계",
    ])
    tax.append([
        2023, "9001011234567", "홍길동", "9001011234567", 7,
        10_000_000, 200_000, 10_000_000, 200_000, 100_000, 200_000, 200_000,
        None, 2_500_000, 2_000_000, 500_000, 500_000,
    ])
    allowance = workbook.create_sheet("수당등록")
    allowance.append(["■ 수당 등록"])
    allowance.append([])
    allowance.append(["수당명", "과세구분", "근로소득유형", "근로소득세부", "코드"])
    allowance.append([
        "야간연장수당(비과세)", "더존비과세", "연장근로",
        "18-1.생산직등야간근로수당", "O01",
    ])
    allowance.append(["비과세식사대(P01)_총계", "기존비과세"])
    allowance.append(["비과세식사대(P01)_현계", "기존비과세"])
    allowance.append(["생산직등야간근로수당(O01)_총계", "기존비과세"])
    allowance.append(["생산직등야간근로수당(O01)_현계", "기존비과세"])
    adjustment = workbook.create_sheet("수당조정(비과세)")
    adjustment.append(["■ 비과세 조정값"])
    adjustment.append([])
    adjustment.append(["구분"])
    adjustment.append([
        "기간", "연번", "사원번호", "이름", "주민번호", "인덱스", "과세차감총액",
    ])
    workbook.save(path)
    workbook.close()


def _make_status(path: Path) -> None:
    workbook = Workbook()
    target = workbook.active
    target.title = "대상자(작업)"
    headers = [
        "사원번호",
        "성명",
        "주민등록번호",
        "입사일",
        "퇴사일",
        "지급년월",
        "급여구분",
        "기본급",
        "연장수당",
        "야간연장수당(비과세)",
        "지급합계",
        "소득세",
        "선지급공제",
        "복지포인트(가지급)",
        "공제합계",
        "차인지급액",
        "직업분류 코드",
        "인덱스",
        "비과세식사대(P01)_총계",
        "비과세식사대(P01)_현계",
        "생산직등야간근로수당(O01)_총계",
        "생산직등야간근로수당(O01)_현계",
    ]
    target.append(headers)
    target.append(["9001011234567", "홍길동", "9001011234567", "2022-12-01", "2023-01-31", "23년01월", "급여", 2_000_000, 0, 100_000, 2_100_000, 10_000, 3_000, 4_000, 17_000, 2_083_000, "9309", 7, 200_000, 200_000, 200_000, 200_000])
    target.append(["9001011234567", "홍길동", "9001011234567", "2022-12-01", "2023-01-31", "23년합계", "급상여", 2_000_000, 0, 100_000, 2_100_000, 10_000, 3_000, 4_000, 17_000, 2_083_000, "9309", 7, 200_000, 200_000, 200_000, 200_000])
    non_target = workbook.create_sheet("비대상자")
    non_target.append([header for header in headers if header != "야간연장수당(비과세)"])
    non_target.append(["9101011234567", "김철수", "9101011234567", "", "", "23년02월", "급여", 3_000_000, 0, 3_000_000, 20_000, 5_000, 6_000, 31_000, 2_969_000, "사무원", None])
    workbook.save(path)
    workbook.close()


def _make_payroll(path: Path) -> None:
    workbook = Workbook()
    for index, month in enumerate(range(1, 13)):
        sheet = workbook.active if index == 0 else workbook.create_sheet()
        sheet.title = f"{month:02d}월"
        sheet.append([
            "사원번호", "성명", "주민등록번호", "급여구분", "기본급", "연장수당",
            "지급합계", "소득세", "공제합계", "차인지급액",
        ])
    workbook["01월"].append([
        "9001011234567", "홍길동", "9001011234567", "급여", 2_000_000, 0,
        2_000_000, 10_000, 10_000, 1_990_000,
    ])
    workbook["02월"].append([
        "9101011234567", "김철수", "9101011234567", "급여", 3_000_000, 0,
        3_000_000, 20_000, 20_000, 2_980_000,
    ])
    workbook.save(path)
    workbook.close()


def _copy_converter(pairs: list[tuple[Path, Path]]) -> None:
    for source, destination in pairs:
        shutil.copy2(source, destination)


def _make_tool(path: Path) -> None:
    workbook = Workbook()
    previous = workbook.active
    previous.title = "prevData"
    headers = [
        "기간", "연번", "사번", "이름", "주민등록번호", "인덱스",
        "18-1.생산직등야간근로수당", "조특법 감면율",
        "감면소득_현계", "감면소득_총계", "과세소득_총계", "비과세소득_총계",
        "과세소득_현계", "비과세소득_현계", "결정세액_소득세",
        "기납부세액_소득세", "차감징수세액_소득세", "결정세액_지방세",
        "기납부세액_지방세", "차감징수세액_지방세", "근로소득금액",
        "차감소득금액", "그밖의소득공제계", "과세표준", "세액감면 계",
        "근로소득", "세액공제 계", "비과세명1_총", "코드1_총", "금액1_총",
        "비과세명1_현", "코드1_현", "금액1_현",
    ]
    previous.append(headers)
    data = workbook.create_sheet("Data")
    data_headers = [
        *headers[:8], "급여_현계", "급여_총계", "과세소득_현계", "과세소득_총계",
        "비과세소득_현계", "비과세소득_총계", "감면소득_현계", "감면소득_총계",
        "결정세액_소득세", "결정세액_지방세", "기납부세액_소득세",
        "기납부세액_지방세", "차감징수세액_소득세", "차감징수세액_지방세",
        "총급여", "근로소득공제", "근로소득금액", "차감소득(역산)",
        "차감소득금액", "그밖의소득공제계", "과세표준", "산출세액",
        "조특법 제30조", "세액감면 계", "근로소득", "세액공제(기타)",
        "세액공제 계", "결정세액",
    ]
    data.append(data_headers)
    formulas = [None] * len(data_headers)
    formulas[6] = "=prevData!G2"
    formulas[7] = "=prevData!H2"
    formulas[29] = "=1000000"
    formulas[30] = "=IF(H2>0,AD2*H2,0)"
    formulas[31] = "=AE2"
    data.append(formulas)
    workbook.create_sheet("Form")
    workbook.create_sheet("TaxTable")
    workbook.save(path)
    workbook.close()


def test_prepares_modified_status_tools_and_monthly_ledgers(tmp_path, caplog):
    caplog.set_level("INFO")
    basic = tmp_path / "테스트_기초정보.xlsx"
    status = tmp_path / "테스트_23년 급상여현황.xlsx"
    payroll = tmp_path / "테스트_23년 급여대장.xlsx"
    tool_template = tmp_path / "Temp_수정 원천징수영수증 생성 Tool.xlsx"
    modified_root = tmp_path / "Input_급상여현황(수정)"
    correction_root = tmp_path / "2.근로소득 수정신고" / "1.Input"
    _make_basic(basic)
    _make_status(status)
    _make_payroll(payroll)
    _make_tool(tool_template)

    result = prepare_correction_inputs(
        company_name="테스트",
        basic_info_path=basic,
        payroll_files=(payroll,),
        payroll_status_files=(status,),
        tool_template_path=tool_template,
        modified_input_directory=modified_root,
        correction_input_directory=correction_root,
        xls_converter=_copy_converter,
    )

    assert result.years == (2023,)
    assert "2023년 급상여현황(수정) 생성 시작" in caplog.messages
    assert any(message.startswith("2023년 급상여현황(수정) 생성 완료:") for message in caplog.messages)
    assert result.modified_status_files[0].read_bytes() == status.read_bytes()
    tool_book = load_workbook(result.withholding_tool_files[0], data_only=False, read_only=True)
    assert tool_book["prevData"]["A2"].value == "2023년"
    assert tool_book["prevData"]["D2"].value == "홍길동"
    assert tool_book["prevData"]["G2"].value == 100_000
    assert tool_book["prevData"]["K2"].value == 10_000_000
    assert tool_book["prevData"]["Y2"].value == 2_000_000
    assert tool_book["prevData"]["AB2"].value == "야간근로수당"
    assert tool_book["prevData"]["AC2"].value == "O01"
    assert tool_book["Data"]["A2"].value == "2023년"
    assert tool_book["Data"]["G2"].value == "=prevData!G2"
    assert tool_book["Data"]["AF2"].value == "=MAX(0,MIN(AD2,IF(N(H2)>0,AE2,N(prevData!Y2))))"
    tool_book.close()
    basic_book = load_workbook(basic, data_only=False, read_only=True)
    adjustment = basic_book["수당조정(비과세)"]
    assert [adjustment.cell(5, column).value for column in range(1, 9)] == [
        "2023년", 1, "9001011234567", "홍길동", "900101-1234567", "7", -100_000, 100_000,
    ]
    adjustment_headers = [cell.value for cell in adjustment[4]]
    assert adjustment.cell(
        5,
        adjustment_headers.index("비과세식사대(P01)_총계") + 1,
    ).value == 200_000
    assert adjustment.cell(
        5,
        adjustment_headers.index("생산직등야간근로수당(O01)_총계") + 1,
    ).value == 100_000
    assert adjustment.cell(
        5,
        adjustment_headers.index("생산직등야간근로수당(O01)_현계") + 1,
    ).value == 100_000
    assert adjustment["E5"].number_format == "@"
    basic_book.close()
    assert any(issue["type"] == "세액감면 수동 검토" for issue in result.issues)
    assert len(result.monthly_payroll_files) == 12

    january = correction_root / "테스트_23년01월_급여대장.xls"
    january_book = load_workbook(BytesIO(january.read_bytes()), data_only=True, read_only=True)
    january_sheet = january_book["Sheet1"]
    assert [cell.value for cell in january_sheet[1]][:8] == [
        "사번",
        "성명",
        "주민번호",
        "입사일",
        "퇴사일",
        "기본급",
        "연장수당",
        "야간연장수당(비과세)",
    ]
    assert january_sheet["A2"].value == "900101-1234567"
    assert january_sheet["D2"].value.date().isoformat() == "2022-12-01"
    assert january_sheet["E2"].value.date().isoformat() == "2023-01-31"
    assert january_sheet["H2"].value == 100_000
    january_headers = [cell.value for cell in january_sheet[1]]
    deduction_total_column = january_headers.index("공제총액")
    assert january_sheet.cell(2, january_headers.index("선지급공제") + 1).value == 3_000
    assert january_sheet.cell(2, january_headers.index("복지포인트(가지급)") + 1).value == 4_000
    assert january_sheet.cell(2, deduction_total_column + 1).value == 17_000
    assert january_headers[deduction_total_column + 1] == "실지급액"
    january_book.close()

    february = correction_root / "테스트_23년02월_급여대장.xls"
    february_book = load_workbook(BytesIO(february.read_bytes()), data_only=True, read_only=True)
    assert february_book["Sheet1"]["B2"].value == "김철수"
    assert february_book["Sheet1"]["D2"].value is None
    assert february_book["Sheet1"]["E2"].value is None
    february_book.close()


def test_excludes_year_with_no_target_rows_and_removes_stale_outputs(tmp_path, caplog):
    caplog.set_level("INFO")
    basic = tmp_path / "테스트_기초정보.xlsx"
    status = tmp_path / "테스트_23년 급상여현황.xlsx"
    tool_template = tmp_path / "Temp_수정 원천징수영수증 생성 Tool.xlsx"
    modified_root = tmp_path / "Input_급상여현황(수정)"
    correction_root = tmp_path / "2.근로소득 수정신고" / "1.Input"
    _make_basic(basic)
    _make_status(status)
    _make_tool(tool_template)

    status_book = load_workbook(status)
    target = status_book["대상자(작업)"]
    target.delete_rows(2, target.max_row)
    status_book.save(status)
    status_book.close()

    modified_root.mkdir()
    correction_root.mkdir(parents=True)
    stale_modified = modified_root / "테스트_23년 급상여현황(수정).xlsx"
    stale_tool = modified_root / "테스트_23년 수정 원천징수영수증 생성 Tool.xlsx"
    stale_monthly = correction_root / "테스트_23년01월_급여대장.xls"
    for path in (stale_modified, stale_tool, stale_monthly):
        path.write_bytes(b"stale")

    result = prepare_correction_inputs(
        company_name="테스트",
        basic_info_path=basic,
        payroll_files=(),
        payroll_status_files=(status,),
        tool_template_path=tool_template,
        modified_input_directory=modified_root,
        correction_input_directory=correction_root,
        xls_converter=_copy_converter,
    )

    assert result.years == ()
    assert result.modified_status_files == ()
    assert result.withholding_tool_files == ()
    assert result.monthly_payroll_files == ()
    assert result.issues == (
        {
            "year": 2023,
            "type": "대상자 없음",
            "detail": "대상자 합계행 0개로 수정신고 입력자료 생성 제외",
        },
    )
    assert "2023년 수정신고 입력자료 생성 제외: 대상자 합계행 0개" in caplog.messages
    assert not stale_modified.exists()
    assert not stale_tool.exists()
    assert not stale_monthly.exists()


def test_tax_reduction_review_distinguishes_rate_and_limit_cases():
    target = {"사원번호": "1", "성명": "홍길동"}

    issues: list[dict[str, object]] = []
    _append_tax_reduction_review_issue(
        issues=issues,
        year=2023,
        target=target,
        tax={"조특법 감면율": 0.7, "산출세액": 1_000_000, "세액감면 계": 700_000},
    )
    assert issues == []

    _append_tax_reduction_review_issue(
        issues=issues,
        year=2023,
        target=target,
        tax={"조특법 감면율": "70%", "산출세액": 3_000_000, "세액감면 계": 2_000_000},
    )
    assert issues[-1]["type"] == "세액감면 수동 검토"
    assert "한도 적용 여부" in str(issues[-1]["reason"])
