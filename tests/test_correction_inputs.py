from io import BytesIO
from pathlib import Path
import shutil

from openpyxl import Workbook, load_workbook

from payroll_amendment.correction_inputs import prepare_correction_inputs


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
    ])
    tax.append([
        2023, "9001011234567", "홍길동", "9001011234567", 7,
        10_000_000, 200_000, 10_000_000, 200_000, 100_000, 200_000, 200_000,
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
    ]
    target.append(headers)
    target.append(["9001011234567", "홍길동", "9001011234567", "2022-12-01", "2023-01-31", "23년01월", "급여", 2_000_000, 0, 100_000, 2_100_000, 10_000, 3_000, 4_000, 17_000, 2_083_000, "9309", 7])
    target.append(["9001011234567", "홍길동", "9001011234567", "2022-12-01", "2023-01-31", "23년합계", "급상여", 2_000_000, 0, 100_000, 2_100_000, 10_000, 3_000, 4_000, 17_000, 2_083_000, "9309", 7])
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
        "18-1.생산직등야간근로수당", "과세소득_총계", "비과세소득_총계",
        "과세소득_현계", "비과세소득_현계", "결정세액_소득세",
        "비과세명1_총", "코드1_총", "금액1_총",
        "비과세명1_현", "코드1_현", "금액1_현",
    ]
    previous.append(headers)
    data = workbook.create_sheet("Data")
    data.append(headers)
    data.append([None] * 6 + ["=prevData!G2"] + [None] * (len(headers) - 7))
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
    assert tool_book["prevData"]["H2"].value == 10_000_000
    assert tool_book["prevData"]["M2"].value == "야간근로수당"
    assert tool_book["prevData"]["N2"].value == "O01"
    assert tool_book["Data"]["A2"].value == "2023년"
    assert tool_book["Data"]["G2"].value == "=prevData!G2"
    tool_book.close()
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
