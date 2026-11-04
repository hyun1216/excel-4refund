from pathlib import Path
import shutil

from openpyxl import Workbook, load_workbook
import pytest

from payroll_amendment.errors import DuplicatePdfIndex
from payroll_amendment.submission_target_lists import generate_submission_target_lists


STATUS_HEADERS = [
    "사원번호",
    "성명",
    "주민등록번호",
    "입사일",
    "퇴사일",
    "지급년월",
    "인덱스",
    "야간연장수당(비과세)",
    "지급합계",
    "비과세소득_현계",
    "과세소득_현계",
    "과세표준",
    "결정세액_소득세",
    "기납부세액_소득세",
    "차감징수세액_소득세",
    "직업분류 코드",
    "직업분류 항목",
]


def _make_status(
    path: Path,
    duplicate_index: bool = False,
    lost_date: str | None = None,
) -> None:
    workbook = Workbook()
    workbook.remove(workbook.active)
    for sheet_name, non_tax in (("대상자(작업)", 500_000), ("대상자", 100_000)):
        sheet = workbook.create_sheet(sheet_name)
        sheet.append(STATUS_HEADERS)
        sheet.append(
            [
                "E001", "홍길동", "9001011234567", "2020-01-02", lost_date,
                "23년01월", None, non_tax, 3_000_000, 0,
                None, None, None, None, None, "9309", "기타 단순 노무 종사자",
            ]
        )
        sheet.append(
            [
                "E001",
                "홍길동",
                "9001011234567",
                "2020-01-02",
                lost_date,
                "23년합계",
                7,
                non_tax,
                30_000_000,
                0,
                30_000_000,
                18_000_000,
                1_200_000,
                1_500_000,
                -300_000,
                "9309",
                "기타 단순 노무 종사자",
            ]
        )
        if duplicate_index:
            sheet.append(
                [
                    "E002",
                    "김영희",
                    "9102022234567",
                    "2021-03-01",
                    None,
                    "23년합계",
                    7,
                    non_tax,
                    20_000_000,
                    0,
                    20_000_000,
                    10_000_000,
                    500_000,
                    700_000,
                    -200_000,
                    "9309",
                    "기타 단순 노무 종사자",
                ]
            )
    workbook.save(path)
    workbook.close()


def _make_basic(path: Path) -> None:
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "건보정보"
    sheet.append(["연번", "일련번호", "증번호", "주민등록번호", "성명", "자격취득일", "자격상실일"])
    sheet.append([1, "01", "12345678901", "900101-1******", "홍길동(재직)", "2020.01.02", None])
    tax = workbook.create_sheet("원천징수추출(23년)")
    tax.append(
        [
            "기간",
            "사원번호",
            "성명",
            "주민등록번호",
            "인덱스",
            "과세소득_총계",
            "과세소득_현계",
            "과세표준",
            "결정세액_소득세",
            "기납부세액_소득세",
            "차감징수세액_소득세",
        ]
    )
    tax.append(
        [
            2023,
            "E001",
            "홍길동",
            "9001011234567",
            7,
            30_000_000,
            30_000_000,
            18_000_000,
            1_200_000,
            1_500_000,
            -300_000,
        ]
    )
    workbook.save(path)
    workbook.close()


def test_generates_nts_and_insurance_target_lists(tmp_path: Path) -> None:
    basic = tmp_path / "테스트_기초정보.xlsx"
    status = tmp_path / "테스트_23년 급상여현황.xlsx"
    _make_basic(basic)
    _make_status(status)

    result = generate_submission_target_lists(
        company_name="테스트",
        basic_info_path=basic,
        payroll_status_files=[status],
        output_root=tmp_path / "output",
        xls_converter=lambda pairs: [shutil.copyfile(source, destination) for source, destination in pairs],
    )

    assert result.years == (2023,)
    assert result.issues == ()
    nts = load_workbook(result.nts_files[0], read_only=True, data_only=True)
    health_handle = result.health_files[0].open("rb")
    legacy_employment_handle = result.legacy_employment_files[0].open("rb")
    general_employment = load_workbook(result.general_employment_files[0], read_only=True, data_only=True)
    health = load_workbook(health_handle, read_only=True, data_only=True)
    employment = load_workbook(result.employment_files[0], read_only=True, data_only=False)
    legacy_employment = load_workbook(
        legacy_employment_handle, read_only=True, data_only=False
    )
    try:
        nts_values = list(nts["제출대상자"].iter_rows(min_row=2, values_only=True))
        nts_headers = list(nts_values[0])
        nts_row = dict(zip(nts_headers, nts_values[1]))
        assert nts_row["PDF 인덱스"] == 7
        assert nts_row["주민등록번호"] == "900101-1234567"
        assert nts_row["추가 비과세금액"] == 400_000
        assert nts_row["기존 총급여"] == 30_000_000
        assert nts_row["수정 총급여"] == 29_600_000

        health_sheet = health.worksheets[0]
        assert [cell.value for cell in health_sheet[1]][:3] == [
            "성명", "주민번호\n(숫자만 13자리)", "증번호\n(숫자만 11자리)"
        ]
        assert [health_sheet.cell(2, column).value for column in range(1, 10)] == [
            "홍길동", "900101-1234567", "12345678901", 1, 2023,
            29_600_000, 12, "2020.01.02", None,
        ]
        assert health_sheet.cell(2, 2).number_format == "@"
        assert health_sheet.cell(2, 3).number_format == "@"
        employment_sheet = employment.worksheets[0]
        assert employment_sheet.max_row == 1
        assert legacy_employment.worksheets[0].max_row == 1
        assert result.employment_files[0].name == "2023년 고용종료자_보수총액수정신고서_테스트.xlsx"
        assert result.general_employment_files[0].name == "2023년 일반_보수총액신고서_테스트.xlsx"
        general_sheet = general_employment.worksheets[0]
        assert [general_sheet.cell(2, column).value for column in range(1, 9)] == [
            "홍길동", "900101-1234567", None, 20200102, 29_600_000,
            20200102, 29_600_000, None,
        ]
    finally:
        nts.close()
        health.close()
        employment.close()
        general_employment.close()
        legacy_employment.close()
        health_handle.close()
        legacy_employment_handle.close()


def test_retired_employee_stays_only_in_correction_form(tmp_path: Path) -> None:
    basic = tmp_path / "테스트_기초정보.xlsx"
    status = tmp_path / "테스트_23년 급상여현황.xlsx"
    _make_basic(basic)
    _make_status(status, lost_date="2023-12-31")

    result = generate_submission_target_lists(
        company_name="테스트",
        basic_info_path=basic,
        payroll_status_files=[status],
        output_root=tmp_path / "output",
        xls_converter=lambda pairs: [
            shutil.copyfile(source, destination) for source, destination in pairs
        ],
    )

    employment = load_workbook(result.employment_files[0], read_only=True, data_only=False)
    legacy_employment_handle = result.legacy_employment_files[0].open("rb")
    legacy_employment = load_workbook(
        legacy_employment_handle, read_only=True, data_only=False
    )
    general_employment = load_workbook(
        result.general_employment_files[0], read_only=True, data_only=True
    )
    try:
        employment_sheet = employment.worksheets[0]
        assert employment_sheet.cell(2, 2).value == "9001011234567"
        assert employment_sheet.cell(2, 2).number_format == "@"
        assert employment_sheet.cell(2, 7).value == 20231231
        assert employment_sheet.cell(2, 10).value == 20231231
        assert employment_sheet.max_column == 12
        legacy_sheet = legacy_employment.worksheets[0]
        assert legacy_sheet.cell(2, 2).value == "900101-1234567"
        assert legacy_sheet.cell(2, 2).number_format == "@"
        assert legacy_sheet.cell(2, 7).value == 20231231
        assert legacy_sheet.cell(2, 10).value == 20231231
        assert legacy_sheet.cell(2, 15).value == "=L2-M2-N2"
        assert general_employment.worksheets[0].max_row == 1
    finally:
        employment.close()
        general_employment.close()
        legacy_employment.close()
        legacy_employment_handle.close()


def test_rejects_duplicate_pdf_index(tmp_path: Path) -> None:
    basic = tmp_path / "테스트_기초정보.xlsx"
    status = tmp_path / "테스트_23년 급상여현황.xlsx"
    _make_basic(basic)
    _make_status(status, duplicate_index=True)

    with pytest.raises(DuplicatePdfIndex):
        generate_submission_target_lists(
            company_name="테스트",
            basic_info_path=basic,
            payroll_status_files=[status],
            output_root=tmp_path / "output",
            xls_converter=lambda pairs: None,
        )
