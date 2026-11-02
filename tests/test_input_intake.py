import json
from pathlib import Path

from openpyxl import Workbook

from payroll_amendment.input_intake import organize_input_batch
from payroll_amendment.errors import SchemaMismatch


def _make_payroll(path: Path) -> None:
    book = Workbook()
    sheet = book.active
    sheet.title = "1월"
    sheet.append(["성명", "사원번호", "기본급", "지급합계"])
    sheet.append(["홍길동", "001", 2_000_000, 2_000_000])
    book.save(path)


def test_organizes_single_company_batch(tmp_path):
    project = tmp_path / "project"
    inbox = project / "input"
    inbox.mkdir(parents=True)
    receipt = inbox / "25년 원천징수영수증.pdf"
    receipt.write_bytes(b"test receipt")
    payroll = inbox / "25년 1월 급여대장.xlsx"
    _make_payroll(payroll)
    employment = inbox / "테스트_근로자고용정보현황조회.xlsx"
    employment.write_bytes(b"employment")
    health = inbox / "테스트_사업장가입자명부.xlsx"
    health.write_bytes(b"health")

    result = organize_input_batch(
        input_directory=inbox,
        project_root=project,
        pdf_classifier=lambda _: "withholding_receipt",
        company_metadata_reader=lambda _: ("주식회사 테스트", "123-45-67890"),
    )

    company = project / "노무_세무 수정신고" / "주식회사 테스트"
    assert result.company_directory == company
    assert (company / "1. 급상여현황 생성" / "1.Input" / "원천징수영수증" / receipt.name).is_file()
    assert (company / "1. 급상여현황 생성" / "1.Input" / "급여대장" / payroll.name).is_file()
    assert (company / employment.name).is_file()
    assert (company / health.name).is_file()
    report = json.loads(result.report_path.read_text(encoding="utf-8-sig"))
    assert report["status"] == "SUCCEEDED"
    assert len(report["moved_files"]) == 4


def test_moves_occupation_mapping_result_to_company_directory(tmp_path):
    project = tmp_path / "project"
    inbox = project / "input"
    inbox.mkdir(parents=True)
    receipt = inbox / "25년 원천징수영수증.pdf"
    receipt.write_bytes(b"test receipt")
    employment = inbox / "SCK_근로자고용정보현황조회_직업분류매핑결과.xlsx"
    employment.write_bytes(b"employment mapping result")

    result = organize_input_batch(
        input_directory=inbox,
        project_root=project,
        pdf_classifier=lambda _: "withholding_receipt",
        company_metadata_reader=lambda _: ("에스씨케이(주)", "123-45-67890"),
    )

    destination = project / "노무_세무 수정신고" / "에스씨케이(주)" / employment.name
    assert destination.is_file()
    assert destination in result.moved_files
    assert not result.unknown_files


def test_existing_destination_blocks_all_moves(tmp_path):
    project = tmp_path / "project"
    inbox = project / "input"
    inbox.mkdir(parents=True)
    receipt = inbox / "25년 원천징수영수증.pdf"
    receipt.write_bytes(b"test receipt")
    payroll = inbox / "25년 1월 급여대장.xlsx"
    _make_payroll(payroll)
    existing = (
        project
        / "노무_세무 수정신고"
        / "주식회사 테스트"
        / "1. 급상여현황 생성"
        / "1.Input"
        / "원천징수영수증"
        / receipt.name
    )
    existing.parent.mkdir(parents=True)
    existing.write_bytes(b"existing")

    try:
        organize_input_batch(
            input_directory=inbox,
            project_root=project,
            pdf_classifier=lambda _: "withholding_receipt",
            company_metadata_reader=lambda _: ("주식회사 테스트", "123-45-67890"),
        )
    except SchemaMismatch:
        pass
    else:
        raise AssertionError("같은 이름의 기존 파일이 있으면 접수를 중단해야 합니다.")

    assert receipt.is_file()
    assert payroll.is_file()
    assert existing.read_bytes() == b"existing"


def test_organizes_nested_year_folders_and_keeps_direct_input_supported(tmp_path):
    project = tmp_path / "project"
    inbox = project / "input"
    receipt = inbox / "원천징수영수증" / "2023년" / "23년_원천징수영수증.pdf"
    receipt.parent.mkdir(parents=True)
    receipt.write_bytes(b"test receipt")
    nested_payroll = inbox / "급여대장" / "23년" / "23년_1월_급여대장.xlsx"
    nested_payroll.parent.mkdir(parents=True)
    _make_payroll(nested_payroll)
    direct_payroll = inbox / "25년_1월_급여대장.xlsx"
    _make_payroll(direct_payroll)

    result = organize_input_batch(
        input_directory=inbox,
        project_root=project,
        pdf_classifier=lambda _: "withholding_receipt",
        company_metadata_reader=lambda _: ("주식회사 테스트", "123-45-67890"),
    )

    moved = {path.name: path for path in result.moved_files}
    assert moved[receipt.name].parent.name == "23년"
    assert moved[nested_payroll.name].parent.name == "23년"
    assert moved[direct_payroll.name].parent.name != "25년"
    assert not (inbox / "원천징수영수증").exists()
    assert not (inbox / "급여대장").exists()
    report = json.loads(result.report_path.read_text(encoding="utf-8-sig"))
    relative_sources = {item["source_relative"] for item in report["moved_files"]}
    assert str(Path("원천징수영수증") / "2023년" / receipt.name) in relative_sources
    assert str(Path("급여대장") / "23년" / nested_payroll.name) in relative_sources
