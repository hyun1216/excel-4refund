import json
from pathlib import Path
import re
from zipfile import ZIP_DEFLATED, ZipFile

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


def _remove_worksheet_dimension(path: Path) -> None:
    with ZipFile(path) as archive:
        contents = {name: archive.read(name) for name in archive.namelist()}
    worksheet = "xl/worksheets/sheet1.xml"
    contents[worksheet] = re.sub(br"<dimension[^>]*/>", b"", contents[worksheet], count=1)
    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        for name, content in contents.items():
            archive.writestr(name, content)


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

    company = project / "노무_세무 수정신고" / "테스트"
    assert result.company_directory == company
    assert (company / "1. 급상여현황 생성" / "1.Input" / "원천징수영수증" / receipt.name).is_file()
    assert (company / "1. 급상여현황 생성" / "1.Input" / "급여대장" / payroll.name).is_file()
    assert (company / employment.name).is_file()
    assert (company / health.name).is_file()
    report = json.loads(result.report_path.read_text(encoding="utf-8-sig"))
    assert report["status"] == "SUCCEEDED"
    assert report["company_aliases"] == ["주식회사 테스트"]
    assert len(report["moved_files"]) == 4


def test_organizes_payroll_without_worksheet_dimension(tmp_path):
    project = tmp_path / "project"
    inbox = project / "input"
    inbox.mkdir(parents=True)
    receipt = inbox / "25년 원천징수영수증.pdf"
    receipt.write_bytes(b"test receipt")
    payroll = inbox / "25년 1월 급여대장.xlsx"
    _make_payroll(payroll)
    _remove_worksheet_dimension(payroll)

    result = organize_input_batch(
        input_directory=inbox,
        project_root=project,
        pdf_classifier=lambda _: "withholding_receipt",
        company_metadata_reader=lambda _: ("주식회사 테스트", "123-45-67890"),
    )

    assert payroll.name in {path.name for path in result.moved_files}
    assert not result.unknown_files


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

    destination = project / "노무_세무 수정신고" / "SCK" / employment.name
    assert destination.is_file()
    assert destination in result.moved_files
    assert not result.unknown_files


def test_allows_pdf_company_aliases_with_same_business_id(tmp_path):
    project = tmp_path / "project"
    inbox = project / "input"
    inbox.mkdir(parents=True)
    old_receipt = inbox / "23년 원천징수영수증.pdf"
    new_receipt = inbox / "24년 원천징수영수증.pdf"
    old_receipt.write_bytes(b"old receipt")
    new_receipt.write_bytes(b"new receipt")
    employment = inbox / "연합시큐리티_근로자고용정보현황조회.xlsx"
    employment.write_bytes(b"employment")
    health = inbox / "연합시큐리티_사업장가입자명부.xlsx"
    health.write_bytes(b"health")

    metadata = {
        old_receipt: ("주식회사코드원에스에스(CODEONESS)", "301-86-16563"),
        new_receipt: ("연합시큐리티", "301-86-16563"),
    }
    result = organize_input_batch(
        input_directory=inbox,
        project_root=project,
        pdf_classifier=lambda _: "withholding_receipt",
        company_metadata_reader=metadata.__getitem__,
    )

    assert result.company_name == "연합시큐리티"
    assert result.company_directory.name == "연합시큐리티"
    report = json.loads(result.report_path.read_text(encoding="utf-8-sig"))
    assert report["company_aliases"] == ["주식회사코드원에스에스(CODEONESS)"]


def test_rejects_conflicting_current_company_names(tmp_path):
    project = tmp_path / "project"
    inbox = project / "input"
    inbox.mkdir(parents=True)
    receipt = inbox / "25년 원천징수영수증.pdf"
    receipt.write_bytes(b"receipt")
    (inbox / "연합시큐리티_근로자고용정보현황조회.xlsx").write_bytes(b"employment")
    (inbox / "코드원에스에스_사업장가입자명부.xlsx").write_bytes(b"health")

    try:
        organize_input_batch(
            input_directory=inbox,
            project_root=project,
            pdf_classifier=lambda _: "withholding_receipt",
            company_metadata_reader=lambda _: ("과거상호", "301-86-16563"),
        )
    except SchemaMismatch as exc:
        assert "현재 상호가 다릅니다" in str(exc)
    else:
        raise AssertionError("고용정보와 사업장가입자명부의 현재 상호가 다르면 중단해야 합니다.")


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


def test_organizes_dated_payroll_pdf_from_year_folder(tmp_path):
    project = tmp_path / "project"
    inbox = project / "input"
    payroll = inbox / "2023" / "202301(2. 급여 + 상여)[20230125].pdf"
    payroll.parent.mkdir(parents=True)
    payroll.write_bytes(b"test payroll")
    receipt = inbox / "2023년 원천징수영수증.pdf"
    receipt.write_bytes(b"test receipt")

    def classify(path: Path) -> str:
        if path == receipt:
            return "withholding_receipt"
        from payroll_amendment.dawon_pipeline import _parse_payroll_name

        _parse_payroll_name(path)
        return "tabular_payroll"

    result = organize_input_batch(
        input_directory=inbox,
        project_root=project,
        pdf_classifier=classify,
        company_metadata_reader=lambda _: ("주식회사 테스트", "123-45-67890"),
    )

    moved = {path.name: path for path in result.moved_files}
    assert moved[payroll.name].parent.name == "23년"
    assert moved[payroll.name].parent.parent.name == "급여대장"
    assert not result.unknown_files
