from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

import pymupdf
import pytest
from openpyxl import Workbook

from payroll_amendment.amended_withholding import generate_company_amended_withholding
from payroll_amendment.errors import SchemaMismatch
from payroll_amendment.withholding_pdf_converter import (
    ConversionResult,
    ConverterConfig,
    _has_extractable_text,
    _index_from_name,
    _pdf_resident_number,
    convert_batch,
)


def _company(tmp_path: Path) -> Path:
    company = tmp_path / "회사"
    (company / "Input_급상여현황(수정)").mkdir(parents=True)
    (company / "1. 급상여현황 생성" / "원천징수(PDF)").mkdir(parents=True)
    return company


def test_generates_normal_and_marked_output_roots(tmp_path):
    company = _company(tmp_path)
    normal = company / "2.근로소득 수정신고" / "2.Output" / "원천징수(PDF)" / "23년" / "수정후.pdf"
    marked = company / "2.근로소득 수정신고" / "2.Output" / "원천징수(PDF)_변경표시" / "23년" / "변경표시.pdf"

    def fake_convert(config):
        staged_normal = company / config.output_dir / "23년" / normal.name
        staged_marked = company / config.marked_output_dir / "23년" / marked.name
        staged_normal.parent.mkdir(parents=True)
        staged_marked.parent.mkdir(parents=True)
        staged_normal.write_bytes(b"normal")
        staged_marked.write_bytes(b"marked")
        (company / config.output_dir / config.report_name).write_text("report", encoding="utf-8")
        conversion = ConversionResult(
            company / "원본.pdf",
            {"normal": staged_normal, "marked": staged_marked},
        )
        return [conversion], [{"상태": "완료", "연도": "23", "PDF": "원본.pdf", "내용": ""}]

    with patch(
        "payroll_amendment.amended_withholding.convert_batch",
        side_effect=fake_convert,
    ) as convert:
        result = generate_company_amended_withholding(company)

    config = convert.call_args.args[0]
    assert config.column_map == {"성명": "이름", "주민번호": "주민등록번호"}
    assert config.workbook_pattern == "*수정 원천징수영수증 생성 Tool.xlsx"
    assert config.year_folder == "{year:02d}년"
    assert result.normal_files == (normal,)
    assert result.marked_files == (marked,)


def test_fails_when_conversion_report_contains_error(tmp_path):
    company = _company(tmp_path)
    with patch(
        "payroll_amendment.amended_withholding.convert_batch",
        return_value=([], [{"상태": "오류", "연도": "23", "PDF": "원본.pdf", "내용": "주민번호 불일치"}]),
    ):
        with pytest.raises(SchemaMismatch, match="주민번호 불일치"):
            generate_company_amended_withholding(company)


def test_pdf_index_supports_leading_and_trailing_filename_styles():
    assert _index_from_name(Path("568_박찬식_발행자보관용.pdf")) == 568
    assert _index_from_name(Path("박찬식_발행자보관용_568.pdf")) == 568


def test_resident_number_ignores_extracted_whitespace_between_characters():
    page = SimpleNamespace(extract_text=lambda: "5\n6\n0\n4\n0\n8\n-\n2\n3\n5\n0\n8\n1\n4")
    reader = SimpleNamespace(pages=[page])

    assert _pdf_resident_number(reader) == "5604082350814"


def test_identity_error_does_not_add_missing_pdf_error(tmp_path):
    excel_root = tmp_path / "tools"
    pdf_root = tmp_path / "pdfs" / "23년"
    excel_root.mkdir()
    pdf_root.mkdir(parents=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["기간", "연번", "이름", "주민등록번호", "인덱스"])
    sheet.append(["2023년", 1, "TEST", "900101-1234567", 1])
    workbook.save(excel_root / "회사_23년 수정 원천징수영수증 생성 Tool.xlsx")
    workbook.close()

    source = pdf_root / "TEST_발행자보관용_1.pdf"
    document = pymupdf.open()
    page = document.new_page()
    page.insert_text((72, 72), "TEST 910101-1234567 EXTRACTABLE TEXT DOCUMENT")
    document.save(source)
    document.close()

    results, report = convert_batch(
        ConverterConfig(
            root=tmp_path,
            excel_dir="tools",
            pdf_dir="pdfs",
            output_dir="output",
            marked_output_dir="marked",
            workbook_pattern="*Tool.xlsx",
            column_map={"성명": "이름", "주민번호": "주민등록번호"},
        )
    )

    errors = [row["내용"] for row in report if row["상태"] == "오류"]
    assert results == []
    assert len(errors) == 1
    assert errors[0].startswith("주민번호 불일치")
    assert not any("해당하는 PDF가 없습니다" in error for error in errors)


def test_image_only_pdf_is_excluded_from_batch(tmp_path):
    excel_root = tmp_path / "tools"
    pdf_root = tmp_path / "pdfs" / "23년"
    excel_root.mkdir()
    pdf_root.mkdir(parents=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Data"
    sheet.append(["기간", "연번", "이름", "주민등록번호", "인덱스"])
    sheet.append(["2023년", 1, "홍길동", "900101-1234567", 1])
    workbook.save(excel_root / "회사_23년 수정 원천징수영수증 생성 Tool.xlsx")

    image_pdf = pdf_root / "1_홍길동_발행자보관용.pdf"
    document = pymupdf.open()
    page = document.new_page()
    page.draw_rect(pymupdf.Rect(20, 20, 200, 200), fill=(0.8, 0.8, 0.8))
    document.save(image_pdf)
    document.close()

    assert _has_extractable_text(image_pdf) is False
    results, report = convert_batch(
        ConverterConfig(
            root=tmp_path,
            excel_dir="tools",
            pdf_dir="pdfs",
            output_dir="output",
            marked_output_dir="marked",
            workbook_pattern="*Tool.xlsx",
            column_map={"성명": "이름", "주민번호": "주민등록번호"},
        )
    )

    assert results == []
    assert sum(row["상태"] == "제외" for row in report) == 1
    assert not any(row["상태"] == "오류" for row in report)
