from __future__ import annotations

from dataclasses import dataclass
import os
from pathlib import Path
import shutil
from uuid import uuid4

from .errors import SchemaMismatch
from .withholding_pdf_converter import ConverterConfig, convert_batch


@dataclass(frozen=True)
class AmendedWithholdingResult:
    normal_files: tuple[Path, ...]
    marked_files: tuple[Path, ...]
    report_path: Path
    excluded_image_pdfs: int = 0


def generate_company_amended_withholding(
    company_directory: Path,
) -> AmendedWithholdingResult:
    company_root = Path(company_directory).resolve(strict=True)
    tool_root = company_root / "Input_급상여현황(수정)"
    pdf_root = company_root / "1. 급상여현황 생성" / "원천징수(PDF)"
    output_root = company_root / "2.근로소득 수정신고" / "2.Output"
    normal_root = output_root / "원천징수(PDF)"
    marked_root = output_root / "원천징수(PDF)_변경표시"
    staging_root = output_root / f".amended_withholding_{uuid4().hex}"
    staging_normal = staging_root / "normal"
    staging_marked = staging_root / "marked"

    if not tool_root.is_dir():
        raise SchemaMismatch(f"수정 원천징수 Tool 폴더가 없습니다: {tool_root}")
    if not pdf_root.is_dir():
        raise SchemaMismatch(f"개인별 원천징수 PDF 폴더가 없습니다: {pdf_root}")

    config = ConverterConfig(
        root=company_root,
        excel_dir=str(tool_root.relative_to(company_root)),
        pdf_dir=str(pdf_root.relative_to(company_root)),
        output_dir=str(staging_normal.relative_to(company_root)),
        marked_output_dir=str(staging_marked.relative_to(company_root)),
        year_folder="{year:02d}년",
        workbook_pattern="*수정 원천징수영수증 생성 Tool.xlsx",
        column_map={"성명": "이름", "주민번호": "주민등록번호"},
        overwrite=True,
    )
    try:
        results, report = convert_batch(config)
        failures = [
            row
            for row in report
            if row["상태"] in {"오류", "부분완료"}
        ]
        if failures:
            details = "; ".join(
                f"{row['연도']}년 {Path(row['PDF']).name if row['PDF'] else ''}: {row['내용']}"
                for row in failures[:10]
            )
            raise SchemaMismatch(f"수정 원천징수 PDF 생성에 실패했습니다: {details}")
        excluded_count = sum(row["상태"] == "제외" for row in report)

        staged_normal = tuple(
            Path(result.outputs["normal"])
            for result in results
            if "normal" in result.outputs
        )
        staged_marked = tuple(
            Path(result.outputs["marked"])
            for result in results
            if "marked" in result.outputs
        )
        if len(staged_normal) != len(staged_marked):
            raise SchemaMismatch(
                "수정후 PDF와 변경표시 PDF 생성 개수가 일치하지 않습니다."
            )
        if not staged_normal and not excluded_count:
            raise SchemaMismatch("수정 원천징수 PDF 생성 대상이 없습니다.")

        normal_files = []
        marked_files = []
        for source in staged_normal:
            destination = normal_root / source.relative_to(staging_normal)
            destination.parent.mkdir(parents=True, exist_ok=True)
            os.replace(source, destination)
            normal_files.append(destination)
        for source in staged_marked:
            destination = marked_root / source.relative_to(staging_marked)
            destination.parent.mkdir(parents=True, exist_ok=True)
            os.replace(source, destination)
            marked_files.append(destination)
        report_path = normal_root / config.report_name
        os.replace(staging_normal / config.report_name, report_path)
        return AmendedWithholdingResult(
            tuple(normal_files),
            tuple(marked_files),
            report_path,
            excluded_count,
        )
    finally:
        shutil.rmtree(staging_root, ignore_errors=True)
