from __future__ import annotations

import argparse
import json
import logging
from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from payroll_amendment.correction_inputs import prepare_correction_inputs
from payroll_amendment.errors import SchemaMismatch


def _single(paths: list[Path], label: str) -> Path:
    if len(paths) != 1:
        raise SchemaMismatch(f"{label} 파일이 {len(paths)}개입니다: {paths}")
    return paths[0]


def main() -> int:
    parser = argparse.ArgumentParser(
        description="완료된 급상여현황으로 급상여현황(수정), Tool, 월별 급여대장을 생성합니다."
    )
    parser.add_argument("--company-dir", type=Path, required=True)
    parser.add_argument("--project-root", type=Path, default=PROJECT_ROOT)
    args = parser.parse_args()

    project_root = args.project_root.resolve(strict=True)
    company_dir = args.company_dir
    if not company_dir.is_absolute():
        company_dir = project_root / company_dir
    company_dir = company_dir.resolve(strict=True)

    basic_info = _single(
        sorted(company_dir.glob("*_기초정보.xlsx")),
        "기초정보",
    )
    company_name = basic_info.name.removesuffix("_기초정보.xlsx")
    payroll_input = company_dir / "2.근로소득 수정신고" / "1.Input"
    payroll_files = sorted(payroll_input.glob("*_??년 급여대장.xlsx"))
    status_dir = company_dir / "1. 급상여현황 생성" / "급상여현황"
    status_files = sorted(
        path
        for path in status_dir.glob("*_??년 급상여현황.xlsx")
        if "(수정)" not in path.name
    )
    if not payroll_files:
        raise SchemaMismatch(f"연도별 급여대장이 없습니다: {payroll_input}")
    if not status_files:
        raise SchemaMismatch(f"완료된 급상여현황이 없습니다: {status_dir}")

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    logging.info("수정신고 입력자료 생성 시작: %s", company_name)
    result = prepare_correction_inputs(
        company_name=company_name,
        basic_info_path=basic_info,
        payroll_files=tuple(payroll_files),
        payroll_status_files=tuple(status_files),
        tool_template_path=project_root / "Resources" / "Template" / "Temp_수정 원천징수영수증 생성 Tool.xlsx",
        modified_input_directory=company_dir / "Input_급상여현황(수정)",
        correction_input_directory=payroll_input,
        withholding_pdf_root=company_dir / "1. 급상여현황 생성" / "원천징수(PDF)",
    )
    logging.info("수정신고 입력자료 생성 완료")
    print(
        json.dumps(
            {
                "modified_payroll_status_files": [str(path) for path in result.modified_status_files],
                "withholding_tool_files": [str(path) for path in result.withholding_tool_files],
                "monthly_payroll_files": [str(path) for path in result.monthly_payroll_files],
                "issues": list(result.issues),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
