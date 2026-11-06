from __future__ import annotations

import argparse
from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from payroll_amendment.amended_withholding import generate_company_amended_withholding


def main() -> int:
    parser = argparse.ArgumentParser(
        description="수정 원천징수 Tool과 개인별 PDF를 연결해 수정후·변경표시 PDF를 생성합니다."
    )
    parser.add_argument("--company-dir", type=Path, required=True)
    args = parser.parse_args()
    result = generate_company_amended_withholding(args.company_dir)
    normal_count = len(result.normal_files)
    marked_count = len(result.marked_files)
    print(
        f"수정 원천징수 PDF 생성 완료: 총 {normal_count + marked_count}개 "
        f"(수정후 {normal_count}개, 변경표시 {marked_count}개, "
        f"이미지 PDF 제외 {result.excluded_image_pdfs}개)"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
