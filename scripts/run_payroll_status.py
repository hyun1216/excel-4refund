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

from payroll_amendment.payroll_status import generate_company_payroll_status


def main() -> int:
    parser = argparse.ArgumentParser(
        description="기초정보 지급항목 표시와 연도별 급상여현황만 다시 생성합니다."
    )
    parser.add_argument("--company-dir", type=Path, required=True)
    parser.add_argument("--resources-dir", type=Path, default=PROJECT_ROOT / "Resources")
    args = parser.parse_args()
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    result = generate_company_payroll_status(
        args.company_dir.resolve(strict=True),
        args.resources_dir.resolve(strict=True),
    )
    print(
        json.dumps(
            {
                "years": result.years,
                "outputs": [str(path) for path in result.output_files],
                "review_report": str(result.report_path),
                "basic_info": str(result.updated_basic_info_path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
