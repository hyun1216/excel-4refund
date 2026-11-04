from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from payroll_amendment.submission_target_lists import generate_company_submission_target_lists


def main() -> int:
    parser = argparse.ArgumentParser(
        description="완료된 급상여현황으로 국세청·공단 제출용 대상자 명단을 생성합니다."
    )
    parser.add_argument("--company-dir", type=Path, required=True)
    parser.add_argument("--project-root", type=Path, default=PROJECT_ROOT)
    args = parser.parse_args()

    project_root = args.project_root.resolve(strict=True)
    company_directory = args.company_dir
    if not company_directory.is_absolute():
        company_directory = project_root / company_directory
    result = generate_company_submission_target_lists(company_directory)
    print(
        json.dumps(
            {
                "years": list(result.years),
                "nts_files": [str(path) for path in result.nts_files],
                "insurance_files": [str(path) for path in result.insurance_files],
                "general_employment_files": [
                    str(path) for path in result.general_employment_files
                ],
                "legacy_employment_files": [
                    str(path) for path in result.legacy_employment_files
                ],
                "report_path": str(result.report_path),
                "issues": list(result.issues),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
