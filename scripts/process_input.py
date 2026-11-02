from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from payroll_amendment.input_intake import organize_input_batch


def main() -> int:
    parser = argparse.ArgumentParser(description="공용 input 접수함의 파일을 거래처 작업 폴더로 분류·이동합니다.")
    parser.add_argument("--project-root", type=Path, default=PROJECT_ROOT)
    parser.add_argument("--input-dir", type=Path)
    args = parser.parse_args()
    project_root = args.project_root.resolve(strict=True)
    input_directory = (args.input_dir or project_root / "input").resolve(strict=True)
    result = organize_input_batch(input_directory=input_directory, project_root=project_root)
    print(
        json.dumps(
            {
                "company_name": result.company_name,
                "company_id": result.company_id,
                "company_directory": str(result.company_directory),
                "moved_files": [str(path) for path in result.moved_files],
                "unknown_files": [str(path) for path in result.unknown_files],
                "report": str(result.report_path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
