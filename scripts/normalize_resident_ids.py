from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from payroll_amendment.resident_excel import normalize_resident_ids_in_workbook


def main() -> int:
    parser = argparse.ArgumentParser(
        description="기존 xlsx의 주민등록번호 열을 하이픈 포함 텍스트로 보정합니다."
    )
    parser.add_argument("workbooks", nargs="+", type=Path)
    args = parser.parse_args()
    results = {
        str(path.resolve()): normalize_resident_ids_in_workbook(path)
        for path in args.workbooks
    }
    print(json.dumps(results, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
