from __future__ import annotations

import json
from pathlib import Path
import shutil

from .errors import RuleNotImplemented
from .infra import safe_name, sha256
from .models import StageResult


def generate_supporting_documents(*, company_id: str, company_name: str, year: int, template_dir: Path, output_dir: Path) -> StageResult:
    specs_path = template_dir / "document_specs.json"
    if not specs_path.exists():
        raise RuleNotImplemented("추가서류 document_specs.json과 승인된 템플릿이 필요합니다.")
    specs = json.loads(specs_path.read_text(encoding="utf-8"))
    artifacts = []
    index = []
    output_dir.mkdir(parents=True, exist_ok=True)
    for spec in specs:
        if not spec.get("required", True):
            continue
        source = template_dir / spec["template"]
        if not source.exists():
            raise RuleNotImplemented(f"추가서류 템플릿 누락: {source.name}")
        filename = spec["filename"].format(company=company_name, year=year)
        target = output_dir / safe_name(filename)
        shutil.copy2(source, target)
        artifacts.append(target)
        index.append({"document_type": spec["document_type"], "file": target.name, "sha256": sha256(target), "year": year})
    index_path = output_dir / "document_index.json"
    index_path.write_text(json.dumps(index, ensure_ascii=False, indent=2), encoding="utf-8")
    artifacts.append(index_path)
    return StageResult("S50", company_id, year, "NEEDS_REVIEW", artifacts, {"documents": len(index)})


def generate_submission_ledger(*, company_id: str, year: int, payroll_artifacts: list[Path], output_dir: Path) -> StageResult:
    ledgers = [path for path in payroll_artifacts if path.suffix.lower() == ".xlsx"]
    if not ledgers:
        raise RuleNotImplemented("제출 임금대장을 만들 월별 급여대장이 없습니다.")
    # 월별 대장은 이미 S20에서 원본 합계를 보존해 생성한다. 기관별 승인 양식 전에는
    # 표준 제출묶음만 작성하고 검토 필요 상태로 둔다.
    output_dir.mkdir(parents=True, exist_ok=True)
    artifacts = []
    for path in ledgers:
        target = output_dir / path.name
        shutil.copy2(path, target); artifacts.append(target)
    reconciliation = output_dir / "ledger_reconciliation.json"
    reconciliation.write_text(json.dumps({"year": year, "files": [p.name for p in artifacts], "status": "기관별 양식 검토 필요"}, ensure_ascii=False, indent=2), encoding="utf-8")
    artifacts.append(reconciliation)
    return StageResult("S60", company_id, year, "NEEDS_REVIEW", artifacts, {"ledgers": len(ledgers)})

