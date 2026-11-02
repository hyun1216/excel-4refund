from __future__ import annotations

from datetime import datetime
import json
from pathlib import Path
import zipfile

from .errors import PayrollAmendmentError
from .infra import sha256
from .models import StageResult


def build_submission_package(*, company_id: str, run_id: str, year_results: dict[int, list[StageResult]], output_dir: Path) -> StageResult:
    blocking = [r for results in year_results.values() for r in results if r.status in {"FAILED", "NEEDS_REVIEW"}]
    if blocking:
        raise PayrollAmendmentError("검토 필요 또는 실패 단계가 있어 제출 ZIP을 만들 수 없습니다.")
    output_dir.mkdir(parents=True, exist_ok=True)
    manifest = []
    package = output_dir / f"제출자료_{company_id}_{run_id}.zip"
    with zipfile.ZipFile(package, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        used = set()
        for year, results in sorted(year_results.items()):
            for result in results:
                if result.stage_code not in {"S40", "S50", "S60"}:
                    continue
                for path in result.artifacts:
                    if not path.is_file() or path.suffix.lower() in {".json", ".csv"}:
                        continue
                    internal = f"{year}/{result.stage_code}/{path.name}"
                    if internal in used:
                        raise PayrollAmendmentError(f"ZIP 내부 경로 중복: {internal}")
                    used.add(internal); archive.write(path, internal)
                    manifest.append({"year": year, "stage": result.stage_code, "path": internal, "sha256": sha256(path), "size": path.stat().st_size})
        archive.writestr("manifest.json", json.dumps({"run_id": run_id, "company_id": company_id, "created_at": datetime.now().astimezone().isoformat(), "files": manifest}, ensure_ascii=False, indent=2))
    with zipfile.ZipFile(package) as check:
        bad = check.testzip()
        if bad or len(check.namelist()) <= 1:
            raise PayrollAmendmentError(f"제출 ZIP 검증 실패: {bad or '파일 없음'}")
    return StageResult("S70", company_id, None, "SUCCEEDED", [package], {"files": len(manifest)})
