from __future__ import annotations

import json
import hashlib
from pathlib import Path
import re
import traceback
import zipfile

from .documents import generate_submission_ledger, generate_supporting_documents
from .errors import PayrollAmendmentError
from .infra import RunStore, make_run_id, safe_name
from .ingestion import ingest_company_inputs
from .models import StageResult, ValidationIssue
from .packaging import build_submission_package
from .payroll import convert_monthly_payroll
from .pdf_render import render_withholding_pdfs
from .withholding import generate_amended_withholding


def _safe_extract(source: Path, destination: Path) -> Path:
    if source.is_dir():
        return source
    if source.suffix.lower() != ".zip":
        raise ValueError("입력은 폴더 또는 ZIP이어야 합니다.")
    destination.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(source) as archive:
        root = destination.resolve()
        total = 0
        for info in archive.infolist():
            if info.is_dir():
                continue
            total += info.file_size
            if total > 4 * 1024 * 1024 * 1024:
                raise ValueError("ZIP 압축 해제 크기가 4GB를 초과합니다.")
            target = (destination / info.filename).resolve()
            if root not in target.parents:
                raise ValueError(f"안전하지 않은 ZIP 경로: {info.filename}")
            # Windows의 일반 경로 길이 제한을 넘는 산출물은 짧은 격리 경로로
            # 옮긴다. 파일명 끝의 PDF 인덱스와 경로상의 연도는 유지한다.
            if len(str(target)) >= 240:
                year_match = re.search(r"(?<!\d)(20\d{2}|\d{2})\s*년", info.filename)
                year_part = year_match.group(1) if year_match else "unknown"
                digest = hashlib.sha256(info.filename.encode("utf-8", errors="surrogatepass")).hexdigest()[:16]
                tail = Path(info.filename).name
                if len(tail) > 100:
                    tail = f"{digest}{Path(tail).suffix}"
                target = (destination / "__long__" / year_part / f"{digest}_{tail}").resolve()
            target.parent.mkdir(parents=True, exist_ok=True)
            with archive.open(info) as src, target.open("wb") as dst:
                while chunk := src.read(1024 * 1024):
                    dst.write(chunk)
    return destination


class RunOrchestrator:
    def __init__(self, workspace: Path, templates: Path | None = None):
        self.workspace = workspace.resolve()
        self.templates = (templates or self.workspace / "templates").resolve()
        self.store = RunStore(self.workspace / "state" / "runs.sqlite3")

    def validate(self, source: Path, years: list[int], company_name: str | None = None) -> dict:
        run_id = make_run_id("VALIDATE")
        run_dir = self.workspace / "runs" / run_id
        prepared = _safe_extract(source, run_dir / "source")
        inputs = ingest_company_inputs(source_dir=prepared, years=years, run_dir=run_dir, company_name=company_name)
        return {"run_id": run_id, "company_id": inputs.company.company_id, "company_name": inputs.company.display_name, "years": years, "targets": {year: len(inputs.exemptions_by_year.get(year, [])) for year in years}, "payroll_files": {year: len(inputs.payroll_files_by_year.get(year, [])) for year in years}, "pdf_files": {year: len(inputs.pdf_files_by_year.get(year, [])) for year in years}, "withholding_excel_files": {year: len(inputs.withholding_excel_files_by_year.get(year, [])) for year in years}, "tool_years": sorted(inputs.tool_workbooks_by_year), "manifest": str(inputs.input_manifest_path)}

    def run(self, source: Path, years: list[int], company_name: str | None = None, *, allow_partial: bool = False) -> dict:
        provisional = safe_name(company_name or "COMPANY")
        run_id = make_run_id(provisional)
        run_dir = self.workspace / "runs" / run_id
        prepared = _safe_extract(source, run_dir / "source")
        inputs = ingest_company_inputs(source_dir=prepared, years=years, run_dir=run_dir, company_name=company_name)
        self.store.create(run_id, inputs.company.company_id, years)
        results: dict[int, list[StageResult]] = {year: [] for year in years}
        s10 = StageResult("S10", inputs.company.company_id, None, "SUCCEEDED", [inputs.input_manifest_path], {"years": len(years), "targets": sum(len(v) for v in inputs.exemptions_by_year.values())})
        self.store.record_stage(run_id, None, "S10", s10.status, s10.to_dict(run_dir))
        for year in years:
            try:
                exemptions = inputs.exemptions_by_year.get(year, [])
                s20 = convert_monthly_payroll(company_id=inputs.company.company_id, company_name=inputs.company.display_name, year=year, payroll_files=inputs.payroll_files_by_year.get(year, []), exemptions=exemptions, output_dir=run_dir / "results" / str(year) / "01_payroll")
                results[year].append(s20); self.store.record_stage(run_id, year, "S20", s20.status, s20.to_dict(run_dir))
                s30 = render_withholding_pdfs(company_id=inputs.company.company_id, year=year, exemptions=exemptions, pdf_files=inputs.pdf_files_by_year.get(year, []), output_dir=run_dir / "results" / str(year) / "02_pdf_images")
                results[year].append(s30); self.store.record_stage(run_id, year, "S30", s30.status, s30.to_dict(run_dir))
                tool = inputs.tool_workbooks_by_year.get(year)
                if tool is None:
                    raise ValueError(f"{year}년 원천징수 생성 Tool이 없습니다.")
                s40 = generate_amended_withholding(company_id=inputs.company.company_id, year=year, exemptions=exemptions, tool_workbook=tool, pdf_files=inputs.pdf_files_by_year.get(year, []), output_dir=run_dir / "results" / str(year) / "03_amended_withholding")
                results[year].append(s40); self.store.record_stage(run_id, year, "S40", s40.status, s40.to_dict(run_dir))
                s50 = generate_supporting_documents(company_id=inputs.company.company_id, company_name=inputs.company.display_name, year=year, template_dir=self.templates / "supporting_docs", output_dir=run_dir / "results" / str(year) / "04_supporting_docs")
                results[year].append(s50); self.store.record_stage(run_id, year, "S50", s50.status, s50.to_dict(run_dir))
                s60 = generate_submission_ledger(company_id=inputs.company.company_id, year=year, payroll_artifacts=s20.artifacts, output_dir=run_dir / "results" / str(year) / "05_submission_ledger")
                results[year].append(s60); self.store.record_stage(run_id, year, "S60", s60.status, s60.to_dict(run_dir))
            except Exception as exc:
                failed_code = next((code for code in ("S20", "S30", "S40", "S50", "S60") if not any(r.stage_code == code for r in results[year])), "S60")
                issue = ValidationIssue(getattr(exc, "code", "E_STAGE_FAILED"), "ERROR", str(exc), year=year)
                failed = StageResult(failed_code, inputs.company.company_id, year, "FAILED", issues=[issue])
                results[year].append(failed); self.store.record_stage(run_id, year, failed_code, "FAILED", failed.to_dict(run_dir))
                (run_dir / "logs").mkdir(parents=True, exist_ok=True)
                (run_dir / "logs" / f"{year}_{failed_code}.log").write_text(traceback.format_exc(), encoding="utf-8")
                if not allow_partial:
                    break
        try:
            s70 = build_submission_package(company_id=inputs.company.company_id, run_id=run_id, year_results=results, output_dir=run_dir / "packages")
            self.store.record_stage(run_id, None, "S70", s70.status, s70.to_dict(run_dir))
            status = "COMPLETED"
        except Exception as exc:
            s70 = StageResult("S70", inputs.company.company_id, None, "FAILED", issues=[ValidationIssue(getattr(exc, "code", "E_PACKAGE"), "ERROR", str(exc))])
            self.store.record_stage(run_id, None, "S70", "FAILED", s70.to_dict(run_dir))
            status = "FAILED"
        self.store.finish(run_id, status)
        report = {"run_id": run_id, "status": status, "company_id": inputs.company.company_id, "years": years, "stages": {year: [r.to_dict(run_dir) for r in stage_results] for year, stage_results in results.items()}, "package": s70.to_dict(run_dir)}
        (run_dir / "run_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        return report
