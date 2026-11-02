from __future__ import annotations

import json
from pathlib import Path

import typer

from .input_intake import organize_input_batch
from .orchestrator import RunOrchestrator
from .payroll_status import generate_company_payroll_status
from .year_parser import parse_years


app = typer.Typer(help="근로소득 수정신고 로컬 자동화")


def _print(value: dict) -> None:
    typer.echo(json.dumps(value, ensure_ascii=False, indent=2, default=str))


@app.command()
def validate(source: Path = typer.Option(..., exists=True), years: str = typer.Option(...), workspace: Path = typer.Option(Path("workspace")), company_name: str | None = None):
    """입력 구조와 연도별 자료 일치 여부만 확인합니다."""
    _print(RunOrchestrator(workspace).validate(source, parse_years(years), company_name))


@app.command()
def run(source: Path = typer.Option(..., exists=True), years: str = typer.Option(...), workspace: Path = typer.Option(Path("workspace")), templates: Path | None = None, company_name: str | None = None, allow_partial: bool = False):
    """S10~S70을 실행합니다. 외부 업로드와 이메일은 포함하지 않습니다."""
    result = RunOrchestrator(workspace, templates).run(source, parse_years(years), company_name, allow_partial=allow_partial)
    _print(result)
    if result["status"] != "COMPLETED":
        raise typer.Exit(1)


@app.command()
def status(run_id: str, workspace: Path = typer.Option(Path("workspace"))):
    """저장된 실행·단계 상태를 조회합니다."""
    _print(RunOrchestrator(workspace).store.status(run_id))


@app.command("payroll-status")
def payroll_status(
    company_dir: Path = typer.Option(..., exists=True, file_okay=False),
    resources_dir: Path = typer.Option(..., exists=True, file_okay=False),
):
    """회사 폴더의 기초정보와 연도별 급여대장으로 급상여현황을 생성합니다."""
    result = generate_company_payroll_status(company_dir, resources_dir)
    _print(
        {
            "years": result.years,
            "outputs": result.output_files,
            "review_report": result.report_path,
            "basic_info": result.updated_basic_info_path,
        }
    )


@app.command("intake")
def intake(
    input_dir: Path = typer.Option(Path("input"), exists=True, file_okay=False),
    project_root: Path = typer.Option(Path("."), exists=True, file_okay=False),
):
    """공용 input 접수함의 파일을 거래처 작업 폴더로 분류·이동합니다."""
    result = organize_input_batch(input_directory=input_dir, project_root=project_root)
    _print(
        {
            "company_name": result.company_name,
            "company_id": result.company_id,
            "company_directory": result.company_directory,
            "moved_files": result.moved_files,
            "unknown_files": result.unknown_files,
            "report": result.report_path,
        }
    )
