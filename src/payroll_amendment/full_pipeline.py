from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import logging
import os
from pathlib import Path
from uuid import uuid4

from .amended_withholding import AmendedWithholdingResult, generate_company_amended_withholding
from .correction_inputs import CorrectionInputResult, prepare_correction_inputs
from .employee_info import discover_employee_info_sources, populate_basic_info_employee_sheets
from .errors import SchemaMismatch
from .input_intake import IntakeResult, organize_input_batch
from .payroll_converter import PayrollConverter, UniversalConversionResult
from .payroll_status import (
    BasicInfoPreparationResult,
    PayrollStatusResult,
    generate_company_payroll_status,
    prepare_company_basic_info_for_review,
)
from .submission_target_lists import (
    SubmissionTargetListResult,
    generate_company_submission_target_lists,
)


logger = logging.getLogger(__name__)


@dataclass(frozen=True)
class FullPipelineResult:
    intake: IntakeResult
    conversion: UniversalConversionResult
    payroll_status: PayrollStatusResult
    basic_info_path: Path
    payroll_files: tuple[Path, ...]
    report_path: Path
    correction_inputs: CorrectionInputResult
    amended_withholding: AmendedWithholdingResult | None
    submission_targets: SubmissionTargetListResult


@dataclass(frozen=True)
class BasicInfoReviewResult:
    intake: IntakeResult
    conversion: UniversalConversionResult
    preparation: BasicInfoPreparationResult
    basic_info_path: Path
    payroll_files: tuple[Path, ...]
    report_path: Path


def _write_json_report(path: Path, report: dict[str, object]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.{uuid4().hex}.tmp")
    try:
        temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8-sig")
        os.replace(temporary, path)
    finally:
        temporary.unlink(missing_ok=True)


def _load_existing_intake(
    company_directory: Path,
    *,
    validate_moved_files: bool = True,
) -> IntakeResult:
    company_root = Path(company_directory).resolve(strict=True)
    report_path = company_root / "1. 급상여현황 생성" / "input_접수보고서.json"
    if not report_path.is_file():
        raise SchemaMismatch(f"기존 입력 접수보고서가 없습니다: {report_path}")
    report = json.loads(report_path.read_text(encoding="utf-8-sig"))
    moved_files = tuple(Path(item["destination"]) for item in report.get("moved_files", []))
    missing = [path for path in moved_files if not path.is_file()]
    if validate_moved_files and missing:
        raise SchemaMismatch(
            "기존 접수 파일이 누락됐습니다: " + ", ".join(path.name for path in missing)
        )
    return IntakeResult(
        company_name=str(report["company_name"]),
        company_id=str(report["company_id"]),
        company_directory=company_root,
        moved_files=moved_files,
        unknown_files=tuple(Path(value) for value in report.get("unknown_files", [])),
        report_path=report_path,
    )


def run_full_pipeline(
    *,
    project_root: Path,
    input_directory: Path | None = None,
    existing_company_directory: Path | None = None,
    start_stage: int = 1,
    review_completed: bool = False,
) -> FullPipelineResult | BasicInfoReviewResult:
    if start_stage not in (1, 3, 5):
        raise ValueError("start_stage는 1, 3 또는 5여야 합니다.")
    if start_stage in (3, 5) and existing_company_directory is None:
        raise ValueError(f"{start_stage}단계 재개에는 existing_company_directory가 필요합니다.")
    if start_stage == 5 and not review_completed:
        raise ValueError("5단계 재개에는 기초정보 사람 검토 완료 확인이 필요합니다.")

    root = Path(project_root).resolve(strict=True)
    resources = root / "Resources"
    template = resources / "Template" / "Temp_기초정보.xlsx"
    withholding_tool_template = resources / "Template" / "Temp_수정 원천징수영수증 생성 Tool.xlsx"
    for required in (
        template,
        withholding_tool_template,
        resources / "지급항목_매칭표.xlsx",
        resources / "직업분류_연계표.xlsx",
    ):
        if not required.is_file():
            raise SchemaMismatch(f"통합 실행 필수 리소스가 없습니다: {required}")

    if existing_company_directory is None:
        input_root = Path(input_directory or root / "input").resolve(strict=True)
        logger.info("전체 실행 [1/8] 입력 접수 및 거래처 폴더 생성 시작")
        intake = organize_input_batch(input_directory=input_root, project_root=root)
        logger.info("전체 실행 [1/8] 완료: %s", intake.company_name)
    else:
        logger.info("전체 실행 [1/8] 기존 입력 재개 시작")
        intake = _load_existing_intake(
            existing_company_directory,
            validate_moved_files=start_stage == 1,
        )
        logger.info("전체 실행 [1/8] 기존 입력 재개 완료: %s", intake.company_name)
    company_root = intake.company_directory
    if start_stage == 1:
        source_input = company_root / "1. 급상여현황 생성" / "1.Input"
        logger.info("전체 실행 [2/8] 원천징수 분할·추출 및 기초정보·급여대장 생성 시작")
        conversion = PayrollConverter().convert_company_input(
            input_root=source_input,
            output_directory=company_root,
            basic_info_template=template,
            company_name=intake.company_name,
        )
        logger.info("전체 실행 [2/8] 완료: 엑셀 %d개", len(conversion.output_files))
        basic_files = [
            path for path in conversion.output_files
            if path.name.endswith("_기초정보.xlsx")
        ]
        payroll_outputs = [
            path for path in conversion.output_files
            if path.name.endswith("년 급여대장.xlsx")
        ]
    else:
        logger.info("전체 실행 [1/8~2/8] 기존 산출물 재사용")
        basic_files = sorted(
            path for path in company_root.glob("*_기초정보.xlsx")
            if not path.name.startswith("~$")
        )
        payroll_outputs = sorted(
            path for path in company_root.glob("*년 급여대장.xlsx")
            if not path.name.startswith("~$")
        )
        conversion_report = company_root / "trans_검증보고서.json"
        if not conversion_report.is_file():
            raise SchemaMismatch(f"기존 변환 검증보고서가 없습니다: {conversion_report}")
        conversion = UniversalConversionResult(
            company_root,
            tuple((*basic_files, *payroll_outputs)),
            conversion_report,
        )
    if len(basic_files) != 1:
        raise SchemaMismatch(f"기초정보 결과가 {len(basic_files)}개입니다: {company_root}")
    if start_stage == 1 and not payroll_outputs:
        raise SchemaMismatch(f"연도별 급여대장 결과가 없습니다: {company_root}")

    basic_info = basic_files[0]
    if start_stage != 5:
        logger.info("전체 실행 [3/8] 직업분류코드 및 건보정보 적재 시작")
        employment_info, health_roster = discover_employee_info_sources(company_root)
        if start_stage == 3 and (
            employment_info is None or "직업분류매핑결과" not in employment_info.name
        ):
            raise SchemaMismatch(
                f"파일명에 '직업분류매핑결과'가 포함된 엑셀 파일이 없습니다: {company_root}"
            )
        employee_info = populate_basic_info_employee_sheets(
            basic_info_path=basic_info,
            employment_info_path=employment_info,
            health_roster_path=health_roster,
            occupation_linkage_path=resources / "직업분류_연계표.xlsx",
        )
        logger.info(
            "전체 실행 [3/8] 완료: 직업분류 입력=%s, 매핑완료=%d, 수동확인=%d",
            employee_info.employment_source,
            employee_info.mapped_rows,
            employee_info.manual_review_rows,
        )
    else:
        logger.info("전체 실행 [1/8~4/8] 사람 검토 완료 기초정보·기존 급여대장 재사용")

    payroll_input = company_root / "2.근로소득 수정신고" / "1.Input"
    payroll_input.mkdir(parents=True, exist_ok=True)
    payroll_files: list[Path] = []
    if payroll_outputs:
        logger.info("전체 실행 [4/8] 파일 이동 시작 (총 %d개)", len(payroll_outputs))
        try:
            for source in payroll_outputs:
                destination = payroll_input / source.name
                os.replace(source, destination)
                payroll_files.append(destination)
        except Exception as exc:
            logger.error("전체 실행 [4/8] 파일 이동 오류: %s", exc)
            raise
        logger.info("전체 실행 [4/8] 파일 이동 완료: %d개", len(payroll_files))
    else:
        payroll_files = sorted(
            path for path in payroll_input.glob("*년 급여대장.xlsx")
            if not path.name.startswith("~$")
        )
        if not payroll_files:
            raise SchemaMismatch(f"기존 연도별 급여대장이 없습니다: {payroll_input}")
        logger.info("전체 실행 [4/8] 기존 급여대장 재사용: %d개", len(payroll_files))

    review_report_path = company_root / "1. 급상여현황 생성" / "기초정보_검토대기.json"
    if start_stage in (1, 3):
        preparation = prepare_company_basic_info_for_review(company_root, resources)
        review_report = {
            "status": "AWAITING_HUMAN_REVIEW",
            "prepared_at": datetime.now().astimezone().isoformat(),
            "company_name": intake.company_name,
            "company_directory": str(company_root),
            "basic_info": str(basic_info),
            "payroll_files": [str(path) for path in payroll_files],
            "years": list(preparation.years),
            "fixed_salary_items": preparation.fixed_count,
            "review_allowance_items": preparation.review_count,
            "next_start_stage": 5,
        }
        _write_json_report(review_report_path, review_report)
        logger.info("기초정보 사람 검토 대기: %s", basic_info)
        return BasicInfoReviewResult(
            intake,
            conversion,
            preparation,
            basic_info,
            tuple(payroll_files),
            review_report_path,
        )

    if not review_report_path.is_file():
        raise SchemaMismatch(f"기초정보 검토 대기 보고서가 없습니다: {review_report_path}")
    review_report = json.loads(review_report_path.read_text(encoding="utf-8-sig"))
    review_report["status"] = "HUMAN_REVIEW_COMPLETED"
    review_report["review_completed_at"] = datetime.now().astimezone().isoformat()
    _write_json_report(review_report_path, review_report)

    logger.info("전체 실행 [5/8] 급상여현황 생성 시작")
    payroll_status = generate_company_payroll_status(
        company_root,
        resources,
        use_reviewed_item_markers=True,
    )
    logger.info("전체 실행 [5/8] 완료: %d개", len(payroll_status.output_files))

    logger.info("전체 실행 [6/8] 수정신고 입력자료 생성 시작 (총 %d개 연도)", len(payroll_status.years))
    correction_inputs = prepare_correction_inputs(
        company_name=intake.company_name,
        basic_info_path=basic_info,
        payroll_files=tuple(payroll_files),
        payroll_status_files=payroll_status.output_files,
        tool_template_path=withholding_tool_template,
        modified_input_directory=company_root / "Input_급상여현황(수정)",
        correction_input_directory=payroll_input,
        withholding_pdf_root=company_root / "1. 급상여현황 생성" / "원천징수(PDF)",
    )
    logger.info(
        "전체 실행 [6/8] 수정신고 입력자료 생성 완료: 수정 급상여현황 %d개, Tool %d개, 월별 급여대장 %d개",
        len(correction_inputs.modified_status_files),
        len(correction_inputs.withholding_tool_files),
        len(correction_inputs.monthly_payroll_files),
    )
    amended_withholding = None
    if correction_inputs.withholding_tool_files:
        logger.info("전체 실행 [6/8] 수정후·변경표시 원천징수 PDF 생성 시작")
        amended_withholding = generate_company_amended_withholding(company_root)
        normal_count = len(amended_withholding.normal_files)
        marked_count = len(amended_withholding.marked_files)
        logger.info(
            "전체 실행 [6/8] 수정 원천징수 PDF 생성 완료: 총 %d개 (수정후 %d개, 변경표시 %d개, 이미지 PDF 제외 %d개)",
            normal_count + marked_count,
            normal_count,
            marked_count,
            amended_withholding.excluded_image_pdfs,
        )
    logger.info("전체 실행 [7/8] 국세청·공단 제출자료 생성 시작")
    submission_targets = generate_company_submission_target_lists(company_root)
    logger.info(
        "전체 실행 [7/8] 국세청·공단 제출자료 생성 완료: 국세청 %d개, 공단 %d개",
        len(submission_targets.nts_files),
        len(submission_targets.insurance_files),
    )
    report = {
        "completed_at": datetime.now().astimezone().isoformat(),
        "company_name": intake.company_name,
        "company_id": intake.company_id,
        "company_directory": str(company_root),
        "start_stage": start_stage,
        "intake_report": str(intake.report_path),
        "conversion_report": str(conversion.report_path),
        "basic_info": str(basic_info),
        "payroll_files": [str(path) for path in payroll_files],
        "payroll_status_files": [str(path) for path in payroll_status.output_files],
        "payroll_status_report": str(payroll_status.report_path),
        "modified_payroll_status_files": [str(path) for path in correction_inputs.modified_status_files],
        "withholding_tool_files": [str(path) for path in correction_inputs.withholding_tool_files],
        "monthly_payroll_files": [str(path) for path in correction_inputs.monthly_payroll_files],
        "correction_input_issues": list(correction_inputs.issues),
        "amended_withholding_files": (
            [str(path) for path in amended_withholding.normal_files]
            if amended_withholding else []
        ),
        "marked_withholding_files": (
            [str(path) for path in amended_withholding.marked_files]
            if amended_withholding else []
        ),
        "amended_withholding_report": (
            str(amended_withholding.report_path) if amended_withholding else None
        ),
        "submission_years": list(submission_targets.years),
        "nts_submission_files": [str(path) for path in submission_targets.nts_files],
        "insurance_submission_files": [
            str(path) for path in submission_targets.insurance_files
        ],
        "submission_report": str(submission_targets.report_path),
        "submission_issues": list(submission_targets.issues),
    }
    report_path = company_root / "1. 급상여현황 생성" / "전체실행_결과.json"
    _write_json_report(report_path, report)
    logger.info("전체 실행 [8/8] 결과 보고서 저장 완료: %s", report_path)
    return FullPipelineResult(
        intake,
        conversion,
        payroll_status,
        basic_info,
        tuple(payroll_files),
        report_path,
        correction_inputs,
        amended_withholding,
        submission_targets,
    )
