from __future__ import annotations

import argparse
from datetime import datetime
import json
import logging
from pathlib import Path
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[1]
SRC_ROOT = PROJECT_ROOT / "src"
if str(SRC_ROOT) not in sys.path:
    sys.path.insert(0, str(SRC_ROOT))

from payroll_amendment.full_pipeline import BasicInfoReviewResult, run_full_pipeline


_PYPDF_WRONG_POINTER_PREFIX = "Ignoring wrong pointing object "


class _PypdfWarningFilter(logging.Filter):
    def __init__(self) -> None:
        super().__init__()
        self._wrong_pointer_reported = False

    def filter(self, record: logging.LogRecord) -> bool:
        message = record.getMessage()
        if record.name != "pypdf._reader" or not message.startswith(_PYPDF_WRONG_POINTER_PREFIX):
            return True
        if self._wrong_pointer_reported:
            return False
        self._wrong_pointer_reported = True
        record.msg = "PDF 내부 객체 참조 이상을 감지했습니다. 이후 동일 경고는 생략합니다. (%s)"
        record.args = (message,)
        return True


def _configure_logging(project_root: Path) -> Path:
    log_directory = project_root / "logs"
    log_directory.mkdir(parents=True, exist_ok=True)
    log_path = log_directory / f"full_pipeline_{datetime.now():%Y%m%d_%H%M%S}.log"
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    stream = logging.StreamHandler(sys.stdout)
    stream.setFormatter(formatter)
    file_handler = logging.FileHandler(log_path, encoding="utf-8")
    file_handler.setFormatter(formatter)
    warning_filter = _PypdfWarningFilter()
    stream.addFilter(warning_filter)
    file_handler.addFilter(warning_filter)
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    root_logger.handlers.clear()
    root_logger.addHandler(stream)
    root_logger.addHandler(file_handler)
    return log_path


def main() -> int:
    parser = argparse.ArgumentParser(
        description="기초정보 생성 후 사람 검토에서 멈추고, 검토 완료 후 급상여현황부터 재개합니다."
    )
    parser.add_argument("--project-root", type=Path, default=PROJECT_ROOT)
    parser.add_argument("--input-dir", type=Path)
    parser.add_argument("--resume-company-dir", type=Path)
    parser.add_argument(
        "--start-stage",
        type=int,
        choices=(1, 3, 5),
        default=1,
        help="1은 신규 기초정보 생성, 3은 직업분류 적재부터 검토 준비, 5는 검토 완료 후 급상여현황부터 재개합니다.",
    )
    parser.add_argument(
        "--review-completed",
        action="store_true",
        help="기초정보 사람 검토가 완료됐음을 확인합니다. --start-stage 5에 필수입니다.",
    )
    args = parser.parse_args()
    if args.start_stage in (3, 5) and args.resume_company_dir is None:
        parser.error(f"--start-stage {args.start_stage}에는 --resume-company-dir이 필요합니다.")
    if args.start_stage == 5 and not args.review_completed:
        parser.error("--start-stage 5에는 --review-completed가 필요합니다.")
    project_root = args.project_root.resolve(strict=True)
    resume_company_directory = args.resume_company_dir
    if resume_company_directory is not None and not resume_company_directory.is_absolute():
        resume_company_directory = project_root / resume_company_directory
    log_path = _configure_logging(project_root)
    logging.info("통합 실행 시작")
    logging.info("진행 로그 파일: %s", log_path)
    try:
        result = run_full_pipeline(
            project_root=project_root,
            input_directory=args.input_dir,
            existing_company_directory=resume_company_directory,
            start_stage=args.start_stage,
            review_completed=args.review_completed,
        )
    except Exception:
        logging.exception("통합 실행 실패")
        raise
    if isinstance(result, BasicInfoReviewResult):
        logging.info("기초정보 생성 완료. 사람 검토 후 5단계로 재개하십시오.")
        print(
            json.dumps(
                {
                    "status": "AWAITING_HUMAN_REVIEW",
                    "company_name": result.intake.company_name,
                    "company_directory": str(result.intake.company_directory),
                    "basic_info": str(result.basic_info_path),
                    "payroll_files": [str(path) for path in result.payroll_files],
                    "review_report": str(result.report_path),
                    "next_start_stage": 5,
                    "log_file": str(log_path),
                },
                ensure_ascii=False,
                indent=2,
            )
        )
        return 0
    logging.info("통합 실행 완료")
    print(
        json.dumps(
            {
                "company_name": result.intake.company_name,
                "company_directory": str(result.intake.company_directory),
                "basic_info": str(result.basic_info_path),
                "payroll_files": [str(path) for path in result.payroll_files],
                "payroll_status_files": [str(path) for path in result.payroll_status.output_files],
                "modified_payroll_status_files": [
                    str(path) for path in result.correction_inputs.modified_status_files
                ],
                "withholding_tool_files": [
                    str(path) for path in result.correction_inputs.withholding_tool_files
                ],
                "monthly_payroll_files": [
                    str(path) for path in result.correction_inputs.monthly_payroll_files
                ],
                "amended_withholding_count": (
                    len(result.amended_withholding.normal_files)
                    if result.amended_withholding else 0
                ),
                "marked_withholding_count": (
                    len(result.amended_withholding.marked_files)
                    if result.amended_withholding else 0
                ),
                "nts_submission_files": [
                    str(path) for path in result.submission_targets.nts_files
                ],
                "insurance_submission_files": [
                    str(path) for path in result.submission_targets.insurance_files
                ],
                "submission_report": str(result.submission_targets.report_path),
                "report": str(result.report_path),
                "log_file": str(log_path),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
