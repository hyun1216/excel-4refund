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

from payroll_amendment.full_pipeline import run_full_pipeline


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
        description="공용 input 접수부터 기초정보·급여대장·급상여현황 생성까지 실행합니다."
    )
    parser.add_argument("--project-root", type=Path, default=PROJECT_ROOT)
    parser.add_argument("--input-dir", type=Path)
    parser.add_argument("--resume-company-dir", type=Path)
    args = parser.parse_args()
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
        )
    except Exception:
        logging.exception("통합 실행 실패")
        raise
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
