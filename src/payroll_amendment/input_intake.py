from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import json
import logging
import os
from pathlib import Path
import re
import shutil
from typing import Callable
from uuid import uuid4

from openpyxl import load_workbook

from .errors import SchemaMismatch
from .infra import safe_name, sha256
from .payroll_converter import PayrollConverter


logger = logging.getLogger(__name__)


EMPLOYMENT_NAME = "근로자고용정보현황조회"
HEALTH_NAME = "사업장가입자명부"
PAYROLL_KEYWORDS = {
    "성명",
    "이름",
    "사원번호",
    "기본급",
    "지급총액",
    "지급합계",
    "실지급액",
    "차인지급액",
}

YEAR_SCOPED_KINDS = {
    "tabular_payroll",
    "withholding_receipt",
    "withholding_ledger",
}
YEAR_DIRECTORY_PATTERN = re.compile(r"^(?P<year>20\d{2}|\d{2})\s*(?:년도?|년)?$")


@dataclass(frozen=True)
class IntakeFile:
    source: Path
    kind: str
    destination: Path


@dataclass(frozen=True)
class IntakeResult:
    company_name: str
    company_id: str
    company_directory: Path
    moved_files: tuple[Path, ...]
    unknown_files: tuple[Path, ...]
    report_path: Path


def _classify_pdf(path: Path) -> str:
    return PayrollConverter._document_signature(path)


def _read_company_metadata(path: Path) -> tuple[str, str]:
    return PayrollConverter._company_metadata(path)


def _looks_like_payroll_workbook(path: Path) -> bool:
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return False
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
    except Exception:
        return False
    try:
        for sheet in workbook.worksheets:
            for values in sheet.iter_rows(min_row=1, max_row=15, values_only=True):
                normalized = {re.sub(r"\s+", "", str(value or "")) for value in values}
                if len(normalized & PAYROLL_KEYWORDS) >= 2:
                    return True
    finally:
        workbook.close()
    return False


def _classify_file(path: Path) -> str:
    name = path.name
    if EMPLOYMENT_NAME in name:
        return "employment_info"
    if HEALTH_NAME in name:
        return "health_roster"
    if path.suffix.lower() == ".pdf":
        return _classify_pdf(path)
    if _looks_like_payroll_workbook(path):
        return "tabular_payroll"
    return "unknown"


def _current_company_names(classified: list[tuple[Path, str]]) -> set[str]:
    names: set[str] = set()
    markers = {
        "employment_info": EMPLOYMENT_NAME,
        "health_roster": HEALTH_NAME,
    }
    for path, kind in classified:
        marker = markers.get(kind)
        if marker is None:
            continue
        company_name = path.name.partition(marker)[0].rstrip(" _-")
        if company_name:
            names.add(company_name)
    return names


def _create_company_directories(company_directory: Path, company_name: str) -> dict[str, Path]:
    income_tax = company_directory / "1. 급상여현황 생성"
    source_input = income_tax / "1.Input"
    correction = company_directory / "2.근로소득 수정신고"
    paths = {
        "company": company_directory,
        "tabular_payroll": source_input / "급여대장",
        "withholding_receipt": source_input / "원천징수영수증",
        "withholding_ledger": source_input / "소득자별근로소득원천징수부",
        "employment_info": company_directory,
        "health_roster": company_directory,
        "payroll_status": income_tax / "급상여현황",
        "withholding_archive": income_tax / "원천징수(PDF)",
        "salary_input": company_directory / "Input_급상여현황(수정)",
        "correction_input": correction / "1.Input",
        "correction_output": correction / "2.Output",
        "correction_withholding": correction / "2.Output" / "원천징수(PDF)",
        "nts_output": correction / "2.Output" / f"★국세청제출자료_{company_name}",
        "insurance_output": correction / "2.Output" / f"★공단제출자료_{company_name}",
    }
    for path in dict.fromkeys(paths.values()):
        path.mkdir(parents=True, exist_ok=True)
    return paths


def _year_directory_from_source(source: Path, input_root: Path) -> str | None:
    for part in source.parent.relative_to(input_root).parts:
        matched = YEAR_DIRECTORY_PATTERN.fullmatch(part.strip())
        if matched:
            year = matched.group("year")
            return f"{year[-2:]}년"
    return None


def _destination_for(
    kind: str,
    paths: dict[str, Path],
    source: Path,
    input_root: Path,
) -> Path:
    destination_root = paths.get(kind)
    if destination_root is None:
        raise SchemaMismatch(f"지원하지 않는 접수 파일 유형입니다: {kind}")
    if kind in YEAR_SCOPED_KINDS:
        year_directory = _year_directory_from_source(source, input_root)
        if year_directory:
            destination_root = destination_root / year_directory
    return destination_root / source.name


def organize_input_batch(
    *,
    input_directory: Path,
    project_root: Path,
    pdf_classifier: Callable[[Path], str] | None = None,
    company_metadata_reader: Callable[[Path], tuple[str, str]] | None = None,
) -> IntakeResult:
    input_root = Path(input_directory).resolve(strict=True)
    root = Path(project_root).resolve(strict=True)
    classifier = pdf_classifier or _classify_pdf
    metadata_reader = company_metadata_reader or _read_company_metadata
    files = sorted(
        path
        for path in input_root.rglob("*")
        if path.is_file()
        and not path.name.startswith(("~$", "."))
        and not any(part.startswith(".") for part in path.relative_to(input_root).parts)
    )
    if not files:
        raise SchemaMismatch(f"접수할 파일이 없습니다: {input_root}")
    logger.info("입력 파일 분류 시작 (총 %d개)", len(files))

    classified: list[tuple[Path, str]] = []
    unknown: list[Path] = []
    receipt_files: list[Path] = []
    try:
        for path in files:
            if path.suffix.lower() == ".pdf":
                kind = classifier(path)
            else:
                kind = _classify_file(path)
            if kind == "unknown":
                unknown.append(path)
                continue
            classified.append((path, kind))
            if kind == "withholding_receipt":
                receipt_files.append(path)
    except Exception as exc:
        logger.error("입력 파일 분류 오류: %s", exc)
        raise
    logger.info("입력 파일 분류 완료: %d개", len(files))
    if not receipt_files:
        raise SchemaMismatch("거래처명과 사업자번호를 확인할 원천징수영수증 PDF가 없습니다.")

    metadata = [metadata_reader(path) for path in receipt_files]
    company_ids = {company_id for _, company_id in metadata}
    if len(company_ids) != 1:
        raise SchemaMismatch(
            "한 접수 배치에 여러 거래처의 원천징수영수증이 섞여 있습니다: "
            + ", ".join(sorted(f"{name}({company_id})" for name, company_id in set(metadata)))
        )
    current_company_names = _current_company_names(classified)
    if len(current_company_names) > 1:
        raise SchemaMismatch(
            "고용정보와 사업장가입자명부의 현재 상호가 다릅니다: "
            + ", ".join(sorted(current_company_names))
        )
    company_name = next(iter(current_company_names), metadata[0][0])
    company_id = next(iter(company_ids))
    company_aliases = sorted({name for name, _ in metadata if name != company_name})
    company_directory = root / "노무_세무 수정신고" / safe_name(company_name)
    planned_roots = {
        "tabular_payroll": company_directory / "1. 급상여현황 생성" / "1.Input" / "급여대장",
        "withholding_receipt": company_directory / "1. 급상여현황 생성" / "1.Input" / "원천징수영수증",
        "withholding_ledger": company_directory / "1. 급상여현황 생성" / "1.Input" / "소득자별근로소득원천징수부",
        "employment_info": company_directory,
        "health_roster": company_directory,
    }
    plan = [
        IntakeFile(
            source,
            kind,
            _destination_for(kind, planned_roots, source, input_root),
        )
        for source, kind in classified
    ]
    duplicate_kinds = {
        kind for kind in ("employment_info", "health_roster")
        if sum(item.kind == kind for item in plan) > 1
    }
    if duplicate_kinds:
        raise SchemaMismatch("한 번에 하나만 허용되는 파일이 중복되었습니다: " + ", ".join(sorted(duplicate_kinds)))
    collisions = [item.destination for item in plan if item.destination.exists()]
    if collisions:
        raise SchemaMismatch("이동 대상에 같은 이름의 파일이 이미 있습니다: " + ", ".join(path.name for path in collisions))

    paths = _create_company_directories(company_directory, company_name)
    logger.info("거래처 폴더 생성 완료: %s", company_directory)
    moved: list[tuple[Path, Path]] = []
    logger.info("파일 이동 시작 (총 %d개)", len(plan))
    try:
        for item in plan:
            destination = item.destination
            destination.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(item.source), str(destination))
            moved.append((item.source, destination))
    except Exception as exc:
        for source, destination in reversed(moved):
            if destination.exists() and not source.exists():
                shutil.move(str(destination), str(source))
        logger.error("파일 이동 오류: %s", exc)
        raise
    logger.info("파일 이동 완료: %d개", len(moved))

    received_at = datetime.now().astimezone().isoformat()
    report = {
        "received_at": received_at,
        "input_directory": str(input_root),
        "company_name": company_name,
        "company_id": company_id,
        "company_aliases": company_aliases,
        "company_directory": str(company_directory),
        "moved_files": [
            {
                "kind": item.kind,
                "source_name": item.source.name,
                "source_relative": str(item.source.relative_to(input_root)),
                "destination": str(item.destination),
                "size": item.destination.stat().st_size,
                "sha256": sha256(item.destination),
            }
            for item in plan
        ],
        "unknown_files": [str(path.relative_to(input_root)) for path in unknown],
        "status": "NEEDS_REVIEW" if unknown else "SUCCEEDED",
    }
    report_directory = paths["company"] / "1. 급상여현황 생성"
    report_path = report_directory / "input_접수보고서.json"
    temporary = report_path.with_name(f".{report_path.name}.{uuid4().hex}.tmp")
    try:
        temporary.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8-sig")
        os.replace(temporary, report_path)
    except Exception:
        for source, destination in reversed(moved):
            if destination.exists() and not source.exists():
                shutil.move(str(destination), str(source))
        raise
    finally:
        temporary.unlink(missing_ok=True)
    for directory in sorted(
        (path for path in input_root.rglob("*") if path.is_dir()),
        key=lambda path: len(path.parts),
        reverse=True,
    ):
        try:
            directory.rmdir()
        except OSError:
            pass
    return IntakeResult(
        company_name,
        company_id,
        company_directory,
        tuple(destination for _, destination in moved),
        tuple(unknown),
        report_path,
    )
