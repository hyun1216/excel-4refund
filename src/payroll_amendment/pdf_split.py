from __future__ import annotations

from dataclasses import dataclass
import json
import os
from pathlib import Path
import re
import shutil
from typing import Literal, Protocol
from uuid import uuid4

from pypdf import PdfReader, PdfWriter

from .errors import PdfSplitError
from .infra import safe_name, sha256
from .models import TaxExemption


FIRST_PAGE_MARKER = "소득세법시행규칙[별지제24호서식(1)]"
FALLBACK_FIRST_PAGE_MARKERS = ("근로소득원천징수영수증", "1.법인명(상호)")
_SOCIAL_NUMBER = re.compile(r"(?<!\d)(\d{6})\s*-?\s*(\d{7})(?!\d)")


@dataclass(frozen=True)
class PdfInspection:
    path: Path
    kind: Literal["individual", "bundle"]
    page_count: int
    first_pages: tuple[int, ...]
    page_texts: tuple[str, ...]


@dataclass(frozen=True)
class PreparedPdfBatch:
    output_files: list[Path]
    source_files: list[Path]
    manifest_path: Path


class WithholdingPdfIdentity(Protocol):
    pdf_index: int
    employee_name: str
    resident_id: str


def _compact(value: str) -> str:
    return re.sub(r"\s+", "", value or "")


def _first_pages(page_texts: list[str] | tuple[str, ...]) -> tuple[int, ...]:
    marker = _compact(FIRST_PAGE_MARKER)
    compacted = tuple(_compact(text) for text in page_texts)
    starts = tuple(index for index, text in enumerate(compacted) if marker in text)
    if starts:
        return starts
    fallback_markers = tuple(_compact(value) for value in FALLBACK_FIRST_PAGE_MARKERS)
    return tuple(
        index
        for index, text in enumerate(compacted)
        if all(value in text for value in fallback_markers)
    )


def classify_page_texts(page_texts: list[str] | tuple[str, ...]) -> tuple[str, tuple[int, ...]]:
    if not page_texts:
        raise PdfSplitError("페이지가 없는 PDF입니다.")
    starts = _first_pages(page_texts)
    if not starts:
        raise PdfSplitError(f"원천징수영수증 첫 페이지 표식을 찾지 못했습니다: {FIRST_PAGE_MARKER}")
    if starts[0] != 0:
        raise PdfSplitError(f"PDF 첫 페이지가 원천징수영수증 서식으로 시작하지 않습니다: {starts[0] + 1}쪽")
    return ("bundle" if len(starts) > 1 else "individual"), starts


def inspect_withholding_pdf(path: Path) -> PdfInspection:
    source = Path(path).resolve(strict=True)
    try:
        reader = PdfReader(str(source))
        if reader.is_encrypted:
            raise PdfSplitError("암호화된 PDF는 분할할 수 없습니다.")
        texts = tuple(page.extract_text() or "" for page in reader.pages)
    except PdfSplitError:
        raise
    except Exception as exc:
        raise PdfSplitError(f"PDF를 읽을 수 없습니다: {source.name}: {exc}") from exc
    kind, starts = classify_page_texts(texts)
    return PdfInspection(source, kind, len(texts), starts, texts)


def _digits(value: object) -> str:
    return re.sub(r"\D", "", str(value or ""))


def _resident_numbers(page_text: str) -> set[str]:
    values = {f"{match.group(1)}{match.group(2)}" for match in _SOCIAL_NUMBER.finditer(page_text or "")}
    return {value for value in values if not value.startswith("110111")}


def _pdf_index(path: Path) -> int | None:
    match = re.search(r"_(\d+)\.pdf$", path.name, re.IGNORECASE)
    if match:
        return int(match.group(1))
    match = re.match(r"(\d+)_", path.name)
    return int(match.group(1)) if match else None


def _target_from_filename(path: Path, exemptions: list[TaxExemption]) -> TaxExemption | None:
    parts = {_compact(part) for part in path.stem.split("_")}
    matches = [
        exemption
        for exemption in exemptions
        if _compact(exemption.employee_id) in parts and _compact(exemption.name) in parts
    ]
    if len(matches) > 1:
        raise PdfSplitError(f"개인 PDF 파일명에서 대상자가 둘 이상 검출됐습니다: {path.name}")
    return matches[0] if matches else None


def _target_maps(exemptions: list[TaxExemption]) -> tuple[dict[int, TaxExemption], dict[str, TaxExemption]]:
    by_index: dict[int, TaxExemption] = {}
    by_resident: dict[str, TaxExemption] = {}
    for exemption in exemptions:
        if not str(exemption.pdf_index).isdigit():
            raise PdfSplitError(f"숫자가 아닌 PDF 인덱스입니다: {exemption.pdf_index}")
        index = int(exemption.pdf_index)
        if index in by_index:
            raise PdfSplitError(f"PDF 분할 대상 인덱스가 중복됩니다: {index}")
        by_index[index] = exemption
        resident = _digits(exemption.resident_id)
        if resident:
            if len(resident) != 13:
                raise PdfSplitError(f"주민번호 형식이 올바르지 않습니다: 인덱스 {index}")
            if resident in by_resident:
                raise PdfSplitError(f"PDF 분할 대상 주민번호가 중복됩니다: 인덱스 {index}")
            by_resident[resident] = exemption
    if not by_index:
        raise PdfSplitError("PDF 분할 대상자가 없습니다.")
    return by_index, by_resident


def _validate_employee_page(page_text: str, exemption: TaxExemption) -> None:
    expected = _digits(exemption.resident_id)
    found = _resident_numbers(page_text)
    if expected and found and expected not in found:
        raise PdfSplitError(
            f"원천징수영수증과 기초정보의 주민번호가 다릅니다: 인덱스 {exemption.pdf_index}"
        )


def withholding_pdf_output_name(pdf_index: int | str, employee_name: str) -> str:
    return safe_name(f"{int(pdf_index)}_{employee_name}_발행자보관용.pdf")


def _output_name(company_name: str, year: int, exemption: TaxExemption) -> str:
    return withholding_pdf_output_name(exemption.pdf_index, exemption.name)


def _write_pages(reader: PdfReader, page_indexes: range, target: Path) -> None:
    writer = PdfWriter()
    indexes = list(page_indexes)
    for page_index in indexes:
        writer.add_page(reader.pages[page_index])
    with target.open("wb") as stream:
        writer.write(stream)
    try:
        if len(PdfReader(str(target)).pages) != len(indexes):
            raise PdfSplitError(f"분할 PDF 페이지 수 검증에 실패했습니다: {target.name}")
    except Exception:
        target.unlink(missing_ok=True)
        raise


def _quarantine(paths: list[Path], quarantine_dir: Path, error: Exception) -> None:
    quarantine_dir.mkdir(parents=True, exist_ok=True)
    copied = []
    for number, source in enumerate(dict.fromkeys(paths), 1):
        if not source.is_file():
            continue
        target = quarantine_dir / f"{number:03d}_{safe_name(source.name)}"
        shutil.copy2(source, target)
        copied.append(target.name)
    report = {
        "status": "QUARANTINED",
        "error_code": getattr(error, "code", "E_PDF_SPLIT"),
        "message": str(error),
        "files": copied,
    }
    (quarantine_dir / "failure.json").write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def archive_prepared_withholding_pdfs(
    *,
    pdf_files: list[Path],
    identities: list[WithholdingPdfIdentity],
    year: int,
    output_dir: Path,
) -> PreparedPdfBatch:
    sources = [Path(path).resolve(strict=True) for path in pdf_files]
    source_by_index: dict[int, Path] = {}
    for source in sources:
        index = _pdf_index(source)
        if index is None:
            raise PdfSplitError(f"개인 원천징수 PDF 인덱스를 찾지 못했습니다: {source.name}")
        if index in source_by_index:
            raise PdfSplitError(f"개인 원천징수 PDF 인덱스가 중복됐습니다: {index}")
        source_by_index[index] = source

    identity_by_index: dict[int, WithholdingPdfIdentity] = {}
    for identity in identities:
        index = int(identity.pdf_index)
        if index in identity_by_index:
            raise PdfSplitError(f"원천징수 추출 인덱스가 중복됐습니다: {index}")
        identity_by_index[index] = identity

    if set(source_by_index) != set(identity_by_index):
        missing = sorted(set(identity_by_index) - set(source_by_index))
        extra = sorted(set(source_by_index) - set(identity_by_index))
        raise PdfSplitError(f"개인 원천징수 PDF와 추출 대상이 다릅니다: 누락={missing}, 초과={extra}")

    destination = Path(output_dir).resolve()
    staging = destination.parent / f".{destination.name}.{uuid4().hex}.tmp"
    records: list[dict[str, object]] = []
    try:
        staging.mkdir(parents=True, exist_ok=False)
        outputs: list[Path] = []
        for index, identity in sorted(identity_by_index.items()):
            source = source_by_index[index]
            reader = PdfReader(str(source))
            if reader.is_encrypted:
                raise PdfSplitError(f"암호화된 PDF는 보관할 수 없습니다: {source.name}")
            _validate_employee_page(reader.pages[0].extract_text() or "", identity)
            target = staging / withholding_pdf_output_name(index, identity.employee_name)
            shutil.copy2(source, target)
            if sha256(source) != sha256(target):
                raise PdfSplitError(f"개인 원천징수 PDF 복사 검증에 실패했습니다: {source.name}")
            outputs.append(target)
            records.append(
                {
                    "pdf_index": index,
                    "employee_name": identity.employee_name,
                    "source": source.name,
                    "output": target.name,
                    "pages": len(reader.pages),
                    "sha256": sha256(target),
                }
            )

        manifest_path = staging / "pdf_preparation_manifest.json"
        manifest_path.write_text(
            json.dumps({"year": year, "outputs": records}, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        destination.parent.mkdir(parents=True, exist_ok=True)
        if destination.exists():
            raise PdfSplitError(f"PDF 보관 출력 폴더가 이미 존재합니다: {destination}")
        os.replace(staging, destination)
        return PreparedPdfBatch(
            [destination / path.name for path in outputs],
            sources,
            destination / manifest_path.name,
        )
    except Exception as exc:
        shutil.rmtree(staging, ignore_errors=True)
        if isinstance(exc, PdfSplitError):
            raise
        raise PdfSplitError(str(exc)) from exc


def prepare_withholding_pdfs(
    *,
    pdf_files: list[Path],
    exemptions: list[TaxExemption],
    company_name: str,
    year: int,
    output_dir: Path,
    quarantine_dir: Path,
) -> PreparedPdfBatch:
    sources = [Path(path).resolve(strict=True) for path in dict.fromkeys(pdf_files)]
    by_index, by_resident = _target_maps(exemptions)
    staging = output_dir.parent / f".{output_dir.name}.{uuid4().hex}.tmp"
    used_sources: list[Path] = []
    records: list[dict[str, object]] = []
    try:
        inspections: list[PdfInspection] = []
        ignored: list[str] = []
        target_indexes = set(by_index)
        for source in sources:
            try:
                inspections.append(inspect_withholding_pdf(source))
            except PdfSplitError:
                if _pdf_index(source) in target_indexes or "원천징수" in str(source):
                    raise
                ignored.append(source.name)

        bundles = [item for item in inspections if item.kind == "bundle"]
        individuals = [item for item in inspections if item.kind == "individual"]
        if bundles and individuals:
            raise PdfSplitError("통합본과 개인별 원천징수 PDF가 함께 들어 있습니다.")
        if len(bundles) > 1:
            raise PdfSplitError("같은 연도에 통합 원천징수 PDF가 두 개 이상입니다.")
        if not inspections:
            raise PdfSplitError(f"{year}년 원천징수영수증 PDF를 찾지 못했습니다.")

        staging.mkdir(parents=True, exist_ok=False)
        produced: dict[int, Path] = {}
        if bundles:
            inspection = bundles[0]
            reader = PdfReader(str(inspection.path))
            starts = inspection.first_pages
            for index, exemption in sorted(by_index.items()):
                if index < 1 or index > len(starts):
                    raise PdfSplitError(
                        f"대상 인덱스가 통합 PDF 사원 수를 벗어납니다: 인덱스={index}, 사원={len(starts)}"
                    )
                start = starts[index - 1]
                end = starts[index] if index < len(starts) else inspection.page_count
                _validate_employee_page(inspection.page_texts[start], exemption)
                target = staging / _output_name(company_name, year, exemption)
                _write_pages(reader, range(start, end), target)
                produced[index] = target
                records.append({
                    "pdf_index": index,
                    "source": inspection.path.name,
                    "source_pages": [start + 1, end],
                    "output": target.name,
                    "sha256": sha256(target),
                })
            used_sources.append(inspection.path)
            input_kind = "bundle"
        else:
            for inspection in individuals:
                index = _pdf_index(inspection.path)
                exemption = by_index.get(index) if index is not None else None
                if exemption is None:
                    exemption = _target_from_filename(inspection.path, exemptions)
                    index = int(exemption.pdf_index) if exemption else None
                if exemption is None:
                    matches = {
                        by_resident[value].pdf_index: by_resident[value]
                        for value in _resident_numbers(inspection.page_texts[0])
                        if value in by_resident
                    }
                    if len(matches) > 1:
                        raise PdfSplitError(f"개인 PDF에서 대상자가 둘 이상 검출됐습니다: {inspection.path.name}")
                    exemption = next(iter(matches.values()), None)
                    index = int(exemption.pdf_index) if exemption else None
                if exemption is None or index is None:
                    ignored.append(inspection.path.name)
                    continue
                if index in produced:
                    raise PdfSplitError(f"동일 대상자의 개인 PDF가 중복됩니다: 인덱스 {index}")
                _validate_employee_page(inspection.page_texts[0], exemption)
                target = staging / _output_name(company_name, year, exemption)
                shutil.copy2(inspection.path, target)
                if len(PdfReader(str(target)).pages) != inspection.page_count:
                    raise PdfSplitError(f"개인 PDF 복사 검증에 실패했습니다: {inspection.path.name}")
                produced[index] = target
                used_sources.append(inspection.path)
                records.append({
                    "pdf_index": index,
                    "source": inspection.path.name,
                    "source_pages": [1, inspection.page_count],
                    "output": target.name,
                    "sha256": sha256(target),
                })
            input_kind = "individual"

        missing = sorted(set(by_index) - set(produced))
        if missing:
            raise PdfSplitError(f"대상자의 원천징수 PDF가 누락됐습니다: 인덱스 {missing}")

        manifest = {
            "year": year,
            "input_kind": input_kind,
            "first_page_marker": FIRST_PAGE_MARKER,
            "source_files": [path.name for path in used_sources],
            "ignored_files": sorted(set(ignored)),
            "outputs": records,
        }
        manifest_path = staging / "pdf_preparation_manifest.json"
        manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
        output_dir.parent.mkdir(parents=True, exist_ok=True)
        if output_dir.exists():
            raise PdfSplitError(f"PDF 준비 출력 폴더가 이미 존재합니다: {output_dir}")
        os.replace(staging, output_dir)
        return PreparedPdfBatch(
            [output_dir / path.name for _, path in sorted(produced.items())],
            used_sources,
            output_dir / manifest_path.name,
        )
    except Exception as exc:
        shutil.rmtree(staging, ignore_errors=True)
        _quarantine(sources, quarantine_dir, exc)
        if isinstance(exc, PdfSplitError):
            raise
        raise PdfSplitError(str(exc)) from exc
