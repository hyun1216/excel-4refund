from __future__ import annotations

import json
from pathlib import Path
import re
import zipfile
import xml.etree.ElementTree as ET

from .basic_info import parse_basic_info
from .errors import SourceNotFound
from .infra import sha256, snapshot_file
from .models import IngestedInputs
from .pdf_split import prepare_withholding_pdfs
from .withholding_pdf_excel import convert_withholding_pdfs_to_excel
from .year_parser import infer_year
from openpyxl import load_workbook


def _workbook_sheets(path: Path) -> set[str]:
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return set()
    try:
        # 시트명 선별에는 셀 데이터를 읽을 필요가 없다. workbook.xml만 읽으면
        # 대용량 결과 파일도 수초 안에 분류할 수 있다.
        with zipfile.ZipFile(path) as archive:
            root = ET.fromstring(archive.read("xl/workbook.xml"))
        return {element.attrib["name"] for element in root.iter() if element.tag.endswith("}sheet") and "name" in element.attrib}
    except Exception:
        return set()


def _pdf_year(path: Path) -> int | None:
    # 파일명 끝의 _10, _25 같은 값은 PDF 인덱스이므로 연도로 해석하지 않는다.
    for part in reversed(path.parent.parts):
        year = infer_year(part)
        if year:
            return year
        # CP949 이름이 ZIP에서 잘못 디코딩된 경우에도 '25년'의 숫자 부분은
        # 보존된다. 짧은 연도 폴더에 한해 2자리 연도를 복구한다.
        if len(part) <= 8:
            short = re.search(r"(?<!\d)(\d{2})(?!\d)", part)
            if short and 20 <= int(short.group(1)) <= 99:
                return 2000 + int(short.group(1))
    match = re.search(r"(?<!\d)(20\d{2})(?:\d{2})?(?!\d)", path.name)
    return int(match.group(1)) if match else None


def _data_row_count(path: Path) -> int:
    try:
        book = load_workbook(path, read_only=True, data_only=True)
        sheet = book["Data"]
        count = sum(1 for row in sheet.iter_rows(min_row=2, max_row=5000, values_only=True) if any(value not in (None, "") for value in row))
        book.close()
        return count
    except Exception:
        return 0


def ingest_company_inputs(*, source_dir: Path, years: list[int], run_dir: Path, company_name: str | None = None) -> IngestedInputs:
    if not source_dir.is_dir():
        raise SourceNotFound(f"입력 폴더가 없습니다: {source_dir}")
    files = [p for p in source_dir.rglob("*") if p.is_file() and not p.name.startswith("~$")]
    workbooks = [p for p in files if p.suffix.lower() in {".xlsx", ".xlsm"}]
    sheet_map = {path: _workbook_sheets(path) for path in workbooks}
    basic_candidates = [p for p, sheets in sheet_map.items() if {"사업장정보", "수당등록", "수당조정(비과세)"} <= sheets]
    def input_rank(path: Path) -> int:
        parts = [part.casefold() for part in path.parts]
        if any(part == "1.input" for part in parts):
            return 3
        if any(part.startswith("input_") and "raw" not in part for part in parts):
            return 2
        if any(part.startswith("input") for part in parts):
            return 1
        return 0

    def basic_score(path: Path) -> tuple[int, int, int]:
        sheets = sheet_map[path]
        return (
            input_rank(path),
            sum(name.startswith("원천징수추출(") and "수정" not in name for name in sheets),
            -len(sheets),
        )
    if not basic_candidates:
        raise SourceNotFound("기초정보 파일을 찾지 못했습니다.")
    ranked = sorted(basic_candidates, key=basic_score, reverse=True)
    if len(ranked) > 1 and basic_score(ranked[0]) == basic_score(ranked[1]):
        raise SourceNotFound(f"기초정보 우선 후보가 중복입니다: {ranked[0].name}, {ranked[1].name}")
    basic = ranked[0]
    company, exemptions = parse_basic_info(basic, set(years), company_name)

    tools: dict[int, Path] = {}
    payroll: dict[int, list[Path]] = {year: [] for year in years}
    for path, sheets in sheet_map.items():
        if path == basic:
            continue
        year = infer_year(path.name)
        if "Data" in sheets:
            if year in years:
                current = tools.get(year)
                candidate_score = (_data_row_count(path), input_rank(path))
                current_score = (_data_row_count(current), input_rank(current)) if current else (-1, -1)
                if current is None or candidate_score > current_score:
                    tools[year] = path
                # 같은 낮은 우선순위 후보는 뒤에 나타날 Input 계열 후보를 위해
                # 보류한다. 최종적으로 가장 높은 경로 우선순위의 첫 파일을 쓴다.
        elif year in years and {"대상자(작업)", "비대상자"} <= sheets:
            payroll[year].append(path)
    pdf_sources: dict[int, list[Path]] = {year: [] for year in years}
    for year in years:
        if payroll[year]:
            best_rank = max(input_rank(path) for path in payroll[year])
            payroll[year] = [path for path in payroll[year] if input_rank(path) == best_rank]
    for path in files:
        if path.suffix.lower() == ".pdf":
            year = _pdf_year(path)
            if year in years:
                pdf_sources[year].append(path)

    prepared = {}
    converted = {}
    pdfs: dict[int, list[Path]] = {year: [] for year in years}
    withholding_excels: dict[int, list[Path]] = {year: [] for year in years}
    for year in years:
        prepared[year] = prepare_withholding_pdfs(
            pdf_files=pdf_sources[year],
            exemptions=exemptions.get(year, []),
            company_name=company.display_name,
            year=year,
            output_dir=run_dir / "prepared" / "withholding" / str(year),
            quarantine_dir=run_dir / "quarantine" / "withholding" / str(year),
        )
        pdfs[year] = prepared[year].output_files
        converted[year] = convert_withholding_pdfs_to_excel(
            pdf_files=pdfs[year],
            exemptions=exemptions.get(year, []),
            year=year,
            output_dir=run_dir / "prepared" / "withholding_excel" / str(year),
            quarantine_dir=run_dir / "quarantine" / "withholding_excel" / str(year),
        )
        withholding_excels[year] = converted[year].output_files

    snapshot = run_dir / "incoming"
    manifest = []
    selected = [basic, *tools.values(), *(p for group in payroll.values() for p in group), *(p for item in prepared.values() for p in item.source_files)]
    for number, source in enumerate(dict.fromkeys(selected), 1):
        target = snapshot / f"{number:05d}_{source.name}"
        snapshot_file(source, target)
        manifest.append({"source_name": source.name, "source_relative": str(source.resolve().relative_to(source_dir.resolve())), "snapshot": str(target.relative_to(run_dir)), "size": source.stat().st_size, "sha256": sha256(source)})
    for item in prepared.values():
        for generated in [*item.output_files, item.manifest_path]:
            manifest.append({"source_name": generated.name, "source_relative": None, "snapshot": str(generated.relative_to(run_dir)), "size": generated.stat().st_size, "sha256": sha256(generated), "generated": True})
    for item in converted.values():
        for generated in [*item.output_files, item.manifest_path]:
            manifest.append({"source_name": generated.name, "source_relative": None, "snapshot": str(generated.relative_to(run_dir)), "size": generated.stat().st_size, "sha256": sha256(generated), "generated": True})
    manifest_path = run_dir / "input_manifest.json"
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return IngestedInputs(
        company,
        years,
        basic,
        tools,
        payroll,
        pdfs,
        withholding_excels,
        exemptions,
        manifest_path,
    )


def _pdf_index(path: Path) -> int | None:
    match = re.search(r"_(\d+)\.pdf$", path.name, re.IGNORECASE)
    return int(match.group(1)) if match else None
