from __future__ import annotations

import json
from pathlib import Path
import re

import fitz
from PIL import Image

from .errors import PdfMatchError
from .infra import safe_name, sha256
from .models import StageResult, TaxExemption, ValidationIssue


def pdf_index(path: Path) -> str | None:
    match = re.search(r"_(\d+)(?:\.[^.]+)$", path.name)
    return str(int(match.group(1))) if match else None


def render_withholding_pdfs(*, company_id: str, year: int, exemptions: list[TaxExemption], pdf_files: list[Path], output_dir: Path, dpi: int = 300, image_format: str = "gif") -> StageResult:
    index: dict[str, Path] = {}
    for path in pdf_files:
        key = pdf_index(path)
        if key is None:
            continue
        if key in index:
            raise PdfMatchError(f"PDF 인덱스 중복: {key}")
        index[key] = path
    artifacts: list[Path] = []
    mapping: list[dict] = []
    for exemption in exemptions:
        source = index.get(str(int(exemption.pdf_index)))
        if source is None:
            raise PdfMatchError(f"PDF 인덱스를 찾을 수 없습니다: {exemption.pdf_index}")
        document = fitz.open(source)
        try:
            folder = output_dir / f"{year}년_{exemption.pdf_index}_{safe_name(exemption.name)}"
            folder.mkdir(parents=True, exist_ok=True)
            images = []
            for page_number, page in enumerate(document, 1):
                pixmap = page.get_pixmap(dpi=dpi, alpha=False)
                target = folder / f"page_{page_number:04d}.{image_format.lower()}"
                image = Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
                image.save(target, format=image_format.upper())
                with Image.open(target) as check:
                    check.verify()
                images.append(target); artifacts.append(target)
            if len(images) != document.page_count:
                raise PdfMatchError(f"페이지 변환 개수 불일치: {source.name}")
            mapping.append({"pdf_index": exemption.pdf_index, "source": source.name, "source_sha256": sha256(source), "pages": document.page_count, "images": [str(p.relative_to(output_dir)) for p in images]})
        finally:
            document.close()
    map_path = output_dir / "pdf_mapping.json"
    map_path.parent.mkdir(parents=True, exist_ok=True)
    map_path.write_text(json.dumps(mapping, ensure_ascii=False, indent=2), encoding="utf-8")
    artifacts.append(map_path)
    return StageResult("S30", company_id, year, "SUCCEEDED", artifacts, {"employees": len(exemptions), "images": len(artifacts) - 1})

