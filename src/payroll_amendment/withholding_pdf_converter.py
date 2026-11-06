"""원천징수영수증 PDF 변환 독립 모듈. 기존 배포 프로그램과 무관하게 복사하여 사용합니다.

설정: ConverterConfig / 단건: convert_pdf / 일괄: convert_batch.
입력 금액을 재계산하지 않고 전달받은 값을 표시하며 추가 페이지를 보존합니다.
"""
from __future__ import annotations

import argparse
import csv
import json
import tempfile
from contextvars import ContextVar
from dataclasses import dataclass, field
import os
from pathlib import Path
import re
import sys
import traceback

import pymupdf as fitz
from openpyxl import load_workbook
from pypdf import PdfReader



@dataclass
class ConverterConfig:
    """경로는 root 기준. column_map은 표준열:외부열. 좌표 단위는 point.
    layout_overrides는 {페이지번호: {항목순번: [좌,우,상,검색라벨]}}.
    페이지/항목 번호는 0부터 시작하며 현재 원천징수영수증 서식이 기본입니다.
    """
    root: Path = field(default_factory=Path.cwd)
    excel_dir: str = "input/금액정보"
    pdf_dir: str = "input/원천징수영수증_PDF"
    output_dir: str = "output"
    marked_output_dir: str | None = None
    sheet_name: str = "Data"
    index_column: str = "인덱스"
    workbook_pattern: str = "*"
    column_map: dict = field(default_factory=dict)
    year_pattern: str = r"(?<!\d)(\d{2}|20\d{2})년"
    index_pattern: str = r"_(\d+)\.pdf$"
    normal_enabled: bool = True
    marked_enabled: bool = True
    overwrite: bool = True
    report_enabled: bool = True
    report_name: str = "처리결과.csv"
    year_folder: str = "{year:02d}년도수행"
    marked_folder: str = "변경표시"
    normal_filename: str = "{연번}_{성명}_{기간}_원천징수영수증(수정후).pdf"
    marked_filename: str = "{연번}_{성명}_{기간}_원천징수영수증(변경표시).pdf"
    normal_sizes: tuple = (6.96, 8.0376568)
    marked_sizes: tuple = (7.4, 8.5)
    number_fonts: tuple = ("helv", "tiro")
    label_size: float = 6.0
    code_size: float = 6.0
    font_path: str | None = None
    raster_scale: float = 4.0
    change_gap: float = 2.0
    old_color: tuple = (0.35, 0.35, 0.35)
    new_color: tuple = (0, 0, 0)
    strike_color: tuple = (0.25, 0.25, 0.25)
    strike_width: float = 0.55
    page_indices: tuple = (0, 1)
    detail_tops: tuple = (403.81, 416.29, 428.29, 440.29, 452.29, 464.29)
    detail_row_step: float = 12.3
    layout_overrides: dict = field(default_factory=dict)

    def path(self, value):
        return (Path(self.root).resolve() / value).resolve()

    def validate(self):
        if not self.normal_enabled and not self.marked_enabled:
            raise ValueError("출력 종류를 하나 이상 선택해야 합니다")
        for sizes in (self.normal_sizes, self.marked_sizes):
            if len(sizes) != 2 or any(float(x) <= 0 for x in sizes):
                raise ValueError("페이지별 글자 크기는 양수 2개여야 합니다")
        if len(self.page_indices) != 2 or min(self.page_indices) < 0 or len(set(self.page_indices)) != 2:
            raise ValueError("수정 페이지 번호가 잘못되었습니다")
        if self.raster_scale <= 0:
            raise ValueError("이미지 배율은 양수여야 합니다")
        for pattern in (self.year_pattern, self.index_pattern):
            if re.compile(pattern).groups < 1:
                raise ValueError("정규식에 추출 그룹이 필요합니다")
        if self.font_path and not Path(self.font_path).is_file():
            raise ValueError("설정한 글꼴 파일이 없습니다")

    @classmethod
    def from_json(cls, path):
        """JSON의 상대 root는 JSON 파일 위치 기준."""
        path = Path(path).resolve()
        data = json.loads(path.read_text(encoding="utf-8-sig"))
        data["root"] = (path.parent / data.get("root", ".")).resolve()
        return cls(**data)


_ACTIVE_CONFIG = ContextVar("converter_config", default=None)


def _config():
    return _ACTIVE_CONFIG.get() or ConverterConfig()


def _mapped_row(row):
    result = dict(row)
    for standard, external in _config().column_map.items():
        if external in row:
            result[standard] = row[external]
    return result


def _layout(values, page_number):
    overrides = _config().layout_overrides.get(str(page_number),
                _config().layout_overrides.get(page_number, {}))
    result = list(values)
    for position, specification in overrides.items():
        i = int(position)
        left, right, top, label = specification
        result[i] = (left, right, top, result[i][3], label)
    return result


PAGE_H = 842
YEAR_PATTERN = re.compile(r"(?<!\d)(\d{2}|20\d{2})년")


def _digits(value: object) -> str:
    return re.sub(r"\D", "", str(value or ""))


def _as_text(value: object) -> str:
    if value is None or value == "":
        return ""
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, (int, float)):
        return f"{round(value):,}"
    return str(value).strip()


def _index_from_name(path: Path) -> int | None:
    match = re.search(_config().index_pattern, path.name, re.IGNORECASE)
    if match is None:
        match = re.match(r"^(\d+)_", path.name)
    return int(match.group(1)) if match else None


def _year_from_name(path: Path) -> int | None:
    match = re.search(_config().year_pattern, path.name)
    if not match:
        return None
    return int(match.group(1)) % 100


def _year_from_pdf_folder(path: Path, pdf_root: Path) -> int | None:
    """Read the processing year from a PDF's parent folder under pdf_root."""
    current = path.parent
    while current != pdf_root and pdf_root in current.parents:
        year = _year_from_name(current)
        if year is not None:
            return year
        current = current.parent
    return None


def _pdf_resident_number(reader: PdfReader) -> str | None:
    text = re.sub(r"\s+", "", reader.pages[0].extract_text() or "")
    matches = re.findall(r"\d{6}-\d{7}", text)
    employee_ids = [_digits(value) for value in matches if not value.startswith("110111")]
    return employee_ids[0] if employee_ids else None


def _has_extractable_text(source: Path) -> bool:
    reader = PdfReader(source)
    text = "".join(
        reader.pages[index].extract_text() or ""
        for index in range(min(2, len(reader.pages)))
    )
    return len(re.findall(r"[0-9A-Za-z가-힣]", text)) >= 20


def _validate_pdf_identity(source: Path, row: dict[str, object]) -> None:
    reader = PdfReader(source)
    expected_resident = _digits(row.get("주민번호") or row.get("주민등록번호"))
    actual_resident = _pdf_resident_number(reader)
    if expected_resident and actual_resident != expected_resident:
        raise ValueError(
            f"주민번호 불일치: Data={expected_resident}, PDF={actual_resident or '없음'}"
        )
    expected_name = re.sub(r"\s+", "", str(row.get("성명") or row.get("이름") or ""))
    page_text = re.sub(r"\s+", "", reader.pages[0].extract_text() or "")
    if expected_name and expected_name not in page_text:
        raise ValueError(f"성명 불일치 또는 PDF에서 성명을 찾지 못했습니다: {expected_name}")


def read_data_rows(workbook: Path) -> dict[int, dict[str, object]]:
    book = load_workbook(workbook, data_only=True, read_only=True)
    try:
        sheet = next(ws for ws in book.worksheets if ws.title.casefold() == _config().sheet_name.casefold())
        headers = [cell.value for cell in sheet[1]]
        rows: dict[int, dict[str, object]] = {}
        for values in sheet.iter_rows(min_row=2, values_only=True):
            row = _mapped_row(dict(zip(headers, values)))
            if row.get(_config().index_column) in (None, ""):
                continue
            index = int(row[_config().index_column])
            if index in rows:
                raise ValueError(f"중복 인덱스: {index}")
            rows[index] = row
        return rows
    finally:
        book.close()


def _find_label_span(
    spans: list[dict], label_key: object, y_hint: float | None = None,
) -> dict | None:
    if label_key in (None, ""):
        return None
    raw_key = str(label_key).strip()
    if raw_key.startswith("@detail|"):
        _, number_text, code = raw_key.split("|", 2)
        occurrence = int(number_text)
        code_matches = [
            span for span in spans
            if re.sub(r"[^0-9A-Za-z]", "", span["text"]).upper()
            == re.sub(r"[^0-9A-Za-z]", "", code).upper()
        ]
        if len(code_matches) >= occurrence:
            return code_matches[occurrence - 1]
        detail_codes = sorted(
            (
                span for span in spans
                if 250 <= span["bbox"][1] <= 650
                and re.fullmatch(r"[A-Z][0-9A-Z]{2}(?:,[A-Z0-9]+)*", span["text"].strip())
            ),
            key=lambda span: span["bbox"][1],
        )
        if occurrence <= len(detail_codes):
            return detail_codes[occurrence - 1]

        # A sparse form may contain only the rows that existed in the original
        # receipt. Allocate a missing Excel detail item to the next table row.
        if detail_codes:
            anchor = detail_codes[0]
            offset = _config().detail_row_step * (occurrence - 1)
        else:
            total_candidates = []
            for span in spans:
                label = re.sub(r"[^0-9A-Za-z가-힣]", "", span["text"])
                if label == "계" and 250 <= span["bbox"][1] <= 450:
                    total_candidates.append(span)
            if not total_candidates:
                return None
            anchor = max(total_candidates, key=lambda span: span["bbox"][1])
            offset = _config().detail_row_step * occurrence
        synthetic = dict(anchor)
        x0, y0, x1, y1 = anchor["bbox"]
        ox, oy = anchor["origin"]
        synthetic["bbox"] = (x0, y0 + offset, x1, y1 + offset)
        synthetic["origin"] = (ox, oy + offset)
        synthetic["synthetic"] = True
        return synthetic
    if raw_key.startswith("#"):
        prefix = raw_key[1:] + "."
        matches = [span for span in spans if span["text"].strip().startswith(prefix)]
        if not matches:
            return None
        return min(matches, key=lambda span: abs(span["bbox"][1] - y_hint)) if y_hint is not None else matches[0]
    key = re.sub(r"[^0-9A-Za-z가-힣]", "", raw_key)
    matches = []
    for span in spans:
        label = re.sub(r"^\s*\d+(?:-\d+)?\.\s*", "", span["text"])
        label = re.sub(r"[^0-9A-Za-z가-힣]", "", label)
        if key in {"급여", "근로소득"}:
            if label == key:
                matches.append(span)
        elif key in label:
            matches.append(span)
    if not matches:
        # Some issuers split one Korean row label into several PDF spans.
        # Recombine spans sharing the same baseline before giving up.
        for anchor in spans:
            ay = anchor["bbox"][1]
            same_line = sorted(
                (span for span in spans if abs(span["bbox"][1] - ay) <= 0.35),
                key=lambda span: span["bbox"][0],
            )
            combined = re.sub(
                r"[^0-9A-Za-z가-힣]", "",
                "".join(span["text"] for span in same_line),
            )
            if key in combined:
                synthetic = dict(same_line[0])
                synthetic["bbox"] = (
                    min(span["bbox"][0] for span in same_line),
                    min(span["bbox"][1] for span in same_line),
                    max(span["bbox"][2] for span in same_line),
                    max(span["bbox"][3] for span in same_line),
                )
                synthetic["synthetic"] = True
                matches.append(synthetic)
    if not matches:
        return None
    return min(matches, key=lambda span: abs(span["bbox"][1] - y_hint)) if y_hint is not None else matches[0]


def _cell_bounds(
    page: fitz.Page, x_hint: float, y: float, fallback_left: float, fallback_right: float,
) -> tuple[float, float]:
    """Return the vertical table boundaries surrounding a point."""
    verticals: list[float] = []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] != "l":
                continue
            p1, p2 = item[1], item[2]
            if abs(p1.x - p2.x) < 0.15 and min(p1.y, p2.y) - 0.5 <= y <= max(p1.y, p2.y) + 0.5:
                verticals.append(p1.x)
    lefts = [x for x in verticals if x < x_hint - 0.2]
    rights = [x for x in verticals if x > x_hint + 0.2]
    if not lefts or not rights:
        return fallback_left, fallback_right
    left, right = max(lefts), min(rights)
    if right - left < 8:
        return fallback_left, fallback_right
    return left, right


def _insert_rasterized_korean_label(
    page: fitz.Page,
    x: float,
    baseline: float,
    text: str,
    fontsize: float,
    korean_font: fitz.Font,
    korean_font_file: str | None,
    max_width: float,
) -> None:
    """Insert a label as pixels so every PDF viewer renders Hangul reliably."""
    width = max(2.0, korean_font.text_length(text, fontsize=fontsize) + 1.0)
    height = fontsize * 1.45 + 1.0

    label_document = fitz.open()
    try:
        label_page = label_document.new_page(width=width, height=height)
        local_baseline = fontsize + 0.4
        font_arguments = {"fontname": "detailfont", "fontfile": korean_font_file} if korean_font_file else {"fontname": "korea"}
        label_page.insert_text(
            (0.2, local_baseline), text,
            fontsize=fontsize, color=_config().new_color, **font_arguments,
        )
        pixmap = label_page.get_pixmap(matrix=fitz.Matrix(_config().raster_scale, _config().raster_scale), alpha=True)
        top = baseline - local_baseline
        target_width = min(width, max_width)
        page.insert_image(
            fitz.Rect(x, top, x + target_width, top + height),
            pixmap=pixmap, keep_proportion=False, overlay=True,
        )
    finally:
        label_document.close()


def _ensure_detail_labels(page: fitz.Page, row: dict[str, object]) -> None:
    """Add Excel-only sparse non-taxable item labels and codes to blank rows."""
    spans = [
        span
        for block in page.get_text("dict")["blocks"]
        for line in block.get("lines", [])
        for span in line.get("spans", [])
    ]
    # Rasterizing removes viewer font dependencies, so the Windows UI font can
    # be used for a closer visual match to the source form.
    malgun_path = Path(os.environ.get("WINDIR", "C:/Windows")) / "Fonts" / "malgun.ttf"
    font_file = _config().font_path or (str(malgun_path) if malgun_path.exists() else None)
    font = fitz.Font(fontfile=font_file) if font_file else fitz.Font("korea")

    for number in range(1, 7):
        code = _as_text(row.get(f"코드{number}_현") or row.get(f"코드{number}_총"))
        if not code:
            continue
        normalized_code = re.sub(r"[^0-9A-Za-z]", "", code).upper()
        if any(
            re.sub(r"[^0-9A-Za-z]", "", span["text"]).upper() == normalized_code
            for span in spans
        ):
            continue

        anchor = _find_label_span(spans, f"@detail|{number}|{code}")
        if anchor is None:
            raise ValueError(f"비과세 항목 {code}를 넣을 빈 행을 찾을 수 없습니다")
        baseline = float(anchor["origin"][1])
        name_left, name_right = _cell_bounds(page, 120.0, baseline, 62.0, 160.0)
        code_left, code_right = _cell_bounds(page, 168.0, baseline, 160.0, 178.0)
        raw_name = _as_text(
            row.get(f"비과세항목명{number}_현")
            or row.get(f"비과세항목명{number}_총")
            or row.get(f"비과세명{number}_현")
            or row.get(f"비과세명{number}_총")
        )
        display_name = re.sub(
            r"^\s*[A-Z][0-9A-Z]{2}\.\s*", "", raw_name,
            flags=re.IGNORECASE,
        )

        if display_name:
            fontsize = _config().label_size
            available = name_right - name_left - 1.2
            _insert_rasterized_korean_label(
                page, name_left + 0.6, baseline,
                display_name, fontsize, font, font_file, available,
            )
        code_fontsize = _config().code_size
        code_width = fitz.get_text_length(code, fontname="helv", fontsize=code_fontsize)
        code_x = code_left + max(0.4, (code_right - code_left - code_width) / 2)
        page.insert_text(
            (code_x, baseline), code, fontname="helv", fontsize=code_fontsize,
            color=_config().new_color, overlay=True,
        )


def _replace_page_text(
    document: fitz.Document,
    page: fitz.Page,
    values: list[tuple],
    page_number: int,
) -> None:
    """Remove old text objects and place replacements on the original baseline."""
    fontsize = _config().normal_sizes[page_number]
    height = 8.5 if page_number == 0 else 9.2
    words = page.get_text("words")
    spans = [
        span
        for block in page.get_text("dict")["blocks"]
        for line in block.get("lines", [])
        for span in line.get("spans", [])
    ]
    baselines: list[float] = []
    adjusted_values: list[tuple[float, float, float, str]] = []

    for value in values:
        left, right, top, text = value[:4]
        label_key = value[4] if len(value) > 4 else None
        label_span = _find_label_span(spans, label_key, top)
        if label_span is not None:
            top = label_span["bbox"][1]
        row_y = label_span["origin"][1] if label_span is not None else top + height / 2
        left, right = _cell_bounds(page, (left + right) / 2, row_y, left, right)
        adjusted_values.append((left, right, top, text))

        candidates = []
        for span in spans:
            x0, y0, x1, _y1 = span["bbox"]
            center_x = (x0 + x1) / 2
            if left <= center_x <= right and re.search(r"\d", span["text"]):
                distance = min(abs(y0 - top), abs(span["origin"][1] - top))
                if distance <= 12:
                    candidates.append((distance, span["origin"][1]))
        if label_span is not None:
            baselines.append(label_span["origin"][1])
        elif candidates:
            baselines.append(min(candidates)[1])
        else:
            baselines.append(top + (7.0 if page_number == 0 else 6.83))

    for left, right, top, _ in adjusted_values:
        band = fitz.Rect(left, top - 0.5, right, top + height)
        for x0, y0, x1, y1, _text, *_ in words:
            word = fitz.Rect(x0, y0, x1, y1)
            if word.intersects(band) and left <= (x0 + x1) / 2 <= right:
                page.add_redact_annot(word, fill=False, cross_out=False)

    page.apply_redactions(images=0, graphics=0, text=0)

    # The source uses a serif Korean font for page-two numerals. PyMuPDF
    # cannot safely reuse its subset encoding for newly inserted characters,
    # so use the metrically similar built-in Times Roman font.
    fontname = _config().number_fonts[page_number]

    for (_left, right, _top, text), baseline in zip(adjusted_values, baselines):
        if text == "":
            continue
        width = fitz.get_text_length(text, fontname=fontname, fontsize=fontsize)
        x = right - width - 0.15
        page.insert_text(
            (x, baseline),
            text,
            fontname=fontname,
            fontsize=fontsize,
            color=_config().new_color,
            overlay=True,
        )


def _mark_page_text_changes(
    document: fitz.Document,
    page: fitz.Page,
    values: list[tuple],
    page_number: int,
) -> None:
    """Show old and new amounts together, striking through the old value."""
    base_fontsize = 5.4 if page_number == 0 else 8.0376568
    words = page.get_text("words")
    spans = [
        span
        for block in page.get_text("dict")["blocks"]
        for line in block.get("lines", [])
        for span in line.get("spans", [])
    ]
    horizontal_rules: list[tuple[float, float, float]] = []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            if item[0] != "l":
                continue
            p1, p2 = item[1], item[2]
            if abs(p1.y - p2.y) < 0.15:
                horizontal_rules.append((p1.y, min(p1.x, p2.x), max(p1.x, p2.x)))
    changes: list[dict[str, object]] = []

    for value in values:
        left, right, top, new_text = value[:4]
        label_key = value[4] if len(value) > 4 else None
        label_span = _find_label_span(spans, label_key, top)
        if label_span is not None:
            top = label_span["bbox"][1]
        row_y = label_span["origin"][1] if label_span is not None else top + 4.0
        left, right = _cell_bounds(page, (left + right) / 2, row_y, left, right)

        candidates = []
        # Label-driven page-two rows can place the amount baseline lower than
        # the label glyph box. Fixed-coordinate page-one rows must stay tight
        # so an empty detail row never captures the total directly above it.
        row_height = 12.0 if label_span is not None else (8.5 if page_number == 0 else 9.2)
        row_top_tolerance = 2.0 if label_span is not None else 0.75
        for span in spans:
            x0, y0, x1, _y1 = span["bbox"]
            center_x = (x0 + x1) / 2
            in_target_row = top - row_top_tolerance <= y0 <= top + row_height
            if left <= center_x <= right and in_target_row and re.search(r"\d", span["text"]):
                distance = min(abs(y0 - top), abs(span["origin"][1] - top))
                candidates.append((distance, span))
        old_span = min(candidates, key=lambda item: item[0])[1] if candidates else None
        old_spans: list[dict] = []
        if old_span is not None:
            old_y = old_span["bbox"][1]
            old_spans = sorted(
                (
                    span for span in spans
                    if left <= (span["bbox"][0] + span["bbox"][2]) / 2 <= right
                    and abs(span["bbox"][1] - old_y) <= 1.2
                    and re.search(r"[0-9,.%\-]", span["text"])
                ),
                key=lambda span: span["bbox"][0],
            )
        old_text = "".join(span["text"] for span in old_spans).strip()
        baseline = (
            label_span["origin"][1]
            if label_span is not None
            else old_span["origin"][1]
            if old_span is not None
            else top + (7.0 if page_number == 0 else 6.83)
        )

        if old_text == new_text:
            continue
        for span in old_spans:
            page.add_redact_annot(fitz.Rect(span["bbox"]), fill=False, cross_out=False)
        changes.append(
            {
                "left": left,
                "right": right,
                "baseline": baseline,
                "old": old_text,
                "new": new_text,
                "base_fontsize": base_fontsize,
            }
        )

    if not changes:
        return
    page.apply_redactions(images=0, graphics=0, text=0)

    fontname = _config().number_fonts[page_number]
    for change in changes:
        left = float(change["left"])
        right = float(change["right"])
        baseline = float(change["baseline"])
        old_text = str(change["old"])
        new_text = str(change["new"])
        # Keep change marks readable instead of shrinking them to the cell width.
        # A long old/new pair may extend to the left of the amount cell.
        gap = _config().change_gap if old_text and new_text else 0.0
        fontsize = _config().marked_sizes[page_number]
        old_width = fitz.get_text_length(old_text, fontname=fontname, fontsize=fontsize)
        new_width = fitz.get_text_length(new_text, fontname=fontname, fontsize=fontsize)

        total_width = old_width + gap + new_width
        start_x = max(4.0, right - total_width - 0.15)

        if old_text:
            page.insert_text(
                (start_x, baseline), old_text, fontname=fontname,
                fontsize=fontsize, color=_config().old_color, overlay=True,
            )
            strike_y = baseline - fontsize * 0.32
            page.draw_line(
                fitz.Point(start_x - 0.2, strike_y),
                fitz.Point(start_x + old_width + 0.2, strike_y),
                color=_config().strike_color, width=_config().strike_width, overlay=True,
            )
        if new_text:
            new_x = start_x + old_width + gap
            page.insert_text(
                (new_x, baseline), new_text, fontname=fontname,
                fontsize=fontsize, color=_config().new_color, overlay=True,
            )


def _page_one_values(row: dict[str, object]) -> list[tuple[float, float, float, str]]:
    left_cell = (178.2, 228.56)
    right_cell = (502.2, 552.56)
    income_tax = (315.0, 356.12)
    local_tax = (405.0, 446.13)
    values = [
        (*left_cell, 303.09, _as_text(row.get("급여_현계")), "급여"),
        (*right_cell, 303.09, _as_text(row.get("급여_총계")), "급여"),
        (*left_cell, 390.81, _as_text(row.get("과세소득_현계")), "계"),
        (*right_cell, 390.81, _as_text(row.get("과세소득_총계")), "계"),
        (*left_cell, 608.73, _as_text(row.get("비과세소득_현계")), "비과세소득계"),
        (*right_cell, 608.73, _as_text(row.get("비과세소득_총계")), "비과세소득계"),
        (*left_cell, 619.81, _as_text(row.get("감면소득_현계")), "감면소득계"),
        (*right_cell, 619.81, _as_text(row.get("감면소득_총계")), "감면소득계"),
        (*income_tax, 647.37, _as_text(row.get("결정세액_소득세")), "결정세액"),
        (*local_tax, 647.37, _as_text(row.get("결정세액_지방세")), "결정세액"),
        (*income_tax, 746.37, _as_text(row.get("차감징수세액_소득세")), "차감징수세액"),
        (*local_tax, 746.37, _as_text(row.get("차감징수세액_지방세")), "차감징수세액"),
    ]
    detail_tops = _config().detail_tops
    for number, top in enumerate(detail_tops, start=1):
        current = _as_text(row.get(f"금액{number}_현"))
        total = _as_text(row.get(f"금액{number}_총"))
        code = _as_text(row.get(f"코드{number}_현") or row.get(f"코드{number}_총"))
        if not any((current, total, code)):
            continue
        label_key = f"@detail|{number}|{code}"
        values.append((*left_cell, top, current, label_key))
        values.append((*right_cell, top, total, label_key))
    return values


def _page_two_values(row: dict[str, object]) -> list[tuple[float, float, float, str]]:
    left = (260.0, 310.45)
    right = (498.0, 547.45)
    rate = row.get("실효세율")
    rate_text = f"{rate:.1%}" if isinstance(rate, (int, float)) else _as_text(rate)
    return [
        (*left, 31.21, _as_text(row.get("총급여")), "총급여"),
        (*left, 46.93, _as_text(row.get("근로소득공제")), "근로소득공제"),
        (*left, 59.65, _as_text(row.get("근로소득금액")), "근로소득금액"),
        (*left, 581.65, _as_text(row.get("차감소득금액")), "차감소득금액"),
        (*left, 775.93, _as_text(row.get("그밖의소득공제계")), "그밖의소득공제계"),
        (*right, 31.57, _as_text(row.get("과세표준")), "과세표준"),
        (*right, 54.01, _as_text(row.get("산출세액")), "산출세액"),
        (*right, 124.10, _as_text(row.get("조특법 제30조")), "제30조"),
        (*right, 175.57, _as_text(row.get("세액감면 계")), "세액감면계"),
        (*right, 199.57, _as_text(row.get("근로소득")), "근로소득"),
        (*right, 775.93, _as_text(row.get("세액공제 계")), "세액공제계"),
        (*right, 790.21, _as_text(row.get("결정세액")), "결정세액"),
        (*right, 803.65, rate_text, "실효세율"),
    ]



@dataclass
class ConversionResult:
    """부분 성공도 생성된 파일 경로와 오류를 각각 반환합니다."""
    source: Path
    outputs: dict = field(default_factory=dict)
    errors: dict = field(default_factory=dict)

    @property
    def status(self):
        return "부분완료" if self.errors and self.outputs else "오류" if self.errors else "완료"


def convert_pdf(source, row, *, config=None, normal_path=None, marked_path=None):
    """계산된 dict를 받아 PDF 하나를 변환합니다. Excel 없이 호출 가능합니다.

    명시한 출력 경로는 호출 위치 기준이며, 생략하면 설정의 파일명 규칙을 씁니다.
    import만으로는 디렉터리나 파일을 만들지 않습니다.
    """
    config = config or ConverterConfig()
    config.validate()
    token = _ACTIVE_CONFIG.set(config)
    source = Path(source).resolve()
    result = ConversionResult(source)
    try:
        row = _mapped_row(row)
        jobs = [("normal", config.normal_enabled, normal_path, config.normal_filename),
                ("marked", config.marked_enabled, marked_path, config.marked_filename)]
        planned = []
        for kind, enabled, explicit, template in jobs:
            if not enabled:
                continue
            destination = Path(explicit).resolve() if explicit else _safe_output(config.path(config.output_dir), template.format_map(row))
            if destination == source:
                raise ValueError("원본 PDF 경로로 출력할 수 없습니다")
            if destination in [p for _, p in planned]:
                raise ValueError("일반본과 변경표시본 출력 경로가 같습니다")
            planned.append((kind, destination))
        for kind, destination in planned:
            temporary = None
            try:
                if destination.exists() and not config.overwrite:
                    raise FileExistsError(f"기존 결과 파일: {destination}")
                destination.parent.mkdir(parents=True, exist_ok=True)
                with fitz.open(source) as document:
                    if len(document) <= max(config.page_indices):
                        raise ValueError("PDF 수정 대상 페이지가 없습니다")
                    for logical, actual in enumerate(config.page_indices):
                        page = document[actual]
                        if logical == 0:
                            _ensure_detail_labels(page, row)
                        values = _page_one_values(row) if logical == 0 else _page_two_values(row)
                        values = _layout(values, logical)
                        handler = _replace_page_text if kind == "normal" else _mark_page_text_changes
                        handler(document, page, values, logical)
                    # 같은 디렉터리의 고유 임시 파일로 저장한 후 교체합니다.
                    descriptor, temporary_name = tempfile.mkstemp(suffix=".pdf", prefix=".convert-", dir=destination.parent)
                    os.close(descriptor)
                    temporary = Path(temporary_name)
                    document.save(temporary, garbage=4, deflate=True, clean=True)
                temporary.replace(destination)
                result.outputs[kind] = destination
            except Exception as exc:
                result.errors[kind] = str(exc)
            finally:
                if temporary is not None and temporary.exists():
                    temporary.unlink()
        return result
    finally:
        _ACTIVE_CONFIG.reset(token)


def _safe_output(root, relative):
    """출력 템플릿의 상위 경로 이동을 차단합니다."""
    root = Path(root).resolve()
    destination = (root / relative).resolve()
    if not destination.is_relative_to(root) or destination == root:
        raise ValueError(f"출력 폴더 밖의 파일명: {relative}")
    return destination


def convert_batch(config=None, *, on_progress=None):
    """연도 폴더 + 파일명 마지막 인덱스로 일괄 처리합니다.

    on_progress(result)는 파일 처리 후 호출합니다. 보고서와 결과 목록을 반환합니다.
    주민번호·성명을 검증하며 중복 연도/인덱스는 오류로 기록합니다.
    """
    config = config or ConverterConfig()
    config.validate()
    token = _ACTIVE_CONFIG.set(config)
    results, report = [], []
    def record(status, detail, year="", workbook="", source="", normal="", marked=""):
        report.append(dict(zip(
            ["연도", "금액파일", "PDF", "상태", "내용", "수정본", "변경표시본"],
            [str(year), str(workbook), str(source), status, detail, str(normal), str(marked)])))
    try:
        excel_root, pdf_root = config.path(config.excel_dir), config.path(config.pdf_dir)
        output_root = config.path(config.output_dir)
        marked_output_root = (
            config.path(config.marked_output_dir)
            if config.marked_output_dir
            else output_root / config.marked_folder
        )
        workbooks = sorted(
            path
            for path in excel_root.glob(config.workbook_pattern)
            if path.suffix.lower() in {".xlsx", ".xlsm"}
            and not path.name.startswith("~$")
        )
        pdfs = sorted(p for p in pdf_root.rglob("*") if p.is_file() and p.suffix.lower() == ".pdf")
        if not workbooks:
            record("오류", "금액정보 파일 없음")
        if not pdfs:
            record("오류", "PDF 파일 없음")
        grouped = {}
        for workbook in workbooks:
            year = _year_from_name(workbook)
            if year is None:
                record("건너뜀", "파일명에서 수행연도를 찾을 수 없음", workbook=workbook)
            else:
                grouped.setdefault(year, []).append(workbook)
        tables = {}
        for year, books in grouped.items():
            if len(books) != 1:
                record("오류", "같은 연도의 금액정보가 2개 이상임", year=year)
                continue
            try:
                tables[year] = read_data_rows(books[0])
            except Exception as exc:
                record("오류", str(exc), year, books[0])
        expected = {
            (year, index)
            for year, rows in tables.items()
            for index in rows
        }
        matched = set()
        excluded = set()
        located = set()
        reserved = set()
        for source in pdfs:
            year = _year_from_pdf_folder(source, pdf_root)
            if year not in tables:
                record("건너뜀", "연도 폴더 또는 해당 연도 금액정보 확인 필요", year or "", source=source)
                continue
            index = _index_from_name(source)
            if index not in tables[year]:
                record("건너뜀", f"Data 시트에 인덱스 {index} 없음", year, source=source)
                continue
            located.add((year, index))
            row = tables[year][index]
            workbook = grouped[year][0]
            try:
                if not _has_extractable_text(source):
                    excluded.add((year, index))
                    record(
                        "제외",
                        "텍스트 추출 불가 이미지 PDF",
                        year,
                        workbook,
                        source,
                    )
                    continue
                _validate_pdf_identity(source, row)
                # 기존 연번 출력의 정수 표기를 유지합니다.
                naming = dict(row)
                naming["연번"] = int(row["연번"])
                naming["성명"] = row.get("성명") or row.get("이름")
                naming["주민번호"] = row.get("주민번호") or row.get("주민등록번호")
                folder = config.year_folder.format(year=year)
                normal = _safe_output(output_root, str(Path(folder) / config.normal_filename.format_map(naming)))
                marked = _safe_output(marked_output_root, str(Path(folder) / config.marked_filename.format_map(naming)))
                targets = [p for p, enabled in [(normal, config.normal_enabled), (marked, config.marked_enabled)] if enabled]
                if any(str(p).casefold() in reserved for p in targets):
                    raise ValueError("출력 파일명 충돌: 같은 연도/인덱스 또는 연번 확인 필요")
                reserved.update(str(p).casefold() for p in targets)
                result = convert_pdf(source, row, config=config, normal_path=normal, marked_path=marked)
                results.append(result)
                matched.add((year, index))
                record(result.status, "; ".join(result.errors.values()), year, workbook, source,
                       result.outputs.get("normal", ""), result.outputs.get("marked", ""))
                if on_progress:
                    on_progress(result)
            except Exception as exc:
                record("오류", str(exc), year, workbook, source)
        for year, index in sorted(expected - matched - excluded - located):
            record("오류", f"인덱스 {index}에 해당하는 PDF가 없습니다", year, grouped[year][0])
        if config.report_enabled:
            report_path = _safe_output(output_root, config.report_name)
            report_path.parent.mkdir(parents=True, exist_ok=True)
            with report_path.open("w", encoding="utf-8-sig", newline="") as stream:
                writer = csv.DictWriter(stream, fieldnames=["연도", "금액파일", "PDF", "상태", "내용", "수정본", "변경표시본"])
                writer.writeheader()
                writer.writerows(report)
        return results, report
    finally:
        _ACTIVE_CONFIG.reset(token)


def main():
    """직접 실행: python withholding_pdf_converter.py --config 설정.json"""
    parser = argparse.ArgumentParser(description="다른 프로젝트에서 재사용하는 원천징수영수증 PDF 변환")
    parser.add_argument("--config", type=Path, help="선택 JSON 설정 파일")
    parser.add_argument("--root", type=Path, help="기준 경로; JSON의 root보다 우선")
    args = parser.parse_args()
    config = ConverterConfig.from_json(args.config) if args.config else ConverterConfig()
    if args.root:
        config.root = args.root.resolve()
    results, report = convert_batch(config)
    errors = sum(item["상태"] in {"오류", "부분완료"} for item in report)
    print(f"처리 {len(results)}건 / 오류 또는 부분완료 {errors}건")
    return 1 if errors else 0


# 다른 프로젝트에서 사용하는 예:
# from withholding_pdf_converter import ConverterConfig, convert_pdf, convert_batch
# config = ConverterConfig(root=Path("D:/다른프로젝트"), output_dir="결과",
#                          marked_sizes=(7.4, 8.5), overwrite=False)
# results, report = convert_batch(config)
# result = convert_pdf("원본.pdf", 계산된_데이터_dict, config=config,
#                      normal_path="일반본.pdf", marked_path="변경표시본.pdf")
# 필요한 패키지: pip install pymupdf openpyxl pypdf
if __name__ == "__main__":
    raise SystemExit(main())


