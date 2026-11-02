from __future__ import annotations

from pathlib import Path
import re

import fitz
from pypdf import PdfReader


def digits(value: object) -> str:
    return re.sub(r"\D", "", str(value or ""))


def display(value: object) -> str:
    if value in (None, ""):
        return ""
    if isinstance(value, bool):
        return "1" if value else "0"
    if isinstance(value, (int, float)):
        return f"{round(value):,}"
    return str(value).strip()


def index_from_name(path: Path) -> int | None:
    match = re.search(r"_(\d+)\.pdf$", path.name, re.IGNORECASE)
    return int(match.group(1)) if match else None


def _resident_number(reader: PdfReader) -> str | None:
    text = reader.pages[0].extract_text() or ""
    compact = re.sub(r"\s+", "", text)
    found = [digits(value) for value in re.findall(r"\d{6}-\d{7}", compact)]
    return next((value for value in found if not value.startswith("110111")), None)


def _replace_text(page: fitz.Page, cells: list[tuple[float, float, float, str]], page_no: int) -> None:
    fontsize = 6.96 if page_no == 0 else 8.04
    height = 8.5 if page_no == 0 else 9.2
    words = page.get_text("words")
    for left, right, top, _ in cells:
        band = fitz.Rect(left, top - 0.5, right, top + height)
        for x0, y0, x1, y1, *_ in words:
            word = fitz.Rect(x0, y0, x1, y1)
            if word.intersects(band) and left <= (x0 + x1) / 2 <= right:
                page.add_redact_annot(word, fill=False, cross_out=False)
    page.apply_redactions(images=0, graphics=0, text=0)
    for _, right, top, text in cells:
        if text:
            width = fitz.get_text_length(text, fontname="helv", fontsize=fontsize)
            baseline = top + (7.0 if page_no == 0 else 8.05)
            page.insert_text((right - width - 0.15, baseline), text, fontname="helv", fontsize=fontsize, color=(0, 0, 0), overlay=True)


def _page_one(row: dict[str, object]) -> list[tuple[float, float, float, str]]:
    left, right = (178.2, 228.56), (502.2, 552.56)
    income_tax, local_tax = (315.0, 356.12), (405.0, 446.13)
    cells = [
        (*left, 303.09, display(row.get("급여_현계"))), (*right, 303.09, display(row.get("급여_총계"))),
        (*left, 390.81, display(row.get("과세소득_현계"))), (*right, 390.81, display(row.get("과세소득_총계"))),
        (*left, 608.73, display(row.get("비과세소득_현계"))), (*right, 608.73, display(row.get("비과세소득_총계"))),
        (*left, 619.81, display(row.get("감면소득_현계"))), (*right, 619.81, display(row.get("감면소득_총계"))),
        (*income_tax, 647.37, display(row.get("결정세액_소득세"))), (*local_tax, 647.37, display(row.get("결정세액_지방세"))),
        (*income_tax, 746.37, display(row.get("차감징수세액_소득세"))), (*local_tax, 746.37, display(row.get("차감징수세액_지방세"))),
    ]
    for number, top in enumerate([403.81, 416.29, 428.29, 440.29, 452.29, 464.29], 1):
        cells.extend([(*left, top, display(row.get(f"금액{number}_현"))), (*right, top, display(row.get(f"금액{number}_총")))])
    return cells


def _page_two(row: dict[str, object]) -> list[tuple[float, float, float, str]]:
    left, right = (260.0, 310.45), (498.0, 547.45)
    rate = row.get("실효세율")
    rate_text = f"{rate:.1%}" if isinstance(rate, (int, float)) else display(rate)
    return [
        (*left, 31.21, display(row.get("총급여"))), (*left, 46.93, display(row.get("근로소득공제"))),
        (*left, 59.65, display(row.get("근로소득금액"))), (*left, 581.65, display(row.get("차감소득금액"))),
        (*left, 775.93, display(row.get("그밖의소득공제계"))), (*right, 31.57, display(row.get("과세표준"))),
        (*right, 54.01, display(row.get("산출세액"))), (*right, 124.10, display(row.get("조특법 제30조"))),
        (*right, 175.57, display(row.get("세액감면 계"))), (*right, 199.57, display(row.get("근로소득"))),
        (*right, 775.93, display(row.get("세액공제 계"))), (*right, 790.21, display(row.get("결정세액"))),
        (*right, 803.65, rate_text),
    ]


def update_pdf(source: Path, destination: Path, row: dict[str, object]) -> None:
    reader = PdfReader(source)
    page_count = len(reader.pages)
    if page_count < 2:
        raise ValueError(f"PDF가 2페이지 미만입니다: {page_count}페이지")
    expected, actual = digits(row.get("주민번호")), _resident_number(reader)
    if not expected:
        raise ValueError("Data 시트 주민번호가 비어 있습니다.")
    if actual != expected:
        raise ValueError(f"주민번호 불일치: Data={expected}, PDF={actual or '없음'}")
    document = fitz.open(source)
    temporary = destination.with_suffix(".tmp.pdf")
    try:
        _replace_text(document[0], _page_one(row), 0)
        _replace_text(document[1], _page_two(row), 1)
        destination.parent.mkdir(parents=True, exist_ok=True)
        document.save(temporary, garbage=4, deflate=True, clean=True)
    finally:
        document.close()
    temporary.replace(destination)
    if len(PdfReader(destination).pages) != page_count:
        destination.unlink(missing_ok=True)
        raise ValueError("저장 후 PDF 재검증에 실패했습니다.")
