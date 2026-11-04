from __future__ import annotations

from concurrent.futures import as_completed, ThreadPoolExecutor
from dataclasses import dataclass
import json
import logging
import os
from pathlib import Path
import re
import shutil
from statistics import median
from typing import Any
from uuid import uuid4

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pdfplumber

from .errors import PdfToExcelError, SchemaMismatch
from .infra import format_resident_id, sha256
from .withholding_basic_info import (
    WITHHOLDING_HEADERS,
    _record,
    _taxfree_headers,
    _write_with_excel,
)
from .withholding_pdf_excel import WithholdingPdfData, parse_withholding_tables


logger = logging.getLogger(__name__)


_RESIDENT = re.compile(r"(?<!\d)(\d{6})-(\d{7})(?!\d)")
_MONEY = re.compile(r"-?\d[\d,]*")
_PAYROLL_FILE = re.compile(r".*?(?P<year>20\d{2})(?P<month>0[1-9]|1[0-2])_(?P<part>\d+)\.pdf$", re.IGNORECASE)
_KOREAN_PAYROLL_FILE = re.compile(
    r".*?(?P<year>20\d{2}|\d{2})\s*년\s*(?P<month>0?[1-9]|1[0-2])\s*월.*?급여대장"
    r"(?:.*?[_\- ](?P<part>\d+))?\.pdf$",
    re.IGNORECASE,
)
_DATED_PAYROLL_FILE = re.compile(
    r".*?(?P<year>20\d{2})(?P<month>0[1-9]|1[0-2])"
    r"(?=.*(?:급여|상여)).*?\[(?:20\d{2})(?:0[1-9]|1[0-2])(?:0[1-9]|[12]\d|3[01])\]\.pdf$",
    re.IGNORECASE,
)
_RECEIPT_FILE = re.compile(r"^(?P<year>20\d{2}).*?(?P<start>\d+)_(?P<end>\d+)\.pdf$", re.IGNORECASE)
_HEADER_ALIASES = {
    "세선지급금": "선지급금",
    "지급합계": "지급총액",
    "공제합계": "공제총액",
    "차인지급액": "실지급액",
}
_TOTAL_HEADERS = ("지급총액", "공제총액", "실지급액")
_TABLE_SIGNATURES = (
    ("성명",),
    ("기본급", "기본급여"),
    ("지급총액", "지급합계"),
    ("공제총액", "공제합계"),
    ("실지급액", "차인지급액"),
)
_GRID_IDENTITY_HEADERS = {"사원번호", "성명", "입사일", "퇴사일", "직급", "부서", "영수인"}
_HEADER_FILL = PatternFill("solid", fgColor="FF1F4E78")
_HEADER_FONT = Font(color="FFFFFFFF", bold=True)
_THIN_BORDER = Border(
    left=Side(style="thin", color="FFD9E2F3"),
    right=Side(style="thin", color="FFD9E2F3"),
    top=Side(style="thin", color="FFD9E2F3"),
    bottom=Side(style="thin", color="FFD9E2F3"),
)
_MONEY_FORMAT = "#,##0;[Red]-#,##0"


@dataclass(frozen=True)
class PayrollEmployeeRecord:
    year: int
    month: int
    employee_code: str
    employee_name: str
    resident_id: str
    payments: dict[str, int]
    deductions: dict[str, int]
    payment_total: int
    deduction_total: int
    net_total: int
    source_file: str
    source_page: int


@dataclass(frozen=True)
class PayrollSourceControl:
    source_file: str
    year: int
    month: int
    employee_count: int
    payment_total: int
    deduction_total: int
    net_total: int
    source_payment_total: int
    source_deduction_total: int
    source_net_total: int
    status: str


@dataclass(frozen=True)
class PayrollPdfResult:
    records: tuple[PayrollEmployeeRecord, ...]
    control: PayrollSourceControl
    payment_headers: tuple[str, ...]
    deduction_headers: tuple[str, ...]


@dataclass(frozen=True)
class DawonPipelineResult:
    output_directory: Path
    basic_info_path: Path
    payroll_paths: dict[int, Path]
    report_path: Path
    receipt_counts: dict[int, int]
    payroll_counts: dict[int, int]


def _append_unique(values: list[str], additions: list[str] | tuple[str, ...]) -> None:
    for value in additions:
        if value and value not in values:
            values.append(value)


def _word_lines(page: Any, cell: tuple[float, float, float, float] | None) -> list[tuple[float, str]]:
    if cell is None:
        return []
    words = page.crop(cell).extract_words(x_tolerance=1, y_tolerance=1)
    groups: list[list[Any]] = []
    for word in words:
        if not groups or abs(float(word["top"]) - float(groups[-1][0]["top"])) > 2:
            groups.append([word])
        else:
            groups[-1].append(word)
    return [
        (
            float(group[0]["top"]) - float(cell[1]),
            "".join(str(word["text"]) for word in sorted(group, key=lambda item: float(item["x0"]))),
        )
        for group in groups
    ]


def _compact_cell(value: object) -> str:
    return re.sub(r"\s+", "", str(value or ""))


def _canonical_header(value: object) -> str:
    compact = _compact_cell(value)
    return _HEADER_ALIASES.get(compact, compact)


def _main_payroll_table(page: Any) -> tuple[Any, list[list[object]]]:
    candidates: list[tuple[int, int, Any, list[list[object]]]] = []
    for table in page.find_tables():
        rows = table.extract()
        compact = "".join(_compact_cell(cell) for row in rows[:8] for cell in row)
        score = sum(any(alias in compact for alias in aliases) for aliases in _TABLE_SIGNATURES)
        size = len(rows) * max((len(row) for row in rows), default=0)
        candidates.append((score, size, table, rows))
    if not candidates:
        raise SchemaMismatch("급여대장 페이지에서 표를 찾지 못했습니다.")
    score, _, table, rows = max(candidates, key=lambda item: (item[0], item[1]))
    if score < 3:
        raise SchemaMismatch("급여대장 본문 표를 헤더명으로 확정할 수 없습니다.")
    return table, rows


def _amount(value: object) -> int | None:
    text = _compact_cell(value).replace("원", "")
    if not text:
        return None
    negative = text.startswith("(") and text.endswith(")")
    if negative:
        text = text[1:-1]
    if not re.fullmatch(r"-?\d[\d,]*", text):
        return None
    number = int(text.replace(",", ""))
    return -number if negative else number


def _is_grid_payroll(rows: list[list[object]]) -> bool:
    headers = {_canonical_header(cell) for row in rows[:8] for cell in row}
    return "사원번호" in headers and "성명" in headers and (
        "지급총액" in headers or "공제총액" in headers or "실지급액" in headers
    )


def _header_slots(page: Any, table: Any) -> dict[int, list[tuple[float, str]]]:
    slots: dict[int, list[tuple[float, str]]] = {}
    for column in range(2, 15):
        cell = table.rows[1].cells[column]
        values = []
        for top, text in _word_lines(page, cell):
            normalized = _HEADER_ALIASES.get(text, text)
            if normalized:
                values.append((top, normalized))
        if not values:
            raise SchemaMismatch(f"급여대장 지급·공제 헤더를 읽지 못했습니다: 열 {column + 1}")
        slots[column] = values
    return slots


def _slot_step(slots: list[tuple[float, str]]) -> float:
    differences = [slots[index + 1][0] - slots[index][0] for index in range(len(slots) - 1)]
    return median(differences) if differences else 11.4


def _cell_amounts(
    page: Any,
    cell: tuple[float, float, float, float] | None,
    slots: list[tuple[float, str]],
) -> dict[str, int]:
    if cell is None:
        return {}
    step = _slot_step(slots)
    result: dict[str, int] = {}
    for relative_top, text in _word_lines(page, cell):
        if not _MONEY.fullmatch(text):
            continue
        slot_index = round((relative_top - 2.0) / step)
        if slot_index < 0 or slot_index >= len(slots):
            raise SchemaMismatch(f"급여대장 금액의 세로 위치를 헤더에 연결할 수 없습니다: {text}")
        expected_top = 2.0 + slot_index * step
        if abs(relative_top - expected_top) > max(3.5, step * 0.35):
            raise SchemaMismatch(f"급여대장 금액의 세로 위치가 허용범위를 벗어났습니다: {text}")
        header = slots[slot_index][1]
        amount = int(text.replace(",", ""))
        result[header] = result.get(header, 0) + amount
    return result


def _parse_payroll_name(path: Path) -> tuple[int, int, int]:
    matched = _PAYROLL_FILE.fullmatch(path.name)
    if matched is not None:
        return int(matched.group("year")), int(matched.group("month")), int(matched.group("part"))
    matched = _DATED_PAYROLL_FILE.fullmatch(path.name)
    if matched is not None:
        return int(matched.group("year")), int(matched.group("month")), 1
    matched = _KOREAN_PAYROLL_FILE.fullmatch(path.name)
    if matched is None:
        raise SchemaMismatch(f"급여대장 파일명에서 연월을 찾지 못했습니다: {path.name}")
    year_text = matched.group("year")
    year = int(year_text) if len(year_text) == 4 else 2000 + int(year_text)
    return year, int(matched.group("month")), int(matched.group("part") or 1)


def _grid_structure(rows: list[list[object]]) -> tuple[int, int, int, int, dict[tuple[int, int], str]]:
    header_start = next(
        (
            row_index
            for row_index, row in enumerate(rows[:10])
            if "사원번호" in {_canonical_header(cell) for cell in row}
            and "성명" in {_canonical_header(cell) for cell in row}
        ),
        None,
    )
    if header_start is None:
        raise SchemaMismatch("급여대장 본문 표에서 사원번호·성명 헤더를 찾지 못했습니다.")
    employee_column = next(
        column for column, value in enumerate(rows[header_start]) if _canonical_header(value) == "사원번호"
    )
    name_column = next(
        column for column, value in enumerate(rows[header_start]) if _canonical_header(value) == "성명"
    )
    data_start = next(
        (
            row_index
            for row_index in range(header_start + 1, len(rows))
            if employee_column < len(rows[row_index])
            and name_column < len(rows[row_index])
            and _compact_cell(rows[row_index][employee_column])
            and (
                _compact_cell(rows[row_index][name_column])
                or _compact_cell(rows[row_index][employee_column]).startswith("합계")
            )
            and any(_amount(value) is not None for value in rows[row_index][2:])
        ),
        None,
    )
    if data_start is None:
        raise SchemaMismatch("급여대장 본문 표에서 첫 직원행을 찾지 못했습니다.")
    record_height = data_start - header_start
    if record_height < 1:
        raise SchemaMismatch("급여대장 헤더와 직원행의 반복 간격을 계산할 수 없습니다.")
    slots: dict[tuple[int, int], str] = {}
    for offset, row_index in enumerate(range(header_start, data_start)):
        for column, value in enumerate(rows[row_index]):
            header = _canonical_header(value)
            if header and header not in _GRID_IDENTITY_HEADERS:
                slots[(offset, column)] = header
    total_positions = {
        header: (offset, column)
        for (offset, column), header in slots.items()
        if header in _TOTAL_HEADERS
    }
    missing = [header for header in _TOTAL_HEADERS if header not in total_positions]
    if missing:
        raise SchemaMismatch(f"급여대장 합계 헤더를 찾지 못했습니다: {missing}")
    payment_total_column = total_positions["지급총액"][1]
    deduction_total_column = total_positions["공제총액"][1]
    if payment_total_column >= deduction_total_column:
        raise SchemaMismatch("급여대장 지급·공제 합계 열 순서가 올바르지 않습니다.")
    return data_start, record_height, employee_column, name_column, slots


def _grid_values(
    rows: list[list[object]],
    block_start: int,
    slots: dict[tuple[int, int], str],
) -> dict[str, int]:
    values: dict[str, int] = {header: 0 for header in slots.values()}
    for (offset, column), header in slots.items():
        row_index = block_start + offset
        if row_index >= len(rows) or column >= len(rows[row_index]):
            continue
        amount = _amount(rows[row_index][column])
        if amount is not None:
            values[header] = values.get(header, 0) + amount
    return values


def _parse_grid_payroll_pdf(source: Path, year: int, month: int) -> PayrollPdfResult:
    records: list[PayrollEmployeeRecord] = []
    payment_headers: list[str] = []
    deduction_headers: list[str] = []
    source_totals: dict[str, int] | None = None

    with pdfplumber.open(source) as document:
        for page_number, page in enumerate(document.pages, 1):
            try:
                _, rows = _main_payroll_table(page)
            except SchemaMismatch as exc:
                raise SchemaMismatch(f"{exc}: {source.name}, {page_number}쪽") from exc
            data_start, record_height, employee_column, name_column, slots = _grid_structure(rows)
            payment_total_column = next(
                column for (offset, column), header in slots.items() if header == "지급총액"
            )
            deduction_total_column = next(
                column for (offset, column), header in slots.items() if header == "공제총액"
            )
            page_payment_headers = [
                header
                for (_, column), header in slots.items()
                if column <= payment_total_column and header not in _TOTAL_HEADERS
            ]
            page_deduction_headers = [
                header
                for (_, column), header in slots.items()
                if payment_total_column < column <= deduction_total_column and header not in _TOTAL_HEADERS
            ]
            _append_unique(payment_headers, page_payment_headers)
            _append_unique(deduction_headers, page_deduction_headers)

            for block_start in range(data_start, len(rows), record_height):
                row = rows[block_start]
                employee_code = _compact_cell(row[employee_column]) if employee_column < len(row) else ""
                employee_name = _compact_cell(row[name_column]) if name_column < len(row) else ""
                if employee_code.startswith("합계"):
                    totals = _grid_values(rows, block_start, slots)
                    if all(header in totals for header in _TOTAL_HEADERS):
                        if source_totals is not None:
                            raise SchemaMismatch(f"급여대장 전체합계가 중복되었습니다: {source.name}")
                        source_totals = totals
                    continue
                if not employee_code or not employee_name:
                    continue
                values = _grid_values(rows, block_start, slots)
                missing_totals = [header for header in _TOTAL_HEADERS if header not in values]
                if missing_totals:
                    raise SchemaMismatch(
                        f"직원 합계가 누락됐습니다: {source.name}, {page_number}쪽, "
                        f"{employee_name}, {missing_totals}"
                    )
                payments = {header: values[header] for header in page_payment_headers if header in values}
                deductions = {header: values[header] for header in page_deduction_headers if header in values}
                payment_total = values["지급총액"]
                deduction_total = values["공제총액"]
                net_total = values["실지급액"]
                payment_gap = payment_total - sum(payments.values())
                deduction_gap = deduction_total - sum(deductions.values())
                if payment_gap < 0:
                    raise SchemaMismatch(
                        f"직원 지급항목 합계가 지급총액을 초과합니다: {source.name}, "
                        f"{page_number}쪽, {employee_name}"
                    )
                if deduction_gap < 0:
                    raise SchemaMismatch(
                        f"직원 공제항목 합계가 공제총액을 초과합니다: {source.name}, "
                        f"{page_number}쪽, {employee_name}"
                    )
                if payment_gap:
                    payments["미분류지급액"] = payment_gap
                    _append_unique(payment_headers, ["미분류지급액"])
                if deduction_gap:
                    deductions["미분류공제액"] = deduction_gap
                    _append_unique(deduction_headers, ["미분류공제액"])
                if payment_total - deduction_total != net_total:
                    raise SchemaMismatch(
                        f"직원 차인지급액이 일치하지 않습니다: {source.name}, "
                        f"{page_number}쪽, {employee_name}"
                    )
                records.append(
                    PayrollEmployeeRecord(
                        year,
                        month,
                        employee_code,
                        employee_name,
                        "",
                        payments,
                        deductions,
                        payment_total,
                        deduction_total,
                        net_total,
                        source.name,
                        page_number,
                    )
                )

    if not records:
        raise SchemaMismatch(f"급여대장에서 직원행을 찾지 못했습니다: {source.name}")
    if source_totals is None:
        raise SchemaMismatch(f"급여대장 전체합계를 읽지 못했습니다: {source.name}")
    payment_total = sum(record.payment_total for record in records)
    deduction_total = sum(record.deduction_total for record in records)
    net_total = sum(record.net_total for record in records)
    status = "PASS" if (
        payment_total == source_totals["지급총액"]
        and deduction_total == source_totals["공제총액"]
        and net_total == source_totals["실지급액"]
    ) else "FAILED"
    control = PayrollSourceControl(
        source.name,
        year,
        month,
        len(records),
        payment_total,
        deduction_total,
        net_total,
        source_totals["지급총액"],
        source_totals["공제총액"],
        source_totals["실지급액"],
        status,
    )
    if status != "PASS":
        raise SchemaMismatch(f"급여대장 전체합계 대사에 실패했습니다: {source.name}")
    return PayrollPdfResult(tuple(records), control, tuple(payment_headers), tuple(deduction_headers))


def _completed_page_slots(
    all_page_slots: list[dict[int, list[tuple[float, str]]]],
    page_index: int,
) -> dict[int, list[tuple[float, str]]]:
    completed: dict[int, list[tuple[float, str]]] = {}
    for column in range(2, 15):
        anchors: list[float] = []
        for page_slots in all_page_slots:
            for top, _ in page_slots[column]:
                if not any(abs(top - anchor) <= 3.5 for anchor in anchors):
                    anchors.append(top)
        anchors.sort()
        column_slots: list[tuple[float, str]] = []
        for anchor in anchors:
            visible = next(
                (item for item in all_page_slots[page_index][column] if abs(item[0] - anchor) <= 3.5),
                None,
            )
            if visible is None:
                for distance in range(1, len(all_page_slots)):
                    candidates = (page_index - distance, page_index + distance)
                    for candidate_index in candidates:
                        if candidate_index < 0 or candidate_index >= len(all_page_slots):
                            continue
                        visible = next(
                            (
                                item
                                for item in all_page_slots[candidate_index][column]
                                if abs(item[0] - anchor) <= 3.5
                            ),
                            None,
                        )
                        if visible is not None:
                            break
                    if visible is not None:
                        break
            if visible is None:
                raise SchemaMismatch(
                    f"\uae09\uc5ec\ub300\uc7a5 \ud5e4\ub354 \uc704\uce58\ub97c \ubcf5\uad6c\ud560 \uc218 \uc5c6\uc2b5\ub2c8\ub2e4: \uc5f4 {column + 1}"
                )
            column_slots.append((anchor, visible[1]))
        completed[column] = column_slots
    return completed


def parse_dawon_payroll_pdf(path: Path) -> PayrollPdfResult:
    source = Path(path).resolve(strict=True)
    year, month, _ = _parse_payroll_name(source)
    with pdfplumber.open(source) as document:
        if not document.pages:
            raise SchemaMismatch(f"급여대장 PDF에 페이지가 없습니다: {source.name}")
        try:
            _, first_rows = _main_payroll_table(document.pages[0])
        except SchemaMismatch as exc:
            raise SchemaMismatch(f"{exc}: {source.name}, 1쪽") from exc
    if _is_grid_payroll(first_rows):
        return _parse_grid_payroll_pdf(source, year, month)

    records: list[PayrollEmployeeRecord] = []
    payment_headers: list[str] = []
    deduction_headers: list[str] = []
    source_totals: dict[str, int] | None = None

    with pdfplumber.open(source) as document:
        all_page_slots: list[dict[int, list[tuple[float, str]]]] = []
        for page_number, page in enumerate(document.pages, 1):
            try:
                table, _ = _main_payroll_table(page)
            except SchemaMismatch as exc:
                raise SchemaMismatch(f"{exc}: {source.name}, {page_number}쪽") from exc
            all_page_slots.append(_header_slots(page, table))

        for page_number, page in enumerate(document.pages, 1):
            try:
                table, rows = _main_payroll_table(page)
            except SchemaMismatch as exc:
                raise SchemaMismatch(f"{exc}: {source.name}, {page_number}쪽") from exc
            if len(rows) < 4 or len(table.rows[1].cells) < 15:
                raise SchemaMismatch(f"급여대장 표 구조가 예상과 다릅니다: {source.name}, {page_number}쪽")
            slots = _completed_page_slots(all_page_slots, page_number - 1)
            page_payment_headers = [header for column in range(2, 9) for _, header in slots[column]]
            page_deduction_headers = [header for column in range(9, 14) for _, header in slots[column]]
            collisions = sorted(set(page_payment_headers) & set(page_deduction_headers))
            if collisions:
                raise SchemaMismatch(
                    f"동일 항목이 지급과 공제에 함께 있습니다: {source.name}, {page_number}쪽, {collisions}"
                )
            _append_unique(payment_headers, page_payment_headers)
            _append_unique(deduction_headers, page_deduction_headers)

            for row_index, row in enumerate(rows):
                first = str(row[0] or "") if row else ""
                second = str(row[1] or "") if len(row) > 1 else ""
                resident_match = _RESIDENT.search(second)
                if resident_match is not None:
                    employee_name = first.split("\n", 1)[0].strip()
                    if not employee_name:
                        raise SchemaMismatch(
                            f"급여대장 성명을 읽지 못했습니다: {source.name}, {page_number}쪽"
                        )
                    resident_digits = "".join(resident_match.groups())
                    values: dict[str, int] = {}
                    for column in range(2, 15):
                        for header, amount in _cell_amounts(
                            page, table.rows[row_index].cells[column], slots[column]
                        ).items():
                            values[header] = values.get(header, 0) + amount
                    payments = {header: values[header] for header in page_payment_headers if header in values}
                    deductions = {header: values[header] for header in page_deduction_headers if header in values}
                    missing_totals = [header for header in _TOTAL_HEADERS if header not in values]
                    if missing_totals:
                        raise SchemaMismatch(
                            f"직원 합계가 누락됐습니다: {source.name}, {page_number}쪽, {employee_name}, {missing_totals}"
                        )
                    payment_total = values["지급총액"]
                    deduction_total = values["공제총액"]
                    net_total = values["실지급액"]
                    payment_gap = payment_total - sum(payments.values())
                    deduction_gap = deduction_total - sum(deductions.values())
                    if payment_gap < 0:
                        raise SchemaMismatch(
                            f"\uc9c1\uc6d0 \uc9c0\uae09\ud56d\ubaa9 \ud569\uacc4\uac00 \uc9c0\uae09\ucd1d\uc561\uc744 \ucd08\uacfc\ud569\ub2c8\ub2e4: {source.name}, {page_number}\ucabd, {employee_name}"
                        )
                    if deduction_gap < 0:
                        raise SchemaMismatch(
                            f"\uc9c1\uc6d0 \uacf5\uc81c\ud56d\ubaa9 \ud569\uacc4\uac00 \uacf5\uc81c\ucd1d\uc561\uc744 \ucd08\uacfc\ud569\ub2c8\ub2e4: {source.name}, {page_number}\ucabd, {employee_name}"
                        )
                    if payment_gap:
                        payments["\ubbf8\ubd84\ub958\uc9c0\uae09\uc561"] = payment_gap
                        _append_unique(payment_headers, ["\ubbf8\ubd84\ub958\uc9c0\uae09\uc561"])
                    if deduction_gap:
                        deductions["\ubbf8\ubd84\ub958\uacf5\uc81c\uc561"] = deduction_gap
                        _append_unique(deduction_headers, ["\ubbf8\ubd84\ub958\uacf5\uc81c\uc561"])
                    if payment_total - deduction_total != net_total:
                        raise SchemaMismatch(
                            f"직원 차인지급액이 일치하지 않습니다: {source.name}, {page_number}쪽, {employee_name}"
                        )
                    records.append(
                        PayrollEmployeeRecord(
                            year,
                            month,
                            resident_digits,
                            employee_name,
                            f"{resident_digits[:6]}-{resident_digits[6:]}",
                            payments,
                            deductions,
                            payment_total,
                            deduction_total,
                            net_total,
                            source.name,
                            page_number,
                        )
                    )
                elif first.startswith("전체"):
                    if source_totals is not None:
                        raise SchemaMismatch(f"급여대장 전체합계가 중복되었습니다: {source.name}")
                    source_totals = {}
                    for column in range(2, 15):
                        source_totals.update(
                            _cell_amounts(page, table.rows[row_index].cells[column], slots[column])
                        )

    if not records:
        raise SchemaMismatch(f"급여대장에서 직원행을 찾지 못했습니다: {source.name}")
    if source_totals is None or any(header not in source_totals for header in _TOTAL_HEADERS):
        raise SchemaMismatch(f"급여대장 전체합계를 읽지 못했습니다: {source.name}")
    payment_total = sum(record.payment_total for record in records)
    deduction_total = sum(record.deduction_total for record in records)
    net_total = sum(record.net_total for record in records)
    status = "PASS" if (
        payment_total == source_totals["지급총액"]
        and deduction_total == source_totals["공제총액"]
        and net_total == source_totals["실지급액"]
    ) else "FAILED"
    control = PayrollSourceControl(
        source.name,
        year,
        month,
        len(records),
        payment_total,
        deduction_total,
        net_total,
        source_totals["지급총액"],
        source_totals["공제총액"],
        source_totals["실지급액"],
        status,
    )
    if status != "PASS":
        raise SchemaMismatch(f"급여대장 전체합계 대사에 실패했습니다: {source.name}")
    return PayrollPdfResult(tuple(records), control, tuple(payment_headers), tuple(deduction_headers))


def _write_payroll_sheet(
    worksheet: Any,
    headers: list[str],
    records: list[PayrollEmployeeRecord],
) -> None:
    worksheet.sheet_view.showGridLines = False
    worksheet.append(headers)
    payment_start = 4
    payment_total_index = headers.index("지급합계")
    deduction_total_index = headers.index("공제합계")
    for record in records:
        values: dict[str, object] = {
            "사원번호": record.employee_code,
            "성명": record.employee_name,
            "주민등록번호": format_resident_id(record.resident_id),
            "급여구분": "급여",
            **record.payments,
            "지급합계": record.payment_total,
            **record.deductions,
            "공제합계": record.deduction_total,
            "차인지급액": record.net_total,
        }
        worksheet.append([values.get(header) for header in headers])
    worksheet.freeze_panes = "E2"
    worksheet.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{max(1, len(records) + 1)}"
    worksheet.row_dimensions[1].height = 22
    for cell in worksheet[1]:
        cell.fill = _HEADER_FILL
        cell.font = _HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = _THIN_BORDER
    for column in range(1, len(headers) + 1):
        letter = get_column_letter(column)
        worksheet.column_dimensions[letter].width = 16 if column > 4 else (16 if column == 3 else 13)
    for row_number in range(2, len(records) + 2):
        worksheet.cell(row_number, 1).number_format = "@"
        worksheet.cell(row_number, 3).number_format = "@"
        for column in range(1, len(headers) + 1):
            cell = worksheet.cell(row_number, column)
            cell.border = _THIN_BORDER
            if column > 4:
                cell.number_format = _MONEY_FORMAT
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="center" if column in {1, 4} else "left")
    worksheet.sheet_properties.pageSetUpPr.fitToPage = True
    worksheet.page_setup.orientation = worksheet.ORIENTATION_LANDSCAPE
    worksheet.page_setup.fitToWidth = 1
    worksheet.page_setup.fitToHeight = 0
    worksheet.print_title_rows = "1:1"
    worksheet.print_area = f"A1:{get_column_letter(len(headers))}{max(1, len(records) + 1)}"


def build_dawon_payroll_workbook(
    *,
    year: int,
    pdf_files: list[Path],
    output_path: Path,
) -> tuple[Path, list[PayrollSourceControl], list[str], list[str]]:
    ordered = sorted(pdf_files, key=lambda path: (_parse_payroll_name(path)[1], _parse_payroll_name(path)[2]))
    results_by_path = {}
    total = len(ordered)
    logger.info("급여대장 PDF 분석 시작: %d년 (총 %d개)", year, total)
    try:
        with ThreadPoolExecutor(max_workers=min(4, total)) as executor:
            futures = {executor.submit(parse_dawon_payroll_pdf, path): path for path in ordered}
            for future in as_completed(futures):
                path = futures[future]
                results_by_path[path] = future.result()
    except Exception as exc:
        logger.error("급여대장 PDF 분석 오류: %d년, %s", year, exc)
        raise
    logger.info("급여대장 PDF 분석 완료: %d년, %d개", year, len(results_by_path))
    results = [results_by_path[path] for path in ordered]
    payment_headers: list[str] = []
    deduction_headers: list[str] = []
    records_by_month: dict[int, list[PayrollEmployeeRecord]] = {month: [] for month in range(1, 13)}
    for result in results:
        _append_unique(payment_headers, result.payment_headers)
        _append_unique(deduction_headers, result.deduction_headers)
        for record in result.records:
            records_by_month[record.month].append(record)

    missing_months = [month for month in range(1, 13) if not records_by_month[month]]
    if missing_months:
        raise SchemaMismatch(f"급여대장 월이 누락되었습니다: {year}년 {missing_months}")

    headers = [
        "사원번호",
        "성명",
        "주민등록번호",
        "급여구분",
        *payment_headers,
        "지급합계",
        *deduction_headers,
        "공제합계",
        "차인지급액",
    ]
    workbook = Workbook()
    workbook.remove(workbook.active)
    for month in range(1, 13):
        sheet = workbook.create_sheet(f"{month:02d}월")
        _write_payroll_sheet(sheet, headers, records_by_month[month])
    workbook.calculation.calcMode = "auto"
    workbook.calculation.fullCalcOnLoad = True
    workbook.calculation.forceFullCalc = True

    destination = Path(output_path).resolve()
    destination.parent.mkdir(parents=True, exist_ok=True)
    temporary = destination.parent / f".{destination.stem}.{uuid4().hex}.tmp.xlsx"
    try:
        workbook.save(temporary)
        checked = load_workbook(temporary, read_only=True, data_only=False)
        try:
            if checked.sheetnames != [f"{month:02d}월" for month in range(1, 13)]:
                raise SchemaMismatch(f"생성 급여대장 월별 시트 검증에 실패했습니다: {year}")
            for month in range(1, 13):
                sheet = checked[f"{month:02d}월"]
                if sheet.max_row != len(records_by_month[month]) + 1:
                    raise SchemaMismatch(f"생성 급여대장 행 수 검증에 실패했습니다: {year}년 {month:02d}월")
                if [cell.value for cell in sheet[1]] != headers:
                    raise SchemaMismatch(f"생성 급여대장 헤더 검증에 실패했습니다: {year}년 {month:02d}월")
        finally:
            checked.close()
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)
    return destination, [result.control for result in results], payment_headers, deduction_headers


def _receipt_range(path: Path) -> tuple[int, int, int]:
    matched = _RECEIPT_FILE.fullmatch(path.name)
    if matched is None:
        raise SchemaMismatch(f"원천징수영수증 파일명에서 연도·인덱스 범위를 찾지 못했습니다: {path.name}")
    return int(matched.group("year")), int(matched.group("start")), int(matched.group("end"))


def parse_dawon_receipt_bundle(path: Path) -> tuple[int, list[WithholdingPdfData]]:
    source = Path(path).resolve(strict=True)
    year, start_index, end_index = _receipt_range(source)
    expected_count = end_index - start_index + 1
    result: list[WithholdingPdfData] = []
    with pdfplumber.open(source) as document:
        if len(document.pages) != expected_count * 3:
            raise SchemaMismatch(
                f"원천징수영수증 페이지 수와 파일명 인덱스 범위가 다릅니다: {source.name}, "
                f"예상={expected_count * 3}, 실제={len(document.pages)}"
            )
        for offset, index in enumerate(range(start_index, end_index + 1)):
            page_tables = [document.pages[offset * 3 + page].extract_tables() for page in range(3)]
            data = parse_withholding_tables(
                path=Path(f"{index}_receipt.pdf"),
                year=year,
                page_tables=page_tables,
                page_count=3,
            )
            if data.pdf_index != index:
                raise SchemaMismatch(f"원천징수영수증 인덱스 검증에 실패했습니다: {source.name}, {index}")
            result.append(data)
    return year, result


def extract_dawon_receipts(pdf_files: list[Path]) -> dict[int, list[WithholdingPdfData]]:
    sources = sorted((Path(path) for path in pdf_files), key=lambda path: _receipt_range(path))
    with ThreadPoolExecutor(max_workers=min(4, len(sources))) as executor:
        batches = list(executor.map(parse_dawon_receipt_bundle, sources))
    by_year: dict[int, list[WithholdingPdfData]] = {}
    for year, batch in batches:
        by_year.setdefault(year, []).extend(batch)
    for year, rows in by_year.items():
        rows.sort(key=lambda item: item.pdf_index)
        indexes = [item.pdf_index for item in rows]
        if indexes != list(range(1, len(rows) + 1)):
            raise SchemaMismatch(f"원천징수영수증 인덱스가 연속되지 않습니다: {year}, {indexes[:5]}...")
        residents = [item.resident_id for item in rows]
        if len(residents) != len(set(residents)):
            raise SchemaMismatch(f"같은 연도 원천징수영수증 주민등록번호가 중복되었습니다: {year}")
    return dict(sorted(by_year.items()))


def build_dawon_basic_info(
    *,
    template_path: Path,
    receipt_data: dict[int, list[WithholdingPdfData]],
    output_path: Path,
    company_name: str,
    company_id: str,
) -> Path:
    template = Path(template_path).resolve(strict=True)
    destination = Path(output_path).resolve()
    all_records = [_record(data, data.resident_id) for rows in receipt_data.values() for data in rows]
    taxfree_headers = _taxfree_headers(all_records)
    headers = WITHHOLDING_HEADERS + taxfree_headers
    current_source = template
    destination.parent.mkdir(parents=True, exist_ok=True)
    for year, rows in sorted(receipt_data.items()):
        records = [_record(data, data.resident_id) for data in rows]
        temporary = _write_with_excel(
            source=current_source,
            destination=destination,
            sheet_name=f"원천징수추출({str(year)[-2:]}년)",
            headers=headers,
            records=records,
            taxfree_headers=taxfree_headers,
            year=year,
            company_name=company_name,
            company_id=company_id,
        )
        os.replace(temporary, destination)
        current_source = destination

    checked = load_workbook(destination, read_only=True, data_only=False)
    try:
        for year, rows in sorted(receipt_data.items()):
            sheet_name = f"원천징수추출({str(year)[-2:]}년)"
            if sheet_name not in checked.sheetnames:
                raise SchemaMismatch(f"생성 기초정보에 시트가 없습니다: {sheet_name}")
            if checked[sheet_name].max_row != len(rows) + 1:
                raise SchemaMismatch(f"생성 기초정보 인원수 검증에 실패했습니다: {year}")
    finally:
        checked.close()
    return destination


def run_dawon_test_pipeline(
    *,
    input_directory: Path,
    output_directory: Path,
    template_path: Path,
    company_name: str = "주식회사 다원로지스틱스",
    company_id: str = "803-81-00806",
) -> DawonPipelineResult:
    source_root = Path(input_directory).resolve(strict=True)
    output_root = Path(output_directory).resolve()
    payroll_root = source_root / "급여대장"
    receipt_root = source_root / "원천징수영수증"
    payroll_files = sorted(payroll_root.rglob("*.pdf"))
    receipt_files = sorted(receipt_root.glob("*.pdf"))
    if not payroll_files:
        raise SchemaMismatch(f"다원 급여대장 PDF가 없습니다: {payroll_root}")
    if not receipt_files:
        raise SchemaMismatch(f"다원 원천징수영수증 PDF가 없습니다: {receipt_root}")

    staging = output_root.parent / f".{output_root.name}.{uuid4().hex}.tmp"
    staging.mkdir(parents=True, exist_ok=False)
    try:
        receipt_data = extract_dawon_receipts(receipt_files)
        basic_info_path = build_dawon_basic_info(
            template_path=template_path,
            receipt_data=receipt_data,
            output_path=staging / f"{company_name}_기초정보.xlsx",
            company_name=company_name,
            company_id=company_id,
        )

        payroll_paths: dict[int, Path] = {}
        payroll_controls: list[PayrollSourceControl] = []
        payroll_headers: dict[int, dict[str, list[str]]] = {}
        payroll_counts: dict[int, int] = {}
        payroll_unclassified: dict[int, dict[str, dict[str, int]]] = {}
        years = sorted({_parse_payroll_name(path)[0] for path in payroll_files})
        for year in years:
            year_files = [path for path in payroll_files if _parse_payroll_name(path)[0] == year]
            target, controls, payment_headers, deduction_headers = build_dawon_payroll_workbook(
                year=year,
                pdf_files=year_files,
                output_path=staging / f"{company_name}_{str(year)[-2:]}년 급여대장.xlsx",
            )
            payroll_paths[year] = target
            payroll_controls.extend(controls)
            payroll_headers[year] = {
                "payment": payment_headers,
                "deduction": deduction_headers,
            }
            checked_payroll = load_workbook(target, read_only=True, data_only=False)
            try:
                payroll_counts[year] = sum(
                    checked_payroll[f"{month:02d}\uc6d4"].max_row - 1 for month in range(1, 13)
                )
                unclassified = {
                    "payment": {"record_count": 0, "amount": 0},
                    "deduction": {"record_count": 0, "amount": 0},
                }
                for worksheet in checked_payroll.worksheets:
                    headers = [cell.value for cell in next(worksheet.iter_rows(min_row=1, max_row=1))]
                    for label, kind in (
                        ("\ubbf8\ubd84\ub958\uc9c0\uae09\uc561", "payment"),
                        ("\ubbf8\ubd84\ub958\uacf5\uc81c\uc561", "deduction"),
                    ):
                        if label not in headers:
                            continue
                        column = headers.index(label) + 1
                        for row in range(2, worksheet.max_row + 1):
                            value = worksheet.cell(row, column).value
                            if value not in (None, 0):
                                unclassified[kind]["record_count"] += 1
                                unclassified[kind]["amount"] += int(value)
                payroll_unclassified[year] = unclassified
            finally:
                checked_payroll.close()

        report = {
            "company_name": company_name,
            "company_id": company_id,
            "input_directory": str(source_root),
            "receipt_counts": {str(year): len(rows) for year, rows in receipt_data.items()},
            "payroll_counts": {str(year): count for year, count in payroll_counts.items()},
            "payroll_headers": payroll_headers,
            "payroll_unclassified": payroll_unclassified,
            "payroll_controls": [control.__dict__ for control in payroll_controls],
            "outputs": [
                {"name": basic_info_path.name, "sha256": sha256(basic_info_path)},
                *[
                    {"name": path.name, "sha256": sha256(path)}
                    for _, path in sorted(payroll_paths.items())
                ],
            ],
        }
        report_path = staging / "DAWON_검증보고서.json"
        report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")

        output_root.mkdir(parents=True, exist_ok=True)
        final_paths: dict[str, Path] = {}
        for path in staging.iterdir():
            target = output_root / path.name
            os.replace(path, target)
            final_paths[path.name] = target
        shutil.rmtree(staging, ignore_errors=True)
        final_basic = final_paths[basic_info_path.name]
        final_payroll = {year: final_paths[path.name] for year, path in payroll_paths.items()}
        final_report = final_paths[report_path.name]
        return DawonPipelineResult(
            output_root,
            final_basic,
            final_payroll,
            final_report,
            {year: len(rows) for year, rows in receipt_data.items()},
            payroll_counts,
        )
    except Exception:
        shutil.rmtree(staging, ignore_errors=True)
        raise
