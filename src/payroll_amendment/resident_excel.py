from __future__ import annotations

import os
from pathlib import Path
import re
from uuid import uuid4

from openpyxl import load_workbook

from .errors import SchemaMismatch
from .infra import format_resident_id


def _is_resident_header(value: object) -> bool:
    label = re.sub(r"\s+", "", str(value or ""))
    return (
        "주민등록번호" in label
        or label.startswith("주민번호")
        or "주민(외국인)등록번호" in label
    )


def normalize_resident_ids_in_workbook(path: Path) -> int:
    """기존 xlsx의 표형 주민번호 열을 하이픈 포함 텍스트로 원자적 갱신한다."""
    target = Path(path).resolve(strict=True)
    if target.suffix.lower() != ".xlsx":
        raise SchemaMismatch(f"기존 주민번호 보정은 xlsx만 지원합니다: {target.name}")
    workbook = load_workbook(target)
    temporary = target.with_name(f".{target.stem}.{uuid4().hex}.tmp.xlsx")
    touched = 0
    try:
        for sheet in workbook.worksheets:
            for header_row in range(1, min(3, sheet.max_row) + 1):
                for header_cell in sheet[header_row]:
                    if not _is_resident_header(header_cell.value):
                        continue
                    for row in range(header_row + 1, sheet.max_row + 1):
                        cell = sheet.cell(row, header_cell.column)
                        if cell.value in (None, "") or cell.data_type == "f":
                            continue
                        formatted = format_resident_id(cell.value)
                        if cell.value != formatted or cell.number_format != "@":
                            cell.value = formatted
                            cell.number_format = "@"
                            touched += 1
        workbook.save(temporary)
        os.replace(temporary, target)
    finally:
        workbook.close()
        temporary.unlink(missing_ok=True)
    return touched
