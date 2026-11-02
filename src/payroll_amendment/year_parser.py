from __future__ import annotations

from datetime import datetime
import re


def parse_years(value: str | list[str | int], *, min_year: int = 2000, max_year: int | None = None) -> list[int]:
    max_year = max_year or datetime.now().year
    values = value if isinstance(value, list) else [part.strip() for part in value.split(",")]
    if not values or any(str(item).strip() == "" for item in values):
        raise ValueError("수행연도가 비어 있습니다.")
    years: list[int] = []
    for item in values:
        text = str(item).strip()
        match = re.fullmatch(r"(?:(20\d{2})|(\d{2}))년?", text)
        if not match:
            raise ValueError(f"잘못된 수행연도: {text}")
        year = int(match.group(1) or f"20{match.group(2)}")
        if not min_year <= year <= max_year:
            raise ValueError(f"허용 범위를 벗어난 수행연도: {year}")
        years.append(year)
    return sorted(set(years))


def infer_year(text: str) -> int | None:
    matches = re.findall(r"(?<!\d)(20\d{2}|\d{2})년", text)
    if not matches:
        # 일부 Windows ZIP은 '년' 두 글자를 OEM 문자로 보존한다. 연도 앞 경계가
        # 명확한 파일명/폴더명에 한해서 두 자리 연도를 복구한다.
        matches = re.findall(r"(?:^|[_\s])(20\d{2}|\d{2})(?=\D)", text)
    if len(set(matches)) != 1:
        return None
    value = int(matches[0])
    return value if value >= 2000 else 2000 + value
