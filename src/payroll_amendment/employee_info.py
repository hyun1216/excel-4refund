from __future__ import annotations

from collections import defaultdict
from dataclasses import asdict, dataclass
import os
from pathlib import Path
import re
from typing import Sequence
from uuid import uuid4

from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill

from .errors import SchemaMismatch


MASTER_SHEET = "고용2025_표준직업_연계표"
OCCUPATION_SHEET = "직업분류코드"
HEALTH_SHEET = "건보정보"
INPUT_OCCUPATION_COLUMN = 11
OUTPUT_START_COLUMN = 16


@dataclass(frozen=True)
class LinkageEntry:
    middle_name: str
    group_name: str
    detail_name: str
    ksco_code: str
    ksco_name: str
    source_row: int


@dataclass(frozen=True)
class MappingDecision:
    code: str
    name: str
    status: str
    score: float
    reason: str


@dataclass(frozen=True)
class EmployeeInfoLoadResult:
    employment_source: str | None
    health_source: str | None
    occupation_rows: int
    mapped_rows: int
    manual_review_rows: int
    not_eligible_rows: int
    missing_occupation_rows: int
    health_rows: int

    def to_report(self) -> dict[str, object]:
        return asdict(self)


@dataclass(frozen=True)
class OccupationMappingResult:
    source_path: Path
    output_path: Path
    processed_rows: int
    mapped_rows: int
    manual_review_rows: int
    not_eligible_rows: int


def normalize_label(value: object) -> str:
    return re.sub(r"[^0-9A-Za-z가-힣]", "", str(value or "").lower())


PREFERRED_EXACT_MAPPINGS = {
    normalize_label("미용·여행·숙박·음식·경비·청소직"): ("9999", "기타 서비스 관련 단순 종사자"),
    normalize_label("택배원 및 기타 운송 종사자"): ("9222", "택배원"),
    normalize_label("청소 종사자"): ("9419", "기타 청소 관련 종사원"),
    normalize_label("자동차 운전원"): ("8729", "기타 자동차 운전원"),
    normalize_label("미용 서비스원"): ("4319", "기타 미용 관련 서비스 종사원"),
    normalize_label("주방장 및 조리사"): ("4519", "기타 조리사"),
    normalize_label("식당 서비스원"): ("4531", "식음료 서비스 종사원"),
    normalize_label("판매 종사자"): ("5212", "상점 판매원"),
    normalize_label("매장 계산원 및 매표원"): ("5214", "매장 계산원 및 요금 정산원"),
    normalize_label("식품 가공 기능원"): ("7109", "기타 식품 가공 관련 기능 종사원"),
    normalize_label("의복 제조원 및 수선원"): ("7229", "기타 의복 제조 관련 기능 종사원"),
    normalize_label("운송장비 정비원"): ("7529", "기타 운송장비 정비원"),
    normalize_label("전기·전자 기기 설치·수리원"): ("7619", "기타 전기‧전자기기 설치 및 수리원"),
    normalize_label("건설구조 기능원"): ("7829", "기타 건설 관련 기능 종사원"),
}


def normalize_code(value: object) -> str:
    if value is None or isinstance(value, bool):
        return ""
    if isinstance(value, float) and value.is_integer():
        value = int(value)
    text = str(value).strip()
    if re.fullmatch(r"\d+(?:\.0+)?", text):
        text = text.split(".", 1)[0]
    return text


def load_linkage(path: Path) -> list[LinkageEntry]:
    source = Path(path).resolve(strict=True)
    workbook = load_workbook(source, read_only=True, data_only=True)
    try:
        if MASTER_SHEET not in workbook.sheetnames:
            raise SchemaMismatch(f"직업분류 연계표에 '{MASTER_SHEET}' 시트가 없습니다.")
        sheet = workbook[MASTER_SHEET]
        expected_headers = {
            3: "중분류명",
            5: "소분류명",
            7: "세분류명",
            13: "세분류2",
            14: "세분류명2",
        }
        for column, expected in expected_headers.items():
            actual = sheet.cell(2, column).value
            if normalize_label(actual) != normalize_label(expected):
                raise SchemaMismatch(
                    f"직업분류 연계표 {sheet.cell(2, column).coordinate} 헤더가 "
                    f"예상과 다릅니다: {actual!r}"
                )

        entries: list[LinkageEntry] = []
        seen: set[tuple[str, str, str, str, str]] = set()
        for row_number in range(3, sheet.max_row + 1):
            middle_name = str(sheet.cell(row_number, 3).value or "").strip()
            group_name = str(sheet.cell(row_number, 5).value or "").strip()
            detail_name = str(sheet.cell(row_number, 7).value or "").strip()
            ksco_code = normalize_code(sheet.cell(row_number, 13).value)
            ksco_name = str(sheet.cell(row_number, 14).value or "").strip()
            if not any((middle_name, group_name, detail_name, ksco_code, ksco_name)):
                continue
            if not all((middle_name, group_name, detail_name, ksco_code, ksco_name)):
                raise SchemaMismatch(
                    f"직업분류 연계표 {row_number}행의 C/E/G/M/N 값이 완전하지 않습니다."
                )
            key = (
                normalize_label(middle_name),
                normalize_label(group_name),
                normalize_label(detail_name),
                ksco_code,
                normalize_label(ksco_name),
            )
            if key in seen:
                continue
            seen.add(key)
            entries.append(
                LinkageEntry(
                    middle_name,
                    group_name,
                    detail_name,
                    ksco_code,
                    ksco_name,
                    row_number,
                )
            )
        if not entries:
            raise SchemaMismatch("직업분류 연계표 C/E/G/M/N열에 매핑 데이터가 없습니다.")
        return entries
    finally:
        workbook.close()


class OccupationNameMapper:
    def __init__(self, entries: Sequence[LinkageEntry]) -> None:
        exact: dict[str, dict[tuple[str, str], LinkageEntry]] = defaultdict(dict)
        middle_exact: dict[str, dict[tuple[str, str], LinkageEntry]] = defaultdict(dict)
        for entry in entries:
            mapping_key = (entry.ksco_code, normalize_label(entry.ksco_name))
            exact[normalize_label(entry.group_name)][mapping_key] = entry
            exact[normalize_label(entry.detail_name)][mapping_key] = entry
            exact[normalize_label(entry.ksco_name)][mapping_key] = entry
            middle_exact[normalize_label(entry.middle_name)][mapping_key] = entry
        self.exact = dict(exact)
        self.middle_exact = dict(middle_exact)

    def map(self, occupation_name: str) -> MappingDecision:
        normalized = normalize_label(occupation_name)
        if not normalized:
            return MappingDecision("", "", "미매핑", 0.0, "K열 직종명 없음")
        preferred = PREFERRED_EXACT_MAPPINGS.get(normalized)
        if preferred:
            return MappingDecision(
                preferred[0],
                preferred[1],
                "매핑완료",
                1.0,
                "사전 지정 매핑",
            )
        candidates = list(self.exact.get(normalized, {}).values())
        if len(candidates) == 1:
            candidate = candidates[0]
            return MappingDecision(
                candidate.ksco_code,
                candidate.ksco_name,
                "매핑완료",
                1.0,
                "직종명 정확히 일치",
            )
        if len(candidates) > 1:
            specific = [
                item
                for item in candidates
                if normalize_label(item.detail_name)
                and normalize_label(item.detail_name) in normalized
            ]
            candidate = (
                max(specific, key=lambda item: (len(normalize_label(item.detail_name)), -item.source_row))
                if specific
                else min(candidates, key=lambda item: (item.source_row, item.ksco_code))
            )
            return MappingDecision(
                candidate.ksco_code,
                candidate.ksco_name,
                "수동확인필요",
                1.0,
                f"동일 직종명에 표준분류 {len(candidates)}개",
            )
        middle_candidates = list(self.middle_exact.get(normalized, {}).values())
        if len(middle_candidates) == 1:
            candidate = middle_candidates[0]
            return MappingDecision(
                candidate.ksco_code,
                candidate.ksco_name,
                "매핑완료",
                1.0,
                "중분류명 정확히 일치",
            )
        if len(middle_candidates) > 1:
            candidate = min(middle_candidates, key=lambda item: (item.source_row, item.ksco_code))
            return MappingDecision(
                candidate.ksco_code,
                candidate.ksco_name,
                "수동확인필요",
                1.0,
                f"동일 중분류명에 표준분류 {len(middle_candidates)}개",
            )
        return MappingDecision(
            "경정청구대상아님",
            "경정청구대상아님",
            "경정청구대상아님",
            0.0,
            "연계표 C/E/G/N열에 동일 직종명 없음",
        )


def create_occupation_mapping_result(
    *,
    source_path: Path,
    linkage_path: Path,
    output_path: Path | None = None,
) -> OccupationMappingResult:
    """원본 고용정보를 보존하고 P:T열에 매핑 결과가 추가된 파일을 생성한다."""
    source_file = Path(source_path).resolve(strict=True)
    destination = Path(output_path) if output_path else source_file.with_name(
        f"{source_file.stem}_직업분류매핑결과.xlsx"
    )
    destination = destination.resolve()
    if destination == source_file:
        raise SchemaMismatch("직업분류 매핑 결과는 원본 파일과 다른 경로여야 합니다.")

    mapper = OccupationNameMapper(load_linkage(linkage_path))
    workbook = load_workbook(source_file)
    temporary = destination.with_name(f".{destination.stem}.{uuid4().hex}.tmp.xlsx")
    try:
        if len(workbook.sheetnames) != 1:
            raise SchemaMismatch("근로자고용정보현황조회 대상 시트를 하나로 특정할 수 없습니다.")
        sheet = workbook[workbook.sheetnames[0]]
        for row, column, expected in (
            (1, 1, "연번"),
            (1, 2, "근로자구분"),
            (1, 3, "주민등록번호"),
            (1, 4, "성명"),
            (2, 10, "고용상태"),
            (2, 11, "직종"),
            (2, 12, "상실사유"),
            (2, 13, "근로자원부번호"),
            (2, 14, "사업개시번호"),
            (2, 15, "미승인하수급인번호"),
        ):
            _validate_header(sheet, row, column, expected, "근로자고용정보현황조회")

        headers = ("표준직업분류코드", "표준직업분류명", "매핑상태", "매핑점수", "검토사유")
        header_fill = PatternFill("solid", fgColor="D9EAF7")
        for offset, header in enumerate(headers):
            cell = sheet.cell(2, OUTPUT_START_COLUMN + offset, header)
            cell.font = Font(bold=True)
            cell.fill = header_fill

        processed = mapped = manual = not_eligible = 0
        for row_number in range(3, sheet.max_row + 1):
            occupation_name = str(sheet.cell(row_number, INPUT_OCCUPATION_COLUMN).value or "").strip()
            if not occupation_name:
                continue
            decision = mapper.map(occupation_name)
            for offset, value in enumerate(
                (decision.code, decision.name, decision.status, decision.score, decision.reason)
            ):
                sheet.cell(row_number, OUTPUT_START_COLUMN + offset, value)
            processed += 1
            if decision.status == "매핑완료":
                mapped += 1
            elif decision.status == "경정청구대상아님":
                not_eligible += 1
            else:
                manual += 1

        for column, width in (("P", 18), ("Q", 28), ("R", 16), ("S", 12), ("T", 38)):
            sheet.column_dimensions[column].width = width

        destination.parent.mkdir(parents=True, exist_ok=True)
        workbook.save(temporary)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise
    finally:
        workbook.close()

    try:
        verification = load_workbook(temporary, read_only=True, data_only=False)
        try:
            verified_sheet = verification[verification.sheetnames[0]]
            if verified_sheet.cell(2, OUTPUT_START_COLUMN).value != "표준직업분류코드":
                raise SchemaMismatch("저장 결과의 직업분류 매핑 헤더 검증에 실패했습니다.")
            verified_rows = sum(
                1
                for (status,) in verified_sheet.iter_rows(
                    min_row=3,
                    min_col=OUTPUT_START_COLUMN + 2,
                    max_col=OUTPUT_START_COLUMN + 2,
                    values_only=True,
                )
                if status
            )
            if verified_rows != processed:
                raise SchemaMismatch(
                    f"직업분류 매핑 저장 행 수가 다릅니다: 처리 {processed}, 저장 {verified_rows}"
                )
        finally:
            verification.close()
        os.replace(temporary, destination)
    finally:
        temporary.unlink(missing_ok=True)

    return OccupationMappingResult(
        source_file,
        destination,
        processed,
        mapped,
        manual,
        not_eligible,
    )


def discover_employee_info_sources(source_root: Path) -> tuple[Path | None, Path | None]:
    root = Path(source_root).resolve(strict=True)
    files = [
        path
        for path in root.rglob("*.xlsx")
        if path.is_file() and not path.name.startswith("~$")
    ]
    employment_candidates = [
        path
        for path in files
        if "근로자고용정보현황조회" in path.name
    ]
    mapped_employment = [path for path in employment_candidates if "직업분류매핑결과" in path.name]
    employment = mapped_employment or [
        path for path in employment_candidates if "직업분류매핑결과" not in path.name
    ]
    health = [path for path in files if "사업장가입자명부" in path.name]
    if len(employment) > 1:
        raise SchemaMismatch(
            "근로자고용정보현황조회 파일을 하나로 확정할 수 없습니다: "
            + ", ".join(path.name for path in employment)
        )
    if len(health) > 1:
        raise SchemaMismatch(
            "사업장가입자명부 파일을 하나로 확정할 수 없습니다: "
            + ", ".join(path.name for path in health)
        )
    return (employment[0] if employment else None, health[0] if health else None)


def _validate_header(sheet, row: int, column: int, expected: str, label: str) -> None:
    actual = sheet.cell(row, column).value
    if normalize_label(actual) != normalize_label(expected):
        raise SchemaMismatch(
            f"{label} {sheet.cell(row, column).coordinate} 헤더가 예상과 다릅니다: {actual!r}"
        )


def _clear_values(sheet, start_row: int, end_column: int) -> None:
    for row in sheet.iter_rows(
        min_row=start_row,
        max_row=sheet.max_row,
        min_col=1,
        max_col=end_column,
    ):
        for cell in row:
            if cell.value is not None:
                cell.value = None


def _load_occupation_sheet(target, source_path: Path, linkage_path: Path) -> tuple[int, int, int, int, int]:
    workbook = load_workbook(source_path, data_only=True)
    try:
        if len(workbook.sheetnames) != 1:
            raise SchemaMismatch("근로자고용정보현황조회 대상 시트를 하나로 특정할 수 없습니다.")
        source = workbook[workbook.sheetnames[0]]
        for row, column, expected in (
            (1, 1, "연번"),
            (1, 2, "근로자구분"),
            (1, 3, "주민등록번호"),
            (1, 4, "성명"),
            (2, 10, "고용상태"),
            (2, 11, "직종"),
            (2, 12, "상실사유"),
            (2, 13, "근로자원부번호"),
            (2, 14, "사업개시번호"),
            (2, 15, "미승인하수급인번호"),
        ):
            _validate_header(source, row, column, expected, "근로자고용정보현황조회")

        pre_mapped = normalize_label(source.cell(2, 16).value) == normalize_label("표준직업분류코드")
        if pre_mapped:
            for column, expected in (
                (17, "표준직업분류명"),
                (18, "매핑상태"),
                (19, "매핑점수"),
                (20, "검토사유"),
            ):
                _validate_header(source, 2, column, expected, "직업분류매핑결과")
            mapper = None
        else:
            mapper = OccupationNameMapper(load_linkage(linkage_path))

        _clear_values(target, 3, 19)
        occupation_rows = mapped = manual = not_eligible = missing = 0
        target_row = 3
        for source_row in range(3, source.max_row + 1):
            source_values = [
                source.cell(source_row, column).value
                for column in range(1, 21 if pre_mapped else 16)
            ]
            if not any(value not in (None, "") for value in source_values):
                continue
            if pre_mapped:
                code = normalize_code(source.cell(source_row, 16).value)
                name = str(source.cell(source_row, 17).value or "").strip()
                status = str(source.cell(source_row, 18).value or "").strip()
            else:
                occupation = str(source.cell(source_row, 11).value or "").strip()
                assert mapper is not None
                decision = mapper.map(occupation)
                code = decision.code
                name = decision.name
                status = decision.status
            values = (
                source.cell(source_row, 1).value,
                source.cell(source_row, 2).value,
                source.cell(source_row, 3).value,
                source.cell(source_row, 4).value,
                source.cell(source_row, 5).value,
                source.cell(source_row, 6).value,
                source.cell(source_row, 7).value,
                source.cell(source_row, 8).value,
                source.cell(source_row, 9).value,
                None,
                source.cell(source_row, 10).value,
                source.cell(source_row, 11).value,
                source.cell(source_row, 12).value,
                source.cell(source_row, 13).value,
                source.cell(source_row, 14).value,
                source.cell(source_row, 15).value,
                source.cell(source_row, 11).value,
                code,
                name,
            )
            for column, value in enumerate(values, 1):
                target.cell(target_row, column, value)
            occupation_rows += 1
            if status == "매핑완료":
                mapped += 1
            elif status == "수동확인필요":
                manual += 1
            elif status == "경정청구대상아님":
                not_eligible += 1
            else:
                missing += 1
            target_row += 1
        return occupation_rows, mapped, manual, not_eligible, missing
    finally:
        workbook.close()


def _load_health_sheet(target, source_path: Path) -> int:
    workbook = load_workbook(source_path, data_only=True)
    try:
        if len(workbook.sheetnames) != 1:
            raise SchemaMismatch("사업장가입자명부 대상 시트를 하나로 특정할 수 없습니다.")
        source = workbook[workbook.sheetnames[0]]
        for row, column, expected in (
            (1, 1, "순번"),
            (1, 2, "일련번호"),
            (2, 3, "증번호"),
            (2, 4, "주민등록번호"),
            (2, 5, "성명"),
            (1, 6, "자격취득일"),
            (1, 7, "자격상실일"),
        ):
            _validate_header(source, row, column, expected, "사업장가입자명부")

        _clear_values(target, 2, 7)
        health_rows = 0
        target_row = 2
        for source_row in range(3, source.max_row + 1):
            values = [source.cell(source_row, column).value for column in range(1, 8)]
            if not any(value not in (None, "") for value in values):
                continue
            for column, value in enumerate(values, 1):
                target.cell(target_row, column, value)
            health_rows += 1
            target_row += 1
        return health_rows
    finally:
        workbook.close()


def populate_basic_info_employee_sheets(
    *,
    basic_info_path: Path,
    employment_info_path: Path | None,
    health_roster_path: Path | None,
    occupation_linkage_path: Path,
) -> EmployeeInfoLoadResult:
    target_path = Path(basic_info_path).resolve(strict=True)
    employment = Path(employment_info_path).resolve(strict=True) if employment_info_path else None
    health = Path(health_roster_path).resolve(strict=True) if health_roster_path else None
    if employment is None and health is None:
        return EmployeeInfoLoadResult(None, None, 0, 0, 0, 0, 0, 0)

    if employment is not None:
        source = load_workbook(employment, read_only=True, data_only=False)
        try:
            if len(source.sheetnames) != 1:
                raise SchemaMismatch("근로자고용정보현황조회 대상 시트를 하나로 특정할 수 없습니다.")
            pre_mapped = normalize_label(
                source[source.sheetnames[0]].cell(2, OUTPUT_START_COLUMN).value
            ) == normalize_label("표준직업분류코드")
        finally:
            source.close()
        if not pre_mapped:
            employment = create_occupation_mapping_result(
                source_path=employment,
                linkage_path=occupation_linkage_path,
            ).output_path

    workbook = load_workbook(target_path)
    temporary = target_path.with_name(f".{target_path.stem}.{uuid4().hex}.tmp.xlsx")
    try:
        missing_sheets = [
            name
            for name in (OCCUPATION_SHEET, HEALTH_SHEET)
            if name not in workbook.sheetnames
        ]
        if missing_sheets:
            raise SchemaMismatch(f"기초정보에 필요한 시트가 없습니다: {missing_sheets}")

        occupation_counts = (0, 0, 0, 0, 0)
        if employment is not None:
            occupation_counts = _load_occupation_sheet(
                workbook[OCCUPATION_SHEET],
                employment,
                occupation_linkage_path,
            )
        health_rows = 0
        if health is not None:
            health_rows = _load_health_sheet(workbook[HEALTH_SHEET], health)

        workbook.save(temporary)
    finally:
        workbook.close()

    try:
        verification = load_workbook(temporary, read_only=True, data_only=False)
        try:
            if employment is not None:
                sheet = verification[OCCUPATION_SHEET]
                saved_rows = sum(
                    1
                    for row in sheet.iter_rows(min_row=3, max_col=19, values_only=True)
                    if any(value not in (None, "") for value in row)
                )
                if saved_rows != occupation_counts[0]:
                    raise SchemaMismatch(
                        f"직업분류코드 저장 행 수가 다릅니다: "
                        f"입력 {occupation_counts[0]}, 저장 {saved_rows}"
                    )
            if health is not None:
                sheet = verification[HEALTH_SHEET]
                saved_rows = sum(
                    1
                    for row in sheet.iter_rows(min_row=2, max_col=7, values_only=True)
                    if any(value not in (None, "") for value in row)
                )
                if saved_rows != health_rows:
                    raise SchemaMismatch(
                        f"건보정보 저장 행 수가 다릅니다: 입력 {health_rows}, 저장 {saved_rows}"
                    )
        finally:
            verification.close()
        os.replace(temporary, target_path)
    finally:
        temporary.unlink(missing_ok=True)

    return EmployeeInfoLoadResult(
        employment.name if employment else None,
        health.name if health else None,
        occupation_counts[0],
        occupation_counts[1],
        occupation_counts[2],
        occupation_counts[3],
        occupation_counts[4],
        health_rows,
    )
