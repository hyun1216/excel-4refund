from __future__ import annotations

from dataclasses import asdict, dataclass, field
from decimal import Decimal
from pathlib import Path
from typing import Any, Literal


@dataclass(frozen=True)
class CompanyInfo:
    company_id: str
    display_name: str
    source_fields: dict[str, str]


@dataclass(frozen=True)
class TaxExemptionLine:
    item_name: str
    amount: Decimal


@dataclass(frozen=True)
class TaxExemption:
    company_id: str
    year: int
    employee_id: str
    name: str
    resident_id: str
    pdf_index: str
    serial_number: str
    taxable_deduction_total: Decimal
    lines: tuple[TaxExemptionLine, ...]
    source_sheet: str
    source_row: int
    source_values: dict[str, Any] = field(default_factory=dict)


@dataclass(frozen=True)
class SourceDocument:
    kind: str
    path: Path
    sha256: str
    year: int | None = None


@dataclass
class ValidationIssue:
    code: str
    severity: Literal["INFO", "WARNING", "ERROR"]
    public_message: str
    sheet: str | None = None
    cell: str | None = None
    year: int | None = None


@dataclass
class StageResult:
    stage_code: str
    company_id: str
    year: int | None
    status: Literal["SUCCEEDED", "FAILED", "SKIPPED", "NEEDS_REVIEW"]
    artifacts: list[Path] = field(default_factory=list)
    metrics: dict[str, int | str] = field(default_factory=dict)
    issues: list[ValidationIssue] = field(default_factory=list)
    input_fingerprints: dict[str, str] = field(default_factory=dict)
    implementation_version: str = "0.1.0"

    def to_dict(self, root: Path | None = None) -> dict[str, Any]:
        data = asdict(self)
        data["artifacts"] = [str(p.relative_to(root) if root and p.is_relative_to(root) else p) for p in self.artifacts]
        return data


@dataclass
class IngestedInputs:
    company: CompanyInfo
    years: list[int]
    basic_info_path: Path
    tool_workbooks_by_year: dict[int, Path]
    payroll_files_by_year: dict[int, list[Path]]
    pdf_files_by_year: dict[int, list[Path]]
    withholding_excel_files_by_year: dict[int, list[Path]]
    exemptions_by_year: dict[int, list[TaxExemption]]
    input_manifest_path: Path

