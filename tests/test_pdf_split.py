from decimal import Decimal
import json
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from pypdf import PdfReader, PdfWriter

from payroll_amendment.errors import PdfSplitError
from payroll_amendment.models import TaxExemption
from payroll_amendment.pdf_split import (
    FIRST_PAGE_MARKER,
    PdfInspection,
    _output_name,
    _target_from_filename,
    _write_pages,
    archive_prepared_withholding_pdfs,
    classify_page_texts,
    prepare_withholding_pdfs,
)


def exemption(index: int) -> TaxExemption:
    return TaxExemption(
        company_id="1234567890",
        year=2025,
        employee_id=f"E{index:03d}",
        name=f"대상{index}",
        resident_id=f"90010{index}-1234567",
        pdf_index=str(index),
        serial_number=str(index),
        taxable_deduction_total=Decimal("0"),
        lines=(),
        source_sheet="수당조정(비과세)",
        source_row=index + 4,
    )


class PdfSplitTests(unittest.TestCase):
    def test_matches_downloaded_individual_filename_by_employee_id_and_name(self) -> None:
        target = exemption(7)
        source = Path(
            "주식회사 주엔_E007_대상7_2025_근로소득원천징수영수증_발행자보관용_20260915_115950.pdf"
        )

        self.assertEqual(target, _target_from_filename(source, [target]))

    @staticmethod
    def _blank_pdf(path: Path, pages: int) -> None:
        writer = PdfWriter()
        for _ in range(pages):
            writer.add_blank_page(width=595, height=842)
        with path.open("wb") as stream:
            writer.write(stream)

    def test_classifies_one_marker_as_individual(self) -> None:
        kind, starts = classify_page_texts([FIRST_PAGE_MARKER, "2페이지", "3페이지"])
        self.assertEqual("individual", kind)
        self.assertEqual((0,), starts)

    def test_classifies_repeated_marker_as_bundle_with_variable_pages(self) -> None:
        kind, starts = classify_page_texts([
            FIRST_PAGE_MARKER,
            "2페이지",
            "3페이지",
            FIRST_PAGE_MARKER,
            "2페이지",
        ])
        self.assertEqual("bundle", kind)
        self.assertEqual((0, 3), starts)

    def test_classifies_title_fallback_when_form_marker_is_not_extractable(self) -> None:
        first_page = "근로소득원천징수영수증 1. 법인명(상호) 주식회사 테스트"
        kind, starts = classify_page_texts([
            first_page,
            "2페이지",
            "3페이지",
            first_page,
            "2페이지",
            "3페이지",
        ])
        self.assertEqual("bundle", kind)
        self.assertEqual((0, 3), starts)

    def test_rejects_missing_or_late_marker(self) -> None:
        with self.assertRaises(PdfSplitError):
            classify_page_texts(["표식 없음"])
        with self.assertRaises(PdfSplitError):
            classify_page_texts(["표지", FIRST_PAGE_MARKER])

    def test_output_name_matches_existing_brity_rule(self) -> None:
        self.assertEqual(
            "7_대상7_발행자보관용.pdf",
            _output_name("테스트", 2025, exemption(7)),
        )

    def test_writes_requested_page_range_and_reopens(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "source.pdf"
            target = Path(folder) / "split.pdf"
            self._blank_pdf(source, 5)
            _write_pages(PdfReader(str(source)), range(1, 4), target)
            self.assertEqual(3, len(PdfReader(str(target)).pages))

    def test_archives_prepared_pdfs_by_employee_name(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            sources = [root / "1_receipt.pdf", root / "2_receipt.pdf"]
            for source in sources:
                self._blank_pdf(source, 2)
            identities = [
                SimpleNamespace(pdf_index=1, employee_name="대상1", resident_id="900101-1234567"),
                SimpleNamespace(pdf_index=2, employee_name="대상2", resident_id="900102-1234567"),
            ]

            result = archive_prepared_withholding_pdfs(
                pdf_files=sources,
                identities=identities,
                year=2025,
                output_dir=root / "25년",
            )

            self.assertEqual(
                ["1_대상1_발행자보관용.pdf", "2_대상2_발행자보관용.pdf"],
                [path.name for path in result.output_files],
            )
            self.assertTrue(all(len(PdfReader(str(path)).pages) == 2 for path in result.output_files))
            manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
            self.assertEqual(2025, manifest["year"])
            self.assertEqual(2, len(manifest["outputs"]))

    def test_prepares_target_files_from_bundle(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "전체_2025년_원천징수.pdf"
            self._blank_pdf(source, 5)
            inspection = PdfInspection(
                source.resolve(),
                "bundle",
                5,
                (0, 3),
                (
                    f"{FIRST_PAGE_MARKER} 900101-1234567",
                    "2페이지",
                    "3페이지",
                    f"{FIRST_PAGE_MARKER} 900102-1234567",
                    "2페이지",
                ),
            )
            with patch("payroll_amendment.pdf_split.inspect_withholding_pdf", return_value=inspection):
                result = prepare_withholding_pdfs(
                    pdf_files=[source],
                    exemptions=[exemption(1), exemption(2)],
                    company_name="테스트회사",
                    year=2025,
                    output_dir=root / "prepared",
                    quarantine_dir=root / "quarantine",
                )
            self.assertEqual([3, 2], [len(PdfReader(str(path)).pages) for path in result.output_files])
            self.assertEqual(
                ["1_대상1_발행자보관용.pdf", "2_대상2_발행자보관용.pdf"],
                [path.name for path in result.output_files],
            )
            manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
            self.assertEqual("bundle", manifest["input_kind"])
            self.assertEqual(2, len(manifest["outputs"]))

    def test_prepares_individual_pdf_using_filename_index(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "기존파일_1.pdf"
            self._blank_pdf(source, 2)
            inspection = PdfInspection(
                source.resolve(),
                "individual",
                2,
                (0,),
                (
                    f"{FIRST_PAGE_MARKER} 900101-1234567",
                    "2페이지",
                ),
            )
            with patch("payroll_amendment.pdf_split.inspect_withholding_pdf", return_value=inspection):
                result = prepare_withholding_pdfs(
                    pdf_files=[source],
                    exemptions=[exemption(1)],
                    company_name="테스트회사",
                    year=2025,
                    output_dir=root / "prepared",
                    quarantine_dir=root / "quarantine",
                )
            self.assertEqual(1, len(result.output_files))
            self.assertEqual(
                "1_대상1_발행자보관용.pdf",
                result.output_files[0].name,
            )
            self.assertEqual(2, len(PdfReader(str(result.output_files[0])).pages))
            manifest = json.loads(result.manifest_path.read_text(encoding="utf-8"))
            self.assertEqual("individual", manifest["input_kind"])

    def test_quarantines_bundle_on_resident_mismatch(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "전체_2025년_원천징수.pdf"
            self._blank_pdf(source, 2)
            inspection = PdfInspection(
                source.resolve(),
                "bundle",
                2,
                (0, 1),
                (
                    f"{FIRST_PAGE_MARKER} 990101-1234567",
                    FIRST_PAGE_MARKER,
                ),
            )
            with patch("payroll_amendment.pdf_split.inspect_withholding_pdf", return_value=inspection):
                with self.assertRaises(PdfSplitError):
                    prepare_withholding_pdfs(
                        pdf_files=[source],
                        exemptions=[exemption(1)],
                        company_name="테스트회사",
                        year=2025,
                        output_dir=root / "prepared",
                        quarantine_dir=root / "quarantine",
                    )
            report = json.loads((root / "quarantine" / "failure.json").read_text(encoding="utf-8"))
            self.assertEqual("QUARANTINED", report["status"])
            self.assertEqual(1, len(report["files"]))


if __name__ == "__main__":
    unittest.main()
