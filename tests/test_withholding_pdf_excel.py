from decimal import Decimal
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from openpyxl import load_workbook
from pypdf import PdfWriter

from payroll_amendment.errors import PdfToExcelError
from payroll_amendment.models import TaxExemption
from payroll_amendment.withholding_pdf_excel import (
    _excel_value,
    _exemptions_by_index,
    _pdf_index,
    _write_workbook,
    convert_withholding_pdfs_to_excel,
    parse_withholding_tables,
)


def sample_tables() -> list[list[list[list[str | None]]]]:
    page_one = [
        [
            ["근로소득 원천징수영수증", None, None, None, None],
            ["소득자", "⑥성명 홍길동", "⑦주민(외국인)등록번호 900101-1234567", None, None],
            [None, "구분", None, "주(현)", "합계"],
            [None, "계", None, "10,000,000", "10,000,000"],
            [None, "비과세식사대", "P01", "2,000,000", "2,000,000"],
            [None, "비과세소득계", None, "2,000,000", "2,000,000"],
            [None, "구분", None, "소득세", "지방소득세"],
            [None, "결정세액", None, "100,000", "10,000"],
            [None, "기납부세액", None, None, None],
            [None, "주(현)근무지", None, "120,000", "12,000"],
            [None, "차감징수세액", None, "-20,000", "-2,000"],
        ],
    ]
    page_two = [
        [
            ["정산명세", "근로소득금액", "8,000,000", "종합소득과세표준", None, "5,000,000"],
            [None, "차감소득금액", "5,000,000", "세액감면계", None, "0"],
            [None, "그 밖의 소득공제계", "0", "근로소득", None, "550,000"],
            [None, None, None, "세액공제계", None, "600,000"],
        ]
    ]
    return [page_one, page_two]


class WithholdingPdfExcelTests(unittest.TestCase):
    def test_reads_brity_prefix_pdf_index(self) -> None:
        self.assertEqual(1773, _pdf_index(Path("1773_지춘화_발행자보관용.pdf")))

    def test_parses_required_values_and_taxfree_detail(self) -> None:
        data = parse_withholding_tables(
            path=Path("회사_E001_홍길동_2025_발행자보관용_7.pdf"),
            year=2025,
            page_tables=sample_tables(),
            page_count=3,
        )

        self.assertEqual(7, data.pdf_index)
        self.assertEqual("홍길동", data.employee_name)
        self.assertEqual("9001011234567", data.resident_id)
        self.assertEqual(10_000_000, data.values["과세소득_총계"])
        self.assertEqual(120_000, data.values["기납부세액_소득세"])
        self.assertEqual(-2_000, data.values["차감징수세액_지방세"])
        self.assertEqual(5_000_000, data.values["과세표준"])
        self.assertEqual(550_000, data.values["근로소득"])
        self.assertEqual(1, len(data.taxfree_incomes))
        self.assertEqual("비과세식사대", data.taxfree_incomes[0].item_name)

    def test_writes_reviewable_workbook_and_reopens(self) -> None:
        data = parse_withholding_tables(
            path=Path("회사_E001_홍길동_2025_발행자보관용_7.pdf"),
            year=2025,
            page_tables=sample_tables(),
            page_count=3,
        )
        with tempfile.TemporaryDirectory() as folder:
            source = Path(folder) / "source.pdf"
            writer = PdfWriter()
            for _ in range(3):
                writer.add_blank_page(width=595, height=842)
            with source.open("wb") as stream:
                writer.write(stream)
            target = Path(folder) / "output.xlsx"
            _write_workbook(data, source, target)

            workbook = load_workbook(target, read_only=True, data_only=False)
            try:
                self.assertEqual(["Table 1", "데이터입력"], workbook.sheetnames)
                self.assertEqual("홍길동", workbook["데이터입력"]["D8"].value)
                self.assertEqual(10_000_000, workbook["데이터입력"]["D26"].value)
                self.assertEqual(2_000_000, workbook["데이터입력"]["D29"].value)
                self.assertEqual("f", workbook["Table 1"]["D15"].data_type)
            finally:
                workbook.close()

            workbook = load_workbook(target, data_only=True)
            try:
                form = workbook["Table 1"]
                self.assertTrue(any(cell.value is not None for row in form.iter_rows() for cell in row))
            finally:
                workbook.close()

    def test_rejects_missing_required_table(self) -> None:
        with self.assertRaises(PdfToExcelError):
            parse_withholding_tables(
                path=Path("sample_1.pdf"),
                year=2025,
                page_tables=[[[["일반 문서"]]], [[["정산명세"]]]],
                page_count=2,
            )

    def test_accepts_page_two_table_without_settlement_title(self) -> None:
        page_tables = sample_tables()
        page_tables[1][0][0][0] = None

        data = parse_withholding_tables(
            path=Path("sample_1.pdf"),
            year=2025,
            page_tables=page_tables,
            page_count=2,
        )

        self.assertEqual(8_000_000, data.values["근로소득금액"])
        self.assertEqual(5_000_000, data.values["과세표준"])

    def test_escapes_formula_like_pdf_text(self) -> None:
        self.assertEqual("'=HYPERLINK(\"bad\")", _excel_value('=HYPERLINK("bad")'))
        self.assertEqual(-2000, _excel_value(-2000))

    def test_batch_quarantines_failed_pdf(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            source = root / "회사_E001_홍길동_2025_발행자보관용_7.pdf"
            source.write_bytes(b"invalid pdf")
            with patch(
                "payroll_amendment.withholding_pdf_excel.extract_withholding_pdf",
                side_effect=PdfToExcelError("필수 항목 누락"),
            ):
                with self.assertRaises(PdfToExcelError):
                    convert_withholding_pdfs_to_excel(
                        pdf_files=[source],
                        exemptions=[TaxExemption(
                            company_id="1234567890",
                            year=2025,
                            employee_id="E001",
                            name="홍길동",
                            resident_id="900101-1234567",
                            pdf_index="7",
                            serial_number="1",
                            taxable_deduction_total=Decimal("0"),
                            lines=(),
                            source_sheet="수당조정(비과세)",
                            source_row=5,
                        )],
                        year=2025,
                        output_dir=root / "output",
                        quarantine_dir=root / "quarantine",
                    )
            self.assertTrue((root / "quarantine" / source.name).is_file())
            self.assertTrue((root / "quarantine" / "failure.json").is_file())

    def test_rejects_duplicate_exemption_index(self) -> None:
        exemption = TaxExemption(
            company_id="1234567890",
            year=2025,
            employee_id="E001",
            name="홍길동",
            resident_id="900101-1234567",
            pdf_index="7",
            serial_number="1",
            taxable_deduction_total=Decimal("0"),
            lines=(),
            source_sheet="수당조정(비과세)",
            source_row=5,
        )
        with self.assertRaises(PdfToExcelError):
            _exemptions_by_index([exemption, exemption])


if __name__ == "__main__":
    unittest.main()
