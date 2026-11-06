from pathlib import Path
from unittest.mock import patch

from payroll_amendment.amended_withholding import AmendedWithholdingResult
from scripts.run_amended_withholding import main


def test_completion_output_is_single_summary_line(capsys):
    result = AmendedWithholdingResult(
        (Path("normal1.pdf"), Path("normal2.pdf")),
        (Path("marked1.pdf"), Path("marked2.pdf")),
        Path("처리결과.csv"),
        3,
    )
    with (
        patch("scripts.run_amended_withholding.generate_company_amended_withholding", return_value=result),
        patch("sys.argv", ["run_amended_withholding.py", "--company-dir", "."]),
    ):
        assert main() == 0

    assert capsys.readouterr().out == (
        "수정 원천징수 PDF 생성 완료: 총 4개 "
        "(수정후 2개, 변경표시 2개, 이미지 PDF 제외 3개)\n"
    )
