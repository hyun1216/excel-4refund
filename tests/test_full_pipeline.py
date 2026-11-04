from pathlib import Path
from unittest.mock import patch

from payroll_amendment.correction_inputs import CorrectionInputResult
from payroll_amendment.full_pipeline import run_full_pipeline
from payroll_amendment.input_intake import IntakeResult
from payroll_amendment.payroll_converter import UniversalConversionResult
from payroll_amendment.payroll_status import PayrollStatusResult


def test_full_pipeline_routes_generated_payroll_and_runs_status(tmp_path):
    project = tmp_path / "project"
    inbox = project / "input"
    resources = project / "Resources"
    template = resources / "Template" / "Temp_기초정보.xlsx"
    template.parent.mkdir(parents=True)
    inbox.mkdir()
    for path in (
        template,
        resources / "Template" / "Temp_수정 원천징수영수증 생성 Tool.xlsx",
        resources / "지급항목_매칭표.xlsx",
        resources / "직업분류_연계표.xlsx",
    ):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"resource")

    company = project / "노무_세무 수정신고" / "주식회사 테스트"
    source_input = company / "1. 급상여현황 생성" / "1.Input"
    source_input.mkdir(parents=True)
    intake_report = company / "input_접수보고서.json"
    intake_report.write_text("{}", encoding="utf-8")
    intake = IntakeResult(
        "주식회사 테스트",
        "123-45-67890",
        company,
        (),
        (),
        intake_report,
    )

    basic = company / "주식회사 테스트_기초정보.xlsx"
    payroll = company / "주식회사 테스트_23년 급여대장.xlsx"
    conversion_report = company / "trans_검증보고서.json"
    for path in (basic, payroll, conversion_report):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"output")
    conversion = UniversalConversionResult(company, (basic, payroll), conversion_report)
    status_output = company / "1. 급상여현황 생성" / "급상여현황" / "주식회사 테스트_23년 급상여현황.xlsx"
    status_report = status_output.parent / "급상여현황_검토보고서.json"
    status_output.parent.mkdir(parents=True)
    status_output.write_bytes(b"status")
    status_report.write_text("{}", encoding="utf-8")
    status = PayrollStatusResult((2023,), (status_output,), status_report, basic)
    correction_inputs = CorrectionInputResult((2023,), (), (), (), ())

    with (
        patch("payroll_amendment.full_pipeline.organize_input_batch", return_value=intake),
        patch("payroll_amendment.full_pipeline.PayrollConverter.convert_company_input", return_value=conversion),
        patch("payroll_amendment.full_pipeline.discover_employee_info_sources", return_value=(None, None)),
        patch("payroll_amendment.full_pipeline.populate_basic_info_employee_sheets"),
        patch("payroll_amendment.full_pipeline.generate_company_payroll_status", return_value=status),
        patch("payroll_amendment.full_pipeline.prepare_correction_inputs", return_value=correction_inputs),
    ):
        result = run_full_pipeline(project_root=project)

    routed = company / "2.근로소득 수정신고" / "1.Input" / payroll.name
    assert routed.is_file()
    assert not payroll.exists()
    assert result.payroll_files == (routed,)
    assert result.report_path.is_file()
    assert result.correction_inputs == correction_inputs


def test_full_pipeline_can_resume_from_employee_info_stage(tmp_path):
    project = tmp_path / "project"
    resources = project / "Resources"
    for path in (
        resources / "Template" / "Temp_기초정보.xlsx",
        resources / "Template" / "Temp_수정 원천징수영수증 생성 Tool.xlsx",
        resources / "지급항목_매칭표.xlsx",
        resources / "직업분류_연계표.xlsx",
    ):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"resource")

    company = project / "노무_세무 수정신고" / "주식회사 테스트"
    report_path = company / "1. 급상여현황 생성" / "input_접수보고서.json"
    report_path.parent.mkdir(parents=True)
    report_path.write_text(
        '{"company_name":"주식회사 테스트","company_id":"123-45-67890",'
        '"moved_files":[{"destination":"누락된_직업분류매핑결과.xlsx"}]}',
        encoding="utf-8",
    )
    basic = company / "주식회사 테스트_기초정보.xlsx"
    payroll = company / "주식회사 테스트_23년 급여대장.xlsx"
    mapped = company / "수정본_직업분류매핑결과.xlsx"
    conversion_report = company / "trans_검증보고서.json"
    for path in (basic, payroll, mapped, conversion_report):
        path.write_bytes(b"existing")

    status_output = company / "1. 급상여현황 생성" / "급상여현황" / "주식회사 테스트_23년 급상여현황.xlsx"
    status_report = status_output.parent / "급상여현황_검토보고서.json"
    status_output.parent.mkdir(parents=True)
    status_output.write_bytes(b"status")
    status_report.write_text("{}", encoding="utf-8")
    status = PayrollStatusResult((2023,), (status_output,), status_report, basic)
    correction_inputs = CorrectionInputResult((2023,), (), (), (), ())

    with (
        patch("payroll_amendment.full_pipeline.PayrollConverter.convert_company_input") as convert,
        patch("payroll_amendment.full_pipeline.populate_basic_info_employee_sheets") as populate,
        patch("payroll_amendment.full_pipeline.generate_company_payroll_status", return_value=status),
        patch("payroll_amendment.full_pipeline.prepare_correction_inputs", return_value=correction_inputs),
    ):
        result = run_full_pipeline(
            project_root=project,
            existing_company_directory=company,
            start_stage=3,
        )

    convert.assert_not_called()
    populate.assert_called_once()
    routed = company / "2.근로소득 수정신고" / "1.Input" / payroll.name
    assert routed.is_file()
    assert not payroll.exists()
    assert result.basic_info_path == basic
    assert result.payroll_files == (routed,)


def test_full_pipeline_stage_three_reuses_already_routed_payroll(tmp_path):
    project = tmp_path / "project"
    resources = project / "Resources"
    for path in (
        resources / "Template" / "Temp_기초정보.xlsx",
        resources / "Template" / "Temp_수정 원천징수영수증 생성 Tool.xlsx",
        resources / "지급항목_매칭표.xlsx",
        resources / "직업분류_연계표.xlsx",
    ):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"resource")

    company = project / "노무_세무 수정신고" / "주식회사 테스트"
    report_path = company / "1. 급상여현황 생성" / "input_접수보고서.json"
    report_path.parent.mkdir(parents=True)
    report_path.write_text(
        '{"company_name":"주식회사 테스트","company_id":"123-45-67890","moved_files":[]}',
        encoding="utf-8",
    )
    basic = company / "주식회사 테스트_기초정보.xlsx"
    payroll = company / "2.근로소득 수정신고" / "1.Input" / "주식회사 테스트_23년 급여대장.xlsx"
    mapped = company / "수정본_직업분류매핑결과.xlsx"
    conversion_report = company / "trans_검증보고서.json"
    for path in (basic, payroll, mapped, conversion_report):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"existing")

    status_output = company / "급상여현황.xlsx"
    status_report = company / "급상여현황_검토보고서.json"
    status = PayrollStatusResult((2023,), (status_output,), status_report, basic)
    correction_inputs = CorrectionInputResult((2023,), (), (), (), ())

    with (
        patch("payroll_amendment.full_pipeline.PayrollConverter.convert_company_input") as convert,
        patch("payroll_amendment.full_pipeline.populate_basic_info_employee_sheets"),
        patch("payroll_amendment.full_pipeline.generate_company_payroll_status", return_value=status),
        patch("payroll_amendment.full_pipeline.prepare_correction_inputs", return_value=correction_inputs),
    ):
        result = run_full_pipeline(
            project_root=project,
            existing_company_directory=company,
            start_stage=3,
        )

    convert.assert_not_called()
    assert result.payroll_files == (payroll,)
