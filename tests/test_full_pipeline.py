import json
from pathlib import Path
from unittest.mock import patch

from payroll_amendment.correction_inputs import CorrectionInputResult
from payroll_amendment.full_pipeline import BasicInfoReviewResult, run_full_pipeline
from payroll_amendment.input_intake import IntakeResult
from payroll_amendment.payroll_converter import UniversalConversionResult
from payroll_amendment.payroll_status import BasicInfoPreparationResult, PayrollStatusResult
from payroll_amendment.submission_target_lists import SubmissionTargetListResult


def _submission_result(company: Path) -> SubmissionTargetListResult:
    return SubmissionTargetListResult(
        (),
        (),
        (),
        (),
        (),
        (),
        company / "2.근로소득 수정신고" / "2.Output" / "제출용_대상자명단_생성보고서.json",
        (),
    )


def test_full_pipeline_routes_generated_payroll_and_stops_for_review(tmp_path):
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
    preparation = BasicInfoPreparationResult((2023,), basic, 2, 1)
    correction_inputs = CorrectionInputResult((2023,), (), (), (), ())
    submission = _submission_result(company)

    with (
        patch("payroll_amendment.full_pipeline.organize_input_batch", return_value=intake),
        patch(
            "payroll_amendment.full_pipeline.PayrollConverter.convert_company_input",
            return_value=conversion,
        ) as convert,
        patch("payroll_amendment.full_pipeline.discover_employee_info_sources", return_value=(None, None)),
        patch("payroll_amendment.full_pipeline.populate_basic_info_employee_sheets"),
        patch(
            "payroll_amendment.full_pipeline.prepare_company_basic_info_for_review",
            return_value=preparation,
        ),
        patch(
            "payroll_amendment.full_pipeline.generate_company_payroll_status",
            return_value=status,
        ) as generate_status,
        patch(
            "payroll_amendment.full_pipeline.prepare_correction_inputs",
            return_value=correction_inputs,
        ) as prepare_correction,
        patch(
            "payroll_amendment.full_pipeline.generate_company_submission_target_lists",
            return_value=submission,
        ) as generate_submission,
    ):
        result = run_full_pipeline(project_root=project)

    convert.assert_called_once_with(
        input_root=source_input,
        output_directory=company,
        basic_info_template=template,
        company_name="주식회사 테스트",
    )
    routed = company / "2.근로소득 수정신고" / "1.Input" / payroll.name
    assert routed.is_file()
    assert not payroll.exists()
    assert isinstance(result, BasicInfoReviewResult)
    assert result.payroll_files == (routed,)
    assert result.report_path.is_file()
    generate_status.assert_not_called()
    prepare_correction.assert_not_called()
    generate_submission.assert_not_called()


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
    preparation = BasicInfoPreparationResult((2023,), basic, 2, 1)
    correction_inputs = CorrectionInputResult((2023,), (), (), (), ())
    submission = _submission_result(company)

    with (
        patch("payroll_amendment.full_pipeline.PayrollConverter.convert_company_input") as convert,
        patch("payroll_amendment.full_pipeline.populate_basic_info_employee_sheets") as populate,
        patch(
            "payroll_amendment.full_pipeline.prepare_company_basic_info_for_review",
            return_value=preparation,
        ),
        patch("payroll_amendment.full_pipeline.generate_company_payroll_status", return_value=status) as generate_status,
        patch("payroll_amendment.full_pipeline.prepare_correction_inputs", return_value=correction_inputs),
        patch(
            "payroll_amendment.full_pipeline.generate_company_submission_target_lists",
            return_value=submission,
        ),
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
    assert isinstance(result, BasicInfoReviewResult)
    assert result.basic_info_path == basic
    assert result.payroll_files == (routed,)
    generate_status.assert_not_called()


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
    preparation = BasicInfoPreparationResult((2023,), basic, 2, 1)
    correction_inputs = CorrectionInputResult((2023,), (), (), (), ())
    submission = _submission_result(company)

    with (
        patch("payroll_amendment.full_pipeline.PayrollConverter.convert_company_input") as convert,
        patch("payroll_amendment.full_pipeline.populate_basic_info_employee_sheets"),
        patch(
            "payroll_amendment.full_pipeline.prepare_company_basic_info_for_review",
            return_value=preparation,
        ),
        patch("payroll_amendment.full_pipeline.generate_company_payroll_status", return_value=status) as generate_status,
        patch("payroll_amendment.full_pipeline.prepare_correction_inputs", return_value=correction_inputs),
        patch(
            "payroll_amendment.full_pipeline.generate_company_submission_target_lists",
            return_value=submission,
        ),
    ):
        result = run_full_pipeline(
            project_root=project,
            existing_company_directory=company,
            start_stage=3,
        )

    convert.assert_not_called()
    assert isinstance(result, BasicInfoReviewResult)
    assert result.payroll_files == (payroll,)
    generate_status.assert_not_called()


def test_full_pipeline_resumes_from_reviewed_basic_info(tmp_path):
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
    intake_report = company / "1. 급상여현황 생성" / "input_접수보고서.json"
    intake_report.parent.mkdir(parents=True)
    intake_report.write_text(
        '{"company_name":"주식회사 테스트","company_id":"123-45-67890","moved_files":[]}',
        encoding="utf-8",
    )
    review_report = intake_report.parent / "기초정보_검토대기.json"
    review_report.write_text('{"status":"AWAITING_HUMAN_REVIEW"}', encoding="utf-8")
    basic = company / "주식회사 테스트_기초정보.xlsx"
    payroll = company / "2.근로소득 수정신고" / "1.Input" / "주식회사 테스트_23년 급여대장.xlsx"
    conversion_report = company / "trans_검증보고서.json"
    for path in (basic, payroll, conversion_report):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"existing")

    status_output = company / "1. 급상여현황 생성" / "급상여현황" / "주식회사 테스트_23년 급상여현황.xlsx"
    status_report = status_output.parent / "급상여현황_검토보고서.json"
    status = PayrollStatusResult((2023,), (status_output,), status_report, basic)
    correction_inputs = CorrectionInputResult((2023,), (), (), (), ())
    submission = _submission_result(company)

    with (
        patch("payroll_amendment.full_pipeline.PayrollConverter.convert_company_input") as convert,
        patch("payroll_amendment.full_pipeline.populate_basic_info_employee_sheets") as populate,
        patch("payroll_amendment.full_pipeline.prepare_company_basic_info_for_review") as prepare_review,
        patch(
            "payroll_amendment.full_pipeline.generate_company_payroll_status",
            return_value=status,
        ) as generate_status,
        patch(
            "payroll_amendment.full_pipeline.prepare_correction_inputs",
            return_value=correction_inputs,
        ),
        patch(
            "payroll_amendment.full_pipeline.generate_company_submission_target_lists",
            return_value=submission,
        ),
    ):
        result = run_full_pipeline(
            project_root=project,
            existing_company_directory=company,
            start_stage=5,
            review_completed=True,
        )

    convert.assert_not_called()
    populate.assert_not_called()
    prepare_review.assert_not_called()
    generate_status.assert_called_once_with(
        company,
        resources,
        use_reviewed_item_markers=True,
    )
    assert result.payroll_files == (payroll,)
    assert result.report_path.is_file()
    assert json.loads(review_report.read_text(encoding="utf-8-sig"))["status"] == "HUMAN_REVIEW_COMPLETED"
