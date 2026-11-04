from pathlib import Path
import tempfile
import unittest

from openpyxl import Workbook, load_workbook

from payroll_amendment.employee_info import (
    create_occupation_mapping_result,
    discover_employee_info_sources,
    populate_basic_info_employee_sheets,
)


class EmployeeInfoTests(unittest.TestCase):
    def test_prefers_and_loads_pre_mapped_employment_file(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            basic_path = root / "회사_기초정보.xlsx"
            raw_path = root / "근로자고용정보현황조회.xlsx"
            mapped_path = root / "수정본_직업분류매핑결과.xlsx"
            linkage_path = root / "직업분류_연계표.xlsx"

            basic = Workbook()
            occupation = basic.active
            occupation.title = "직업분류코드"
            occupation.append(["연번", "근로자구분", "주민등록번호", "성명"])
            occupation.append([None] * 16 + ["직종", "직업분류 코드", "직업분류 항목"])
            basic.create_sheet("건보정보")
            basic.save(basic_path)

            for path, mapped in ((raw_path, False), (mapped_path, True)):
                book = Workbook()
                sheet = book.active
                sheet.append(["연번", "근로자구분", "주민등록번호", "성명", "고용보험"])
                headers = [None, None, None, None, "취득일", "상실일", "휴직시작일", "휴직종료일", "전근일", "고용상태", "직종", "상실사유", "근로자원부번호", "사업개시번호", "미승인하수급인번호"]
                if mapped:
                    headers.extend(["표준직업분류코드", "표준직업분류명", "매핑상태", "매핑점수", "검토사유"])
                sheet.append(headers)
                values = [1, "상용", "900101-1234567", "가명", "2023-01-01", None, None, None, None, "재직", "미용·여행·숙박·음식·경비·청소직", None, "1", None, None]
                if mapped:
                    values.extend(["9999", "기타 서비스 관련 단순 종사자", "매핑완료", 1.0, None])
                sheet.append(values)
                book.save(path)

            linkage = Workbook()
            linkage_sheet = linkage.active
            linkage_sheet.title = "고용2025_표준직업_연계표"
            for column, value in ((3, "중분류명"), (5, "소분류명"), (7, "세분류명"), (13, "세분류2"), (14, "세분류명2")):
                linkage_sheet.cell(2, column, value)
            linkage.save(linkage_path)

            employment, _ = discover_employee_info_sources(root)
            self.assertEqual(mapped_path, employment)
            result = populate_basic_info_employee_sheets(
                basic_info_path=basic_path,
                employment_info_path=employment,
                health_roster_path=None,
                occupation_linkage_path=linkage_path,
            )

            self.assertEqual(1, result.mapped_rows)
            output = load_workbook(basic_path, read_only=True, data_only=True)
            try:
                self.assertEqual("9999", output["직업분류코드"]["R3"].value)
                self.assertEqual("기타 서비스 관련 단순 종사자", output["직업분류코드"]["S3"].value)
            finally:
                output.close()

    def test_loads_occupation_mapping_and_health_roster(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            basic_path = root / "회사_기초정보.xlsx"
            employment_path = root / "근로자고용정보현황조회_고용.xlsx"
            health_path = root / "사업장가입자명부.xlsx"
            linkage_path = root / "직업분류_연계표.xlsx"

            basic = Workbook()
            occupation = basic.active
            occupation.title = "직업분류코드"
            occupation.append(["연번", "근로자구분", "주민등록번호", "성명"])
            occupation.append([None] * 16 + ["직종", "직업분류 코드", "직업분류 항목"])
            health = basic.create_sheet("건보정보")
            health.append(["순번", "일련번호", "증번호", "주민등록번호", "성명", "자격취득일", "자격상실일"])
            basic.save(basic_path)

            employment = Workbook()
            source = employment.active
            source.append(["연번", "근로자구분", "주민등록번호", "성명", "고용보험"])
            source.append([None, None, None, None, "취득일", "상실일", "휴직시작일", "휴직종료일", "전근일", "고용상태", "직종", "상실사유", "근로자원부번호", "사업개시번호", "미승인하수급인번호"])
            source.append([1, "상용", "900101-1234567", "가명", "2023-01-01", None, None, None, None, "재직", "경리 사무원", None, "1", None, None])
            source.append([2, "상용", "900102-1234567", "가명2", "2023-02-01", None, None, None, None, "재직", "미등록직종", None, "2", None, None])
            employment.save(employment_path)

            roster = Workbook()
            roster_sheet = roster.active
            roster_sheet.append(["순번", "일련번호", "가입자", None, None, "자격취득일", "자격상실일"])
            roster_sheet.append([None, None, "증번호", "주민등록번호", "성명"])
            roster_sheet.append([1, "A1", "H1", "900101-1234567", "가명", "2023-01-01", None])
            roster.save(health_path)

            linkage = Workbook()
            linkage_sheet = linkage.active
            linkage_sheet.title = "고용2025_표준직업_연계표"
            linkage_sheet.cell(2, 3, "중분류명")
            linkage_sheet.cell(2, 5, "소분류명")
            linkage_sheet.cell(2, 7, "세분류명")
            linkage_sheet.cell(2, 13, "세분류2")
            linkage_sheet.cell(2, 14, "세분류명2")
            linkage_sheet.cell(3, 3, "경영·행정·사무직")
            linkage_sheet.cell(3, 5, "회계 사무원")
            linkage_sheet.cell(3, 7, "경리 사무원")
            linkage_sheet.cell(3, 13, "3125")
            linkage_sheet.cell(3, 14, "경리 사무원")
            linkage.save(linkage_path)

            result = populate_basic_info_employee_sheets(
                basic_info_path=basic_path,
                employment_info_path=employment_path,
                health_roster_path=health_path,
                occupation_linkage_path=linkage_path,
            )

            self.assertEqual(2, result.occupation_rows)
            self.assertEqual(1, result.mapped_rows)
            self.assertEqual(1, result.not_eligible_rows)
            self.assertEqual(1, result.health_rows)
            self.assertEqual(
                "근로자고용정보현황조회_고용_직업분류매핑결과.xlsx",
                result.employment_source,
            )
            self.assertTrue(
                (root / "근로자고용정보현황조회_고용_직업분류매핑결과.xlsx").is_file()
            )

            output = load_workbook(basic_path, read_only=True, data_only=True)
            try:
                occupation = output["직업분류코드"]
                self.assertEqual("재직", occupation["K3"].value)
                self.assertEqual("경리 사무원", occupation["Q3"].value)
                self.assertEqual("3125", occupation["R3"].value)
                self.assertEqual("경리 사무원", occupation["S3"].value)
                self.assertEqual("경정청구대상아님", occupation["R4"].value)
                health = output["건보정보"]
                self.assertEqual("900101-1234567", health["D2"].value)
                self.assertEqual("가명", health["E2"].value)
            finally:
                output.close()

    def test_creates_mapping_result_with_preferred_and_middle_classification(self) -> None:
        with tempfile.TemporaryDirectory() as folder:
            root = Path(folder)
            employment_path = root / "근로자고용정보현황조회.xlsx"
            linkage_path = root / "직업분류_연계표.xlsx"

            employment = Workbook()
            source = employment.active
            source.append(["연번", "근로자구분", "주민등록번호", "성명", "고용보험"])
            source.append([None, None, None, None, "취득일", "상실일", "휴직시작일", "휴직종료일", "전근일", "고용상태", "직종", "상실사유", "근로자원부번호", "사업개시번호", "미승인하수급인번호"])
            source.append([1, "상용", "900101-1234567", "가명", None, None, None, None, None, "재직", "미용·여행·숙박·음식·경비·청소직"])
            source.append([2, "상용", "900102-1234567", "가명2", None, None, None, None, None, "재직", "경영·행정·사무직"])
            employment.save(employment_path)

            linkage = Workbook()
            sheet = linkage.active
            sheet.title = "고용2025_표준직업_연계표"
            for column, value in ((3, "중분류명"), (5, "소분류명"), (7, "세분류명"), (13, "세분류2"), (14, "세분류명2")):
                sheet.cell(2, column, value)
            sheet.cell(3, 3, "경영·행정·사무직")
            sheet.cell(3, 5, "회계 사무원")
            sheet.cell(3, 7, "경리 사무원")
            sheet.cell(3, 13, "3125")
            sheet.cell(3, 14, "경리 사무원")
            linkage.save(linkage_path)

            result = create_occupation_mapping_result(
                source_path=employment_path,
                linkage_path=linkage_path,
            )

            self.assertTrue(employment_path.is_file())
            self.assertEqual(2, result.processed_rows)
            self.assertEqual(2, result.mapped_rows)
            output = load_workbook(result.output_path, read_only=True, data_only=True)
            try:
                mapped = output.active
                self.assertEqual("9999", mapped["P3"].value)
                self.assertEqual("사전 지정 매핑", mapped["T3"].value)
                self.assertEqual("3125", mapped["P4"].value)
                self.assertEqual("중분류명 정확히 일치", mapped["T4"].value)
            finally:
                output.close()


if __name__ == "__main__":
    unittest.main()
