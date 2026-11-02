# 근로소득 수정신고 로컬 자동화

Google Drive 업로드와 이메일 발송을 제외한 로컬 처리 코어입니다.

## 설치

```powershell
git clone https://github.com/hyun1216/excel-4refund.git
cd excel-4refund
py -m pip install -e ".[test]"
```

Python 3.11 이상과 Microsoft Excel이 필요합니다. Tool 이미지 삽입 단계는
Windows의 Microsoft Excel 및 `pywin32`가 필요할 수 있습니다.

설치 확인:

```powershell
python -m pytest -q
```

## 프로젝트 구성

- `src/payroll_amendment`: 실제 실행 코드
- `scripts`: 전체 실행 및 단계별 재실행 스크립트
- `Resources`: 실행 템플릿, 항목 매칭표 및 직업분류 연계표
- `input`: 공용 입력 위치(저장소에는 빈 폴더만 유지)
- `config`: 기본 설정
- `docs`: 설계·운영 문서
- `templates`: 후속 문서 및 변환 템플릿
- `tests`: 자동화 테스트

`노무_세무 수정신고`, `logs`, `output`, `outputs`, `temp`, `tmp` 등 실제 업체
자료와 실행 산출물은 Git에 업로드하지 않습니다.
메일 주소와 환경별 계정값이 들어 있는 `Resources/Config.xlsx`도 보안상 로컬에서만
관리하며 저장소에는 업로드하지 않습니다. 현재 Python 실행 코드는 이 파일을 사용하지
않습니다.

## 전체 실행

공용 `input` 자료부터 기초정보, 급여대장, 급상여현황, 수정 급상여현황 및
원천징수 Tool까지 실행합니다.

```powershell
python scripts/run_full_pipeline.py
```

이미 급상여현황까지 생성된 회사 폴더에서 수정 급상여현황과 Tool만 생성하려면:

```powershell
python scripts/run_correction_inputs.py --company-dir "C:\path\to\회사폴더"
```

## 입력 폴더

한 업체의 입력을 한 폴더에 둡니다. 하위 폴더명은 자유롭지만 파일은 다음 조건을 만족해야 합니다.

- 기초정보 XLSX: `사업장정보`, `수당등록`, `수당조정(비과세)` 시트 포함
- 연도별 수정 급상여현황 XLSX: `대상자(작업)`, `비대상자` 시트 포함
- 연도별 원천징수 생성 Tool XLSX: `Data` 시트 포함, 파일명에 `25년`과 같은 연도 포함
- 원천징수 PDF: 개인별 PDF 또는 전체 사원 통합 PDF, 파일명 또는 상위 폴더에 연도 포함

원천징수 PDF는 `소득세법시행규칙[별지제24호서식(1)]` 제목이 있는 페이지를 개인별
첫 페이지로 판정합니다. 제목이 한 번이면 개인본, 두 번 이상이면 통합본으로 처리합니다.
통합본은 각 제목 페이지부터 다음 제목 직전까지 분할하며 결과 파일은 기존 Brity와 같은
`인덱스_성명_발행자보관용.pdf` 형식으로 저장합니다. 표식 누락,
통합본·개인본 혼재, 대상 인덱스 누락, 주민번호 불일치가 발생하면 원본을 변경하지 않고
실행 폴더의 `quarantine/withholding/<연도>`에 복사한 뒤 해당 연도 처리를 중단합니다.

분할된 개인별 PDF는 알PDF 같은 외부 프로그램을 실행하지 않고 Python에서 표를 직접
추출합니다. 결과는 같은 파일명의 `.xlsx`로
`prepared/withholding_excel/<연도>`에 저장되며 다음 시트를 포함합니다.

- `Table 1`: 지춘화 기준 Excel의 셀·병합·서식을 재사용하고 PDF 텍스트와 금액을 편집 가능한 셀 값으로 배치
- `추출값`: 11번 기초자료 적재에서 사용할 성명·주민번호·과세/비과세·세액·공제 항목
- `원본표`: PDF 페이지별 표를 셀 단위로 보존한 검토 자료
- `검증`: 원본 파일명·SHA-256·페이지 수·추출 필드 수

필수 표나 항목 누락, 파일 인덱스·성명·주민번호 불일치, 생성 Excel 재개방 검증 실패는
`quarantine/withholding_excel/<연도>`에 해당 PDF와 `failure.json`을 남기고 처리를 중단합니다.

원본은 실행별 `workspace/runs/<RUN_ID>/incoming`에 해시 검증 후 복사되며 변경하지 않습니다.

## 실행

```powershell
payroll-amendment validate --source "C:\입력폴더" --years "2024,2025"
payroll-amendment run --source "C:\입력폴더" --years "2024,2025"
payroll-amendment status RUN_ID
```

ZIP 입력도 지원합니다. 실행 결과는 `workspace/runs/<RUN_ID>` 아래에 저장됩니다.

## 단계

- S10 입력 검증, 원천징수 PDF 분할·Excel 직접 추출 및 스냅샷
- S20 월별 급여대장 생성
- S30 원천징수 PDF 전 페이지 GIF 변환
- S40 수정 원천징수영수증 생성
- S50 추가서류 생성
- S60 제출 임금대장 생성
- S70 검증된 로컬 제출 ZIP 생성

S50은 `templates/supporting_docs/document_specs.json`에 승인된 빈 양식이 등록되어야 실행됩니다. 검토가 필요한 결과가 있으면 S70은 제출 ZIP을 만들지 않습니다.

## 외부 전달

Google Drive 및 이메일 기능은 이번 범위에서 구현하지 않았습니다.
