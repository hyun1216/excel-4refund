# 근로소득 수정신고 로컬 자동화

근로소득 수정신고에 필요한 원본 자료를 업체별로 접수하고, 기초정보·연도별 급여대장·급상여현황·수정신고 입력자료를 생성하는 Windows용 Python 프로젝트입니다.

현재 운영 중심 흐름은 `scripts/run_full_pipeline.py`입니다. Google Drive 업로드, 이메일 발송, 수정 원천징수영수증의 최종 PDF 생성, 추가 제출서류와 최종 제출 ZIP은 이 실행 흐름에 아직 완전히 연결되어 있지 않습니다.

## 1. 개발 환경과 설치

- Windows
- Python 3.11 이상
- Microsoft Excel
  - 월별 급여대장을 `.xls`로 변환하는 단계에 필요합니다.
  - 원천징수 Tool에 PDF 이미지를 넣는 단계에 필요합니다.
- Python 패키지: `openpyxl`, `PyMuPDF`, `Pillow`, `pdfplumber`, `pydantic`, `PyYAML`, `pypdf`, `typer`
- 테스트 실행 시 `pytest`

```powershell
git clone https://github.com/hyun1216/excel-4refund.git
cd excel-4refund
py -m pip install -e ".[test]"
python -m pytest -q
```

`run_app.bat`과 `validate_inputs.bat`은 아래의 운영 중심 Excel 파이프라인이 아니라 별도 S10~S70 워크스페이스 파이프라인을 실행합니다. 두 흐름의 차이는 [8. 별도 S10~S70 파이프라인](#8-별도-s10s70-파이프라인)을 참고하십시오.

## 2. 프로젝트 구성

| 경로 | 역할 |
|---|---|
| `src/payroll_amendment` | 입력 접수, 변환, 대상 판정, 산출물 생성 등 실제 구현 |
| `scripts/run_full_pipeline.py` | 운영 중심 전체 실행 및 기존 업체 폴더 재개 |
| `scripts/process_input.py` | 공용 `input` 접수만 실행 |
| `scripts/run_payroll_status.py` | 기초정보와 급여대장으로 급상여현황만 재생성 |
| `scripts/run_correction_inputs.py` | 수정 급상여현황, 원천징수 Tool, 월별 급여대장만 재생성 |
| `Resources` | 기초정보·Tool 템플릿, 지급항목 매칭표, 직업분류 연계표 |
| `input` | 공용 원본 접수함. 저장소에는 빈 폴더만 유지 |
| `config` | S10~S70 파이프라인 기본 설정 |
| `templates` | S10~S70 후속 문서 및 변환 템플릿 |
| `docs` | 설계, 갭 분석, 운영 참고 문서 |
| `tests` | 자동화 테스트 |

실제 업체 자료와 실행 산출물이 들어가는 `노무_세무 수정신고`, `logs`, `workspace`, `output`, `outputs`, `temp`, `tmp`는 Git에 올리지 않습니다. 계정·메일 정보가 포함될 수 있는 `Resources/Config.xlsx`도 로컬에서만 관리하며, 현재 `scripts/run_full_pipeline.py`는 이 파일을 읽지 않습니다.

## 3. 운영 중심 전체 흐름

```text
공용 input
  → 업체명 판별 및 업체 작업 폴더 생성
  → 원천징수 PDF 분할·추출
  → 기초정보와 연도별 급여대장 생성
  → 직업분류코드·건보정보 적재
  → 급상여현황 생성 및 대상자 판정
  → 수정 급상여현황·원천징수 Tool·월별 급여대장 생성
  → 전체실행_결과.json 저장
```

전체 실행:

```powershell
python scripts/run_full_pipeline.py
```

기본 입력 위치 대신 다른 접수함을 사용하려면:

```powershell
python scripts/run_full_pipeline.py --input-dir "D:\업체자료\접수"
```

실행 로그는 `logs/full_pipeline_YYYYMMDD_HHMMSS.log`에 저장됩니다. 성공하면 콘솔에 주요 결과 경로가 JSON으로 출력됩니다.

### 전체 실행의 7개 파트

| 파트 | 처리 내용 | 주요 결과 |
|---:|---|---|
| 1/7 | 공용 입력 접수 또는 기존 접수 결과 로드 | 업체 폴더, `input_접수보고서.json` |
| 2/7 | 원천징수 분할·추출, 기초정보·급여대장 변환 | `{업체명}_기초정보.xlsx`, 연도별 급여대장, 변환 보고서 |
| 3/7 | 직업분류코드와 건강보험 정보 적재 | 갱신된 기초정보 파일 |
| 4/7 | 연도별 급여대장을 수정신고 입력 폴더로 이동 | `2.근로소득 수정신고/1.Input` |
| 5/7 | 지급항목 분류, 대상자 판정, 급상여현황 생성 | `1. 급상여현황 생성/급상여현황` |
| 6/7 | 수정신고용 입력자료 생성 | `Input_급상여현황(수정)`, Tool, 월별 `.xls` |
| 7/7 | 전체 결과 요약 저장 | `1. 급상여현황 생성/전체실행_결과.json` |

## 4. 입력 자료와 업체 폴더

공용 `input`에는 한 업체의 자료만 넣고 실행합니다. 실행 중 다음 기준으로 파일을 분류합니다.

- 원천징수영수증 PDF
- 월별·연도별 급여대장 PDF 또는 Excel
- 소득자별 근로소득원천징수부 PDF
- 근로자 고용정보 현황 또는 직업분류 매핑 결과 Excel
- 건강보험 사업장가입자명부 Excel

정상 접수 후 원본 급여·원천징수 자료는 다음 구조로 이동합니다.

```text
노무_세무 수정신고/
└─ {업체명}/
   ├─ 1. 급상여현황 생성/
   │  ├─ 1.Input/
   │  │  ├─ 급여대장/
   │  │  ├─ 원천징수영수증/
   │  │  └─ 소득자별근로소득원천징수부/
   │  ├─ 급상여현황/
   │  ├─ 원천징수(PDF)/
   │  ├─ input_접수보고서.json
   │  └─ 전체실행_결과.json
   ├─ {업체명}_기초정보.xlsx
   ├─ Input_급상여현황(수정)/
   └─ 2.근로소득 수정신고/
      └─ 1.Input/
```

입력 접수는 파일을 복사하지 않고 업체 폴더로 이동합니다. 목적지에 같은 파일이 이미 있으면 전체 이동을 시작하지 않으므로, 충돌 파일을 확인한 뒤 다시 실행해야 합니다.

원천징수 PDF는 법정 서식 제목을 기준으로 개인본과 통합본을 판별합니다. 통합본은 개인별 PDF로 분할하고, 표와 주요 값을 직접 추출합니다. 표식 누락, 개인본·통합본 혼재, 인덱스·성명·주민번호 불일치 또는 생성 Excel 검증 실패가 발생하면 관련 파일과 `failure.json`을 격리 폴더에 남기고 해당 처리를 중단합니다.

## 5. 파트별 재수행 방법

재수행 전에는 Excel에서 해당 업체의 결과 파일을 모두 닫으십시오. 열린 파일은 Windows에서 교체하거나 변환하지 못할 수 있습니다. 중요한 수동 수정본은 별도 백업한 뒤 실행하십시오.

### 5.1 입력 접수만 수행

```powershell
python scripts/process_input.py
```

다른 접수함을 사용하려면:

```powershell
python scripts/process_input.py --input-dir "D:\업체자료\접수"
```

- 공용 입력을 업체 폴더 구조로 분류·이동하고 `input_접수보고서.json`을 만듭니다.
- 동일 목적지 파일이 있으면 기존 파일을 덮어쓰지 않고 접수를 중단합니다.
- 이미 접수가 끝난 업체의 후속 처리는 이 명령이 아니라 다음의 `--resume-company-dir`을 사용합니다.

### 5.2 접수 이후 전체 파트 재수행

```powershell
python scripts/run_full_pipeline.py `
  --resume-company-dir "D:\RPA\Python_4refund\노무_세무 수정신고\{업체명}"
```

- 1/7의 공용 `input` 이동만 생략합니다.
- 기존 `input_접수보고서.json`의 업체 정보와 이동 파일 목록을 검증합니다.
- 2/7부터 7/7까지는 모두 다시 실행합니다. 실패 지점부터 이어가는 체크포인트 방식은 아닙니다.
- 기초정보, 급여대장, 급상여현황, 수정 급상여현황, Tool, 월별 급여대장과 결과 보고서는 같은 경로의 결과를 다시 생성하거나 교체할 수 있습니다.
- 원본 입력을 추가·교체했다면 `input_접수보고서.json`의 목록과 실제 파일 구성이 일치하는지 먼저 확인하십시오.

전체 실행이 중간에 실패했을 때의 기본 복구 명령도 위 명령입니다. `logs`의 최신 로그와 각 검토·실패 JSON을 먼저 확인하고 원인을 수정한 뒤 재수행합니다.

### 5.3 급상여현황만 재수행

```powershell
python scripts/run_payroll_status.py `
  --company-dir "D:\RPA\Python_4refund\노무_세무 수정신고\{업체명}"
```

동일 기능의 설치형 CLI 명령:

```powershell
payroll-amendment payroll-status `
  --company-dir "D:\RPA\Python_4refund\노무_세무 수정신고\{업체명}" `
  --resources-dir "D:\RPA\Python_4refund\Resources"
```

재수행 범위:

- 업체 루트의 `{업체명}_기초정보.xlsx`를 찾습니다.
- `2.근로소득 수정신고/1.Input`의 연도별 급여대장을 읽습니다.
- `Resources/지급항목_매칭표.xlsx`로 항목을 다시 분류합니다.
- 기초정보의 대상 연도·항목 표시를 갱신합니다.
- `1. 급상여현황 생성/급상여현황`의 연도별 급상여현황과 `급상여현황_검토보고서.json`을 다시 생성합니다.

지급항목 매칭표, 기초정보, 급여대장 또는 대상 판정 로직이 바뀐 경우 사용합니다. 이 명령은 수정 급상여현황·Tool·월별 `.xls`까지 갱신하지 않으므로, 성공 후 5.4 명령도 이어서 실행해야 후속 결과가 일치합니다.

### 5.4 수정신고 입력자료만 재수행

```powershell
python scripts/run_correction_inputs.py `
  --company-dir "D:\RPA\Python_4refund\노무_세무 수정신고\{업체명}"
```

재수행 범위:

- 업체 루트의 기초정보 파일 1개를 사용합니다.
- `1. 급상여현황 생성/급상여현황`의 연도별 급상여현황을 읽습니다.
- `Input_급상여현황(수정)`에 연도별 수정 급상여현황과 수정 원천징수 Tool을 생성합니다.
- `2.근로소득 수정신고/1.Input`에 연도별 01~12월 급여대장 `.xls`를 생성합니다.
- 원천징수 PDF가 있으면 Tool 이미지 적재에도 사용합니다.

중간 오류가 발생하면 이번 실행에서 생성 목록에 포함된 수정 급상여현황, Tool, 월별 급여대장을 삭제하고 실패를 반환합니다. 원인을 수정한 뒤 같은 명령을 다시 실행하십시오.

### 5.5 재수행 순서 요약

| 변경 또는 장애 위치 | 실행할 명령 | 후속 실행 |
|---|---|---|
| 공용 `input` 접수 전 | `scripts/process_input.py` 또는 전체 실행 | 전체 실행 필요 |
| 원본 PDF·급여대장·기초정보 생성 단계 | `run_full_pipeline.py --resume-company-dir ...` | 명령 하나로 후속 전체 재생성 |
| 지급항목 매칭표·대상 판정·급상여현황 | `scripts/run_payroll_status.py` | `scripts/run_correction_inputs.py` 실행 |
| 수정 급상여현황·Tool·월별 `.xls` | `scripts/run_correction_inputs.py` | 없음 |
| 접수 보고서 또는 접수 원본 누락 | 공용 입력을 복원한 뒤 신규 전체 실행 검토 | 단순 재개 불가할 수 있음 |

## 6. 결과 확인과 장애 대응

우선 확인할 파일:

1. `logs/full_pipeline_*.log`
2. `1. 급상여현황 생성/input_접수보고서.json`
3. 변환 단계의 `trans_검증보고서.json`
4. `1. 급상여현황 생성/급상여현황/급상여현황_검토보고서.json`
5. `1. 급상여현황 생성/전체실행_결과.json`
6. 격리 폴더의 `failure.json`

`급상여현황_검토보고서.json`의 상태가 `NEEDS_REVIEW`이면 `issues`를 확인해야 합니다. 미연결 원천징수, 직업분류 미연결, 지급항목 미분류·충돌, 복수 사원번호 등은 결과를 사람이 검토해야 한다는 의미입니다.

현재 재개 기능의 한계:

- `--resume-company-dir`은 입력 접수만 건너뜁니다.
- 단계별·연도별 성공 체크포인트를 재사용하지 않습니다.
- 규칙 파일 하나가 바뀌어도 영향 단계만 자동 판별해 무효화하지 않습니다.
- 따라서 실패 원인을 해결한 뒤 안전한 공개 스크립트 단위로 다시 실행해야 합니다.

## 7. 테스트와 변경 검증

전체 테스트:

```powershell
python -m pytest -q
```

파트별 주요 테스트:

```powershell
python -m pytest -q tests/test_input_intake.py
python -m pytest -q tests/test_payroll_converter.py tests/test_withholding_pdf_excel.py
python -m pytest -q tests/test_employee_info.py tests/test_payroll_status.py
python -m pytest -q tests/test_correction_inputs.py
python -m pytest -q tests/test_run_full_pipeline_logging.py tests/test_full_pipeline.py
```

실제 운영 파일을 사용한 검증에서는 자동화 테스트 외에도 다음을 확인합니다.

- 연도별 직원 수와 대상자·비대상자 수
- 지급합계, 공제합계, 차인지급액 대사
- 기초정보의 직업분류코드·건보정보 미연결 건
- 급상여현황 검토 보고서의 `issues`
- 수정 급상여현황과 월별 급여대장의 합계
- Tool의 Data 시트와 원천징수 이미지 연결

## 8. 별도 S10~S70 파이프라인

패키지 CLI와 배치 파일에는 원본 스냅샷부터 로컬 제출 ZIP까지를 목표로 한 별도 워크스페이스 파이프라인이 있습니다.

```powershell
payroll-amendment validate --source "C:\입력폴더" --years "2024,2025"
payroll-amendment run --source "C:\입력폴더" --years "2024,2025"
payroll-amendment status RUN_ID
```

- 실행 결과: `workspace/runs/<RUN_ID>`
- `validate_inputs.bat`: 입력 검증 명령의 대화형 실행기
- `run_app.bat`: S10~S70 명령의 대화형 실행기
- `status`: 저장된 실행과 단계 상태 조회 전용이며 기존 RUN_ID를 이어서 실행하는 명령은 아닙니다.

S10~S70의 설계 단계는 입력 검증, 월별 급여대장, PDF 렌더링, 수정 원천징수, 추가서류, 제출 임금대장, ZIP 생성입니다. 다만 승인된 추가서류 템플릿과 검증 조건에 따라 후속 단계가 중단될 수 있으며, 운영 중심 `scripts/run_full_pipeline.py`와 결과 폴더 및 구현 완성도가 다릅니다. 인수인계 시 두 흐름을 하나의 완성된 파이프라인으로 오해하지 않도록 구분해야 합니다.

## 9. 개발자 인수인계 핵심 사항

- 운영 진입점은 우선 `scripts/run_full_pipeline.py`로 봅니다.
- 업체별 기초정보 파일은 차수별 복제본이 아니라 업체 루트의 단일 마스터입니다.
- 원본 입력은 `1. 급상여현황 생성/1.Input` 아래에서 보존하고 생성 결과를 원본 폴더에 섞지 않습니다.
- 생성 파일은 임시 파일에 먼저 저장하고 검증 후 교체하는 구현이 많지만, 재수행 시 같은 이름의 기존 결과가 교체될 수 있으므로 수동 수정본은 백업해야 합니다.
- 지급항목 분류 기준은 `Resources/지급항목_매칭표.xlsx`, 직업분류 연결 기준은 `Resources/직업분류_연계표.xlsx`입니다.
- 연도별 급여대장은 `2.근로소득 수정신고/1.Input`, 급상여현황은 `1. 급상여현황 생성/급상여현황`, 수정 결과는 `Input_급상여현황(수정)`에 생성됩니다.
- 현재 자동 재개는 접수 생략 수준입니다. 단계·연도별 체크포인트, 승인 이력, 종단 대사, 최종 제출물 연결은 후속 개발 과제입니다.
- 외부 전달 기능과 민감정보가 포함된 로컬 설정 파일은 저장소 범위 밖입니다.

상세 설계와 현황은 `docs/`의 갭 분석, 입출력 경로 보고서, 통합 개발 계획을 함께 참고하십시오. 문서와 코드가 다르면 현재 테스트를 통과하는 코드 동작을 기준으로 판단하고 관련 문서를 함께 갱신하십시오.
