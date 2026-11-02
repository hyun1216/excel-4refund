@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"

echo 근로소득 수정신고 로컬 실행기
set /p "SOURCE=입력 ZIP 또는 폴더 경로: "
set /p "YEARS=수행연도 (예: 2024,2025): "
if not defined SOURCE goto :invalid
if not defined YEARS goto :invalid

py -m pip install -e .
if errorlevel 1 goto :failed
py -m payroll_amendment run --source "%SOURCE%" --years "%YEARS%" --workspace "%~dp0workspace"
if errorlevel 1 goto :failed

echo.
echo 완료되었습니다. workspace\runs 폴더를 확인하세요.
pause
exit /b 0

:invalid
echo 입력값이 비어 있습니다.
pause
exit /b 2

:failed
echo 실행 중 오류가 발생했습니다. workspace\runs 아래 run_report.json과 logs를 확인하세요.
pause
exit /b 1
