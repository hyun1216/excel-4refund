@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set /p "SOURCE=입력 ZIP 또는 폴더 경로: "
set /p "YEARS=수행연도 (예: 2024,2025): "
py -m pip install -e .
if errorlevel 1 exit /b 1
py -m payroll_amendment validate --source "%SOURCE%" --years "%YEARS%" --workspace "%~dp0workspace"
pause
