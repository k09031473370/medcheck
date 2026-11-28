@echo off
rem OCR v5.3 Excel (88 columns): import postal code and address into T_KOJIN2 (preview, then y)
setlocal
cd /d "%~dp0"
if "%~1"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0jusho_v53.ps1"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0jusho_v53.ps1" -File "%~1"
)
pause
