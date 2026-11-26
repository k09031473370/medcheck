@echo off
rem OCR v5 Excel (80 columns): import past medical history into T_KIOU (preview, then y)
setlocal
cd /d "%~dp0"
if "%~1"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kiou_v5.ps1"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kiou_v5.ps1" -File "%~1"
)
pause
