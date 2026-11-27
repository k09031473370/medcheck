@echo off
rem OCR v5 Excel: import front side (form_import) and medical/family history (kiou_v5) in one go. Preview first, then y.
setlocal
cd /d "%~dp0"
if "%~1"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0jushinhyo_all.ps1"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0jushinhyo_all.ps1" -File "%~1"
)
pause
