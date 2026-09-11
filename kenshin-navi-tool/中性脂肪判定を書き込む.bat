@echo off
rem Write triglyceride judgement (H mark / lipid group letter / reference range) into exported result-report Excel files
setlocal
chcp 932 >nul
cd /d "%~dp0"
if "%~1"=="" (
  set /p FD="Folder: "
) else (
  set "FD=%~1"
)
set /p FN="CSV: "
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0’†«‰–b”»’è‚ğ‘‚«‚Ş.ps1" -Folder "%FD%" -Csv "%FN%"
pause
