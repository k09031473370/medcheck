@echo off
rem Compare exported result-report Excel files with the Toshinkyo CSV (read-only)
setlocal
chcp 932 >nul
cd /d "%~dp0"
if "%~1"=="" (
  set /p FD="Folder: "
) else (
  set "FD=%~1"
)
set /p FN="CSV: "
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0åãâ Excelè∆çá.ps1" -Folder "%FD%" -Csv "%FN%"
pause
