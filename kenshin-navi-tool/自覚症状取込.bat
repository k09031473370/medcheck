@echo off
rem Import subjective symptoms from the Toshinkyo CSV
setlocal
chcp 932 >nul
cd /d "%~dp0"
if "%~1"=="" (
  set /p FN="CSV: "
) else (
  set "FN=%~1"
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0é©äoè«èÛéÊçû.ps1" -Csv "%FN%"
pause
