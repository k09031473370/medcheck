@echo off
rem Build the Kenshin-Navi booking-import Excel from the reception app's JSON (read-only on the DB)
setlocal
chcp 932 >nul
cd /d "%~dp0"
if "%~1"=="" (
  set /p FN="JSON: "
) else (
  set "FN=%~1"
)
set /p YM="Date (e.g. 2026/10/02, blank = all): "
if "%YM%"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0—\–ñæExcelì¬.ps1" -Json "%FN%"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0—\–ñæExcelì¬.ps1" -Json "%FN%" -Ymd "%YM%"
)
pause
