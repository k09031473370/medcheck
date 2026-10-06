@echo off
rem Fussa: run the whole day-of import in one go (drag the reception app's JSON onto this file)
setlocal
chcp 932 >nul
cd /d "%~dp0"
if "%~1"=="" (
  set /p FN="JSON: "
) else (
  set "FN=%~1"
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ïüê∂éÊçû.ps1" -Json "%FN%"
pause
