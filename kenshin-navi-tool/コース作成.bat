@echo off
rem Create new courses for a client by copying template courses (writes to the DB after a preview)
setlocal
chcp 932 >nul
cd /d "%~dp0"
if "%~1"=="" (
  set /p DT="Dantai CD: "
) else (
  set "DT=%~1"
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ÉRÅ[ÉXçÏê¨.ps1" -Dantai "%DT%"
pause
