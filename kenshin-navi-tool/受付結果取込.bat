@echo off
rem Apply the reception app's navi.json to Kenshin-Navi (preview first, then write)
setlocal
chcp 932 >nul
cd /d "%~dp0"
if "%~1"=="" (
  set /p FN="JSON: "
) else (
  set "FN=%~1"
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ó•tŒ‹‰Êæ.ps1" -Json "%FN%"
pause
