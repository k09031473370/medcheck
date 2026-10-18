@echo off
rem Ishiya 9/9-9/10: import past medical history from the 123-column input backup (preview, then y)
setlocal
cd /d "%~dp0"
if "%~1"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kiou123.ps1"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kiou123.ps1" -File "%~1"
)
pause
