@echo off
rem Add an open-password to every Excel file in a folder (copies; originals untouched)
setlocal
chcp 932 >nul
cd /d "%~dp0"
if "%~1"=="" (
  set /p FD="Folder: "
) else (
  set "FD=%~1"
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Excelにパスワード一括.ps1" -Folder "%FD%"
pause
