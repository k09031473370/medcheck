@echo off
rem Copy result-report Excel files with katakana names in the file name
setlocal
chcp 932 >nul
cd /d "%~dp0"
if "%~1"=="" (
  set /p FD="Folder: "
) else (
  set "FD=%~1"
)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0結果ファイル名をカナに.ps1" -Folder "%FD%"
pause
