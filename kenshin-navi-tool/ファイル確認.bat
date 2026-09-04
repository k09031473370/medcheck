@echo off
setlocal
rem Check an import file against the mapping. Read only - no database.
cd /d "%~dp0"
if "%~1"=="" (
  echo.
  set /p FN="check file: "
) else (
  set "FN=%~1"
)
if "%FN%"=="" goto :end
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ファイル確認.ps1" -Csv "%FN%"
:end
echo.
pause
