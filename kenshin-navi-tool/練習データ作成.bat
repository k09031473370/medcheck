@echo off
rem Make practice import data from the DB (read-only). See make_testdata.ps1
setlocal
chcp 932 >nul
echo ================================================
echo  Practice data generator (DB is read-only)
echo ================================================
echo.
set "Y="
set /p Y=Exam date (blank = list dates) : 
if "%Y%"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_testdata.ps1" -List
  goto :done
)
set "N="
set /p N=Uketsuke No, comma separated (blank = all) : 
if "%N%"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_testdata.ps1" -Ymd "%Y%"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0make_testdata.ps1" -Ymd "%Y%" -KenNo "%N%"
)
:done
echo.
pause
