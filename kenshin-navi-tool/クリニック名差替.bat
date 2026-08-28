@echo off
rem Swap clinic name on a generated report (see clinic_swap.ps1)
setlocal
chcp 932 >nul
set /p F=Report file or folder: 
set /p C=Clinic key (blank = list): 
if "%C%"=="" (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0clinic_swap.ps1" -In "%F%"
) else (
  powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0clinic_swap.ps1" -In "%F%" -Clinic "%C%"
)
echo.
pause
