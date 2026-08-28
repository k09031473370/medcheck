@echo off
rem Add the union name to generated reports (see clinic_swap.ps1)
setlocal
chcp 932 >nul
set /p F=Report file or folder: 
set /p K=Union name: 
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0clinic_swap.ps1" -In "%F%" -Kumiai "%K%"
echo.
echo If the preview looks right, run again with -Commit to write.
pause
