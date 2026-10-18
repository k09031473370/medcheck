@echo off
rem Ishiya 9/3: doctor decision - chest X-ray F to D for 4638,4655,4670,4683 (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0xray_d.ps1" -Dates 2026/09/03 -UkeNo 4638,4655,4670,4683
pause
