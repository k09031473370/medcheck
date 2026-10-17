@echo off
rem Ishiya 9/3, 9/9, 9/10: reflect HbA1c / PLT / eGFR judgments into group and overall judgments (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kukuri_fix.ps1" -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
pause
