@echo off
rem Ishiya 9/9 only: reflect HbA1c / PLT / eGFR judgments into group and overall judgments (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kukuri_fix.ps1" -Dates 2026/09/09 -UkeFrom 4631 -UkeTo 4831
pause
