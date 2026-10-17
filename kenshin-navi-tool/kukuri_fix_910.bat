@echo off
rem Ishiya 9/10 only (run AFTER auto-judgment): reflect HbA1c / PLT / eGFR / TG judgments into group and overall judgments
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kukuri_fix.ps1" -Dates 2026/09/10 -UkeFrom 4631 -UkeTo 4831
pause
