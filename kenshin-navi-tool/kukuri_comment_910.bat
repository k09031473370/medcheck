@echo off
rem Ishiya 9/10 only (run AFTER auto-judgment): add HbA1c / eGFR / TG comments (D/F) to empty doctor-instruction slots
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kukuri_comment.ps1" -Dates 2026/09/10 -UkeFrom 4631 -UkeTo 4831
pause
