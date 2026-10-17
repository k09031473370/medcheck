@echo off
rem Ishiya 9/3, 9/9, 9/10: add HbA1c / eGFR comments (D/F) to empty doctor-instruction slots (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kukuri_comment.ps1" -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
pause
