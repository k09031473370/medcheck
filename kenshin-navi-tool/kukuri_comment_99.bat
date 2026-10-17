@echo off
rem Ishiya 9/9 only: add HbA1c / eGFR comments (D/F) to empty doctor-instruction slots (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kukuri_comment.ps1" -Dates 2026/09/09 -UkeFrom 4631 -UkeTo 4831
pause
