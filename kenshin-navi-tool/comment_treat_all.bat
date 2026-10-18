@echo off
rem Ishiya: comments for J (treated) groups -> continue treatment text (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0comment_treat.ps1" -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
pause
