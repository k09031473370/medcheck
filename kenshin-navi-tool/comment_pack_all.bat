@echo off
rem Ishiya: close blank gaps in doctor comments (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0comment_pack.ps1" -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
pause
