@echo off
rem Ishiya 9/3: save comments and all judgments BEFORE auto-judgment (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0comment_keep.ps1" -Save -Dates 2026/09/03 -UkeFrom 4631 -UkeTo 4831
pause
