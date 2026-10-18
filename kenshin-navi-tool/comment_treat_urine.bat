@echo off
rem Ishiya: delete urine glucose work-up comments for people treated for diabetes (J), doctor decision (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0comment_treat.ps1" -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831 -UrineGlucose
pause
