@echo off
rem Ishiya 3 days: medication answers (yes) -> past history hypertension/diabetes/dyslipidemia under drug treatment (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kiou_med.ps1" -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
pause
