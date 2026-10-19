@echo off
rem Dock society judgment vs navi for Ishiya 3 days (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0dock_hantei.ps1" -Dates 2026/09/03,2026/09/09,2026/09/10 -UkeFrom 4631 -UkeTo 4831
pause
