@echo off
rem Ishiya 3 days: save all T_KENSA rows now (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ishiya_snapshot.ps1" -Save
pause
