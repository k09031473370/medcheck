@echo off
rem Ishiya 3 days: restore T_KENSA to the saved snapshot (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ishiya_snapshot.ps1" -Restore
pause
