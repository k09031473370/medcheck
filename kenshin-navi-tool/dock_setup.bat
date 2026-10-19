@echo off
rem Build dock judgment table from KenshinNavi reference-value settings (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0dock_setup.ps1"
pause
