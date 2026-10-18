@echo off
rem Ishiya: verify external review points (read-only, CSV in out folder)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check113.ps1"
pause
