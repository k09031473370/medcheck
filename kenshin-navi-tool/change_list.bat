@echo off
rem Ishiya: change list after 2026/10/18 re-judgment (read-only, CSV in out folder)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0change_list.ps1" -Since 20261018_1130
pause
