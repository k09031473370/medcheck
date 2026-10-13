@echo off
rem Check that everyone on 9/3, 9/9, 9/10 has blood and form values (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check84.ps1"
