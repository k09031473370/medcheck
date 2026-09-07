@echo off
rem Compare reception numbers across mobile-clinic days (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check39.ps1"
