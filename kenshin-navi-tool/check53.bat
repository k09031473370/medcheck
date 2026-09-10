@echo off
rem Inspect the report-name tables (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check53.ps1"
