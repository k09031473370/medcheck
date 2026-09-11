@echo off
rem Dump the clinic reference ranges for comparison with the academy table (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check59.ps1"
