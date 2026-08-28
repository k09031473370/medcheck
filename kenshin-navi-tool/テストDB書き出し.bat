@echo off
rem Test DB export - see export_testdb.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0export_testdb.ps1" -Ask
echo.
pause
