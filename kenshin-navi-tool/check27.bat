@echo off
rem Verify the test person is back to the same state as the other 97
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check27.ps1"
