@echo off
rem Verify the judgement conversion took effect (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check45.ps1"
