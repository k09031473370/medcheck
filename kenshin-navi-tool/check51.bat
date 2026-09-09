@echo off
rem Why blood sugar / triglyceride are missing on the report (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check51.ps1"
