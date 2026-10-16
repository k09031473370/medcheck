@echo off
rem Does 001282 (fasting flag) affect glucose / TG judgment? (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check95.ps1"
