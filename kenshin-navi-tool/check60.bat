@echo off
rem Would judging triglycerides (R06-0002) drop anyone from A in lipid/overall judgement? (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check60.ps1"
