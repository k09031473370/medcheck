@echo off
rem List the people imported for Fussa 10/2 so we can find who is missing (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check72.ps1"
