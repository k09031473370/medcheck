@echo off
rem Read the set definition tables to know what rows an option creates (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check75.ps1"
