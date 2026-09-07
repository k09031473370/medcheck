@echo off
rem Check the reception-number state for a visit date (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check38.ps1"
