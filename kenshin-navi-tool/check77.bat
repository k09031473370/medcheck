@echo off
rem Check what breaks when a visit date is moved (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check77.ps1"
