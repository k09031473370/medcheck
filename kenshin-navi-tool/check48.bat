@echo off
rem Find where past-history and symptoms are stored (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check48.ps1"
