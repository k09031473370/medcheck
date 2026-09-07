@echo off
rem Look for stray keystrokes landed in result fields (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check42.ps1"
