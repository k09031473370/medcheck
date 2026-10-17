@echo off
rem Undo a kukuri_fix run (restores the previous judgment letters)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kukuri_fix.ps1" -Undo ask
pause
