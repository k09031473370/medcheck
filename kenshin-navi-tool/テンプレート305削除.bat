@echo off
rem Remove template 305 added by template_copy (keeps a copy in out)
chcp 932 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0template_copy.ps1" -Undo 305
pause
