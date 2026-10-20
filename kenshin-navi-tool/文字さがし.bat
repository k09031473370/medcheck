@echo off
rem Find where a word (default: seikyo) is in xlsx files. Read-only
chcp 932 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0word_find.ps1"
pause
