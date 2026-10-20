@echo off
rem Replace fixed text on exported result sheets (writes a copy with suffix). Template untouched
chcp 932 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0text_swap.ps1"
pause
