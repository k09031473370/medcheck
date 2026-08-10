@echo off
rem 自作の健診ビューア(読むだけ)。ブラウザが開きます
chcp 932 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kensin_web.ps1"
pause
