@echo off
rem Make template 305 = copy of 301 with clinic text swapped. 301 is not changed
chcp 932 >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0template_copy.ps1" -Src 301 -No 305
pause
