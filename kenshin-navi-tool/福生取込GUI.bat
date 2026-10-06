@echo off
rem Fussa day-of import, GUI version (double-click; pick or drag the reception app's JSON)
setlocal
cd /d "%~dp0"
start "" powershell -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0ïüê∂éÊçûGUI.ps1" %1
