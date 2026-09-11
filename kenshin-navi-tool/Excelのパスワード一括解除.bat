@echo off
rem Remove the open-password from every Excel file in this folder (copies into "解除済み")
setlocal
chcp 932 >nul
cd /d "%~dp0"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Excelのパスワード一括解除.ps1" -Folder "%~dp0."
