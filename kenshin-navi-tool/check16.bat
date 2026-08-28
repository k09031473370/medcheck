@echo off
rem Check if KENSHIN navi can output the Form No.6 report
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check16.ps1"
