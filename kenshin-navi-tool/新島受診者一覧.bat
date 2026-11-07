@echo off
rem Niijima: write scheduled examinees (Nov) to xlsx in out folder. Read-only DB
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0niijima_list.ps1"
pause
