@echo off
rem Compare tool-written T_KENSA rows against real rows, all columns
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check24.ps1"
