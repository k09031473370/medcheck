@echo off
rem Ishiya: compare kyokai (610xxx) judgments with group judgments (600xxx) (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check105.ps1"
