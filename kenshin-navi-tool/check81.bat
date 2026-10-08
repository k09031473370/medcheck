@echo off
rem People still without a reception number on 9/3, 9/9, 9/10 (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check81.ps1"
