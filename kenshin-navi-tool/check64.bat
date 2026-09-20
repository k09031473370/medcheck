@echo off
rem Final check before copying the Toshinkyo courses to the new client (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check64.ps1"
