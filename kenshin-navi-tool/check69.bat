@echo off
rem Check whether the October bookings exist in Kenshin-Navi (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check69.ps1"
