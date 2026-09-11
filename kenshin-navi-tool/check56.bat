@echo off
rem Find where the eyesight judgement criteria are configured (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check56.ps1"
