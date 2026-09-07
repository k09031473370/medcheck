@echo off
rem Find items where only some people are missing a judgement (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check41.ps1"
