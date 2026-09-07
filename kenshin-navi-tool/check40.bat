@echo off
rem Find people whose automatic judgement was missed (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check40.ps1"
