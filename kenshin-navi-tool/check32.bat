@echo off
rem Find out what the reception step writes (run twice)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check32.ps1"
