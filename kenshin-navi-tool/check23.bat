@echo off
rem Scan judgement criteria rows for malformed strings
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check23.ps1"
