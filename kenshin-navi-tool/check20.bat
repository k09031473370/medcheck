@echo off
rem Find which imported value breaks auto judgement
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check20.ps1"
