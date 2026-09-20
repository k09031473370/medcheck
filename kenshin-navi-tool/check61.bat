@echo off
rem Look up how the existing courses are set up, to build a new one (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check61.ps1"
