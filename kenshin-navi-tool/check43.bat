@echo off
rem Check whether C-or-worse findings reach the overall judgement (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check43.ps1"
