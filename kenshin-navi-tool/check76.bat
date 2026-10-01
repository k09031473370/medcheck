@echo off
rem Collect the child items of each option from real data (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check76.ps1"
