@echo off
rem Validate the Fussa company map against T_DANTAI1 (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check80.ps1"
