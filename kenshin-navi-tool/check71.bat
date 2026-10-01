@echo off
rem Verify the Fussa booking import (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check71.ps1"
