@echo off
rem Dump one person's judgments and comments (read-only). Run before and after auto-judgment and compare.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check101.ps1"
