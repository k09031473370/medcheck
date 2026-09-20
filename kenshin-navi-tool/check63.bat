@echo off
rem Dump the template course contents (items, price, conditions) to build a new course (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check63.ps1"
