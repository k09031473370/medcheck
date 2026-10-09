@echo off
rem Albumin / total bilirubin: item code and whether the 9/3, 9/9, 9/10 courses have a slot (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check83.ps1"
