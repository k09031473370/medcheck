@echo off
rem Find out why the ECG judgement is empty (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check35.ps1"
