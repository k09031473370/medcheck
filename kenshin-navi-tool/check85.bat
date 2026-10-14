@echo off
rem Chest X-ray import pre-check: masters (ZK020/ZK021), items and slots for Ishiya (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check85.ps1"
