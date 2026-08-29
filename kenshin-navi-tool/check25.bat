@echo off
rem Who has empty SEQ1 - decisive test for the auto judgement crash
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check25.ps1"
