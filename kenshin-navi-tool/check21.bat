@echo off
rem Find items nobody else ever filled (auto judgement crash suspects)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check21.ps1"
