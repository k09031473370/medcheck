@echo off
rem Check the course of 4742 / 4795 and the courses of the others on 9/3, 9/9, 9/10 (read-only)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check82.ps1"
