@echo off
rem Choose a saved file, list changed judgments, and put back comments that disappeared (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0comment_keep.ps1" -Restore ask
pause
