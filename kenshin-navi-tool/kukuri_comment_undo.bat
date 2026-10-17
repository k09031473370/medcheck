@echo off
rem Undo a kukuri_comment run (restores the previous comment text)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kukuri_comment.ps1" -Undo ask
pause
