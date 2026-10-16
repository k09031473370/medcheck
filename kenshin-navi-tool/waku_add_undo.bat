@echo off
rem Undo slots added by waku_add (choose a record file; only rows still empty are deleted)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0waku_add.ps1" -Undo ask
pause
