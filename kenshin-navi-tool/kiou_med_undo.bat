@echo off
rem Undo a kiou_med import (choose the record)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kiou_med.ps1" -Undo ask
pause
