@echo off
rem Undo a kiou123 import (choose the record)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kiou123.ps1" -Undo ask
pause
