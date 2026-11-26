@echo off
rem Undo a kiou_v5 import (choose the record)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0kiou_v5.ps1" -Undo ask
pause
