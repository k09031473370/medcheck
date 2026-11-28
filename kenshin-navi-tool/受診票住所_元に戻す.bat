@echo off
rem Undo a jusho_v53 import (choose the record)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0jusho_v53.ps1" -Undo ask
pause
