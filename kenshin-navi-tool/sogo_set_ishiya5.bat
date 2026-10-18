@echo off
rem Ishiya: doctor decision - align total judgment to worst group judgment for 5 people (preview, then y)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0sogo_set.ps1" -Dates 2026/09/09,2026/09/10 -Set "4770=D>B,4778=D>C,4784=D>C,4785=G>F,4802=G>F"
pause
