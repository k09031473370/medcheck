@echo off
rem Ishiya: add stool test slots (OPJ010 set + 069245/069246) to the 5 people who submitted stool (preview, then y to write)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0waku_add.ps1" -Items OPJ010,069245,069246 -Dates 2026/09/03,2026/09/09 -UkeNo 4634,4635,4658,4693,4702
pause
