@echo off
rem Ishiya 9/3,9/9,9/10: add empty HbA1c + platelet slots to KY101/KY102 people (preview, then y to write)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0waku_add.ps1" -Items 068689,069010 -Dates 2026/09/03,2026/09/09,2026/09/10 -Course KY101,KY102 -UkeFrom 4631 -UkeTo 4831
pause
