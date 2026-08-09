@echo off
rem 会社宛の請求一覧CSVを作る (ダブルクリックで実行)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0seikyu_list.ps1"
pause
