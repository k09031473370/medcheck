@echo off
rem 健診結果 取込ツール (ダブルクリックで起動)
start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0form_import_gui.ps1"
