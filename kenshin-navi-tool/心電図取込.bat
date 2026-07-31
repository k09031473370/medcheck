@echo off
setlocal
rem 心電図CSVの取込 (CSVをこのファイルにドラッグ＆ドロップ、またはダブルクリック)
cd /d "%~dp0"

set "CSVPATH=%~1"
if "%CSVPATH%"=="" (
  echo 心電計から出たCSVファイルを、このバットにドラッグ^&ドロップしてください。
  echo または、下にファイルのパスを貼り付けてEnterを押してください。
  echo.
  set /p CSVPATH="CSVのパス: "
)
if "%CSVPATH%"=="" goto :end
set "CSVPATH=%CSVPATH:"=%"

echo.
echo ================== プレビュー (書き込みません) ==================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ecg_import.ps1" -Csv "%CSVPATH%"
echo.
echo 上の内容を確認してください。
set /p YN="この内容で書き込みますか? (y = 書き込む / それ以外 = やめる): "
if /i not "%YN%"=="y" (
  echo 書き込みませんでした。
  goto :end
)
echo.
echo ================== 書込実行 ==================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0ecg_import.ps1" -Csv "%CSVPATH%" -Commit
echo.
echo 健診ナビで「自動判定」を実行してください。

:end
echo.
pause
