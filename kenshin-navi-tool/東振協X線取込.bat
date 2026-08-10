@echo off
rem 東振協の胸部X線コードを一括で取り込む
rem 受付NO と 5文字コード だけのExcel/CSVを、このbatにドラッグ&ドロップ
setlocal
chcp 932 >nul
if "%~1"=="" (
  echo 取り込むファイルを、このバッチにドラッグ^&ドロップしてください。
  pause
  exit /b
)
set "YMD="
set /p YMD=受診日を入れてEnter (例 2026/07/31) : 
echo.
echo プレビューを出します(まだ書き込みません)...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0form_import.ps1" -Csv "%~1" -KenYmd "%YMD%"
echo.
set "ANS="
set /p ANS=上の内容で書き込みますか? (y = 書き込む) : 
if /i not "%ANS%"=="y" ( echo 書き込みませんでした。 & pause & exit /b )
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0form_import.ps1" -Csv "%~1" -KenYmd "%YMD%" -Commit
echo.
pause
