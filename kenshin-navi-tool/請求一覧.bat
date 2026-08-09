@echo off
rem 会社宛の請求一覧CSVを作る (ダブルクリックで実行)
setlocal
chcp 932 >nul
echo ============================================
echo  請求一覧を作ります (健診ナビは読むだけです)
echo ============================================
echo.
set "YMD="
set /p YMD=受診日を入れてEnter (例 2026/06/26) : 
if "%YMD%"=="" (
  echo 受診日が入力されませんでした。終了します。
  pause
  exit /b
)
set "KAISHA="
set /p KAISHA=会社名の一部を入れてEnter (例 サンテック) : 
if "%KAISHA%"=="" (
  echo 会社名が入力されませんでした。終了します。
  pause
  exit /b
)
echo.
echo 集計しています...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0seikyu_list.ps1" -Ymd "%YMD%" -Dantai "%KAISHA%"
echo.
pause
