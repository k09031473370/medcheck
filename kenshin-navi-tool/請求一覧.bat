@echo off
rem 会社宛の請求明細を作る (ダブルクリックで実行)
setlocal
chcp 932 >nul
echo ================================================
echo  健診費用の請求明細を作ります
echo  (健診ナビは読むだけです。データは変わりません)
echo ================================================
echo.
echo  受診日は 2026/06/26 のように1日分でも、
echo  2026/06 のように1か月分でも指定できます。
echo.
set "YMD="
set /p YMD=受診日または年月を入れてEnter : 
if "%YMD%"=="" (
  echo 入力されませんでした。終了します。
  pause
  exit /b
)
echo.
echo  会社名は一部でOKです(例 サンテック)。
echo  分からないときは空のままEnterを押すと、その期間に来た会社の一覧が出ます。
echo.
set "KAISHA="
set /p KAISHA=会社名の一部を入れてEnter : 
echo.
echo 集計しています...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0seikyu_list.ps1" -Ymd "%YMD%" -Dantai "%KAISHA%"
echo.
pause
