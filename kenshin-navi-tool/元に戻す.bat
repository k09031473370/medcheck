@echo off
setlocal
rem 書込前の状態に戻す (ダブルクリックで実行)
cd /d "%~dp0"

echo ================== バックアップの一覧 ==================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0restore_backup.ps1" -List
echo.
echo 戻したいファイル名を、上の一覧からコピーして貼り付けてください。
echo (例  T_KENSA_2004717_20260731_094500.csv)
echo 何も入れずにEnterを押すと終了します。
echo.
set /p FN="ファイル名: "
if "%FN%"=="" goto :end
set "FN=%FN:"=%"

echo.
echo ================== 何が戻るかの確認 (まだ戻しません) ==================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0restore_backup.ps1" -File "%~dp0backup\%FN%"
echo.
set /p YN="この内容で戻しますか? (y = 戻す / それ以外 = やめる): "
if /i not "%YN%"=="y" (
  echo 戻しませんでした。
  goto :end
)
echo.
echo ================== 復元 ==================
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0restore_backup.ps1" -File "%~dp0backup\%FN%" -Commit
echo.
echo 健診ナビで「自動判定」をやり直してください。

:end
echo.
pause
