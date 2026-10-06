@echo off
rem Fussa day-of import, GUI version (double-click; pick or drag the reception app's JSON)
setlocal
cd /d "%~dp0"
start "" powershell -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -Command "Add-Type -AssemblyName System.Windows.Forms; try { & '%~dp0福生取込GUI.ps1' %1 } catch { $m = $_.Exception.Message + [Environment]::NewLine + $_.InvocationInfo.PositionMessage; try { [IO.File]::WriteAllText('%~dp0log\福生取込GUI_起動エラー.txt', $m) } catch {}; [Windows.Forms.MessageBox]::Show($m, '福生取込GUI を開けませんでした', 'OK', 'Error') }"
