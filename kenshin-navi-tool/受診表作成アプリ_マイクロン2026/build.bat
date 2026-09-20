@echo off
chcp 932 > nul
cd /d "%~dp0"
echo Folder : %CD%
echo.
if not exist "main.py" (
    echo NG : main.py ga kono folder ni arimasen.
    echo      ZIP no naka kara chokusetsu jikkou shiteimasen ka?
    echo      ZIP wo folder ni tenkai shite kara build.bat wo jikkou shite kudasai.
    echo.
    pause
    exit /b
)
echo Building main.exe from main.py ...
echo.
python -m pip install --upgrade pyinstaller reportlab
if exist "main.spec" (
    python -m PyInstaller --noconfirm main.spec
) else (
    python -m PyInstaller --noconfirm --onefile --console --name main main.py
)
if exist "dist\main.exe" (
    copy /Y "dist\main.exe" "main.exe"
    echo.
    echo OK : main.exe wo sakusei shimashita.
) else (
    echo.
    echo NG : shippai shimashita. ue no message wo kakunin shite kudasai.
)
pause
