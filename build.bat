@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found in PATH.
    pause
    exit /b 1
)

echo Installing dependencies...
python -m pip install --upgrade pip setuptools
if errorlevel 1 goto :fail

python -m pip install -r requirements.txt
if errorlevel 1 goto :fail

echo Cleaning previous build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo Building AcademicFigureCropper.exe...
python -m PyInstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name AcademicFigureCropper ^
    --icon icon.ico ^
    --add-data "icon.ico;." ^
    --collect-all tkinterdnd2 ^
    main.py
if errorlevel 1 goto :fail

echo.
echo Build finished: dist\AcademicFigureCropper.exe
pause
exit /b 0

:fail
echo.
echo Build failed.
pause
exit /b 1
