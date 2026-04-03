@echo off
setlocal
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found in PATH.
    pause
    exit /b 1
)

set "BUILD_VENV=%~dp0.build-venv"
set "VENV_PYTHON=%BUILD_VENV%\Scripts\python.exe"

if not exist "%VENV_PYTHON%" (
    echo Creating isolated build environment...
    python -m venv "%BUILD_VENV%"
    if errorlevel 1 goto :fail
)

echo Upgrading packaging tools...
"%VENV_PYTHON%" -m pip install --upgrade pip
if errorlevel 1 goto :fail

"%VENV_PYTHON%" -m pip install --upgrade --force-reinstall "setuptools<81"
if errorlevel 1 goto :fail

echo Installing build dependencies...
"%VENV_PYTHON%" -m pip install -r requirements.txt
if errorlevel 1 goto :fail

"%VENV_PYTHON%" -c "import pkg_resources, PyInstaller, fitz, PIL, numpy, tkinterdnd2" >nul 2>nul
if errorlevel 1 (
    echo Required build dependencies are missing.
    goto :fail
)

echo Cleaning previous build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo Building AcademicFigureCropper.exe...
"%VENV_PYTHON%" -m PyInstaller --noconfirm --clean "AcademicFigureCropper.spec"
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
