@echo off
echo Building Environics Data Tools Executable...
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo Error: Python is not installed or not in PATH
    pause
    exit /b 1
)

echo Installing required packages...
pip install pyinstaller
pip install -r requirements.txt

echo.
echo Building executable with PyInstaller...
pyinstaller build_exe.spec --clean --noconfirm

if exist "dist\EnvironicsDataTools.exe" (
    echo.
    echo ========================================
    echo BUILD SUCCESSFUL!
    echo ========================================
    echo.
    echo Your executable has been created at:
    echo %cd%\dist\EnvironicsDataTools.exe
    echo.
    echo You can now distribute this single .exe file.
    echo It contains all your Python programs and dependencies.
    echo.
) else (
    echo.
    echo ========================================
    echo BUILD FAILED!
    echo ========================================
    echo.
    echo Please check the output above for errors.
    echo.
)

pause
