@echo off
echo ========================================================
echo  Glucose Converter - Windows Setup
echo ========================================================
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python is not installed or not in PATH
    echo.
    echo Please install Python from: https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation!
    echo.
    pause
    exit /b 1
)

echo [OK] Python is installed
python --version
echo.

REM Install requirements
echo Installing required packages...
echo.

pip install --upgrade pip
pip install openpyxl
pip install tkinterdnd2

echo.
echo ========================================================
echo  Setup Complete!
echo ========================================================
echo.
echo You can now run the application using:
echo   1. GUI Version: python glucose_converter_gui.py
echo   2. Command Line: python glucose_converter.py [csv_file]
echo   3. Create EXE: python build_exe.py
echo.
echo Or simply double-click "run_gui.bat" to start the GUI
echo.
pause