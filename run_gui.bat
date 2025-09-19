@echo off
title Glucose Converter GUI

echo Starting Glucose Converter GUI...
echo.

REM Run the GUI application
python glucose_converter_gui.py

if errorlevel 1 (
    echo.
    echo [ERROR] Failed to start the application
    echo.
    echo Please run "setup_windows.bat" first to install requirements
    echo.
    pause
)