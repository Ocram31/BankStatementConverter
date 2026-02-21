@echo off
cd /d "%~dp0"

REM --- Check Python ---
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed.
    echo.
    echo Please download and install Python from:
    echo   https://www.python.org/downloads/
    echo.
    echo IMPORTANT: Tick "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

REM --- Auto-setup on first run ---
if not exist "venv\Scripts\activate.bat" (
    echo First run detected - setting up automatically...
    echo.
    call setup.bat
    if %errorlevel% neq 0 (
        echo Setup failed. Please check the errors above.
        pause
        exit /b 1
    )
    echo.
)

call venv\Scripts\activate.bat
pythonw converter_gui.py
