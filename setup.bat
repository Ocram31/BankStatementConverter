@echo off
echo ============================================================
echo   Bank Statement Converter - First Time Setup (Windows)
echo ============================================================
echo.

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

echo Found Python:
python --version
echo.

echo Creating virtual environment...
python -m venv venv
if %errorlevel% neq 0 (
    echo ERROR: Could not create virtual environment.
    pause
    exit /b 1
)

echo Activating virtual environment...
call venv\Scripts\activate.bat

echo Installing dependencies...
pip install -r requirements.txt
if %errorlevel% neq 0 (
    echo ERROR: Could not install dependencies.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo   Setup complete!
echo.
echo   To convert (command line): double-click run.bat
echo   To convert (GUI):         double-click converter.bat
echo.
echo   Optional: For OCR of FNB image-based fee descriptions:
echo     pip install pytesseract Pillow
echo     (also install Tesseract-OCR from github.com/tesseract-ocr)
echo ============================================================
echo.
pause
