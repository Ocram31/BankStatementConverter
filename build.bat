@echo off
echo ============================================================
echo   Building Standalone Executable
echo ============================================================
echo.

cd /d "%~dp0"

REM --- Check Python ---
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed.
    pause
    exit /b 1
)

REM --- Ensure venv exists ---
if not exist "venv\Scripts\activate.bat" (
    echo Setting up virtual environment first...
    call setup.bat
)

call venv\Scripts\activate.bat

REM --- Install PyInstaller ---
echo Installing PyInstaller...
pip install pyinstaller >nul 2>nul

REM --- Build GUI executable (uses spec file for tkinterdnd2/ttkthemes) ---
echo Building standalone executable...
echo.

pyinstaller converter.spec

if %errorlevel% neq 0 (
    echo.
    echo ERROR: Build failed. Check the errors above.
    pause
    exit /b 1
)

REM --- Build CLI executable ---
echo.
echo Building command-line executable...
echo.

pyinstaller ^
    --onefile ^
    --console ^
    --name "convert" ^
    --add-data "VERSION;." ^
    --hidden-import pdfplumber ^
    --hidden-import pdfplumber.page ^
    --hidden-import pdfplumber.table ^
    --hidden-import pdfplumber.utils ^
    --hidden-import pdfminer ^
    --hidden-import pdfminer.high_level ^
    --hidden-import PIL ^
    --hidden-import PIL.Image ^
    --hidden-import PIL.ImageEnhance ^
    convert.py

if %errorlevel% neq 0 (
    echo.
    echo ERROR: CLI build failed. Check the errors above.
    pause
    exit /b 1
)

REM --- Create distribution folder ---
echo.
echo Creating distribution package...

mkdir dist\BankStatementConverter-dist 2>nul
mkdir dist\BankStatementConverter-dist\pdfs 2>nul
mkdir dist\BankStatementConverter-dist\csv 2>nul
copy dist\BankStatementConverter.exe dist\BankStatementConverter-dist\ >nul 2>nul
copy dist\convert.exe dist\BankStatementConverter-dist\ >nul 2>nul
copy README.md dist\BankStatementConverter-dist\ >nul 2>nul
copy USER_MANUAL.md dist\BankStatementConverter-dist\ >nul 2>nul
copy VERSION dist\BankStatementConverter-dist\ >nul 2>nul

echo.
echo ============================================================
echo   Build complete!
echo.
echo   Standalone files in: dist\
echo     BankStatementConverter.exe  — GUI (double-click to run)
echo     convert.exe                — Command line
echo.
echo   Distribution package in: dist\BankStatementConverter-dist\
echo     Zip this folder to share with users.
echo     No Python installation required.
echo ============================================================
echo.
pause
