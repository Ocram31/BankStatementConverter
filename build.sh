#!/bin/bash
echo "============================================================"
echo "  Building Standalone Executable"
echo "============================================================"
echo

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# --- Check Python ---
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed."
    exit 1
fi

# --- Ensure venv exists ---
if [ ! -d "venv" ]; then
    echo "Setting up virtual environment first..."
    chmod +x setup.sh 2>/dev/null
    ./setup.sh
fi

source venv/bin/activate

# --- Install PyInstaller ---
echo "Installing PyInstaller..."
pip install pyinstaller >/dev/null 2>&1

# --- Build GUI executable (uses spec file for tkinterdnd2/ttkthemes) ---
echo "Building standalone executable..."
echo

pyinstaller converter.spec

if [ $? -ne 0 ]; then
    echo
    echo "ERROR: Build failed. Check the errors above."
    exit 1
fi

# --- Build CLI executable ---
echo
echo "Building command-line executable..."
echo

pyinstaller \
    --onefile \
    --console \
    --name "convert" \
    --add-data "VERSION:." \
    --hidden-import pdfplumber \
    --hidden-import pdfplumber.page \
    --hidden-import pdfplumber.table \
    --hidden-import pdfplumber.utils \
    --hidden-import pdfminer \
    --hidden-import pdfminer.high_level \
    --hidden-import PIL \
    --hidden-import PIL.Image \
    --hidden-import PIL.ImageEnhance \
    convert.py

if [ $? -ne 0 ]; then
    echo
    echo "ERROR: CLI build failed. Check the errors above."
    exit 1
fi

# --- Create distribution folder ---
echo
echo "Creating distribution package..."

DIST_DIR="dist/BankStatementConverter-dist"
mkdir -p "$DIST_DIR/pdfs" "$DIST_DIR/csv"
cp dist/BankStatementConverter "$DIST_DIR/" 2>/dev/null
cp dist/convert "$DIST_DIR/" 2>/dev/null
cp README.md USER_MANUAL.md VERSION "$DIST_DIR/" 2>/dev/null

echo
echo "============================================================"
echo "  Build complete!"
echo
echo "  Standalone files in: dist/"
echo "    BankStatementConverter  — GUI (double-click to run)"
echo "    convert                 — Command line"
echo
echo "  Distribution package in: dist/BankStatementConverter-dist/"
echo "    Zip this folder to share with users."
echo "    No Python installation required."
echo "============================================================"
