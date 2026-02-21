#!/bin/bash
echo "============================================================"
echo "  Bank Statement Converter - First Time Setup (Mac/Linux)"
echo "============================================================"
echo

if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed."
    echo
    echo "Install it with:"
    echo "  Ubuntu/Debian: sudo apt install python3 python3-pip python3-venv"
    echo "  Mac:           brew install python3"
    echo "  Or download from: https://www.python.org/downloads/"
    exit 1
fi

echo "Found Python:"
python3 --version
echo

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "Creating virtual environment..."
python3 -m venv venv
if [ $? -ne 0 ]; then
    echo "ERROR: Could not create virtual environment."
    echo "On Ubuntu/Debian, try: sudo apt install python3-venv"
    exit 1
fi

echo "Activating virtual environment..."
source venv/bin/activate

echo "Installing dependencies..."
pip install -r requirements.txt
if [ $? -ne 0 ]; then
    echo "ERROR: Could not install dependencies."
    exit 1
fi

# Check for tkinter (needed for GUI)
if ! python3 -c "import tkinter" 2>/dev/null; then
    echo
    echo "NOTE: tkinter is not installed. The GUI will not work."
    echo "Install it with:"
    echo "  Ubuntu/Debian: sudo apt install python3-tk"
    echo "  Fedora/RHEL:   sudo dnf install python3-tkinter"
    echo "  Mac:           brew install python-tk"
    echo
    echo "The command-line converter (run.sh) will still work without it."
fi

echo
echo "============================================================"
echo "  Setup complete!"
echo
echo "  To convert (command line): ./run.sh"
echo "  To convert (GUI):         ./converter.sh"
echo
echo "  Optional extras (already included in requirements.txt):"
echo "    pikepdf  — decrypt password-protected PDFs"
echo
echo "  For OCR of FNB image-based fee descriptions:"
echo "    pip install pytesseract Pillow"
echo "    sudo apt install tesseract-ocr tesseract-ocr-afr"
echo "============================================================"
