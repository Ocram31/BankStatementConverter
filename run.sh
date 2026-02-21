#!/bin/bash
echo "============================================================"
echo "  Bank Statement PDF to CSV Converter"
echo "============================================================"
echo
echo "  PDFs from:  pdfs/"
echo "  CSVs to:    csv/"
echo

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

# --- Check Python ---
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed."
    echo
    echo "Install it with:"
    echo "  Ubuntu/Debian: sudo apt install python3 python3-pip python3-venv"
    echo "  Mac:           brew install python3"
    echo "  Or download from: https://www.python.org/downloads/"
    exit 1
fi

# --- Auto-setup on first run ---
if [ ! -d "venv" ]; then
    echo "First run detected - setting up automatically..."
    echo
    chmod +x setup.sh 2>/dev/null
    ./setup.sh
    if [ $? -ne 0 ]; then
        echo "Setup failed. Please check the errors above."
        exit 1
    fi
    echo
    echo "Setup complete! Starting conversion..."
    echo
fi

# Check if pdfs/ folder has any PDFs
if [ -d "pdfs" ] && [ -z "$(ls pdfs/*.pdf 2>/dev/null)" ]; then
    echo "No PDF files found in the pdfs/ folder."
    echo "Copy your bank statement PDFs into: $SCRIPT_DIR/pdfs/"
    echo
fi

source venv/bin/activate
python3 convert.py "$@"

echo
echo "============================================================"
echo "  Done. CSV files are in the 'csv' folder."
echo "============================================================"
