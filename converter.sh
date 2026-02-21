#!/bin/bash
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
fi

source venv/bin/activate
python3 converter_gui.py &
