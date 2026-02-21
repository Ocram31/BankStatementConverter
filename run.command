#!/bin/bash
# macOS: double-click this file to run the command-line converter.
# It will auto-install dependencies on first run.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"
chmod +x run.sh setup.sh 2>/dev/null
exec ./run.sh
