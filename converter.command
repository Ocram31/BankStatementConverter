#!/bin/bash
# macOS: double-click this file to open the converter GUI.
# It will auto-install dependencies on first run.
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"
chmod +x converter.sh setup.sh run.sh 2>/dev/null
exec ./converter.sh
