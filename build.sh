#!/bin/bash
# Build standalone executables for DBC Compare Tool
# Outputs:
#   dist/dbc_compare_gui.app  - GUI application (macOS .app bundle, double-click)
#   dist/dbc_compare          - CLI version (terminal usage)

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "=== DBC Compare Tool - Build ==="
echo ""

# Check dependencies
echo "[1/4] Installing dependencies..."
python3 -m pip install --quiet pyinstaller openpyxl

# Build GUI app (onedir for proper .app bundle on macOS)
echo "[2/4] Building GUI application..."
python3 -m PyInstaller \
    --onedir \
    --clean \
    --name dbc_compare_gui \
    --windowed \
    --add-data "dbc_compare.py:." \
    dbc_compare_gui.py

# Build CLI tool (onefile for single binary)
echo "[3/4] Building CLI tool..."
python3 -m PyInstaller \
    --onefile \
    --clean \
    --name dbc_compare \
    --console \
    dbc_compare.py

echo "[4/4] Build complete!"
echo ""
echo "Executables:"
if [ -d "$SCRIPT_DIR/dist/dbc_compare_gui.app" ]; then
    echo "  GUI: $SCRIPT_DIR/dist/dbc_compare_gui.app  (double-click to run)"
else
    echo "  GUI: $SCRIPT_DIR/dist/dbc_compare_gui/dbc_compare_gui  (run from terminal)"
fi
echo "  CLI: $SCRIPT_DIR/dist/dbc_compare  (terminal usage)"
