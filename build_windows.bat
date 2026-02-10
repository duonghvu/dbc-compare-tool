@echo off
REM Build standalone executables for DBC Compare Tool (Windows)
REM Outputs:
REM   dist\dbc_compare_gui.exe  - GUI application (double-click to run)
REM   dist\dbc_compare.exe      - CLI version (terminal usage)

echo === DBC Compare Tool - Build (Windows) ===
echo.

echo [1/4] Installing dependencies...
python -m pip install --quiet pyinstaller openpyxl

echo [2/4] Building GUI application...
python -m PyInstaller --onefile --clean --name dbc_compare_gui --windowed --add-data "dbc_compare.py;." dbc_compare_gui.py

echo [3/4] Building CLI tool...
python -m PyInstaller --onefile --clean --name dbc_compare --console dbc_compare.py

echo [4/4] Build complete!
echo.
echo Executables:
echo   GUI: dist\dbc_compare_gui.exe   (double-click to run)
echo   CLI: dist\dbc_compare.exe        (terminal usage)
echo.
pause
