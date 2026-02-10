# DBC Compare Tool

A **cross-platform visual application** to compare two versions of CAN DBC database files and generate comparison reports in **Excel**, **HTML**, and **PDF** formats with categorized change sheets.

Runs on **Windows**, **macOS**, and **Linux**.

![GUI Screenshot](docs/screenshot.png)

---

## Features

- **Visual GUI** — browse folders, click Compare, view results
- **CLI mode** — scriptable command-line interface for automation
- **Cross-platform** — standalone executables for Windows (.exe), macOS (.app), Linux
- Parses `.dbc` files (Vector CANdb++ format)
- Compares **all messages and signals** side-by-side
- Covers **every DBC attribute** (37 columns per side)
- **Multi-format output** — Excel (.xlsx), HTML, and PDF per bus
- **8-sheet Excel** with categorized changes:
  - Comparison Summary dashboard
  - New / Removed / Modified Messages
  - New / Removed / Modified Signals
  - Full Comparison (flat side-by-side)
- **HTML report** — standalone, responsive, opens in any browser
- **PDF report** — print-ready landscape document with all sections
- Color-coded differences (pink for changed cells, green tint for context)
- Professional styling with dark header theme and alternating row shading
- Automatically matches DBC files by bus name across folders

---

## Quick Start

### Option 1: GUI Application (recommended)

**Run with Python:**
```bash
pip install openpyxl fpdf2
python3 dbc_compare_gui.py
```

**Or build a standalone executable** (no Python needed to run):

| Platform | Build Command | Output |
|---|---|---|
| **Windows** | `build_windows.bat` | `dist\dbc_compare_gui.exe` |
| **macOS** | `./build.sh` | `dist/dbc_compare_gui.app` |
| **Linux** | `./build.sh` | `dist/dbc_compare_gui` |

Then just **double-click** the executable to launch.

### Option 2: Command Line (CLI)

```bash
pip install openpyxl fpdf2
python3 dbc_compare.py <old_folder> <new_folder>
```

---

## How to Use (GUI)

1. **Launch** the application (double-click the executable or run `python3 dbc_compare_gui.py`)
2. Click **Browse** to select the **Old Version** folder containing DBC files
3. Click **Browse** to select the **New Version** folder containing DBC files
4. (Optional) Set a custom **Output Folder**, or leave empty for auto-naming
5. Click **Compare**
6. View the results summary table showing differences per bus
7. Click **Open Output Folder** to access the generated Excel files

---

## How to Use (CLI)

```
python3 dbc_compare.py <old_folder> <new_folder> [--output <output_folder>]
```

**Example:**
```bash
python3 dbc_compare.py old_version_folder new_version_folder
```

**Output:**
```
SUMMARY
============================================================
Bus               Total Rows   Differences
--------------- ------------ -------------
CAN_CH1                 5160          1309
CAN_CH2                 2281           643
CAN_CH3                 1988           945
...
TOTAL                  14826          4983
```

---

## Excel Output Structure (8 sheets per file)

Each generated `.xlsx` file contains 8 sheets for comprehensive change analysis:

### Sheet 1: Comparison Summary

Overview statistics showing total message counts and categorized change counts.

| Metric | Example |
|--------|---------|
| Total Messages in Old Version | 256 |
| Total Messages in New Version | 257 |
| New Messages | 1 |
| Removed Messages | 0 |
| Modified Messages | 10 |
| New Signals | 35 |
| Removed Signals | 16 |
| Modified Signals | 44 |

### Sheets 2-3: New / Removed Messages

Messages that exist only in the new or old version (8 columns):

```
Message ID | Message Name | Message Length | Signals | Send Type | Cycle Time | Tx ECU | Rx ECUs
```

### Sheet 4: Modified Messages

Side-by-side comparison of messages with changes (8 fields x 2 versions = 16 columns). Red-highlighted cells indicate which fields changed.

### Sheets 5-6: New / Removed Signals

Signals that exist only in the new or old version (12 columns):

```
Message ID | Message Name | Signal Name | Start Bit | Length | Initial | Minimum | Maximum | Unit | Factor | Value Table | Receivers
```

### Sheet 7: Modified Signals

3 fixed identifier columns + 9 paired old/new attribute columns (21 columns total). Red-highlighted cells indicate which attributes changed.

### Sheet 8: Full Comparison

Comprehensive flat side-by-side table with all 37 attributes per side (74 columns total), covering every DBC attribute.

---

## Columns Compared (37 per side)

### Message-level (19 columns)

| Column | DBC Source | Description |
|---|---|---|
| Id | BO_ | Message ID (hex) |
| Name | BO_ / SystemMessageLongSymbol | Message name |
| SendType | GenMsgSendType | Cyclic, IfActive, etc. |
| CycleTime | GenMsgCycleTime | Normal cycle time (ms) |
| CycleTimeFast | GenMsgCycleTimeFast | Fast cycle time (ms) |
| CycleTimeActive | GenMsgCycleTimeActive | Active cycle time (ms) |
| NrOfRepetition | GenMsgNrOfRepetition | Repetition count |
| DelayTime | GenMsgDelayTime | Delay time (ms) |
| StartDelayTime | GenMsgStartDelayTime | Start delay time (ms) |
| Dlc | BO_ | Data length code |
| VFrameFormat | VFrameFormat | StandardCAN / CAN_FD |
| CANFD_BRS | CANFD_BRS | Bit rate switch |
| Transmitter | BO_ | Transmitting node |
| ILSupport | GenMsgILSupport | Interaction layer support |
| NmMessage | NmMessage | Network management flag |
| DiagRequest | DiagRequest | Diagnostic request |
| DiagResponse | DiagResponse | Diagnostic response |
| DiagState | DiagState | Diagnostic state |
| MsgComment | CM_ BO_ | Message comment |

### Signal-level (18 columns)

| Column | DBC Source | Description |
|---|---|---|
| Name | SG_ / SystemSignalLongSymbol | Signal name |
| Start | SG_ | Start bit (natural position) |
| Length | SG_ | Bit length |
| Endian | SG_ @0/@1 | big (Motorola) / little (Intel) |
| Signed | SG_ +/- | unsigned / signed |
| Scale | SG_ factor | Scale factor |
| Offset | SG_ offset | Offset value |
| Min | SG_ min | Minimum physical value |
| Max | SG_ max | Maximum physical value |
| Unit | SG_ unit | Unit string |
| InvalidValue | BA_ InvalidValue | Signal invalid value (e.g. "FFh") |
| StartValue | BA_ GenSigStartValue | Initial / failsafe value |
| InactiveValue | BA_ GenSigInactiveValue | Inactive value |
| SendType | BA_ GenSigSendType | Signal send type |
| TimeoutTime | BA_ GenSigTimeoutTime_ALL | Timeout time (ms) |
| ValueDesc | VAL_ | Value descriptions (e.g. "0: Off, 1: On") |
| Comment | CM_ SG_ | Signal comment |
| Receivers | SG_ receivers | Receiver node list |

### Color Scheme in Excel Output

| Color | Meaning |
|---|---|
| **Light Green** (E2EFDA) | Row has differences — unchanged cells provide context |
| **Pink** (FFC7CE) | Specific cell that changed |
| **Light Gray** (F2F2F2) | Alternating row shading for readability |
| **Dark Blue Header** (1F4E79) | Column headers with white text |
| No color | Row is identical between versions |

---

## Building Standalone Executables

### Prerequisites
- Python 3.7+
- pip

### Windows
```cmd
build_windows.bat
```
Produces:
- `dist\dbc_compare_gui.exe` — GUI app (double-click)
- `dist\dbc_compare.exe` — CLI tool

### macOS
```bash
chmod +x build.sh
./build.sh
```
Produces:
- `dist/dbc_compare_gui.app` — GUI app (double-click)
- `dist/dbc_compare` — CLI tool

### Linux (Ubuntu/Debian)
```bash
# Install tkinter if not present
sudo apt-get install python3-tk

chmod +x build.sh
./build.sh
```
Produces:
- `dist/dbc_compare_gui` — GUI app
- `dist/dbc_compare` — CLI tool

---

## DBC File Naming Convention

The tool matches old vs new DBC files by **bus prefix** extracted from filenames:

```
<number>_<BUS_NAME>_<rest_of_filename>.dbc
```

Examples:
- `01_CAN_CH1_Project_v1.0.dbc` → bus prefix: `01_CAN_CH1`
- `02_CAN_CH2_Project_v2.0.dbc` → bus prefix: `02_CAN_CH2`

Files with matching bus prefixes are compared automatically.

---

## Project Structure

```
dbc-compare-tool/
  dbc_compare_gui.py   # GUI application source
  dbc_compare.py       # Core comparison engine + CLI
  requirements.txt     # Python dependencies
  build.sh             # Build script (macOS/Linux)
  build_windows.bat    # Build script (Windows)
  CHECKSUM.sha256      # Author verification hash
  README.md
  LICENSE
```

---

## Requirements

- **Python 3.7+**
- **openpyxl** — Excel file generation
- **fpdf2** — PDF report generation
- **tkinter** — GUI (included with Python on Windows/macOS; `sudo apt install python3-tk` on Linux)
- **PyInstaller** — only needed for building standalone executables

---

## License

MIT License
