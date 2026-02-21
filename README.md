# Bank Statement PDF to CSV Converter

**Version 2.0.0**

Converts South African bank statement PDFs into CSV files for import into
Sage, Pastel, Xero, QuickBooks, or any accounting software.

## Quick Start (One Click)

**Just double-click `converter.bat` (Windows) or `converter.sh` (Mac/Linux).**

That's it. On first run it automatically installs everything, then opens the GUI.
Drop your PDFs in, click Convert, done.

> **No Python?** Download a standalone binary from the Releases page — no
> installation needed at all.

## macOS Installation (Standalone App)

1. Download `BankStatementConverter-Mac.dmg` from the [Releases page](../../releases).
2. Double-click the DMG to mount it, then drag **BankStatementConverter** to your
   **Applications** folder.
3. **First launch:** right-click the app → **Open** → click **Open** in the dialog.
   (This is required once because the app is unsigned — Gatekeeper will warn you,
   but the app is safe to run.)

> On subsequent launches you can open it normally from Launchpad or Spotlight.

Alternatively, download `BankStatementConverter-Mac.zip`, unzip it, drag the `.app`
to Applications, and right-click → Open on first launch.

### Alternative: Command line

Double-click `run.bat` (Windows) or `./run.sh` (Mac/Linux). It auto-installs
on first run, then converts all PDFs in the `pdfs/` folder.

## Features

- **One-click install** — auto-setup on first run, no separate install step
- **6 output formats:** Generic, Sage, Sage Split, Xero, QuickBooks, QuickBooks Split
- **Auto-detection:** Identifies bank and statement type automatically
- **Balance verification:** Every transaction reconciled against running balance
- **Duplicate detection:** Flags potential duplicate transactions
- **Password-protected PDFs:** Supports encrypted statements (ID number as password)
- **Drag-and-drop GUI:** Drop PDFs directly into the converter window
- **Standalone exe available:** No Python needed (see Building below)

## Supported Banks

| Bank | Support |
|------|---------|
| ABSA | Dedicated parsers (Cheque, Credit Card, Current Account) |
| FNB | Dedicated parser (Business Cheque) |
| Standard Bank, Nedbank, Capitec, Investec, Discovery, TymeBank, African Bank | Auto-detect (generic parser) |

## Command Line Options

```
python3 convert.py                          # Convert all PDFs in pdfs/
python3 convert.py -f sage_split            # Sage 4-column format
python3 convert.py -f xero                  # Xero format
python3 convert.py -f quickbooks_split      # QuickBooks 4-column format
python3 convert.py --password 8501015800080 # Decrypt protected PDFs
python3 convert.py -o ~/Desktop/output      # Custom output folder
```

## Folder Structure

```
conversion/
  pdfs/               Put your bank statement PDFs here
  csv/                Converted CSV files appear here
  convert.py          Converter engine
  converter_gui.py    Graphical interface
  setup.bat/.sh       One-time setup (called automatically)
  run.bat/.sh         Command-line launcher (auto-installs)
  converter.bat/.sh   GUI launcher (auto-installs)
  build.bat/.sh       Build standalone exe (optional)
  USER_MANUAL.md      Full user manual
```

## Building Standalone Executables

To create a standalone `.exe` / binary that requires no Python:

**Windows:** Double-click `build.bat`
**Mac / Linux:** `./build.sh`

This creates a `dist/BankStatementConverter/` folder containing:
- `BankStatementConverter` (.exe on Windows) — GUI, double-click to run
- `convert` (.exe on Windows) — command-line version
- `pdfs/` and `csv/` folders
- Documentation

Zip the folder and share — recipients need nothing installed.

## Requirements (for development)

- Python 3.12+
- pdfplumber (installed automatically)
- Optional: pikepdf (password-protected PDFs), tkinterdnd2 (drag-and-drop), PyInstaller (building exe)

## Full Documentation

See [USER_MANUAL.md](USER_MANUAL.md) for detailed instructions, troubleshooting,
and import guides for each accounting package.
