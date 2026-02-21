# Bank Statement PDF to CSV Converter

## User Manual

This tool converts PDF bank statements into CSV files that can be imported
into Sage, Pastel, Xero, QuickBooks, or any accounting software.

---

## Table of Contents

1. [Standalone Download (easiest)](#standalone-download-easiest)
2. [One-Time Setup (source version)](#one-time-setup-source-version)
3. [How to Convert Bank Statements](#how-to-convert-bank-statements)
4. [What the Output Looks Like](#what-the-output-looks-like)
5. [Output Formats](#output-formats)
6. [Advanced Usage](#advanced-usage)
7. [Supported Banks](#supported-banks)
8. [Importing into Accounting Software](#importing-into-accounting-software)
9. [Password-Protected PDFs](#password-protected-pdfs)
10. [Troubleshooting](#troubleshooting)
11. [Quick Reference Card](#quick-reference-card)

---

## Standalone Download (easiest)

**No Python or setup needed.** Download the ready-to-run app for your platform:

**Download page:** https://github.com/Ocram31/BankStatementConverter/releases/latest

| Platform | Download | How to open |
|----------|----------|-------------|
| **Windows** | `BankStatementConverter-Windows.zip` | Unzip, double-click **`BankStatementConverter.exe`** |
| **Mac** | `BankStatementConverter-Mac.dmg` | Open DMG, drag app to Applications, right-click → **Open** first time |
| **Linux** | `BankStatementConverter-Linux.zip` | Unzip, run `./BankStatementConverter` in Terminal |

> **Mac users:** macOS will warn you the first time — see [Troubleshooting: Mac Gatekeeper](#mac-cannot-be-opened-because-it-is-from-an-unidentified-developer--cannot-verify-it-is-free-of-malware).

> **Windows users:** SmartScreen may warn you — click **"More info"** → **"Run anyway"**. See [Troubleshooting: Windows SmartScreen](#windows-windows-protected-your-pc-smartscreen).

If you downloaded the standalone app, **skip to [How to Convert Bank Statements](#how-to-convert-bank-statements)**.

---

## One-Time Setup (source version)

> **Already downloaded the standalone app above?** Skip this entire section.

> **Auto-setup:** If you have Python installed, you can skip Step 2 entirely.
> Just double-click `converter.bat` or `run.bat` — it will set up automatically
> on first run.

You only need to install Python once on your computer.

### Step 1: Install Python

#### Windows:

1. Open your web browser and go to: **https://www.python.org/downloads/**
2. Click the big yellow button that says **"Download Python 3.x.x"**.
3. Run the downloaded file.
4. **IMPORTANT:** On the first screen of the installer, tick the checkbox
   that says **"Add Python to PATH"** at the bottom. This is essential.
5. Click **"Install Now"** and wait for it to finish.
6. Click **"Close"**.

#### Mac:

1. Open **Terminal** (press Cmd + Space, type "Terminal", press Enter).
2. Check if Python is installed: type `python3 --version` and press Enter.
3. If you see `Python 3.x.x`, it's already installed. Skip to Step 2.
4. If not, go to **https://www.python.org/downloads/** and install it.

#### Linux (Ubuntu / Debian):

1. Open Terminal (press **Ctrl + Alt + T**).
2. Check if Python is installed: type `python3 --version` and press Enter.
3. If you see `Python 3.x.x`, it's already installed. Skip to Step 2.
4. If not, run:
   ```
   sudo apt update
   sudo apt install python3 python3-pip python3-venv
   ```

---

### Step 2: Run the Converter (auto-setup)

Just double-click **`converter.bat`** (Windows) or **`converter.sh`**
(Mac/Linux). On first run, it automatically:
- Creates a private workspace (virtual environment)
- Installs all required libraries
- Opens the converter

**That's it. No separate setup step needed.**

If you prefer to set up manually, you can still run `setup.bat` / `setup.sh`
first, but it's not required — the launchers handle it automatically.

---

## How to Convert Bank Statements

You have two options: a **graphical window** (easiest) or the
**command line** (fastest for batch processing). Both produce
the same results.

---

### Option A: Graphical Window (easiest)

1. **Windows:** Double-click **`converter.bat`**
   **Mac:** Double-click **`converter.command`**
   **Linux:** Run `./converter.sh` in Terminal

2. A window will open with buttons:
   - Click **"Browse for PDFs..."** to select your bank statement PDFs.
   - Or click **"Add all PDFs in folder..."** to select a whole folder.

3. Choose your **Output format** from the dropdown:
   - **Generic** — simple Date, Description, Amount (works with most software)
   - **Sage Split** — separate Money in / Money out columns (for Sage/Pastel)
   - **Xero** — Xero-specific format with Payee column
   - **QuickBooks Split** — separate Credit / Debit columns
   - See [Output Formats](#output-formats) for full list.

4. If your PDF is password-protected, enter the password in the
   **PDF password** field (SA banks often use your ID number).

5. **Optional features** (checkboxes below the format dropdown):
   - **Auto-categorise** — labels each transaction (e.g. Salary, Groceries,
     Fuel, Insurance). On by default.
   - **Smart filenames** — names output files with the bank and date range
     (e.g. `ABSA_Current_2026-01-01_to_2026-01-31.csv`). On by default.
   - **Excel output** — creates an `.xlsx` file in addition to CSV. Off by default.
   - **Open folder when done** — opens the output folder after conversion. On by default.

   These settings are remembered between sessions.

6. Click the **"Convert"** button.

7. Watch the log at the bottom — it shows progress, verification
   results, category breakdown, and any duplicate warnings for each file.

8. When it says **"ALL DONE"**, a popup confirms how many files were
   converted. Your CSV files are in the `csv` folder next to the app.

---

### Option B: Command Line (fastest)

1. Download your PDF bank statements from your bank's internet banking.

2. Copy the PDF files into the **`pdfs`** folder inside the converter folder.

3. Run the converter:

   **Windows:** Double-click **`run.bat`**

   **Mac / Linux:** Open Terminal and type:
   ```
   cd /path/to/conversion
   ./run.sh
   ```

4. The tool will:
   - Find all PDF files in the folder automatically
   - Detect which bank each one is from
   - Convert each one to CSV
   - Check that all transactions add up correctly
   - Warn about potential duplicates
   - Show you a summary table

5. Your CSV files will be in the **`csv`** subfolder, ready to import
   into your accounting software.

---

## What the Output Looks Like

When you run the converter, you'll see something like this:

```
======================================================================
PDF Bank Statement to CSV Converter for SA Accounting
======================================================================

Processing: ABSA Jan.pdf
  Parser: absa_current (auto-detected)
  Output: ABSA Jan.csv (format: generic)

  Transactions: 66
  Date range: 01/01/2026 to 31/01/2026
  Total debits:  R    -440489.50
  Total credits: R     508944.48
  Net movement:  R      68454.98
  Balance check: PASS (all 66 consecutive balances reconcile)

Processing: FNB Jan.pdf
  Parser: fnb (auto-detected)
  Output: FNB Jan.csv (format: generic)

  Transactions: 169
  Balance check: PASS (all 169 consecutive balances reconcile)

======================================================================
SUMMARY REPORT
Output format: generic
----------------------------------------------------------------------
File                                Parser           Txns            Net Status
----------------------------------------------------------------------
ABSA Jan.pdf                        absa_current       66  R    68454.98 PASS
FNB Jan.pdf                         fnb               169  R    46797.33 PASS
----------------------------------------------------------------------
TOTAL                                                 235  R   115252.31
======================================================================

======================================================================
ALL VERIFICATIONS PASSED
======================================================================
```

**What to look for:**

| Message | Meaning |
|---------|---------|
| Balance check: **PASS** | Every transaction adds up correctly |
| **ALL VERIFICATIONS PASSED** | Everything is good, your CSV files are ready |
| Balance check: **FAIL** | Some transactions didn't add up — check them manually |
| **WARNING: potential duplicate(s)** | Same date+amount+description found — may need review |
| **ERROR** | Something went wrong with that file (see Troubleshooting) |

---

## Output Formats

Choose the format that matches your accounting software:

| Format | Command | Columns | Best for |
|--------|---------|---------|----------|
| **Generic** (default) | `--format generic` | Date, Description, Amount | Excel, any software |
| **Sage** | `--format sage` | Date, Description, Amount | Sage (zero amounts blank) |
| **Sage Split** | `--format sage_split` | Date, Description, Money in, Money out | Sage/Pastel 4-column import |
| **Xero** | `--format xero` | *Date, *Amount, Payee, Description, Reference | Xero precoded format |
| **QuickBooks** | `--format quickbooks` | Date, Description, Amount | QuickBooks (no commas) |
| **QuickBooks Split** | `--format quickbooks_split` | Date, Description, Credit, Debit | QuickBooks 4-column |

### Amount handling by format:

- **Generic / Sage / QuickBooks:** Single signed Amount column. Negative = money out, positive = money in.
- **Sage Split:** Separate unsigned "Money in" and "Money out" columns. Only one is filled per row.
- **QuickBooks Split:** Separate unsigned "Credit" and "Debit" columns. Only one is filled per row.
- **Xero:** Signed amount with separate Payee field extracted from description.

### Encoding:

- **Generic / Sage / Sage Split:** UTF-8 with BOM (for Excel compatibility on Windows)
- **Xero / QuickBooks / QuickBooks Split:** UTF-8 without BOM

---

## Advanced Usage

These are optional — the basic method above covers most needs.

### Choosing an output format

**Command line:**
```
python3 convert.py --format sage_split
python3 convert.py -f xero
python3 convert.py -f quickbooks_split
```

**GUI:** Select from the "Output format" dropdown before clicking Convert.

### Converting specific files only

If you don't want to convert all PDFs in the folder:

**Windows:** Open Command Prompt, then:
```
cd C:\Users\YourName\Documents\conversion
venv\Scripts\activate
python convert.py "ABSA Jan.pdf" "FNB Jan.pdf"
```

**Mac / Linux:**
```
cd ~/Documents/conversion
source venv/bin/activate
python3 convert.py "ABSA Jan.pdf" "FNB Jan.pdf"
```

**Important:** If the file name has spaces, put it in quotes.

### Saving CSV files to a different folder

**Windows:**
```
python convert.py -o C:\Users\YourName\Desktop\statements
```

**Mac / Linux:**
```
python3 convert.py -o ~/Desktop/statements
```

The folder will be created automatically if it doesn't exist.

### Opening password-protected PDFs

SA banks often encrypt statement PDFs. Your ID number is usually the password:

**Command line:**
```
python3 convert.py --password 8501015800080
```

**GUI:** Enter the password in the "PDF password" field before clicking Convert.

You need `pikepdf` installed for this feature (see [Password-Protected PDFs](#password-protected-pdfs)).

---

## Supported Banks

### Fully tested (dedicated parsers):

| Bank | Account Types |
|------|--------------|
| ABSA | Tjekrekeningstaat (Cheque Account) |
| ABSA | Credit Card (Kredietkaart) |
| ABSA | Current Account (Lopende Rekening) |
| FNB  | Business Cheque Account (Besigheids Tjekrekening) |

### Auto-detect (generic parser):

These banks are recognised and will be parsed automatically. The tool
reads the column headers in the PDF and figures out the layout:

| Bank |
|------|
| Standard Bank |
| Nedbank |
| Capitec |
| Investec |
| Discovery Bank |
| TymeBank |
| African Bank |

Any other bank with a standard columnar PDF layout (columns for Date,
Description, Amount/Debit/Credit, Balance) will also be attempted.

---

## Importing into Accounting Software

### Sage / Pastel

**Recommended format:** `sage_split` (4-column with Money in / Money out)

1. Run: `python3 convert.py --format sage_split` (or select "Sage Split" in the GUI)
2. Open Sage/Pastel.
3. Go to the bank statement import function.
4. Browse to the `csv` folder.
5. Select your CSV file.
6. Map the columns:
   - Column 1 = Date
   - Column 2 = Description
   - Column 3 = Money in (credits)
   - Column 4 = Money out (debits)
7. Set the date format to **DD/MM/YYYY**.
8. Click Import.

Alternatively, use `--format sage` for a simpler 3-column format (Date, Description, Amount).

### Xero

**Recommended format:** `xero`

1. Run: `python3 convert.py --format xero` (or select "Xero" in the GUI)
2. In Xero, go to **Accounting > Bank Accounts**.
3. Select your bank account.
4. Click **Import a Statement**.
5. Select the CSV file from the `csv` folder.
6. Xero will auto-detect the columns (*Date, *Amount, Payee, Description, Reference).
7. Click Import.

### QuickBooks

**Recommended format:** `quickbooks_split` (4-column with Credit / Debit)

1. Run: `python3 convert.py --format quickbooks_split` (or select "QuickBooks Split" in the GUI)
2. In QuickBooks, go to **Banking > Upload Transactions**.
3. Select your bank account.
4. Browse to the CSV file in the `csv` folder.
5. Map the columns:
   - Column 1 = Date
   - Column 2 = Description
   - Column 3 = Credit
   - Column 4 = Debit
6. Click Import.

Alternatively, use `--format quickbooks` for a simpler 3-column format.

### Excel

**Recommended format:** `generic` (default)

1. Run: `python3 convert.py` (default format)
2. Open Excel.
3. Open the CSV file from the `csv` folder.
4. The file uses UTF-8 with BOM, so special characters (Afrikaans) display correctly.

---

## Password-Protected PDFs

SA banks commonly encrypt statement PDFs. The password is usually your
**ID number** (13 digits).

### Setup

Install the `pikepdf` library (one-time):

**Windows:**
```
venv\Scripts\activate
pip install pikepdf
```

**Mac / Linux:**
```
source venv/bin/activate
pip install pikepdf
```

### Usage

**Command line:**
```
python3 convert.py --password 8501015800080
```

**GUI:** Enter the password in the "PDF password" field.

The tool automatically detects whether a PDF is encrypted. If it's not
encrypted, the password is simply ignored.

---

## Troubleshooting

### Mac: "Cannot be opened because it is from an unidentified developer" / "Cannot verify it is free of malware"

macOS blocks apps that aren't signed with an Apple certificate. To open the converter:

1. **Don't** double-click the app.
2. **Right-click** (or Control+click) on `BankStatementConverter`.
3. Click **"Open"** from the menu.
4. A dialog appears — click **"Open"** again.
5. macOS remembers your choice — it won't ask again.

If that doesn't work (on newer macOS versions), open **Terminal** and run:

```
xattr -cr /path/to/BankStatementConverter
```

**Tip:** Type `xattr -cr ` (with a space at the end), then drag the BankStatementConverter
folder from Finder onto the Terminal window — it fills in the path for you. Press Enter.

This removes the quarantine flag that macOS adds to downloaded files. You only need to do
this once.

---

### Windows: "Windows protected your PC" (SmartScreen)

Windows may show a blue SmartScreen warning for downloaded apps. To open the converter:

1. Click **"More info"** on the SmartScreen popup.
2. Click **"Run anyway"**.
3. Windows remembers your choice — it won't ask again.

---

### "Setup complete!" didn't appear when running setup

- **Windows:** Make sure Python is installed and "Add to PATH" was ticked.
  Reinstall Python if needed.
- **Linux:** You may need to install the venv package:
  `sudo apt install python3-venv`

### "No PDF files found"

Your PDF files aren't in the `pdfs/` folder. Copy them into `pdfs/` first,
then run the converter again.

### "PDF appears to be password-protected"

The PDF is encrypted. Two options:
- Install `pikepdf` and use `--password` (see [Password-Protected PDFs](#password-protected-pdfs))
- Open the PDF in your browser/reader, enter the password, and "Print to PDF"
  to create an unencrypted copy.

### "Wrong password for PDF"

The password you provided didn't work. SA banks typically use:
- Your **ID number** (13 digits, e.g. 8501015800080)
- Your **account number**
- A password you set in internet banking

### "Could not extract text from PDF"

The PDF is a scanned image, not digital text. This tool works with PDFs
downloaded from internet banking (which contain selectable text). It does
not work with photos or scanned paper statements.

### "Could not detect column headers"

The generic parser couldn't find column headers (Date, Description,
Debit/Credit, Balance) in the PDF. This bank uses an unusual layout —
contact your developer to add support.

### "Balance check: FAIL"

Some transactions didn't reconcile. The CSV is still created, but check
the flagged transactions manually. This usually means the PDF has a
formatting quirk on certain pages.

### "WARNING: potential duplicate(s)"

Transactions with the same date, amount, and description were found.
This is usually legitimate (e.g. multiple identical payments) but worth
checking to ensure nothing was double-counted.

### "ERROR" on a specific file

That file couldn't be processed. Other files still convert normally.
Common causes:
- The PDF is corrupted or password-protected
- The PDF is from a bank with a very unusual layout

### "'python' / 'python3' is not recognized"

Python isn't installed or wasn't added to PATH.
- **Windows:** Reinstall Python, tick **"Add Python to PATH"**.
- **Mac / Linux:** Install with `brew install python3` or `sudo apt install python3`.

### "No module named 'tkinter'" (GUI only)

tkinter is needed for the graphical window. Install it:
- **Ubuntu/Debian:** `sudo apt install python3-tk`
- **Fedora/RHEL:** `sudo dnf install python3-tkinter`
- **Mac:** `brew install python-tk`
- **Windows:** Reinstall Python, tick **"tcl/tk and IDLE"** in the installer.

The command-line converter (`run.bat` / `run.sh`) works without tkinter.

### "No module named 'pdfplumber'"

The virtual environment isn't activated. Run the setup script again, or
activate manually:
- **Windows:** `venv\Scripts\activate` then `pip install pdfplumber`
- **Mac / Linux:** `source venv/bin/activate` then `pip install pdfplumber`

---

## Tips

- **Always verify:** Look for "ALL VERIFICATIONS PASSED" at the end.
  If you see it, your CSV files are accurate.
- **Your PDFs are safe:** The tool never modifies or deletes your PDF
  files. Your originals are always untouched.
- **Batch processing:** Drop all your PDFs from all banks into the `pdfs/`
  folder at once. The tool processes them all and auto-detects each bank.
- **Monthly routine:** Download statements, copy into `pdfs/` folder,
  double-click `run.bat` (or `./run.sh`), import CSVs into Sage. Under a minute.
- **Check duplicates:** If you see duplicate warnings, review those
  transactions — they're usually legitimate but worth confirming.

---

## Quick Reference Card

### First time:

| OS | What to do |
|----|-----------|
| Windows | Install Python (tick "Add to PATH"), then double-click **converter.bat** — it sets up automatically |
| Mac | Install Python if needed, then double-click **converter.command** — auto-setup on first run |
| Linux | Install Python if needed, then run `./converter.sh` in Terminal — auto-setup on first run |

### Every time you convert:

| OS | GUI (easiest) | Command line (fastest) |
|----|--------------|----------------------|
| Windows | Double-click **converter.bat** | Double-click **run.bat** |
| Mac | Double-click **converter.command** | Double-click **run.command** |
| Linux | Run `./converter.sh` in Terminal | Run `./run.sh` in Terminal |

### Command line flags:

| Flag | Example | Purpose |
|------|---------|---------|
| `--format` / `-f` | `-f sage_split` | Choose output format |
| `--output-dir` / `-o` | `-o ./csv` | Change output folder |
| `--parser` / `-p` | `-p fnb` | Force a specific bank parser |
| `--password` | `--password 8501015800080` | Decrypt password-protected PDFs |

### Files in the converter folder:

| File | What it is |
|------|-----------|
| `converter.bat` | Opens the GUI — Windows |
| `converter.command` | Opens the GUI — Mac (double-click in Finder) |
| `converter.sh` | Opens the GUI — Linux |
| `run.bat` | Command-line converter — Windows |
| `run.command` | Command-line converter — Mac |
| `run.sh` | Command-line converter — Linux |
| `setup.bat` / `setup.sh` | Manual setup (optional — launchers auto-setup) |
| `convert.py` | The converter engine (don't modify) |
| `converter_gui.py` | The graphical interface (don't modify) |
| `requirements.txt` | List of dependencies (used by setup) |
| `pdfs/` | Input folder — put your bank statement PDFs here |
| `csv/` | Output folder — your CSV files go here |
| `VERSION` | Current version number |
| `USER_MANUAL.md` | This manual |
