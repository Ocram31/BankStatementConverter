#!/usr/bin/env python3
"""
Bank Statement PDF to CSV Converter — Graphical Interface.

Cross-platform GUI wrapper around convert.py.
Uses ttk themed widgets for native look on Windows, Mac, and Linux.
"""

import json
import os
import sys
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from decimal import Decimal
from pathlib import Path

# Optional drag-and-drop support
try:
    import tkinterdnd2
    HAS_DND = True
except ImportError:
    HAS_DND = False

def _get_app_dir():
    """Directory where the executable (or script) lives.

    In frozen (PyInstaller) mode, this is where the .exe sits — the right
    place for user-facing folders like pdfs/ and csv/.
    """
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# Ensure convert.py is importable from the same directory
SCRIPT_DIR = _get_app_dir()
if not getattr(sys, 'frozen', False):
    sys.path.insert(0, str(SCRIPT_DIR))

import convert

# Ensure working directories exist next to the executable
(SCRIPT_DIR / "pdfs").mkdir(exist_ok=True)
(SCRIPT_DIR / "csv").mkdir(exist_ok=True)

# Format labels for the dropdown
FORMAT_OPTIONS = [
    ("Generic (Date, Description, Amount)", "generic"),
    ("Sage (no zero amounts)", "sage"),
    ("Sage Split (Money in / Money out)", "sage_split"),
    ("Xero (*Date, *Amount, Payee, Description)", "xero"),
    ("QuickBooks (Date, Description, Amount)", "quickbooks"),
    ("QuickBooks Split (Credit / Debit)", "quickbooks_split"),
]

# Settings file path (next to converter)
_SETTINGS_PATH = SCRIPT_DIR / "settings.json"

_DEFAULT_SETTINGS = {
    "format": "generic",
    "categorise": True,
    "smart_names": True,
    "xlsx": False,
    "open_folder": True,
    "output_dir": str(SCRIPT_DIR / "csv"),
}


def _load_settings():
    """Load settings from JSON file, falling back to defaults."""
    settings = dict(_DEFAULT_SETTINGS)
    if _SETTINGS_PATH.exists():
        try:
            with open(_SETTINGS_PATH, 'r', encoding='utf-8') as f:
                saved = json.load(f)
            settings.update(saved)
        except (json.JSONDecodeError, OSError):
            pass
    return settings


def _save_settings(settings):
    """Save settings to JSON file next to converter."""
    try:
        with open(_SETTINGS_PATH, 'w', encoding='utf-8') as f:
            json.dump(settings, f, indent=2)
    except OSError:
        pass


def _pick_theme(style):
    """Apply the best available ttk theme for the platform."""
    try:
        from ttkthemes import ThemedStyle
        themed = ThemedStyle(style.master)
        if sys.platform == "win32":
            themed.set_theme("vista")
        elif sys.platform == "darwin":
            themed.set_theme("aqua")
        else:
            themed.set_theme("arc")
        return
    except ImportError:
        pass

    available = style.theme_names()
    for preferred in ("clam", "alt", "vista", "aqua", "default"):
        if preferred in available:
            style.theme_use(preferred)
            return


class ConverterGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Bank Statement PDF to CSV Converter")

        # Set window icon
        self._set_icon()

        self.pdf_files = []
        self.converting = False

        # Load persisted settings
        self._settings = _load_settings()
        self.output_dir = Path(self._settings.get("output_dir", str(SCRIPT_DIR / "csv")))

        # Apply theme
        self.style = ttk.Style(self.root)
        _pick_theme(self.style)

        # Platform-appropriate fonts
        if sys.platform == "win32":
            _UI_FONT = "Segoe UI"
            _MONO_FONT = "Consolas"
        elif sys.platform == "darwin":
            _UI_FONT = "Helvetica Neue"
            _MONO_FONT = "Menlo"
        else:
            _UI_FONT = "Helvetica"
            _MONO_FONT = "Monospace"

        self._ui_font = _UI_FONT
        self._mono_font = _MONO_FONT

        # Custom styles
        self.style.configure("Title.TLabel", font=(_UI_FONT, 15, "bold"))
        self.style.configure("Subtitle.TLabel", font=(_UI_FONT, 10), foreground="#666666")
        self.style.configure("Status.TLabel", font=(_UI_FONT, 10))
        self.style.configure("Count.TLabel", font=(_UI_FONT, 9), foreground="#888888")
        self.style.configure(
            "Convert.TButton", font=(_UI_FONT, 11, "bold"), padding=(20, 6),
        )
        self.style.configure("Small.TLabel", font=(_UI_FONT, 9), foreground="#888888")

        self._build_ui()

        # Enable drag-and-drop if tkinterdnd2 is available
        self._setup_drag_and_drop()

        # Keyboard shortcuts
        self.root.bind("<Delete>", lambda e: self._remove_selected())
        self.root.bind("<BackSpace>", lambda e: self._remove_selected())

    def _set_icon(self):
        """Set the window icon from icon.png (bundled or local).

        Uses tkinter's built-in PhotoImage (reads PNG natively, no PIL needed).
        """
        try:
            if getattr(sys, 'frozen', False):
                icon_dir = Path(sys._MEIPASS)
            else:
                icon_dir = Path(__file__).parent
            icon_path = icon_dir / 'icon.png'
            if icon_path.exists():
                self._icon_photo = tk.PhotoImage(file=str(icon_path))
                self.root.iconphoto(True, self._icon_photo)
        except Exception:
            pass  # No icon — not critical

    def _build_ui(self):
        # Main container with padding
        main = ttk.Frame(self.root, padding=12)
        main.pack(fill=tk.BOTH, expand=True)

        # --- Header with icon ---
        header_row = ttk.Frame(main)
        header_row.pack(fill=tk.X)

        # Load orchid icon for header (pre-scaled 36px version)
        try:
            if getattr(sys, 'frozen', False):
                icon_dir = Path(sys._MEIPASS)
            else:
                icon_dir = Path(__file__).parent
            icon_path = icon_dir / 'icon_small.png'
            if icon_path.exists():
                self._header_icon = tk.PhotoImage(file=str(icon_path))
                # Match label background to theme so PNG transparency blends in
                bg = self.style.lookup('TFrame', 'background') or self.root.cget('bg')
                tk.Label(header_row, image=self._header_icon, bg=bg, bd=0).pack(side=tk.LEFT, padx=(0, 8))
        except Exception:
            pass

        ttk.Label(
            header_row, text="Bank Statement PDF to CSV Converter",
            style="Title.TLabel",
        ).pack(side=tk.LEFT)
        subtitle = "Drag & drop PDFs here, or use the buttons below."
        if not HAS_DND:
            subtitle = "Select PDF bank statements, choose format, then click Convert."
        ttk.Label(main, text=subtitle, style="Subtitle.TLabel").pack(anchor=tk.W, pady=(0, 10))

        # --- Button bar ---
        btn_bar = ttk.Frame(main)
        btn_bar.pack(fill=tk.X)

        ttk.Button(
            btn_bar, text="Browse PDFs...",
            command=self._browse_files,
        ).pack(side=tk.LEFT)

        ttk.Button(
            btn_bar, text="Add Folder...",
            command=self._browse_folder,
        ).pack(side=tk.LEFT, padx=(6, 0))

        ttk.Button(
            btn_bar, text="Remove Selected",
            command=self._remove_selected,
        ).pack(side=tk.LEFT, padx=(6, 0))

        ttk.Button(
            btn_bar, text="Clear All",
            command=self._clear_files,
        ).pack(side=tk.LEFT, padx=(6, 0))

        self.lbl_count = ttk.Label(btn_bar, text="0 files selected", style="Count.TLabel")
        self.lbl_count.pack(side=tk.RIGHT)

        # --- Drop zone + file list ---
        self.drop_frame = tk.Frame(
            main, bg="#e8f0fe", highlightbackground="#4a90d9",
            highlightthickness=2, highlightcolor="#4a90d9",
        )
        self.drop_frame.pack(fill=tk.BOTH, expand=False, pady=(6, 0))

        self.drop_label = tk.Label(
            self.drop_frame,
            text="Drop PDF files here",
            font=(self._ui_font, 10), fg="#4a90d9", bg="#e8f0fe",
            pady=6,
        )
        self.drop_label.pack(fill=tk.X)

        list_inner = tk.Frame(self.drop_frame, bg="#e8f0fe")
        list_inner.pack(fill=tk.BOTH, expand=True, padx=2, pady=(0, 2))

        list_scroll = ttk.Scrollbar(list_inner, orient=tk.VERTICAL)
        self.file_listbox = tk.Listbox(
            list_inner, height=5, selectmode=tk.EXTENDED,
            font=(self._mono_font, 9),
            yscrollcommand=list_scroll.set,
            relief=tk.FLAT, borderwidth=0, highlightthickness=0,
            selectbackground="#4a90d9",
        )
        list_scroll.config(command=self.file_listbox.yview)
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        list_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # --- Options ---
        opts = ttk.LabelFrame(main, text="Options", padding=8)
        opts.pack(fill=tk.X, pady=(10, 0))

        # Output directory
        out_row = ttk.Frame(opts)
        out_row.pack(fill=tk.X)

        ttk.Label(out_row, text="Save CSVs to:").pack(side=tk.LEFT)
        self.lbl_outdir = ttk.Label(out_row, text=str(self.output_dir), foreground="#0066cc")
        self.lbl_outdir.pack(side=tk.LEFT, padx=(6, 0))
        ttk.Button(out_row, text="Change...", command=self._change_output_dir).pack(side=tk.RIGHT)

        # Format selector
        fmt_row = ttk.Frame(opts)
        fmt_row.pack(fill=tk.X, pady=(6, 0))

        ttk.Label(fmt_row, text="Output format:").pack(side=tk.LEFT)
        saved_fmt = self._settings.get("format", "generic")
        saved_label = FORMAT_OPTIONS[0][0]
        for label, key in FORMAT_OPTIONS:
            if key == saved_fmt:
                saved_label = label
                break
        self.format_var = tk.StringVar(value=saved_label)
        self.format_combo = ttk.Combobox(
            fmt_row, textvariable=self.format_var,
            values=[label for label, _ in FORMAT_OPTIONS],
            state="readonly", width=45,
        )
        self.format_combo.set(saved_label)
        self.format_combo.pack(side=tk.LEFT, padx=(6, 0))

        # Password
        pw_row = ttk.Frame(opts)
        pw_row.pack(fill=tk.X, pady=(6, 0))

        ttk.Label(pw_row, text="PDF password:").pack(side=tk.LEFT)
        self.password_var = tk.StringVar(value="")
        self.password_entry = ttk.Entry(pw_row, textvariable=self.password_var, show="*", width=30)
        self.password_entry.pack(side=tk.LEFT, padx=(6, 0))
        ttk.Label(pw_row, text="(leave blank if not encrypted)", style="Small.TLabel").pack(
            side=tk.LEFT, padx=(6, 0)
        )

        # --- Feature checkboxes ---
        feat_row = ttk.Frame(opts)
        feat_row.pack(fill=tk.X, pady=(8, 0))

        ttk.Label(feat_row, text="Features:").pack(side=tk.LEFT)

        self.categorise_var = tk.BooleanVar(value=self._settings.get("categorise", True))
        self.chk_categorise = ttk.Checkbutton(
            feat_row, text="Auto-categorise", variable=self.categorise_var,
        )
        self.chk_categorise.pack(side=tk.LEFT, padx=(10, 0))

        self.smart_names_var = tk.BooleanVar(value=self._settings.get("smart_names", True))
        self.chk_smart_names = ttk.Checkbutton(
            feat_row, text="Smart filenames", variable=self.smart_names_var,
        )
        self.chk_smart_names.pack(side=tk.LEFT, padx=(10, 0))

        self.xlsx_var = tk.BooleanVar(value=self._settings.get("xlsx", False))
        self.chk_xlsx = ttk.Checkbutton(
            feat_row, text="Excel output", variable=self.xlsx_var,
        )
        self.chk_xlsx.pack(side=tk.LEFT, padx=(10, 0))

        self.open_folder_var = tk.BooleanVar(value=self._settings.get("open_folder", True))
        self.chk_open_folder = ttk.Checkbutton(
            feat_row, text="Open folder when done", variable=self.open_folder_var,
        )
        self.chk_open_folder.pack(side=tk.LEFT, padx=(10, 0))

        # --- Convert bar ---
        convert_bar = ttk.Frame(main)
        convert_bar.pack(fill=tk.X, pady=(10, 0))

        self.btn_convert = ttk.Button(
            convert_bar, text="   Convert   ",
            command=self._start_convert, style="Convert.TButton",
        )
        self.btn_convert.pack(side=tk.LEFT)

        self.btn_open_csv = ttk.Button(
            convert_bar, text="Open CSV Folder",
            command=self._open_csv_folder,
        )
        self.btn_open_csv.pack(side=tk.LEFT, padx=(10, 0))

        self.lbl_status = ttk.Label(convert_bar, text="Ready", style="Status.TLabel")
        self.lbl_status.pack(side=tk.RIGHT)

        # --- Progress bar ---
        self.progress = ttk.Progressbar(main, mode="determinate", length=200)
        self.progress.pack(fill=tk.X, pady=(6, 0))

        # --- Log output ---
        log_label = ttk.Label(main, text="Output Log:", font=(self._ui_font, 9, "bold"))
        log_label.pack(anchor=tk.W, pady=(8, 0))

        log_frame = ttk.Frame(main)
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(2, 0))

        log_scroll = ttk.Scrollbar(log_frame, orient=tk.VERTICAL)
        self.log_text = tk.Text(
            log_frame, height=10,
            font=(self._mono_font, 9),
            state=tk.DISABLED, wrap=tk.WORD,
            bg="#1e1e1e", fg="#d4d4d4",
            insertbackground="white",
            relief=tk.FLAT, borderwidth=1,
            yscrollcommand=log_scroll.set,
        )
        log_scroll.config(command=self.log_text.yview)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        log_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        # Colour tags
        self.log_text.tag_config("pass", foreground="#4caf50")
        self.log_text.tag_config("fail", foreground="#f44336")
        self.log_text.tag_config("warn", foreground="#ff9800")
        self.log_text.tag_config("info", foreground="#64b5f6")
        self.log_text.tag_config("heading", foreground="#ffffff",
                                  font=(self._mono_font, 9, "bold"))

    # ------------------------------------------------------------------
    # Settings persistence
    # ------------------------------------------------------------------

    def _persist_settings(self):
        """Save current checkbox/format states to settings.json."""
        settings = {
            "format": self._get_selected_format(),
            "categorise": self.categorise_var.get(),
            "smart_names": self.smart_names_var.get(),
            "xlsx": self.xlsx_var.get(),
            "open_folder": self.open_folder_var.get(),
            "output_dir": str(self.output_dir),
        }
        _save_settings(settings)

    # ------------------------------------------------------------------
    # Drag and drop
    # ------------------------------------------------------------------

    def _setup_drag_and_drop(self):
        """Wire up drag-and-drop on the drop zone and file list."""
        if not HAS_DND:
            self.drop_label.config(
                text="Drop PDF files here  (install tkinterdnd2 for drag-and-drop)",
                fg="#999999",
            )
            return

        for widget in (self.drop_frame, self.drop_label, self.file_listbox):
            widget.drop_target_register(tkinterdnd2.DND_FILES)
            widget.dnd_bind("<<Drop>>", self._handle_drop)
            widget.dnd_bind("<<DragEnter>>", self._on_drag_enter)
            widget.dnd_bind("<<DragLeave>>", self._on_drag_leave)

    def _handle_drop(self, event):
        """Handle files dropped onto the window."""
        files = self._parse_drop_data(event.data)
        added = 0
        for f in files:
            p = Path(f)
            if p.is_dir():
                for pdf in sorted(p.glob("*.pdf")):
                    if pdf not in self.pdf_files:
                        self.pdf_files.append(pdf)
                        added += 1
            elif p.suffix.lower() == ".pdf" and p not in self.pdf_files:
                self.pdf_files.append(p)
                added += 1
        if added:
            self._refresh_file_list()
        self._on_drag_leave(None)

    def _parse_drop_data(self, data):
        """Parse tkinterdnd2 drop data into file paths."""
        files = []
        i = 0
        while i < len(data):
            if data[i] == '{':
                end = data.index('}', i)
                files.append(data[i + 1:end])
                i = end + 2
            elif data[i] == ' ':
                i += 1
            else:
                space = data.find(' ', i)
                if space == -1:
                    files.append(data[i:])
                    break
                files.append(data[i:space])
                i = space + 1
        return files

    def _on_drag_enter(self, event):
        self.drop_frame.config(bg="#cce0ff")
        self.drop_label.config(bg="#cce0ff", text="Release to add files")

    def _on_drag_leave(self, event):
        self.drop_frame.config(bg="#e8f0fe")
        self.drop_label.config(bg="#e8f0fe", text="Drop PDF files here")

    # ------------------------------------------------------------------
    # File management
    # ------------------------------------------------------------------

    def _browse_files(self):
        files = filedialog.askopenfilenames(
            title="Select PDF Bank Statements",
            filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")],
            initialdir=str(SCRIPT_DIR / "pdfs"),
        )
        for f in files:
            p = Path(f)
            if p not in self.pdf_files:
                self.pdf_files.append(p)
        self._refresh_file_list()

    def _browse_folder(self):
        folder = filedialog.askdirectory(
            title="Select folder containing PDFs",
            initialdir=str(SCRIPT_DIR / "pdfs"),
        )
        if folder:
            for f in sorted(Path(folder).glob("*.pdf")):
                if f not in self.pdf_files:
                    self.pdf_files.append(f)
            self._refresh_file_list()

    def _remove_selected(self):
        selected = list(self.file_listbox.curselection())
        if not selected:
            return
        for idx in reversed(selected):
            if idx < len(self.pdf_files):
                del self.pdf_files[idx]
        self._refresh_file_list()

    def _clear_files(self):
        self.pdf_files.clear()
        self._refresh_file_list()

    def _refresh_file_list(self):
        self.file_listbox.delete(0, tk.END)
        for p in self.pdf_files:
            size_kb = p.stat().st_size / 1024 if p.exists() else 0
            self.file_listbox.insert(tk.END, f"  {p.name:<50s}  ({size_kb:.0f} KB)")
        count = len(self.pdf_files)
        self.lbl_count.config(text=f"{count} file{'s' if count != 1 else ''} selected")

    def _change_output_dir(self):
        folder = filedialog.askdirectory(
            title="Select output folder for CSV files",
            initialdir=str(self.output_dir),
        )
        if folder:
            self.output_dir = Path(folder)
            self.lbl_outdir.config(text=str(self.output_dir))

    def _open_csv_folder(self):
        path = str(self.output_dir)
        if sys.platform == "win32":
            os.startfile(path)
        elif sys.platform == "darwin":
            os.system(f'open "{path}"')
        else:
            os.system(f'xdg-open "{path}" 2>/dev/null &')

    def _get_selected_format(self):
        selected_label = self.format_combo.get()
        for label, key in FORMAT_OPTIONS:
            if label == selected_label:
                return key
        return "generic"

    # ------------------------------------------------------------------
    # Conversion
    # ------------------------------------------------------------------

    def _log(self, text, tag=None):
        def _insert():
            self.log_text.config(state=tk.NORMAL)
            if tag:
                self.log_text.insert(tk.END, text, tag)
            else:
                self.log_text.insert(tk.END, text)
            self.log_text.see(tk.END)
            self.log_text.config(state=tk.DISABLED)
        self.root.after(0, _insert)

    def _set_status(self, text, color="#888888"):
        def _update():
            self.lbl_status.config(text=text, foreground=color)
        self.root.after(0, _update)

    def _set_progress(self, value, maximum):
        def _update():
            self.progress.config(maximum=maximum, value=value)
        self.root.after(0, _update)

    def _set_converting(self, active):
        def _update():
            self.converting = active
            # ttk buttons use state differently
            for btn in (self.btn_convert, self.btn_open_csv):
                btn.state(["disabled"] if active else ["!disabled"])
            self.format_combo.config(state="disabled" if active else "readonly")
            self.password_entry.config(state="disabled" if active else "normal")
            # Disable/enable feature checkboxes
            for chk in (self.chk_categorise, self.chk_smart_names,
                        self.chk_xlsx, self.chk_open_folder):
                chk.state(["disabled"] if active else ["!disabled"])
        self.root.after(0, _update)

    def _start_convert(self):
        if not self.pdf_files:
            messagebox.showwarning(
                "No files selected",
                "Please select PDF bank statements first.\n\n"
                "Click 'Browse PDFs...' or 'Add Folder...' to add files.",
            )
            return

        if self.converting:
            return

        # Clear log
        self.log_text.config(state=tk.NORMAL)
        self.log_text.delete("1.0", tk.END)
        self.log_text.config(state=tk.DISABLED)

        # Reset progress
        self._set_progress(0, len(self.pdf_files))

        # Make output dir
        self.output_dir.mkdir(parents=True, exist_ok=True)

        thread = threading.Thread(target=self._run_conversion, daemon=True)
        thread.start()

    def _run_conversion(self):
        self._set_converting(True)

        files = list(self.pdf_files)
        total = len(files)
        success = 0
        total_txns = 0
        errors = 0
        fmt = self._get_selected_format()
        password = self.password_var.get().strip() or None
        do_categorise = self.categorise_var.get()
        do_smart_names = self.smart_names_var.get()
        do_xlsx = self.xlsx_var.get()
        do_open_folder = self.open_folder_var.get()

        self._log("=" * 58 + "\n", "heading")
        self._log("  PDF Bank Statement to CSV Converter\n", "heading")
        self._log(f"  Format: {fmt}  |  Files: {total}\n", "heading")
        self._log("=" * 58 + "\n\n", "heading")

        for i, pdf_path in enumerate(files, 1):
            self._set_status(f"Converting {i}/{total}: {pdf_path.name}...", color="#1976d2")
            self._set_progress(i - 1, total)

            self._log(f"[{i}/{total}] {pdf_path.name}\n", "info")

            if not pdf_path.exists():
                self._log("  ERROR: File not found\n\n", "fail")
                errors += 1
                continue

            # Handle password-protected PDFs
            actual_path = str(pdf_path)
            decrypted_tmp = None
            try:
                actual_path, decrypted_tmp = convert._decrypt_pdf_if_needed(
                    str(pdf_path), password=password
                )
                if decrypted_tmp:
                    self._log("  Decrypted password-protected PDF\n", "info")
            except (ValueError, Exception) as e:
                self._log(f"  ERROR: {e}\n\n", "fail")
                errors += 1
                continue

            try:
                # Detect bank
                parser_key = convert.detect_parser(actual_path)
                fmt_entry = next(
                    (f for f in convert.BANK_FORMATS if f['key'] == parser_key), None
                )
                bank_label = fmt_entry['label'] if fmt_entry else parser_key
                self._log(f"  Detected: {bank_label}\n")

                # Parse
                parse_fn = convert.PARSERS[parser_key]
                transactions, opening_bal = parse_fn(actual_path)

                if not transactions:
                    self._log("  WARNING: No transactions extracted\n\n", "warn")
                    errors += 1
                    continue

                # Auto-categorise
                if do_categorise:
                    transactions = convert._apply_categories(transactions)
                    categorised = sum(1 for t in transactions if t.category)
                    self._log(f"  Categorised: {categorised}/{len(transactions)} transactions\n")

                # Generate output filename
                output_name = convert._generate_output_filename(
                    str(pdf_path), parser_key,
                    transactions=transactions, smart=do_smart_names,
                )
                output_path = self.output_dir / output_name

                # Write CSV
                convert.write_csv(
                    transactions, str(output_path),
                    fmt=fmt, categorise=do_categorise,
                )

                # Write Excel if enabled
                if do_xlsx:
                    xlsx_name = Path(output_name).stem + '.xlsx'
                    xlsx_path = self.output_dir / xlsx_name
                    convert.write_xlsx(
                        transactions, str(xlsx_path),
                        fmt=fmt, categorise=do_categorise,
                    )
                    self._log(f"  Excel:        {xlsx_name}\n", "pass")

                # Summary stats
                total_deb = sum(t.amount for t in transactions if t.amount < 0)
                total_cred = sum(t.amount for t in transactions if t.amount > 0)
                net = total_deb + total_cred

                self._log(f"  Transactions: {len(transactions)}\n")
                self._log(f"  Period:       {transactions[0].date} to {transactions[-1].date}\n")
                self._log(f"  Debits:       R{total_deb:>13}\n")
                self._log(f"  Credits:      R{total_cred:>13}\n")
                self._log(f"  Net:          R{net:>13}\n")

                # Balance reconciliation
                bal_txns = [t for t in transactions if t.balance is not None]
                if len(bal_txns) >= 2:
                    mismatches = 0
                    for j in range(1, len(bal_txns)):
                        prev_bal = bal_txns[j - 1]
                        curr_bal = bal_txns[j]
                        expected = prev_bal.balance + curr_bal.amount
                        if abs(expected - curr_bal.balance) > Decimal("0.02"):
                            mismatches += 1

                    if mismatches == 0:
                        self._log(
                            f"  Balance:      PASS ({len(bal_txns)} reconciled)\n", "pass"
                        )
                    else:
                        self._log(
                            f"  Balance:      FAIL ({mismatches} mismatches)\n", "fail"
                        )
                else:
                    self._log("  Balance:      N/A (no balance data)\n")

                # Duplicate check
                dupes = convert._detect_duplicates(transactions)
                if dupes:
                    self._log(f"  Duplicates:   {len(dupes)} potential\n", "warn")

                # Category summary in log
                if do_categorise:
                    cat_summary = convert.generate_category_summary(transactions)
                    if cat_summary:
                        self._log(f"\n  Category Breakdown:\n", "info")
                        for row in cat_summary:
                            cat_name = row['category']
                            cnt = row['count']
                            net_val = row['net']
                            self._log(f"    {cat_name:<28} {cnt:>4} txns  R{net_val:>12}\n")

                self._log(f"  Saved:        {output_name}\n", "pass")
                success += 1
                total_txns += len(transactions)

            except Exception as e:
                self._log(f"  ERROR: {e}\n", "fail")
                errors += 1
            finally:
                if decrypted_tmp and os.path.exists(decrypted_tmp):
                    os.remove(decrypted_tmp)

            self._log("\n")

        self._set_progress(total, total)

        # Final summary
        self._log("=" * 58 + "\n", "heading")
        if errors == 0 and success > 0:
            self._log(
                f"  ALL DONE — {success} file{'s' if success != 1 else ''}"
                f" converted, {total_txns} transactions\n", "pass"
            )
            self._set_status(f"Done — {success} converted", color="#2e7d32")
        elif success > 0:
            self._log(f"  {success} converted, {errors} failed\n", "warn")
            self._set_status(f"{success} converted, {errors} failed", color="#e65100")
        else:
            self._log(f"  No files converted ({errors} error{'s' if errors != 1 else ''})\n", "fail")
            self._set_status("Conversion failed", color="#d32f2f")
        self._log("=" * 58 + "\n", "heading")

        self._set_converting(False)

        # Persist settings after successful conversion
        self._persist_settings()

        # Success popup + auto-open folder
        if success > 0:
            def _post_conversion():
                messagebox.showinfo(
                    "Conversion Complete",
                    f"{success} file{'s' if success != 1 else ''} converted, "
                    f"{total_txns} transactions processed.",
                )
                if do_open_folder:
                    self._open_csv_folder()
            self.root.after(100, _post_conversion)


def main():
    global HAS_DND
    # Use TkinterDnD root if available (enables drag-and-drop).
    # Falls back to plain Tk if the tkdnd native extension fails to load
    # (e.g. when tkinterdnd2 is compiled for Tcl 8.x but the runtime uses Tcl 9.x).
    if HAS_DND:
        try:
            root = tkinterdnd2.TkinterDnD.Tk()
        except (RuntimeError, Exception):
            HAS_DND = False
            root = tk.Tk()
    else:
        root = tk.Tk()

    # Window sizing and centering
    width, height = 800, 760
    root.update_idletasks()
    screen_w = root.winfo_screenwidth()
    screen_h = root.winfo_screenheight()
    x = (screen_w - width) // 2
    y = (screen_h - height) // 2
    root.geometry(f"{width}x{height}+{x}+{y}")
    root.minsize(650, 600)

    ConverterGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
