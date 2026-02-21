# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec for Bank Statement PDF to CSV Converter.

Build with:  pyinstaller converter.spec
Output:      dist/BankStatementConverter  (or .exe on Windows)
"""
import sys

block_cipher = None
is_mac = sys.platform == 'darwin'

# Collect optional package data before Analysis
extra_datas = []
extra_binaries = []
extra_hiddenimports = []

try:
    from PyInstaller.utils.hooks import collect_all
    datas, binaries, hiddenimports = collect_all('tkinterdnd2')
    extra_datas += datas
    extra_binaries += binaries
    extra_hiddenimports += hiddenimports
except Exception:
    print("NOTE: tkinterdnd2 not found — drag-and-drop will be disabled")

try:
    from PyInstaller.utils.hooks import collect_data_files
    extra_datas += collect_data_files('ttkthemes')
except Exception:
    print("NOTE: ttkthemes not found — will use built-in themes")

try:
    datas, binaries, hiddenimports = collect_all('openpyxl')
    extra_datas += datas
    extra_binaries += binaries
    extra_hiddenimports += hiddenimports
except Exception:
    print("NOTE: openpyxl not found — Excel output will be disabled")

try:
    datas, binaries, hiddenimports = collect_all('pikepdf')
    extra_datas += datas
    extra_binaries += binaries
    extra_hiddenimports += hiddenimports
except Exception:
    print("NOTE: pikepdf not found — password-protected PDFs will be disabled")

a = Analysis(
    ['converter_gui.py', 'convert.py'],
    pathex=['.'],
    binaries=extra_binaries,
    datas=[('VERSION', '.'), ('icon.png', '.'), ('icon_small.png', '.')] + extra_datas,
    hiddenimports=[
        'convert',
        'pdfplumber',
        'pdfminer',
        'pdfminer.high_level',
        'pdfminer.layout',
        'pdfminer.pdfparser',
        'pdfminer.pdfdocument',
        'pdfminer.pdfpage',
        'pdfminer.pdfinterp',
        'pdfminer.converter',
        'pdfminer.utils',
        'charset_normalizer',
        'cryptography',
        'pikepdf',
        'ttkthemes',
        'tkinterdnd2',
        'pytesseract',
        'PIL',
        'PIL.Image',
        'PIL.ImageEnhance',
        'PIL.ImageTk',
        'PIL._tkinter_finder',
        'openpyxl',
    ] + extra_hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'scipy',
        'pandas',
        'pytest',
        'IPython',
        'jupyter',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='BankStatementConverter',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='icon.ico',
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='BankStatementConverter',
)

# macOS: create a .app bundle (wraps COLLECT so all files are inside)
if is_mac:
    app = BUNDLE(
        coll,
        name='BankStatementConverter.app',
        icon='icon.png',
        bundle_identifier='za.co.bankconverter',
        info_plist={
            'CFBundleDisplayName': 'BankStatementConverter',
            'CFBundleShortVersionString': '2.0.0',
            'NSHighResolutionCapable': True,
        },
    )
