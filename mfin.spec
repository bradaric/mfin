# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for mfin Windows GUI
# Build with: pyinstaller mfin.spec

a = Analysis(
    ['windows_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'camelot', 'pdfplumber', 'pdfminer', 'pdfminer.high_level',
        'charset_normalizer', 'charset_normalizer.md__mypyc',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='mfin',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window
    icon=None,
)
