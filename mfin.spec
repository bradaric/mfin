# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for mfin Windows GUI
# Build with: pyinstaller mfin.spec

import os

# Include the shell extension DLL if it has been built.
_shell_ext = os.path.join('shell_extension', 'target', 'release', 'mfin_shell.dll')
_extra_binaries = [(_shell_ext, '.')] if os.path.isfile(_shell_ext) else []

a = Analysis(
    ['windows_gui.py'],
    pathex=[],
    binaries=_extra_binaries,
    datas=[],
    hiddenimports=[
        'camelot', 'pdfplumber', 'pdfminer', 'pdfminer.high_level',
        'charset_normalizer',
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
