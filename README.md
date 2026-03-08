# mfin

Extract fiscal tables from Serbian Ministry of Finance (MFIN) bilten PDFs into Excel files.

## Requirements

- Python 3.10+
- Dependencies: `pip install -r requirements.txt`

## Usage (Linux/macOS CLI)

Process all PDFs in `bilteni/`:

```bash
python extract_tables.py
```

Process specific files:

```bash
python extract_tables.py bilteni/2025-10-*.pdf
```

Interactive file picker:

```bash
python extract_tables.py -i
```

Output goes to `tabele/<bilten-id>/*.xlsx`.

## Windows — Right-Click Integration

Right-click any PDF on Windows → "Extract Tables (mfin)" → Excel files appear
next to the PDF. No Python or command line needed on the target machine.

### Step 1: Build the .exe (one-time, on a Windows machine with Python)

You need one Windows machine with Python 3.10+ to build the executable.
The build bundles Python and all dependencies into a single standalone file.

```bash
pip install -r requirements.txt -r requirements-build.txt
pyinstaller mfin.spec
```

The `requirements-build.txt` forces `charset-normalizer` to install as pure
Python (no mypyc compiled extensions), which avoids a known PyInstaller
bundling issue with mypyc hash-named modules. If you see
`No module named '...__mypyc'` at runtime, reinstall with:

```bash
pip install charset-normalizer --no-binary charset-normalizer --force-reinstall
```

This creates `dist\mfin.exe` (~50-100MB). The build machine and target machine
should both be 64-bit Windows 10 or 11.

### Step 2: Install on any Windows machine (no Python needed)

Copy `dist\mfin.exe` to the target machine (USB, network share, etc.), then
double-click it. A setup window appears with **Install** and **Uninstall**
buttons. Click **Install** to register the right-click menu.

No admin privileges required — it installs for the current user only.

**Windows 11 note:** On Windows 11 the menu item is hidden behind "Show more
options" by default. The setup dialog has a checkbox (enabled by default on
Win11) that restores the classic full context menu so the item is visible
directly. This applies system-wide to all right-click menus and restarts
Explorer.

### Step 3: Use it

Right-click any `.pdf` file → "Extract Tables (mfin)".

A progress window shows extraction status. Output goes to a `tabele\` folder
next to the PDF.

### Uninstall

Double-click `mfin.exe` again and click **Uninstall**.
