# Windows Right-Click Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Let Windows users right-click a PDF to extract fiscal tables via a single `.exe` with a tkinter progress window.

**Architecture:** Refactor `extract_tables.py` to support a progress callback, create `windows_gui.py` as the Windows entry point (tkinter UI + registry install/uninstall), bundle with PyInstaller.

**Tech Stack:** Python 3.12, tkinter (stdlib), winreg (stdlib), PyInstaller

---

### Task 1: Refactor process_pdf to support a progress callback

Currently `process_pdf()` and helpers use bare `print()`. Add an optional `log` callback parameter so output can be redirected to a GUI.

**Files:**
- Modify: `extract_tables.py` (functions: `process_pdf`, `find_table_pages`)

**Step 1: Add `log` parameter to `process_pdf` and `find_table_pages`**

In `extract_tables.py`, change `process_pdf` signature and all its `print()` calls:

```python
def process_pdf(pdf_path, output_dir, log=None):
    """Process a single PDF: extract all tables and write to Excel files."""
    if log is None:
        log = print
    bilten_id = derive_bilten_id(os.path.basename(pdf_path))
    out_dir = os.path.join(output_dir, bilten_id)
    os.makedirs(out_dir, exist_ok=True)

    log(f"\nProcessing: {pdf_path} -> {out_dir}/")
    # ... replace all print(...) with log(...) in this function
```

Similarly in `find_table_pages`:

```python
def find_table_pages(page_texts, log=None):
    if log is None:
        log = print
    # ... replace print(f"  WARNING: ...") with log(f"  WARNING: ...")
```

And in `main()`, pass `log=print` explicitly (or leave default).

**Step 2: Verify CLI still works**

Run: `cd /home/sasa/workshop/mfin/.worktrees/windows-gui && .venv/bin/python extract_tables.py bilteni/*.pdf 2>&1 | head -30`

Expected: Same output as before — the refactor is purely additive.

**Step 3: Commit**

```bash
git add extract_tables.py
git commit -m "refactor: add log callback to process_pdf for GUI integration"
```

---

### Task 2: Create the tkinter progress window

**Files:**
- Create: `windows_gui.py`

**Step 1: Write `windows_gui.py` with the progress window**

```python
"""Windows GUI entry point for mfin PDF table extraction.

Usage:
    mfin.exe <path.pdf>       — process PDF with progress window
    mfin.exe --install        — register right-click context menu
    mfin.exe --uninstall      — remove right-click context menu
"""

import sys
import os
import threading
import tkinter as tk
from tkinter import scrolledtext

from extract_tables import process_pdf


class ProgressWindow:
    """A small tkinter window that shows extraction progress."""

    def __init__(self, pdf_path):
        self.pdf_path = pdf_path
        self.root = tk.Tk()
        self.root.title("mfin — Extracting Tables")
        self.root.geometry("500x350")
        self.root.resizable(True, True)

        # Filename label
        label = tk.Label(
            self.root,
            text=f"Processing: {os.path.basename(pdf_path)}",
            anchor="w",
            padx=10,
            pady=5,
        )
        label.pack(fill="x")

        # Scrolling log area
        self.log_area = scrolledtext.ScrolledText(
            self.root, wrap="word", state="disabled", height=15
        )
        self.log_area.pack(fill="both", expand=True, padx=10, pady=(0, 5))

        # Close button (disabled until done)
        self.close_btn = tk.Button(
            self.root, text="Close", state="disabled", command=self.root.destroy
        )
        self.close_btn.pack(pady=(0, 10))

        # Handle window close via X button
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self._done = False

    def log(self, message):
        """Thread-safe logging to the text area."""
        self.root.after(0, self._append_log, message)

    def _append_log(self, message):
        self.log_area.config(state="normal")
        self.log_area.insert("end", message + "\n")
        self.log_area.see("end")
        self.log_area.config(state="disabled")

    def _on_close(self):
        if self._done:
            self.root.destroy()

    def _mark_done(self, success):
        self._done = True
        self.close_btn.config(state="normal")
        if success:
            self.log("\n--- Done! ---")
        else:
            self.log("\n--- Failed (see errors above) ---")

    def run(self):
        """Start processing in a background thread, run the GUI main loop."""
        thread = threading.Thread(target=self._process, daemon=True)
        thread.start()
        self.root.mainloop()

    def _process(self):
        try:
            output_dir = os.path.join(os.path.dirname(self.pdf_path), "tabele")
            process_pdf(self.pdf_path, output_dir, log=self.log)
            self.root.after(0, self._mark_done, True)
        except Exception as e:
            self.log(f"\nERROR: {e}")
            self.root.after(0, self._mark_done, False)


def run_gui(pdf_path):
    """Launch the progress window for a single PDF."""
    window = ProgressWindow(pdf_path)
    window.run()
```

**Step 2: Test the GUI manually on Linux**

Run: `.venv/bin/python -c "from windows_gui import ProgressWindow; print('import OK')"`

Expected: `import OK` (verifies no syntax/import errors). Full GUI test requires a PDF but the import validates structure.

**Step 3: Commit**

```bash
git add windows_gui.py
git commit -m "feat: add tkinter progress window for Windows GUI"
```

---

### Task 3: Add registry install/uninstall to windows_gui.py

**Files:**
- Modify: `windows_gui.py`

**Step 1: Add registry functions**

Append to `windows_gui.py`:

```python
REGISTRY_PATH = r"Software\Classes\SystemFileAssociations\.pdf\shell\MfinExtract"


def install_context_menu():
    """Register the right-click context menu entry for PDF files (per-user)."""
    if sys.platform != "win32":
        print("Context menu installation is only supported on Windows.")
        sys.exit(1)

    import winreg

    exe_path = os.path.abspath(sys.executable)
    # When frozen by PyInstaller, sys.executable is the .exe itself
    # When running as script, point to this script via pythonw
    if not getattr(sys, 'frozen', False):
        exe_path = f'"{sys.executable}" "{os.path.abspath(__file__)}"'
    else:
        exe_path = f'"{exe_path}"'

    try:
        # Create shell key with display name
        key = winreg.CreateKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH)
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, "Extract Tables (mfin)")
        winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, "")
        winreg.CloseKey(key)

        # Create command subkey
        cmd_key = winreg.CreateKey(
            winreg.HKEY_CURRENT_USER, REGISTRY_PATH + r"\command"
        )
        winreg.SetValueEx(cmd_key, "", 0, winreg.REG_SZ, f'{exe_path} "%1"')
        winreg.CloseKey(cmd_key)

        print("Context menu installed successfully.")
        print("Right-click any PDF to see 'Extract Tables (mfin)'.")
    except OSError as e:
        print(f"Failed to install context menu: {e}")
        sys.exit(1)


def uninstall_context_menu():
    """Remove the right-click context menu entry."""
    if sys.platform != "win32":
        print("Context menu removal is only supported on Windows.")
        sys.exit(1)

    import winreg

    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH + r"\command")
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, REGISTRY_PATH)
        print("Context menu removed successfully.")
    except FileNotFoundError:
        print("Context menu entry not found (already removed?).")
    except OSError as e:
        print(f"Failed to remove context menu: {e}")
        sys.exit(1)
```

**Step 2: Add `main()` entry point**

Append to `windows_gui.py`:

```python
def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    arg = sys.argv[1]

    if arg == "--install":
        install_context_menu()
    elif arg == "--uninstall":
        uninstall_context_menu()
    elif os.path.isfile(arg) and arg.lower().endswith(".pdf"):
        run_gui(arg)
    else:
        print(f"Unknown argument or file not found: {arg}")
        sys.exit(1)


if __name__ == "__main__":
    main()
```

**Step 3: Verify import still works**

Run: `.venv/bin/python -c "from windows_gui import main; print('OK')"`

Expected: `OK`

**Step 4: Commit**

```bash
git add windows_gui.py
git commit -m "feat: add registry install/uninstall for Windows context menu"
```

---

### Task 4: Add PyInstaller build configuration

**Files:**
- Create: `mfin.spec`
- Modify: `requirements.txt` (add pyinstaller)

**Step 1: Add pyinstaller to requirements**

Add a `requirements-build.txt` (keeps build deps separate from runtime):

```
pyinstaller
```

**Step 2: Install pyinstaller**

Run: `.venv/bin/pip install pyinstaller`

**Step 3: Create `mfin.spec`**

```python
# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for mfin Windows GUI
# Build with: pyinstaller mfin.spec

a = Analysis(
    ['windows_gui.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=['camelot', 'pdfplumber', 'pdfminer', 'pdfminer.high_level'],
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
```

**Step 4: Verify spec is valid**

Run: `.venv/bin/python -c "exec(open('mfin.spec').read()); print('Spec OK')"`

Expected: `Spec OK` (it's valid Python)

**Step 5: Commit**

```bash
git add mfin.spec requirements-build.txt
git commit -m "feat: add PyInstaller build config for Windows .exe"
```

---

### Task 5: Add build instructions to README

**Files:**
- Modify: `README.md`

**Step 1: Update README with Windows build/install instructions**

Add a Windows section to README.md covering:
- How to build the .exe: `pip install -r requirements-build.txt && pyinstaller mfin.spec`
- How to install the context menu: `mfin.exe --install`
- How to uninstall: `mfin.exe --uninstall`
- How it works: right-click PDF → "Extract Tables (mfin)" → output in `tabele/` next to PDF

**Step 2: Commit**

```bash
git add README.md
git commit -m "docs: add Windows build and install instructions"
```

---

### Task 6: End-to-end verification

**Step 1: Verify the CLI refactor didn't break anything**

Run: `.venv/bin/python extract_tables.py bilteni/*.pdf 2>&1 | head -40`

Expected: Same extraction output as before.

**Step 2: Verify the GUI module loads cleanly**

Run: `.venv/bin/python -c "from windows_gui import ProgressWindow, install_context_menu, uninstall_context_menu, main; print('All imports OK')"`

Expected: `All imports OK`

**Step 3: Verify PyInstaller can analyze the module (dry run)**

Run: `.venv/bin/pyinstaller --noconfirm mfin.spec 2>&1 | tail -10`

Note: This builds a Linux binary (not useful as .exe) but validates that PyInstaller can resolve all imports and the spec is correct. The actual Windows .exe must be built on Windows.

**Step 4: Commit any fixes if needed**
