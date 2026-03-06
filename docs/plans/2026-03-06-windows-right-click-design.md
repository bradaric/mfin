# Windows Right-Click PDF Extraction

## Goal

Allow Windows users to right-click a PDF and extract fiscal tables into Excel files, without opening a command line. Distributed as a single `.exe` with no Python dependency.

## Architecture

A single `mfin.exe` (built with PyInstaller) serves three purposes:

- `mfin.exe --install` — registers the right-click context menu (per-user, no admin)
- `mfin.exe --uninstall` — removes the context menu entry
- `mfin.exe <path.pdf>` — processes the PDF with a tkinter progress window

## Right-Click Integration

Registry key: `HKEY_CURRENT_USER\Software\Classes\SystemFileAssociations\.pdf\shell\MfinExtract`

- Display name: "Extract Tables (mfin)"
- Command: `"<path-to-exe>" "%1"`
- Per-user scope (HKCU) — no admin/UAC prompt needed

## Progress Window

Small tkinter window (~400x300):

- Filename being processed
- Scrolling log area (mirrors current console output)
- "Close" button activates when done or on error

No console window — PyInstaller `--noconsole` flag.

## Output Location

Output next to the source PDF: `<pdf_dir>/tabele/<bilten_id>/*.xlsx`

## Files

| File | Change |
|------|--------|
| `extract_tables.py` | Refactor `process_pdf` to accept a progress callback instead of bare `print` |
| `windows_gui.py` | New — tkinter progress window, registry install/uninstall, entry point |
| `mfin.spec` | New — PyInstaller build config |

## Constraints

- All extraction logic stays unchanged
- Linux CLI usage stays unchanged
- The `-i`/`--interactive` TUI mode stays unchanged
- Single-file right-click only (no multi-select for now)
