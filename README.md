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

### Install

1. Download `mfin.exe` from releases (or build it yourself, see below)
2. Run: `mfin.exe --install`
3. Right-click any PDF → "Extract Tables (mfin)"

Output appears in a `tabele/` folder next to the PDF.

### Uninstall

```
mfin.exe --uninstall
```

### Build from Source (Windows)

```bash
pip install -r requirements.txt -r requirements-build.txt
pyinstaller mfin.spec
```

The `.exe` is created at `dist/mfin.exe`.
