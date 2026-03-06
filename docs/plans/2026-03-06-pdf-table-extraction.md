# PDF Table Extraction - Fiscal Section Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Extract all tables from Section II (ФИСКАЛНА КРЕТАЊА) of Serbian Ministry of Finance bulletins (PDF) into organized Excel files.

**Architecture:** A single Python script (`extract_tables.py`) driven by a JSON-like config that maps page numbers to table definitions. Uses `camelot-py` (stream mode) for extraction and `pandas`/`openpyxl` for Excel output. Multi-page horizontal tables are merged column-wise on row labels. Output organized as `tabele/<bilten-id>/` with one `.xlsx` per subsection, each table as a separate sheet.

**Tech Stack:** Python 3.12, camelot-py (stream), pandas, openpyxl, pdfplumber (for TOC validation)

---

## Verified Assumptions

- Both PDFs (Oct 2025, Nov 2025) have **identical structure**: 96 pages, same tables at same page numbers
- All tables in Section II live on pages 39-77
- Page structure is stable across issues (same TOC layout)

## Table Inventory (pages are 1-indexed)

### File: `00_fiskalna_kretanja.xlsx`
| Sheet | Table | Pages | Type |
|-------|-------|-------|------|
| Табела 1 | Консолидовани биланс државе 2005-2025 | 39, 40, 41 | 3-page horizontal merge |
| Табела 2 | Консолидовани биланс по нивоима власти | 42 | single page |

### File: `01_budzet_rs.xlsx`
| Sheet | Table | Pages | Type |
|-------|-------|-------|------|
| Табела 3 | Примања и издаци буџета РС 2005-2025 | 48, 49, 50 | 3-page horizontal merge |
| Табела 4 | Порески приходи | 51 | single page |
| Табела 5 | ПДВ и акцизе | 52 | single page |
| Табела 6 | Непорески приходи | 54 | single page |
| Табела 7 | Укупни издаци буџета РС | 55, 56 | 2-page horizontal merge |
| Табела 8 | Расходи за запослене | 57 | single page |
| Табела 9 | Отплата камата | 59 | single page |
| Табела 10 | Субвенције | 60 | single page |
| Табела 11 | Донације и трансфери | 61 | single page |

### File: `02_budzet_vojvodine.xlsx`
| Sheet | Table | Pages | Type |
|-------|-------|-------|------|
| Табела 1 | Примања буџета Војводине | 64 | single page |
| Табела 2 | Издаци буџета Војводине | 65 | single page |

### File: `03_budzet_opstina.xlsx`
| Sheet | Table | Pages | Type |
|-------|-------|-------|------|
| Табела 1 | Примања буџета општина и градова | 68 | single page |
| Табела 2 | Издаци буџета општина и градова | 69 | single page |

### File: `04_ooso.xlsx`
| Sheet | Table | Pages | Type |
|-------|-------|-------|------|
| Табела 1 | Примања РФПИО | 72 | single page |
| Табела 2 | Издаци РФПИО | 73 | single page |
| Табела 3 | Примања РФЗО | 74 | single page |
| Табела 4 | Издаци РФЗО | 75 | single page |
| Табела 5 | Примања НСЗ | 76 | single page |
| Табела 6 | Издаци НСЗ | 77 | single page |

**Total: 20 tables across 5 Excel files per PDF**

---

### Task 1: Project Setup

**Files:**
- Create: `requirements.txt`
- Create: `extract_tables.py` (skeleton)

**Step 1: Create requirements.txt**

```
camelot-py[base]
opencv-python-headless
pandas
openpyxl
pdfplumber
```

**Step 2: Create skeleton extract_tables.py with config**

The config defines all 20 tables with their page numbers, output file, sheet name, and merge type.

```python
#!/usr/bin/env python3
"""Extract fiscal tables from MFIN bilten PDFs into Excel files."""

import sys
import os
import re
import camelot
import pandas as pd

# Table definitions: each entry maps to one sheet in an output Excel file.
# Pages are 1-indexed. For multi-page horizontal tables, pages are listed in order.
TABLE_CONFIG = {
    "00_fiskalna_kretanja": [
        {"sheet": "Табела 1", "pages": [39, 40, 41], "merge": "horizontal"},
        {"sheet": "Табела 2", "pages": [42], "merge": None},
    ],
    "01_budzet_rs": [
        {"sheet": "Табела 3", "pages": [48, 49, 50], "merge": "horizontal"},
        {"sheet": "Табела 4", "pages": [51], "merge": None},
        {"sheet": "Табела 5", "pages": [52], "merge": None},
        {"sheet": "Табела 6", "pages": [54], "merge": None},
        {"sheet": "Табела 7", "pages": [55, 56], "merge": "horizontal"},
        {"sheet": "Табела 8", "pages": [57], "merge": None},
        {"sheet": "Табела 9", "pages": [59], "merge": None},
        {"sheet": "Табела 10", "pages": [60], "merge": None},
        {"sheet": "Табела 11", "pages": [61], "merge": None},
    ],
    "02_budzet_vojvodine": [
        {"sheet": "Табела 1", "pages": [64], "merge": None},
        {"sheet": "Табела 2", "pages": [65], "merge": None},
    ],
    "03_budzet_opstina": [
        {"sheet": "Табела 1", "pages": [68], "merge": None},
        {"sheet": "Табела 2", "pages": [69], "merge": None},
    ],
    "04_ooso": [
        {"sheet": "Табела 1", "pages": [72], "merge": None},
        {"sheet": "Табела 2", "pages": [73], "merge": None},
        {"sheet": "Табела 3", "pages": [74], "merge": None},
        {"sheet": "Табела 4", "pages": [75], "merge": None},
        {"sheet": "Табела 5", "pages": [76], "merge": None},
        {"sheet": "Табела 6", "pages": [77], "merge": None},
    ],
}
```

**Step 3: Verify skeleton runs**

Run: `python extract_tables.py --help` (or just import check)

**Step 4: Commit**

```bash
git init
git add requirements.txt extract_tables.py
git commit -m "feat: project skeleton with table config and requirements"
```

---

### Task 2: Single-Page Table Extraction

**Files:**
- Modify: `extract_tables.py`

**Step 1: Implement `extract_single_page(pdf_path, page_num)` function**

Uses camelot stream mode to extract the table from a single page. Returns a pandas DataFrame. Handles the common post-processing:
- Drop completely empty rows/columns
- Strip whitespace from all cells
- Fix year-label rows (e.g. "2023" appearing as a standalone row - merge it into the period column of adjacent rows)

```python
def extract_single_page(pdf_path, page_num):
    """Extract a table from a single PDF page using camelot stream mode."""
    tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
    if not tables:
        print(f"  WARNING: No table found on page {page_num}")
        return pd.DataFrame()
    # Take the largest table if multiple found
    df = max(tables, key=lambda t: t.df.size).df
    # Drop fully empty rows and columns
    df = df.dropna(how='all').reset_index(drop=True)
    df = df.loc[:, ~df.isna().all()]
    # Strip whitespace
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    # Drop rows that are fully empty strings
    df = df[~df.apply(lambda row: all(v == '' for v in row), axis=1)].reset_index(drop=True)
    return df
```

**Step 2: Test on a known single-page table (Војводина Табела 1, page 64)**

```python
# Quick test
df = extract_single_page('bilteni/2025-10-DpSCfe_6952453171db7.pdf', 64)
print(df.shape)
print(df.head(15))
```

Expected: ~68 rows x 10-12 cols with clean data matching the PDF visually.

**Step 3: Commit**

```bash
git add extract_tables.py
git commit -m "feat: single-page table extraction with camelot stream"
```

---

### Task 3: Multi-Page Horizontal Table Merging

**Files:**
- Modify: `extract_tables.py`

**Step 1: Implement `extract_horizontal_merge(pdf_path, page_nums)` function**

For tables like Табела 1 (Консолидовани биланс) that span 3 pages horizontally:
- Extract each page separately
- The first column(s) on each page are the row labels (same across pages)
- Data columns differ per page (different years)
- Merge by matching row labels column-wise

Strategy:
1. Extract each page as a DataFrame
2. On the first page, identify the label column(s) - typically column 0 (and sometimes column 1 for sub-labels)
3. On continuation pages, the same label column(s) appear - use them to align
4. Concatenate horizontally (join on row index after alignment)

```python
def extract_horizontal_merge(pdf_path, page_nums):
    """Extract a multi-page horizontal table and merge column-wise."""
    dfs = []
    for pg in page_nums:
        df = extract_single_page(pdf_path, pg)
        if not df.empty:
            dfs.append(df)
    if not dfs:
        return pd.DataFrame()
    if len(dfs) == 1:
        return dfs[0]
    # First page is the base. Subsequent pages share label column(s) but add data columns.
    # We concatenate by index position (row order matches across pages).
    base = dfs[0]
    for extra in dfs[1:]:
        # Skip label columns on continuation pages (typically first 1-2 cols that match base labels)
        # Find how many leading cols are labels by checking if they contain mostly non-numeric text
        skip = find_label_cols_count(extra)
        data_cols = extra.iloc[:, skip:]
        # Align row count to base (trim or pad)
        min_rows = min(len(base), len(data_cols))
        data_cols = data_cols.iloc[:min_rows].reset_index(drop=True)
        base = base.iloc[:min_rows].reset_index(drop=True)
        # Append columns
        new_col_start = base.shape[1]
        for i, col in enumerate(data_cols.columns):
            base[new_col_start + i] = data_cols[col].values
    return base


def find_label_cols_count(df):
    """Heuristic: count leading columns that are mostly non-numeric (label columns)."""
    count = 0
    for col in df.columns:
        vals = df[col].dropna().astype(str)
        numeric_count = vals.apply(lambda v: bool(re.match(r'^-?[\d,.]+$', v.replace(' ', '')))).sum()
        if numeric_count / max(len(vals), 1) < 0.5:
            count += 1
        else:
            break
    return max(count, 1)  # At least 1 label column
```

**Step 2: Test on Табела 1 (pages 39-41)**

```python
df = extract_horizontal_merge('bilteni/2025-10-DpSCfe_6952453171db7.pdf', [39, 40, 41])
print(f"Shape: {df.shape}")
print(df.head(5))
```

Expected: ~45 rows with columns spanning 2005-2025 plus index columns.

**Step 3: Test on Табела 3 (pages 48-50) and Табела 7 (pages 55-56)**

Verify horizontal merge works for different table structures.

**Step 4: Commit**

```bash
git add extract_tables.py
git commit -m "feat: multi-page horizontal table merging"
```

---

### Task 4: Main Processing Loop and Excel Output

**Files:**
- Modify: `extract_tables.py`

**Step 1: Implement main processing function**

```python
def derive_bilten_id(pdf_filename):
    """Extract bilten identifier like '2025-10' from filename."""
    match = re.search(r'(\d{4}-\d{2})', pdf_filename)
    return match.group(1) if match else os.path.splitext(pdf_filename)[0]


def process_pdf(pdf_path, output_dir):
    """Process a single PDF: extract all tables and write to Excel files."""
    bilten_id = derive_bilten_id(os.path.basename(pdf_path))
    out_dir = os.path.join(output_dir, bilten_id)
    os.makedirs(out_dir, exist_ok=True)

    print(f"\nProcessing: {pdf_path} -> {out_dir}/")

    for xlsx_name, table_defs in TABLE_CONFIG.items():
        xlsx_path = os.path.join(out_dir, f"{xlsx_name}.xlsx")
        print(f"\n  Writing: {xlsx_path}")

        with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
            for tdef in table_defs:
                sheet = tdef["sheet"]
                pages = tdef["pages"]
                merge = tdef["merge"]

                print(f"    {sheet} (pages {pages})...", end=" ")

                if merge == "horizontal":
                    df = extract_horizontal_merge(pdf_path, pages)
                else:
                    df = extract_single_page(pdf_path, pages[0])

                if df.empty:
                    print("EMPTY!")
                    continue

                # Truncate sheet name to 31 chars (Excel limit)
                safe_sheet = sheet[:31]
                df.to_excel(writer, sheet_name=safe_sheet, index=False, header=False)
                print(f"{df.shape[0]} rows x {df.shape[1]} cols")

    print(f"\nDone: {bilten_id}")
```

**Step 2: Implement CLI entry point**

```python
def main():
    import glob
    output_dir = os.path.join(os.path.dirname(__file__), "tabele")
    pdf_dir = os.path.join(os.path.dirname(__file__), "bilteni")

    pdfs = sorted(glob.glob(os.path.join(pdf_dir, "*.pdf")))
    if not pdfs:
        print(f"No PDFs found in {pdf_dir}")
        sys.exit(1)

    print(f"Found {len(pdfs)} PDF(s) in {pdf_dir}")
    for pdf_path in pdfs:
        process_pdf(pdf_path, output_dir)

    print("\nAll done.")


if __name__ == "__main__":
    main()
```

**Step 3: Run on October PDF only first (for speed)**

```bash
cd /home/sasa/workshop/mfin
.venv/bin/python extract_tables.py
```

Expected: Creates `tabele/2025-10/` with 5 xlsx files, each with correct sheets.

**Step 4: Spot-check a few Excel files**

Open and verify:
- `tabele/2025-10/02_budzet_vojvodine.xlsx` Табела 1 sheet should have ~68 data rows x 10 cols
- `tabele/2025-10/04_ooso.xlsx` should have 6 sheets
- `tabele/2025-10/00_fiskalna_kretanja.xlsx` Табела 1 should have wide columns (2005-2025 span)

**Step 5: Run on both PDFs**

Verify `tabele/2025-11/` is also created with same structure.

**Step 6: Commit**

```bash
git add extract_tables.py
git commit -m "feat: full extraction pipeline with Excel output"
```

---

### Task 5: Quality Check and Fix Edge Cases

**Files:**
- Modify: `extract_tables.py`

**Step 1: Visual spot-check**

Compare extracted Excel data against PDF values for:
- A few cells in Табела 1 (consolidated balance) - check 2005, 2015, 2025 columns
- Табела 3 row counts match
- Табела 2 (by government level) has all 12 column headers
- ООСО tables have monthly breakdowns through the expected month

**Step 2: Fix any misalignment or missing data issues found**

Common issues to watch for:
- Camelot splitting one table into multiple fragments
- Header rows being mixed with data
- Year-label rows (2023, 2024, 2025) causing row misalignment in horizontal merges
- Footnote rows being included at bottom

**Step 3: Final run on both PDFs**

```bash
.venv/bin/python extract_tables.py
```

**Step 4: Commit**

```bash
git add extract_tables.py
git commit -m "fix: edge cases in table extraction"
```

---

## Summary

| Task | What | Est. Complexity |
|------|------|----------------|
| 1 | Project setup + config | Simple |
| 2 | Single-page extraction | Medium |
| 3 | Horizontal merge for multi-page tables | Medium-Hard |
| 4 | Main loop + Excel output | Medium |
| 5 | QA + edge case fixes | Variable |

Output: `tabele/{2025-10,2025-11}/` each containing 5 `.xlsx` files with 20 total table sheets.
