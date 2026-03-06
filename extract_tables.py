#!/usr/bin/env python3
"""Extract fiscal tables from MFIN bilten PDFs into Excel files."""

import sys
import os
import re
import glob
import curses

import camelot
import pandas as pd
import pdfplumber


# Table signatures used to find pages dynamically.
# Each entry: (title_pattern, is_continuation_pattern)
# A "continuation" page has "наставак" in the title and belongs to the previous table.
TABLE_PATTERNS = {
    "00_fiskalna_kretanja": [
        {"id": "Табела 1", "pattern": r"Табела 1[:\.]?\s*Консолидовани биланс државе у периоду"},
        {"id": "Табела 2", "pattern": r"Табела 2\.?\s*Консолидовани биланс државе по нивоима"},
    ],
    "01_budzet_rs": [
        {"id": "Табела 3", "pattern": r"[TТ]абела 3\.?\s*Примања и издаци буџета"},
        {"id": "Табела 4", "pattern": r"[TТ]абела 4\.?\s*Порески приходи"},
        {"id": "Табела 5", "pattern": r"[TТ]абела 5\.?\s*Порез на додату вредност"},
        {"id": "Табела 6", "pattern": r"[TТ]абела 6\.?\s*Непорески приходи"},
        {"id": "Табела 7", "pattern": r"[TТ]абела 7\.?\s*Укупни издаци буџета"},
        {"id": "Табела 8", "pattern": r"[TТ]абела 8\.?\s*Укупни расходи за запослене"},
        {"id": "Табела 9", "pattern": r"[TТ]абела 9\.?\s*Расходи по основу отплате камата"},
        {"id": "Табела 10", "pattern": r"[TТ]абела 10\.?\s*Субвенције из буџета"},
        {"id": "Табела 11", "pattern": r"[TТ]абела 11\.?\s*Донације и трансфери из буџета"},
    ],
    "02_budzet_vojvodine": [
        {"id": "Табела 1", "pattern": r"Табела 1\.?\s*Примања буџета Војводине"},
        {"id": "Табела 2", "pattern": r"Табела 2\.?\s*Издаци буџета Војводине"},
    ],
    "03_budzet_opstina": [
        {"id": "Табела 1", "pattern": r"Табела 1\.?\s*Примања буџета општина"},
        {"id": "Табела 2", "pattern": r"Табела 2\.?\s*Издаци буџета општина"},
    ],
    "04_ooso": [
        {"id": "Табела 1", "pattern": r"Табела 1\.?\s*Примања РФПИО"},
        {"id": "Табела 2", "pattern": r"Табела 2\.?\s*Издаци РФПИО"},
        {"id": "Табела 3", "pattern": r"Табела 3\.?\s*Примања Републичког фонда за здравствено"},
        {"id": "Табела 4", "pattern": r"Табела 4\.?\s*Издаци Републичког фонда за здравствено"},
        {"id": "Табела 5", "pattern": r"Табела 5\.?\s*Примања Националне службе"},
        {"id": "Табела 6", "pattern": r"Табела 6\.?\s*Издаци Националне службе"},
    ],
}

# Tables that can span multiple pages horizontally (continuation pages have "наставак" in title)
MULTI_PAGE_TABLES = {
    ("00_fiskalna_kretanja", "Табела 1"),
    ("01_budzet_rs", "Табела 3"),
    ("01_budzet_rs", "Табела 7"),
}


def scan_pages(pdf_path):
    """Scan PDF and return mapping of page numbers to their text content (first 200 chars)."""
    page_texts = {}
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = (page.extract_text() or '')[:300]
            page_texts[i + 1] = text  # 1-indexed
    return page_texts


def find_table_pages(page_texts):
    """Find which pages contain which tables. Returns dict: (xlsx_name, table_id) -> [page_numbers]."""
    table_pages = {}

    for xlsx_name, table_defs in TABLE_PATTERNS.items():
        for tdef in table_defs:
            tid = tdef["id"]
            pattern = tdef["pattern"]
            pages = []
            for pg_num, text in sorted(page_texts.items()):
                # Normalize text for matching
                text_norm = text.replace('\n', ' ')
                if re.search(pattern, text_norm):
                    pages.append(pg_num)
            if pages:
                table_pages[(xlsx_name, tid)] = pages
            else:
                print(f"  WARNING: Could not find {xlsx_name}/{tid}")

    return table_pages


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
    # Collapse multi-line rows split by camelot
    label_cols = find_label_cols_count(df)
    df = collapse_multiline_rows(df, label_cols=label_cols)
    # Split merged year+month labels into separate columns for consistency
    df, _ = split_year_month_column(df, label_cols)
    return df


def _is_new_item_label(label):
    """Check if a label looks like the start of a new table row (not a continuation)."""
    if re.match(r'^\d{4}(\s|$)', label):  # Year like "2023"
        return True
    if re.match(r'^\d+\.', label):  # Numbered item like "1." or "1.5"
        return True
    if re.match(r'^[IVX]+[\s.]', label):  # Roman numeral section like "III "
        return True
    return False


def _should_merge_next(acc, next_row, label_cols):
    """Determine if next_row is a continuation of acc (split by camelot).

    Camelot splits wrapped cells into multiple rows, distributing data values
    across them based on vertical position. The key signal is that the rows
    have complementary (non-overlapping) data columns.
    """
    next_label = str(next_row[0]).strip()

    acc_data = acc[label_cols:]
    next_data = next_row[label_cols:]

    acc_empty = [str(v).strip() == '' for v in acc_data]
    next_empty = [str(v).strip() == '' for v in next_data]

    acc_filled = sum(1 for e in acc_empty if not e)
    next_filled = sum(1 for e in next_empty if not e)
    overlap = sum(1 for ae, ne in zip(acc_empty, next_empty) if not ae and not ne)

    # If accumulated row has no data at all, it's an incomplete label row.
    # Merge with next unless next is clearly a new item.
    if acc_filled == 0:
        if next_label and _is_new_item_label(next_label):
            return False
        return True

    # Too much overlap means these are separate rows
    if overlap > 2:
        return False

    # Empty label with some data: overflow from previous row
    if next_label == '' and next_filled > 0:
        return True

    # If next label starts a new numbered/sectioned item, don't merge
    if next_label and _is_new_item_label(next_label):
        return False

    # Continuation label with no data: pure label wrap
    if next_label and next_filled == 0:
        return True

    # Continuation label with sparse complementary data
    if next_label and next_filled <= 2 and overlap <= 1:
        return True

    return False


def _merge_rows(acc, next_row, label_cols):
    """Merge next_row into acc: combine labels, prefer non-empty data values."""
    merged = list(acc)

    # Combine labels
    next_label = str(next_row[0]).strip()
    if next_label:
        curr_label = str(merged[0]).strip()
        merged[0] = (curr_label + ' ' + next_label) if curr_label else next_label

    # Merge data: fill in empty cells from next_row
    for c in range(label_cols, min(len(merged), len(next_row))):
        if str(merged[c]).strip() == '' and str(next_row[c]).strip() != '':
            merged[c] = next_row[c]

    return merged


def collapse_multiline_rows(df, label_cols=1):
    """Collapse rows where camelot split a multi-line cell into separate rows.

    Handles multiple patterns:
    - Label-only row + data-only row + optional label continuation
    - Row with partial data + overflow row with complementary data
    - Any combination where adjacent rows have non-overlapping data columns
    """
    if df.empty or len(df) < 2:
        return df

    rows = [list(r) for r in df.itertuples(index=False, name=None)]
    result = []
    i = 0
    while i < len(rows):
        acc = list(rows[i])
        i += 1

        # Greedily merge following rows that belong to the same logical row
        while i < len(rows):
            if _should_merge_next(acc, rows[i], label_cols):
                acc = _merge_rows(acc, rows[i], label_cols)
                i += 1
            else:
                break

        result.append(acc)

    return pd.DataFrame(result, columns=df.columns).reset_index(drop=True)


MONTH_NAMES = (
    'Јануар', 'Фебруар', 'Март', 'Maрт', 'Април', 'Мај', 'Јун',
    'Јул', 'Август', 'Септембар', 'Октобар', 'Новембар', 'Децембар',
)
_YEAR_MONTH_RE = re.compile(
    r'^(\d{4})\s+(' + '|'.join(MONTH_NAMES) + r'|Укупно)(.*)', re.IGNORECASE
)


def split_year_month_column(df, label_cols):
    """Split merged 'YYYY Month' values into separate year + month columns.

    Handles two cases:
    1. Col 0 has merged values (e.g. "2024 Јун") — inserts a new year column.
    2. Col 0 is already a year column but col 1 has merged values
       (e.g. "2005 Укупно") — splits year into col 0, keeps rest in col 1.
    """
    if df.empty or df.shape[1] < 2:
        return df, label_cols

    col0 = df.iloc[:, 0].astype(str)
    col0_has_merged = col0.apply(lambda v: bool(_YEAR_MONTH_RE.match(v.strip()))).any()

    if col0_has_merged:
        # Case 1: col 0 has merged year+month — insert a new year column
        years = []
        labels = []
        for val in col0:
            m = _YEAR_MONTH_RE.match(val.strip())
            if m:
                years.append(m.group(1))
                labels.append(m.group(2) + m.group(3))
            else:
                years.append('')
                labels.append(val)

        new_df = df.copy()
        new_df.iloc[:, 0] = labels
        new_df.insert(0, 'year_col', years)
        new_df.columns = range(len(new_df.columns))
        return new_df, label_cols + 1

    # Case 2: col 1 has merged year+month while col 0 is empty for those rows
    col1 = df.iloc[:, 1].astype(str)
    col1_has_merged = col1.apply(lambda v: bool(_YEAR_MONTH_RE.match(v.strip()))).any()

    if col1_has_merged:
        new_df = df.copy()
        for idx, val in col1.items():
            m = _YEAR_MONTH_RE.match(val.strip())
            if m and str(new_df.iloc[idx, 0]).strip() == '':
                new_df.iloc[idx, 0] = m.group(1)
                new_df.iloc[idx, 1] = m.group(2) + m.group(3)
        return new_df, label_cols

    return df, label_cols


def find_label_cols_count(df):
    """Heuristic: count leading columns that are mostly non-numeric (label columns)."""
    count = 0
    for col in df.columns:
        vals = df[col].dropna().astype(str)
        vals = vals[vals != '']
        numeric_count = vals.apply(
            lambda v: bool(re.match(r'^-?[\d,.\s]+$', v.replace(' ', '')))
        ).sum()
        if numeric_count / max(len(vals), 1) < 0.5:
            count += 1
        else:
            break
    return max(count, 1)  # At least 1 label column


def _find_alignment_offset(base, extra):
    """Find row offset to align data rows between base and continuation page.

    Returns offset such that extra.iloc[offset:] aligns with base data rows.
    Looks for the first substantive row label in base and finds it in extra.
    """
    for i in range(min(10, len(base))):
        label = str(base.iloc[i, 0]).strip()
        if label and len(label) > 5 and not label.lower().startswith('табела'):
            target = label[:20]
            for j in range(min(10, len(extra))):
                extra_label = str(extra.iloc[j, 0]).strip()
                if extra_label.startswith(target):
                    return j - i
    return 0


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
    base = dfs[0]
    for extra in dfs[1:]:
        # Skip label columns on continuation pages
        skip = find_label_cols_count(extra)

        # Align rows: continuation pages may have different header row counts
        offset = _find_alignment_offset(base, extra)
        if offset > 0:
            # Extra has more header rows - trim its top to align
            data_cols = extra.iloc[offset:, skip:]
        elif offset < 0:
            # Base has more header rows - pad extra with empty rows at top
            data_cols = extra.iloc[:, skip:]
            padding = pd.DataFrame(
                [[''] * data_cols.shape[1]] * abs(offset),
                columns=data_cols.columns,
            )
            data_cols = pd.concat([padding, data_cols], ignore_index=True)
        else:
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


def pick_files_tui(pdf_paths):
    """Show a terminal UI for selecting PDF files to process.

    Arrow keys to navigate, Space to toggle, Enter to confirm, 'a' to toggle all, 'q' to quit.
    Returns list of selected file paths.
    """
    labels = [os.path.basename(p) for p in pdf_paths]
    selected = [False] * len(labels)

    def draw(stdscr):
        curses.curs_set(0)
        curses.use_default_colors()
        curses.init_pair(1, curses.COLOR_BLACK, curses.COLOR_CYAN)
        current = 0

        while True:
            stdscr.clear()
            h, w = stdscr.getmaxyx()
            title = "Select PDF files to process (Space=toggle, a=all, Enter=confirm, q=quit)"
            stdscr.addnstr(0, 0, title, w - 1, curses.A_BOLD)

            for i, label in enumerate(labels):
                if i + 2 >= h - 1:
                    stdscr.addnstr(h - 1, 0, f"  ... and {len(labels) - i} more (scroll down)", w - 1)
                    break
                marker = "[x]" if selected[i] else "[ ]"
                line = f"  {marker} {label}"
                attr = curses.color_pair(1) | curses.A_BOLD if i == current else 0
                stdscr.addnstr(i + 2, 0, line, w - 1, attr)

            count = sum(selected)
            footer = f"  {count} file(s) selected"
            if 2 + len(labels) + 1 < h:
                stdscr.addnstr(2 + len(labels) + 1, 0, footer, w - 1)
            stdscr.refresh()

            key = stdscr.getch()
            if key == curses.KEY_UP and current > 0:
                current -= 1
            elif key == curses.KEY_DOWN and current < len(labels) - 1:
                current += 1
            elif key == ord(' '):
                selected[current] = not selected[current]
            elif key == ord('a'):
                toggle = not all(selected)
                for i in range(len(selected)):
                    selected[i] = toggle
            elif key in (curses.KEY_ENTER, 10, 13):
                return [pdf_paths[i] for i, s in enumerate(selected) if s]
            elif key == ord('q'):
                return []

    return curses.wrapper(draw)


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

    # Step 1: Scan pages to find table locations
    print("  Scanning pages...")
    page_texts = scan_pages(pdf_path)
    table_pages = find_table_pages(page_texts)

    # Step 2: Extract and write
    for xlsx_name, table_defs in TABLE_PATTERNS.items():
        xlsx_path = os.path.join(out_dir, f"{xlsx_name}.xlsx")
        print(f"\n  Writing: {xlsx_path}")

        with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
            for tdef in table_defs:
                tid = tdef["id"]
                key = (xlsx_name, tid)
                pages = table_pages.get(key, [])

                if not pages:
                    print(f"    {tid}: SKIPPED (not found)")
                    continue

                is_multi = key in MULTI_PAGE_TABLES
                print(f"    {tid} (pages {pages}, {'horizontal merge' if is_multi and len(pages) > 1 else 'single'})...", end=" ")

                if is_multi and len(pages) > 1:
                    df = extract_horizontal_merge(pdf_path, pages)
                else:
                    # For single-page tables, or multi-page tables with only 1 page found
                    df = extract_single_page(pdf_path, pages[0])

                if df.empty:
                    print("EMPTY!")
                    continue

                safe_sheet = tid[:31]
                df.to_excel(writer, sheet_name=safe_sheet, index=False, header=False)
                print(f"{df.shape[0]} rows x {df.shape[1]} cols")

    print(f"\nDone: {bilten_id}")


def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, "tabele")
    pdf_dir = os.path.join(script_dir, "bilteni")

    interactive = '-i' in sys.argv or '--interactive' in sys.argv
    args = [a for a in sys.argv[1:] if a not in ('-i', '--interactive')]

    if args:
        # Process specific file(s) passed as arguments
        pdfs = args
    else:
        # Process all PDFs in bilteni/
        pdfs = sorted(glob.glob(os.path.join(pdf_dir, "*.pdf")))

    if not pdfs:
        print(f"No PDFs found in {pdf_dir}")
        sys.exit(1)

    if interactive:
        pdfs = pick_files_tui(pdfs)
        if not pdfs:
            print("No files selected.")
            sys.exit(0)

    print(f"Processing {len(pdfs)} PDF(s)")
    for pdf_path in pdfs:
        process_pdf(pdf_path, output_dir)

    print("\nAll done.")


if __name__ == "__main__":
    main()
