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


def find_table_pages(page_texts, log=None):
    """Find which pages contain which tables. Returns dict: (xlsx_name, table_id) -> [page_numbers]."""
    if log is None:
        log = print
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
                log(f"  WARNING: Could not find {xlsx_name}/{tid}")

    return table_pages


def _reconstruct_headers(camelot_table, pdf_path, page_num):
    """Reconstruct header cells using pdfplumber word positions.

    Camelot stream mode often misses text at the top of header cells.
    This uses pdfplumber to extract all words in the header area and
    maps them to camelot's column positions to build complete headers.
    """
    ct = camelot_table
    df = ct.df
    col_ranges = ct.cols  # list of (x0, x1) tuples

    # Find the first data row in camelot's output
    data_row_idx = _find_data_start_row(df)
    if data_row_idx == 0:
        return df  # Can't distinguish headers

    # Get the y-coordinate boundary between headers and data (PDF coords, bottom-left origin)
    header_bottom_y = ct.rows[data_row_idx][0]
    table_top_y = ct._bbox[3]

    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[page_num - 1]  # 0-indexed
        page_height = float(page.height)
        words = page.extract_words(keep_blank_chars=True, x_tolerance=3, y_tolerance=3)

    # Map words to columns, filtering to header area only
    # col_headers[col_idx] = [(y_pdf, text), ...] sorted top to bottom
    col_headers = {i: [] for i in range(len(col_ranges))}

    for w in words:
        w_top_pdf = page_height - w['top']
        w_bottom_pdf = page_height - w['bottom']
        w_x_center = (w['x0'] + w['x1']) / 2
        text = w['text'].strip()

        if not text:
            continue
        # Must be in header area (above first data row)
        # Extend well above table_top since camelot often underestimates
        # the header boundary; cap at 50pt above table top to avoid page chrome
        if w_bottom_pdf < header_bottom_y or w_top_pdf > table_top_y + 50:
            continue
        # Skip title text
        if _TABELA_RE.match(text):
            continue

        # Find which column this word belongs to
        for ci, (cx0, cx1) in enumerate(col_ranges):
            if cx0 - 5 <= w_x_center <= cx1 + 5:
                col_headers[ci].append((w_top_pdf, text))
                break

    # Build label text for each column (skip formula-like lines)
    for ci in col_headers:
        entries = sorted(col_headers[ci], key=lambda x: -x[0])  # top to bottom
        if not entries:
            col_headers[ci] = ''
            continue

        # Group by y-position (words within 3pt are on the same line)
        lines = []
        current_line = [entries[0][1]]
        current_y = entries[0][0]
        for y, text in entries[1:]:
            if abs(y - current_y) < 3:
                current_line.append(text)
            else:
                lines.append(' '.join(current_line))
                current_line = [text]
                current_y = y
        lines.append(' '.join(current_line))

        # Keep only label lines (not formulas/column indices)
        label_lines = [l.strip() for l in lines
                       if not re.match(r'^[\d\s+=+]+$', l.strip())]
        col_headers[ci] = '\n'.join(label_lines) if label_lines else ''

    # Build label row from pdfplumber
    new_labels = [''] * len(col_ranges)
    for ci, val in col_headers.items():
        new_labels[ci] = val

    # Check if pdfplumber found meaningful headers
    total_labels = sum(1 for v in new_labels if v)
    if total_labels < 2:
        return df  # Not enough header data, keep camelot's version

    # Collect formula/index text per column from camelot's header rows
    col_formula = [''] * len(col_ranges)
    for r in range(data_row_idx):
        for c in range(df.shape[1]):
            v = str(df.iloc[r, c]).strip()
            if not v:
                continue
            # Check if this cell is formula-like (contains "=", plain index, or continuation like "+ 8")
            if re.match(r'^[\d\s+=+]+$', v):
                if col_formula[c]:
                    col_formula[c] += '\n' + v
                else:
                    col_formula[c] = v

    # Fix formula continuations: lines starting with "+" should join the nearest
    # adjacent column that has a formula ending with "+" or containing "="
    for c in range(len(col_formula)):
        f = col_formula[c]
        if not f:
            continue
        # Check if this is ONLY a continuation (all lines start with "+")
        lines = f.split('\n')
        if all(l.strip().startswith('+') for l in lines):
            # Find nearest column with a formula containing "="
            for neighbor in [c + 1, c - 1]:
                if 0 <= neighbor < len(col_formula) and '=' in col_formula[neighbor]:
                    col_formula[neighbor] += ' ' + f.replace('\n', ' ')
                    col_formula[c] = ''
                    break

    # Combine label + formula into single header row (joined with \n)
    header_row = [''] * len(col_ranges)
    for ci in range(len(col_ranges)):
        parts = [p for p in [new_labels[ci], col_formula[ci]] if p]
        header_row[ci] = '\n'.join(parts)

    # Build result: single header row + data rows
    header_df = pd.DataFrame([header_row], columns=df.columns)
    data_rows = df.iloc[data_row_idx:].reset_index(drop=True)
    result = pd.concat([header_df, data_rows], ignore_index=True)
    return result


_FOOTER_RE = re.compile(
    r'(БИЛТЕН\s+[Јј]авних\s+финансија|Министарство\s+финансија)',
    re.IGNORECASE,
)


def _drop_footer_rows(df):
    """Remove rows that contain page footer text picked up by camelot."""
    if df.empty:
        return df
    mask = df.apply(
        lambda row: any(_FOOTER_RE.search(str(v)) for v in row),
        axis=1,
    )
    return df[~mask].reset_index(drop=True)


def extract_single_page(pdf_path, page_num):
    """Extract a table from a single PDF page using camelot stream mode."""
    tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
    if not tables:
        print(f"  WARNING: No table found on page {page_num}")
        return pd.DataFrame()
    # Take the largest table if multiple found
    ct = max(tables, key=lambda t: t.df.size)
    # Reconstruct headers using pdfplumber word positions
    df = _reconstruct_headers(ct, pdf_path, page_num)
    # Drop fully empty rows and columns
    df = df.dropna(how='all').reset_index(drop=True)
    df = df.loc[:, ~df.isna().all()]
    # Strip whitespace
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    # Drop page footer rows (page numbers, "БИЛТЕН ...", "Министарство финансија")
    df = _drop_footer_rows(df)
    # Drop rows that are fully empty strings
    df = df[~df.apply(lambda row: all(v == '' for v in row), axis=1)].reset_index(drop=True)
    # Collapse multi-line rows split by camelot
    label_cols = find_label_cols_count(df)
    df = collapse_multiline_rows(df, label_cols=label_cols)
    # Merge columns split by camelot's handling of merged header cells
    df = merge_split_columns(df)
    # Split merged year+month labels into separate columns for consistency
    label_cols = find_label_cols_count(df)
    df, _ = split_year_month_column(df, label_cols)
    return df


def _is_new_item_label(label):
    """Check if a label looks like the start of a new table row (not a continuation)."""
    if re.match(r'^\d{4}(\s|$)', label):  # Year like "2023"
        return True
    if re.match(r'^\d+\.\d*\)', label):  # Closing paren like "1.6)" — NOT a new item
        return False
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


def _merge_rows(acc, next_row, label_cols, header=False):
    """Merge next_row into acc: combine labels, prefer non-empty data values.

    When header=True, concatenate overlapping cell values with newline
    instead of discarding the later value.
    """
    merged = list(acc)

    # Combine labels
    next_label = str(next_row[0]).strip()
    if next_label:
        curr_label = str(merged[0]).strip()
        sep = '\n' if header else ' '
        merged[0] = (curr_label + sep + next_label) if curr_label else next_label

    for c in range(label_cols, min(len(merged), len(next_row))):
        curr_val = str(merged[c]).strip()
        next_val = str(next_row[c]).strip()
        if not curr_val and next_val:
            merged[c] = next_row[c]
        elif curr_val and next_val and header:
            # In header rows, join multi-line cell content
            merged[c] = curr_val + '\n' + next_val

    return merged


def collapse_multiline_rows(df, label_cols=1):
    """Collapse rows where camelot split a multi-line cell into separate rows.

    Handles multiple patterns:
    - Label-only row + data-only row + optional label continuation
    - Row with partial data + overflow row with complementary data
    - Any combination where adjacent rows have non-overlapping data columns

    Header rows (before the first data row) use newline-join for overlapping
    cells so that multi-line header text like "Сектор\\nдржаве" is preserved.
    """
    if df.empty or len(df) < 2:
        return df

    data_start = _find_data_start_row(df)

    rows = [list(r) for r in df.itertuples(index=False, name=None)]
    result = []
    i = 0
    while i < len(rows):
        acc = list(rows[i])
        start_i = i
        i += 1

        # Greedily merge following rows that belong to the same logical row
        while i < len(rows):
            if _should_merge_next(acc, rows[i], label_cols):
                in_header = (i < data_start)
                acc = _merge_rows(acc, rows[i], label_cols, header=in_header)
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


def _find_data_start_row(df):
    """Find the first row that looks like actual data (formatted numbers, not formula indices).

    Skips formula/index rows like '1 = 2 + 9', '3', '4' which have small plain
    numbers.  Real data rows have formatted monetary values like '3.798.170,1'.
    """
    for i in range(min(20, len(df))):
        row = df.iloc[i]
        data_vals = 0
        for v in row:
            s = str(v).strip().replace('\xa0', '')
            if not s:
                continue
            # Must look like a formatted number (with thousands/decimal separators)
            # not just a bare digit like "3" or a formula like "1 = 2 + 9"
            if re.match(r'^-?[\d]+[.,][\d.,]+$', s.replace(' ', '')):
                data_vals += 1
        if data_vals >= 3:
            return i
    return 0


def merge_split_columns(df):
    """Merge adjacent column pairs split by camelot's stream mode.

    When a PDF has merged header cells, camelot often creates two columns:
    one with the header text (empty in data rows) and one with data values
    (empty in header rows).  This detects such H_+_D pairs and merges them.
    Also merges standalone header-only columns into their left neighbour.
    """
    if df.empty or df.shape[1] < 3:
        return df

    data_start = _find_data_start_row(df)
    if data_start == 0:
        return df  # Can't distinguish headers from data

    n_cols = df.shape[1]

    # Classify columns: 'H_' = header-only, '_D' = data-only, 'HD' = both
    col_type = []
    for c in range(n_cols):
        has_header = False
        for r in range(data_start):
            v = df.iloc[r, c]
            if pd.notna(v) and str(v).strip():
                has_header = True
                break
        has_data = False
        for r in range(data_start, min(data_start + 5, len(df))):
            v = df.iloc[r, c]
            if pd.notna(v) and str(v).strip():
                has_data = True
                break
        if has_header and has_data:
            col_type.append('HD')
        elif has_header:
            col_type.append('H_')
        elif has_data:
            col_type.append('_D')
        else:
            col_type.append('__')

    # Build merge plan: list of column groups to merge
    groups = []
    c = 0
    while c < n_cols:
        if col_type[c] == 'H_' and c + 1 < n_cols and col_type[c + 1] == '_D':
            # Header-only + data-only pair
            groups.append((c, c + 1))
            c += 2
        elif col_type[c] == 'H_' and groups:
            # Standalone header-only column: merge into previous group's header
            prev = groups[-1]
            groups[-1] = (*prev, c) if isinstance(prev, tuple) else (prev, c)
            c += 1
        else:
            groups.append((c,))
            c += 1

    if len(groups) == n_cols:
        return df  # Nothing to merge

    # Build merged dataframe
    rows = []
    for r in range(len(df)):
        new_row = []
        for grp in groups:
            if len(grp) == 1:
                new_row.append(df.iloc[r, grp[0]])
            elif len(grp) == 2:
                a, b = grp
                va = df.iloc[r, a]
                vb = df.iloc[r, b]
                sa = str(va).strip() if pd.notna(va) else ''
                sb = str(vb).strip() if pd.notna(vb) else ''
                if sa and sb:
                    new_row.append(sa + ' ' + sb)
                elif sa:
                    new_row.append(va)
                elif sb:
                    new_row.append(vb)
                else:
                    new_row.append('')
            else:
                # 3+ columns merged (H_ + _D + trailing H_ columns)
                vals = []
                for idx in grp:
                    v = df.iloc[r, idx]
                    s = str(v).strip() if pd.notna(v) else ''
                    if s:
                        vals.append(s)
                new_row.append(' '.join(vals) if vals else '')
        rows.append(new_row)

    result = pd.DataFrame(rows)
    result.columns = range(len(result.columns))
    return result


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


def _normalize_label(val):
    """Normalize a row label for matching: strip, collapse whitespace, remove footnotes."""
    s = str(val).strip()
    # Remove footnote text (starts with "* " and is long explanatory text)
    if s.startswith('*') and len(s) > 20:
        return ''
    s = re.sub(r'\s+', ' ', s)
    return s


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


def _build_label_index(df, label_cols):
    """Build a mapping from normalized row label -> row index for matching."""
    index = {}
    for i in range(len(df)):
        parts = []
        for c in range(label_cols):
            v = _normalize_label(df.iloc[i, c])
            if v:
                parts.append(v)
        label = ' '.join(parts)
        if label:
            index[label] = i
    return index


def extract_horizontal_merge(pdf_path, page_nums):
    """Extract a multi-page horizontal table and merge column-wise.

    Uses label-based alignment so that rows missing on continuation pages
    get empty data columns rather than shifting all subsequent rows.
    """
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
    base_label_cols = find_label_cols_count(base)

    for extra in dfs[1:]:
        extra_label_cols = find_label_cols_count(extra)

        # Find header rows (before data) on each side
        base_data_start = _find_data_start_row(base)
        extra_data_start = _find_data_start_row(extra)

        # Detect duplicate label columns: continuation pages may repeat label columns
        # (e.g. "Период"/"Укупно") that find_label_cols_count misses.
        # Only skip columns that are both header-matched AND non-numeric in data rows.
        skip = extra_label_cols
        base_headers = set()
        for c in range(base_label_cols, base.shape[1]):
            for r in range(base_data_start):
                v = str(base.iloc[r, c]).strip().split('\n')[0]
                if v:
                    base_headers.add(v)
                    break
        for c in range(extra_label_cols, extra.shape[1]):
            header_val = ''
            for r in range(extra_data_start):
                v = str(extra.iloc[r, c]).strip().split('\n')[0]
                if v:
                    header_val = v
                    break
            if not (header_val and header_val in base_headers):
                break  # Header doesn't match base — not a duplicate
            # Check if this column has numeric data (→ real data column, not a label)
            num_count = 0
            check_count = 0
            for r in range(extra_data_start, min(extra_data_start + 10, len(extra))):
                v = str(extra.iloc[r, c]).strip().replace(' ', '')
                if v:
                    check_count += 1
                    if re.match(r'^-?[\d,.\s]+$', v):
                        num_count += 1
            if check_count > 0 and num_count / check_count > 0.5:
                break  # Numeric data column — not a duplicate
            skip = c + 1

        # Align header rows by offset (positional — headers are consistent)
        offset = _find_alignment_offset(base, extra)

        # Number of new data columns from the continuation page
        n_new_cols = extra.shape[1] - skip

        # Pre-fill new columns with empty strings
        new_col_start = base.shape[1]
        for i in range(n_new_cols):
            base[new_col_start + i] = ''

        # Fill header rows positionally (they are consistent across pages)
        header_rows = max(base_data_start, extra_data_start)
        for r in range(header_rows):
            extra_r = r + offset if offset > 0 else r
            base_r = r + abs(offset) if offset < 0 else r
            if 0 <= extra_r < len(extra) and 0 <= base_r < len(base):
                for i in range(n_new_cols):
                    val = extra.iloc[extra_r, skip + i]
                    s = str(val).strip() if pd.notna(val) else ''
                    if s:
                        base.iloc[base_r, new_col_start + i] = val

        # Use consistent label columns for matching (max of base and extra)
        match_label_cols = max(base_label_cols, skip)

        # Build label index for data rows of the continuation page
        extra_label_index = {}
        for j in range(extra_data_start, len(extra)):
            parts = []
            for c in range(min(match_label_cols, extra.shape[1])):
                v = _normalize_label(extra.iloc[j, c])
                if v:
                    parts.append(v)
            label = ' '.join(parts)
            if label:
                extra_label_index[label] = j

        # Match base data rows to extra data rows by label
        for base_r in range(base_data_start, len(base)):
            parts = []
            for c in range(min(match_label_cols, base.shape[1])):
                v = _normalize_label(base.iloc[base_r, c])
                if v:
                    parts.append(v)
            base_label = ' '.join(parts)
            if not base_label:
                continue

            # Try exact match first, then prefix match
            extra_r = extra_label_index.get(base_label)
            if extra_r is None:
                for elabel, eidx in extra_label_index.items():
                    if elabel.startswith(base_label[:20]) or base_label.startswith(elabel[:20]):
                        extra_r = eidx
                        break

            if extra_r is not None:
                for i in range(n_new_cols):
                    base.iloc[base_r, new_col_start + i] = extra.iloc[extra_r, skip + i]

    return base


_TABELA_RE = re.compile(r'[TТ]абела\s+\d+[\s.:].{10,}', re.IGNORECASE)


def _extract_title_from_page(page_text):
    """Extract the table title line from page text (first 300 chars)."""
    for line in page_text.split('\n'):
        line = line.strip()
        if _TABELA_RE.match(line):
            return line
    # Try joining first two lines (title may wrap)
    lines = [l.strip() for l in page_text.split('\n') if l.strip()]
    if len(lines) >= 2 and _TABELA_RE.match(lines[0] + ' ' + lines[1]):
        return lines[0] + ' ' + lines[1]
    return None


def consolidate_title_row(df, title):
    """Ensure the table has a dedicated title row as the first row.

    If the extracted data already contains the title mixed into header cells,
    remove it from there.  Then insert a clean title-only row at position 0.
    """
    if df.empty or not title:
        return df

    # Remove any existing title cell from the first few rows
    for r in range(min(4, len(df))):
        for c in range(df.shape[1]):
            v = df.iloc[r, c]
            if pd.notna(v) and isinstance(v, str) and _TABELA_RE.match(v.strip()):
                # Remove only the title line, keep any other content (e.g. header text
                # that got merged into the same cell via newline)
                lines = v.strip().split('\n')
                remaining = [l for l in lines if not _TABELA_RE.match(l.strip())]
                df.iloc[r, c] = '\n'.join(remaining) if remaining else ''
                break

    # Insert title row at position 0
    title_row = pd.DataFrame([[''] * df.shape[1]], columns=df.columns)
    title_row.iloc[0, 0] = title
    df = pd.concat([title_row, df], ignore_index=True)

    # Drop rows that became fully empty after title removal
    df = df[~df.apply(lambda row: all(str(v).strip() == '' or pd.isna(v) for v in row), axis=1)]
    df = df.reset_index(drop=True)

    return df


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


def process_pdf(pdf_path, output_dir, log=None):
    """Process a single PDF: extract all tables and write to Excel files."""
    if log is None:
        log = print
    bilten_id = derive_bilten_id(os.path.basename(pdf_path))
    out_dir = os.path.join(output_dir, bilten_id)
    os.makedirs(out_dir, exist_ok=True)

    log(f"\nProcessing: {pdf_path} -> {out_dir}/")

    # Step 1: Scan pages to find table locations
    log("  Scanning pages...")
    page_texts = scan_pages(pdf_path)
    table_pages = find_table_pages(page_texts, log=log)

    # Step 2: Extract and write
    for xlsx_name, table_defs in TABLE_PATTERNS.items():
        xlsx_path = os.path.join(out_dir, f"{xlsx_name}.xlsx")
        log(f"\n  Writing: {xlsx_path}")

        with pd.ExcelWriter(xlsx_path, engine='openpyxl') as writer:
            for tdef in table_defs:
                tid = tdef["id"]
                key = (xlsx_name, tid)
                pages = table_pages.get(key, [])

                if not pages:
                    log(f"    {tid}: SKIPPED (not found)")
                    continue

                is_multi = key in MULTI_PAGE_TABLES
                log(f"    {tid} (pages {pages}, {'horizontal merge' if is_multi and len(pages) > 1 else 'single'})...")

                if is_multi and len(pages) > 1:
                    df = extract_horizontal_merge(pdf_path, pages)
                else:
                    # For single-page tables, or multi-page tables with only 1 page found
                    df = extract_single_page(pdf_path, pages[0])

                if df.empty:
                    log("    EMPTY!")
                    continue

                # Consolidate table title into a dedicated first row
                title = _extract_title_from_page(page_texts[pages[0]])
                df = consolidate_title_row(df, title)

                safe_sheet = tid[:31]
                df.to_excel(writer, sheet_name=safe_sheet, index=False, header=False)
                log(f"    {df.shape[0]} rows x {df.shape[1]} cols")

    log(f"\nDone: {bilten_id}")


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
