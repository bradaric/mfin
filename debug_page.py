#!/usr/bin/env python3
"""Diagnostic script to inspect what camelot/pdfplumber produce for a specific page."""
import sys
import camelot
import pdfplumber
import pandas as pd
from extract_tables import (
    extract_single_page, _find_data_start_row, find_label_cols_count,
    collapse_multiline_rows, _normalize_label, extract_horizontal_merge,
    _should_merge_next,
)


def inspect_page(pdf_path, page_num):
    print(f"\n{'='*80}")
    print(f"PAGE {page_num}")
    print(f"{'='*80}")

    # 1. Raw camelot output
    print("\n--- Raw camelot output ---")
    tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
    ct = max(tables, key=lambda t: t.df.size)
    raw_df = ct.df
    # Find rows containing "трансфери" or "домаћинствима"
    for idx, row in raw_df.iterrows():
        row_text = ' '.join(str(v) for v in row)
        if 'трансфер' in row_text.lower() or 'домаћинств' in row_text.lower():
            print(f"  Row {idx}: {[str(v)[:60] for v in row]}")

    # 2. After extract_single_page (full pipeline)
    print("\n--- After extract_single_page ---")
    df = extract_single_page(pdf_path, page_num)
    label_cols = find_label_cols_count(df)
    data_start = _find_data_start_row(df)
    print(f"  label_cols={label_cols}, data_start={data_start}, shape={df.shape}")

    for idx, row in df.iterrows():
        row_text = ' '.join(str(v) for v in row)
        if 'трансфер' in row_text.lower() or 'домаћинств' in row_text.lower():
            label = _normalize_label(row.iloc[0])
            data_filled = sum(1 for c in range(label_cols, len(row)) if str(row.iloc[c]).strip())
            print(f"  Row {idx}: label={repr(label)}, data_filled={data_filled}")
            print(f"    raw col0={repr(str(row.iloc[0]))}")
            if label_cols > 1:
                print(f"    raw col1={repr(str(row.iloc[1]))}")
            # Show first few data values
            data_vals = [str(row.iloc[c]) for c in range(label_cols, min(label_cols+5, len(row))) if str(row.iloc[c]).strip()]
            print(f"    first data: {data_vals}")

    # 3. Show rows around "Социјална помоћ" to check for incorrect merging
    print("\n--- Rows around Социјална помоћ ---")
    for idx, row in df.iterrows():
        label = _normalize_label(row.iloc[0])
        if any(kw in label.lower() for kw in ['социјалн', 'трансфер', 'домаћинств', 'текући расход']):
            data_filled = sum(1 for c in range(label_cols, len(row)) if str(row.iloc[c]).strip())
            print(f"  Row {idx}: label={repr(label)}, data_filled={data_filled}")


def inspect_horizontal_merge(pdf_path, page_nums):
    print(f"\n{'='*80}")
    print(f"HORIZONTAL MERGE (pages {page_nums})")
    print(f"{'='*80}")

    # Extract each page
    dfs = []
    for pg in page_nums:
        df = extract_single_page(pdf_path, pg)
        if not df.empty:
            dfs.append((pg, df))

    # Show label index for each page
    for pg, df in dfs:
        label_cols = find_label_cols_count(df)
        data_start = _find_data_start_row(df)
        print(f"\n--- Page {pg} label index (label_cols={label_cols}, data_start={data_start}) ---")
        for j in range(data_start, len(df)):
            parts = []
            for c in range(label_cols):
                v = _normalize_label(df.iloc[j, c])
                if v:
                    parts.append(v)
            label = ' '.join(parts)
            if any(kw in label.lower() for kw in ['социјалн', 'трансфер', 'домаћинств', 'текући расход', 'помоћ']):
                data_filled = sum(1 for c in range(label_cols, len(df.columns)) if str(df.iloc[j, c]).strip())
                print(f"  Row {j}: label={repr(label)}, data_filled={data_filled}")

    # Now run the actual horizontal merge and inspect the result
    print("\n--- After horizontal merge ---")
    result = extract_horizontal_merge(pdf_path, page_nums)
    label_cols = find_label_cols_count(result)
    for idx, row in result.iterrows():
        label = _normalize_label(row.iloc[0])
        if any(kw in label.lower() for kw in ['социјалн', 'трансфер', 'домаћинств', 'текући расход']):
            data_filled = sum(1 for c in range(label_cols, len(row)) if str(row.iloc[c]).strip())
            # Count empty columns to identify gaps
            empty_ranges = []
            start = None
            for c in range(label_cols, len(row)):
                if str(row.iloc[c]).strip() == '':
                    if start is None:
                        start = c
                else:
                    if start is not None:
                        empty_ranges.append(f"{start}-{c-1}")
                        start = None
            if start is not None:
                empty_ranges.append(f"{start}-{len(row)-1}")
            print(f"  Row {idx}: label={repr(label)}, data_filled={data_filled}/{len(row)-label_cols}, empty_ranges={empty_ranges}")


def inspect_raw_collapse(pdf_path, page_num):
    """Show what happens step-by-step during collapse_multiline_rows."""
    print(f"\n{'='*80}")
    print(f"RAW COLLAPSE DEBUG FOR PAGE {page_num}")
    print(f"{'='*80}")

    tables = camelot.read_pdf(pdf_path, pages=str(page_num), flavor='stream')
    ct = max(tables, key=lambda t: t.df.size)
    from extract_tables import _reconstruct_headers, _drop_footer_rows
    df = _reconstruct_headers(ct, pdf_path, page_num)
    df = df.dropna(how='all').reset_index(drop=True)
    df = df.loc[:, ~df.isna().all()]
    df = df.map(lambda x: x.strip() if isinstance(x, str) else x)
    df = _drop_footer_rows(df)
    df = df[~df.apply(lambda row: all(v == '' for v in row), axis=1)].reset_index(drop=True)

    label_cols = find_label_cols_count(df)
    print(f"label_cols={label_cols}")

    # Show rows before collapse, focusing on area of interest
    print("\n--- Pre-collapse rows (near трансфери/домаћинствима) ---")
    rows = [list(r) for r in df.itertuples(index=False, name=None)]
    for i, row in enumerate(rows):
        label = str(row[0]).strip()
        if any(kw in label.lower() for kw in ['социјалн', 'трансфер', 'домаћинств', 'текући']):
            data_filled = sum(1 for v in row[label_cols:] if str(v).strip())
            print(f"  Row {i}: col0={repr(label)}, data_filled={data_filled}")
            # Check merge decision with previous
            if i > 0:
                prev_label = str(rows[i-1][0]).strip()
                # Simulate _should_merge_next
                try:
                    merge = _should_merge_next(rows[i-1], row, label_cols)
                    print(f"    _should_merge_next(prev={repr(prev_label[:40])}, this) = {merge}")
                except Exception as e:
                    print(f"    _should_merge_next error: {e}")
            if i + 1 < len(rows):
                next_label = str(rows[i+1][0]).strip()
                next_data_filled = sum(1 for v in rows[i+1][label_cols:] if str(v).strip())
                print(f"    Next row: col0={repr(next_label)}, data_filled={next_data_filled}")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python debug_page.py <pdf_path> [page_nums...]")
        print("Example: python debug_page.py bilteni/2025-11.pdf 39 40 41")
        sys.exit(1)

    pdf_path = sys.argv[1]
    pages = [int(p) for p in sys.argv[2:]] if len(sys.argv) > 2 else [39, 40, 41]

    for pg in pages:
        inspect_page(pdf_path, pg)
        inspect_raw_collapse(pdf_path, pg)

    if len(pages) >= 2:
        inspect_horizontal_merge(pdf_path, pages)
