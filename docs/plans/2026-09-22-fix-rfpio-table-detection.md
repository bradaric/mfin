# RFPIO Table Detection Fix Implementation Plan

> **For agentic workers:** REQUIRED: Use subagent-driven-development (if subagents available) or executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Detect the June 2026 RFPIO income table when the PDF title ends the acronym with a Latin `O` instead of Cyrillic `О`, and keep extraction working when pandas 3 encounters an all-missing leading column.

**Architecture:** Keep page scanning and `find_table_pages()` unchanged. Make the single `04_ooso`/`Табела 1` signature explicitly tolerate either Unicode glyph in the final character of `РФПИО`. In `find_label_cols_count()`, replace the dtype-sensitive pandas reduction with Python's integer-returning `sum()` so empty string-typed series are safe; protect both defects with focused standard-library tests.

**Tech Stack:** Python 3.10+, `re`, `unittest`, pdfplumber, Camelot, pandas, openpyxl

---

## Chunk 1: Regression Fix and Verification

### Task 1: Add a failing page-detection regression test

**Files:**
- Create: `tests/test_table_detection.py`
- Reference: `extract_tables.py:18-50`
- Reference: `extract_tables.py:70-91`

- [x] **Step 1: Create the test directory**

Run:

```bash
mkdir -p tests
```

Expected: `tests/` exists without modifying any application files.

- [x] **Step 2: Create a focused standard-library test**

Create `tests/test_table_detection.py` with this content:

```python
import unittest

from extract_tables import find_table_pages


class FindTablePagesTests(unittest.TestCase):
    def test_finds_rfpio_income_title_with_cyrillic_or_latin_o(self):
        key = ("04_ooso", "Табела 1")

        for final_o in ("О", "O"):
            with self.subTest(final_o=final_o):
                page_texts = {
                    72: f"Табела 1. Примања РФПИ{final_o} у мил. динара"
                }

                table_pages = find_table_pages(page_texts, log=lambda _message: None)

                self.assertEqual(table_pages.get(key), [72])


if __name__ == "__main__":
    unittest.main()
```

The synthetic page number is arbitrary; the title mirrors one-based PDF page 70 of `bilteni/2026-06-YGxhTM_6a90314926bb9.pdf`. Testing both glyphs prevents fixing June 2026 by breaking bulletins that use the original Cyrillic spelling.

- [x] **Step 3: Run the test and confirm the Latin-glyph case fails**

Run:

```bash
.venv/bin/python -m unittest discover -s tests -p 'test_table_detection.py' -v
```

Observed after dependencies were installed and before the pattern change: the Latin `final_o='O'` subtest failed with `AssertionError: None != [72]`; the Cyrillic `final_o='О'` subtest passed.

### Task 2: Make the RFPIO income signature glyph-tolerant

**Files:**
- Modify: `extract_tables.py:43`
- Test: `tests/test_table_detection.py`

- [x] **Step 1: Apply the minimal pattern change**

Change only the `04_ooso` definition for `Табела 1`:

```python
{"id": "Табела 1", "pattern": r"Табела 1\.?\s*Примања РФПИ[ОO]"},
```

Do not normalize all scanned PDF text and do not modify the other table signatures. The character class documents and contains the known PDF encoding variation.

- [x] **Step 2: Run the focused test**

Run:

```bash
.venv/bin/python -m unittest discover -s tests -p 'test_table_detection.py' -v
```

Expected: `test_finds_rfpio_income_title_with_cyrillic_or_latin_o` passes for both subtests.

- [x] **Step 3: Run complete unit-test discovery**

Run:

```bash
.venv/bin/python -m unittest discover -s tests -v
```

Expected: all discovered tests pass.

### Task 3: Make label-column counting safe with pandas 3

**Files:**
- Modify: `tests/test_table_detection.py`
- Modify: `extract_tables.py:714-727`

**Confirmed root cause:** During June 2026 extraction, `_drop_footer_rows()` leaves the first column of `03_budzet_opstina/Табела 2` entirely missing. With pandas 3.0.6, `dropna().astype(str)` produces an empty `StringDtype` series; applying the numeric predicate preserves that dtype, and pandas returns `''` from `.sum()`. Dividing that string by the denominator raises `TypeError`. The counting operation must return an integer independently of pandas dtype inference.

- [ ] **Step 1: Add a failing pandas 3 regression test**

Update the imports in `tests/test_table_detection.py`:

```python
import pandas as pd

from extract_tables import find_label_cols_count, find_table_pages
```

Add a separate test class after `FindTablePagesTests`:

```python
class FindLabelColsCountTests(unittest.TestCase):
    def test_counts_all_missing_leading_column_as_label(self):
        df = pd.DataFrame({0: [None, None], 1: ["1", "2"]})

        self.assertEqual(find_label_cols_count(df), 1)
```

- [ ] **Step 2: Run the focused tests and confirm the new test fails**

Run:

```bash
.venv/bin/python -m unittest discover -s tests -p 'test_table_detection.py' -v
```

Expected: the RFPIO title test passes, while `test_counts_all_missing_leading_column_as_label` errors at `extract_tables.py:723` with `TypeError: unsupported operand type(s) for /: 'str' and 'int'` under pandas 3.

- [ ] **Step 3: Replace the dtype-sensitive reduction**

In `find_label_cols_count()`, replace only the `vals.apply(...).sum()` expression with Python's built-in `sum()` over the same predicate:

```python
        numeric_count = sum(
            bool(re.match(r'^-?[\d,.\s]+$', v.replace(' ', '')))
            for v in vals
        )
```

Do not add a pandas version pin or special-case branch. Built-in `sum()` returns integer zero for an empty iterator while preserving the existing result for populated columns.

- [ ] **Step 4: Run the focused tests**

Run:

```bash
.venv/bin/python -m unittest discover -s tests -p 'test_table_detection.py' -v
```

Expected: both tests pass.

- [ ] **Step 5: Run complete unit-test discovery**

Run:

```bash
.venv/bin/python -m unittest discover -s tests -v
```

Expected: all discovered tests pass.

### Task 4: Verify the June 2026 extraction end to end

**Files:**
- Read: `bilteni/2026-06-YGxhTM_6a90314926bb9.pdf`
- Generate, but do not commit: `tabele/2026-06/04_ooso.xlsx`

- [ ] **Step 1: Process only the affected bulletin**

Run:

```bash
set -o pipefail; .venv/bin/python extract_tables.py bilteni/2026-06-YGxhTM_6a90314926bb9.pdf 2>&1 | tee /tmp/mfin-2026-06.log
```

Expected: the command exits successfully; scanning does not emit `WARNING: Could not find 04_ooso/Табела 1`; extraction logs `Табела 1 (pages [70], single)`; and `tabele/2026-06/04_ooso.xlsx` is created. Other generated files under `tabele/2026-06/` are verification artifacts covered by `.gitignore` and must not be committed.

- [ ] **Step 2: Assert that extraction completed without a post-processing fallback**

Run:

```bash
.venv/bin/python -c "from pathlib import Path; log=Path('/tmp/mfin-2026-06.log').read_text(); assert 'WARNING: post-processing failed' not in log; assert 'Табела 1 (pages [70], single)' in log; print('June 2026 extraction log checks passed')"
```

Expected: the command prints `June 2026 extraction log checks passed`. The successful extraction command and this log assertion jointly reject both an unhandled regression and any guarded post-processing fallback.

- [ ] **Step 3: Assert that the workbook contains the extracted sheet and title**

Run:

```bash
.venv/bin/python -c "from openpyxl import load_workbook; p='tabele/2026-06/04_ooso.xlsx'; w=load_workbook(p, read_only=True, data_only=True); assert 'Табела 1' in w.sheetnames; title=str(w['Табела 1']['A1'].value); assert title == 'Табела 1. Примања РФПИO у мил. динара'; print(p, w.sheetnames, title)"
```

Expected: the command prints the workbook path, a sheet list containing `Табела 1`, and its RFPIO income title without raising an assertion.

- [ ] **Step 4: Confirm the implementation diff is scoped**

Run:

```bash
git status --short
git diff -- extract_tables.py tests/test_table_detection.py
```

Expected: implementation changes are limited to the RFPIO signature, the dtype-safe numeric count in `find_label_cols_count()`, and the two regression tests. Do not stage `.claude/` or generated extraction artifacts.

- [ ] **Step 5: Commit the implementation**

```bash
git add extract_tables.py tests/test_table_detection.py
git commit -m "fix(extraction): handle RFPIO title and empty columns"
```

The plan amendment is committed separately before implementation resumes; do not stage it again with the implementation.
