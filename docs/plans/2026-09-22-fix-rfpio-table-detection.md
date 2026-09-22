# RFPIO Table Detection Fix Implementation Plan

> **For agentic workers:** REQUIRED: Use subagent-driven-development (if subagents available) or executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Detect the June 2026 RFPIO income table when the PDF title ends the acronym with a Latin `O` instead of Cyrillic `О`, without weakening unrelated table matching.

**Architecture:** Keep page scanning and `find_table_pages()` unchanged. Make the single `04_ooso`/`Табела 1` signature explicitly tolerate either Unicode glyph in the final character of `РФПИО`, and protect both known title spellings with a focused standard-library unit test.

**Tech Stack:** Python 3.10+, `re`, `unittest`, pdfplumber, Camelot, pandas, openpyxl

---

## Chunk 1: Regression Fix and Verification

### Task 1: Add a failing page-detection regression test

**Files:**
- Create: `tests/test_table_detection.py`
- Reference: `extract_tables.py:18-50`
- Reference: `extract_tables.py:70-91`

- [ ] **Step 1: Create the test directory**

Run:

```bash
mkdir -p tests
```

Expected: `tests/` exists without modifying any application files.

- [ ] **Step 2: Create a focused standard-library test**

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

The synthetic title mirrors physical page 72 of `bilteni/2026-06-YGxhTM_6a90314926bb9.pdf`. Testing both glyphs prevents fixing June 2026 by breaking bulletins that use the original Cyrillic spelling.

- [ ] **Step 3: Run the test and confirm the Latin-glyph case fails**

Run:

```bash
.venv/bin/python -m unittest discover -s tests -p 'test_table_detection.py' -v
```

Expected: the test fails because `table_pages.get(("04_ooso", "Табела 1"))` is `None` for the `final_o='O'` subtest. The Cyrillic `final_o='О'` subtest should not fail.

### Task 2: Make the RFPIO income signature glyph-tolerant

**Files:**
- Modify: `extract_tables.py:43`
- Test: `tests/test_table_detection.py`

- [ ] **Step 1: Apply the minimal pattern change**

Change only the `04_ooso` definition for `Табела 1`:

```python
{"id": "Табела 1", "pattern": r"Табела 1\.?\s*Примања РФПИ[ОO]"},
```

Do not normalize all scanned PDF text and do not modify the other table signatures. The character class documents and contains the known PDF encoding variation.

- [ ] **Step 2: Run the focused test**

Run:

```bash
.venv/bin/python -m unittest discover -s tests -p 'test_table_detection.py' -v
```

Expected: `test_finds_rfpio_income_title_with_cyrillic_or_latin_o` passes for both subtests.

- [ ] **Step 3: Run complete unit-test discovery**

Run:

```bash
.venv/bin/python -m unittest discover -s tests -v
```

Expected: all discovered tests pass.

### Task 3: Verify the June 2026 extraction end to end

**Files:**
- Read: `bilteni/2026-06-YGxhTM_6a90314926bb9.pdf`
- Generate, but do not commit: `tabele/2026-06/04_ooso.xlsx`

- [ ] **Step 1: Process only the affected bulletin**

Run:

```bash
.venv/bin/python extract_tables.py bilteni/2026-06-YGxhTM_6a90314926bb9.pdf
```

Expected: scanning does not emit `WARNING: Could not find 04_ooso/Табела 1`; extraction logs `Табела 1 (pages [72], single)`; and `tabele/2026-06/04_ooso.xlsx` is created. Other generated files under `tabele/2026-06/` are verification artifacts covered by `.gitignore` and must not be committed.

- [ ] **Step 2: Assert that the workbook contains the extracted sheet and title**

Run:

```bash
.venv/bin/python -c "from openpyxl import load_workbook; p='tabele/2026-06/04_ooso.xlsx'; w=load_workbook(p, read_only=True, data_only=True); assert 'Табела 1' in w.sheetnames; title=str(w['Табела 1']['A1'].value); assert title == 'Табела 1. Примања РФПИO у мил. динара'; print(p, w.sheetnames, title)"
```

Expected: the command prints the workbook path, a sheet list containing `Табела 1`, and its RFPIO income title without raising an assertion.

- [ ] **Step 3: Confirm the implementation diff is scoped**

Run:

```bash
git status --short
git diff -- extract_tables.py tests/test_table_detection.py
```

Expected: implementation changes are limited to the one signature in `extract_tables.py` and the new regression test. Do not stage `.claude/` or generated extraction artifacts.

- [ ] **Step 4: Commit the implementation**

```bash
git add extract_tables.py tests/test_table_detection.py
git commit -m "fix(extraction): detect RFPIO title with Latin O"
```
