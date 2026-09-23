# Robust Table Detection Implementation Plan

> **For agentic workers:** REQUIRED: Use subagent-driven-development (if subagents available) or executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make fiscal table page detection survive Cyrillic/Latin lookalike edits, a Latin or mixed-script `Tabela`/`Табела` token, and whitespace or case changes, without another per-signature regex patch.

**Architecture:** Keep `TABLE_PATTERNS` as canonical Cyrillic title prefixes, not regexes. Fold both the scanned page text and each title before a substring compare: NFKC, drop invisible characters, collapse whitespace, casefold, rewrite the table-name token to `табела`, then map Latin lookalikes to Cyrillic. Strip spaces and light punctuation only in the comparison key so `Табела 1.`, `Табела 1:`, and `Tabela1` still match, while `Табела 1` does not match `Табела 10`. Use the same fold for title-line recognition so page detection and title capture do not diverge. Stored title text is never rewritten.

**Tech Stack:** Python 3.10+, `re`, `unicodedata`, `unittest`. No new dependencies.

---

## Decisions (locked)

The requester waived interactive approval. Do not reopen these.

1. **Lookalike folding, not full Serbian transliteration.** Latin `B` is the visual twin of Cyrillic `В`, but the orthographic twin of `Б`. Mapping both ways is impossible. Observed misses (`РФПИО`/`РФПИO`, `Табела`/`Taбела`) are lookalikes. Full Latin retypes of the descriptive title (`Primanja budzeta Vojvodine`) stay unmatched on purpose.
2. **The table-name token is special.** `Tabela`, `TABELA`, `Табела`, and per-letter mixes including non-lookalikes `b`/`б` and `l`/`л` all become `табела` before the lookalike map. Do this only for that 6-letter token. Do not extend it to other words.
3. **Lookalike map, after casefold, Latin to Cyrillic:** `a->а`, `b->в`, `c->с`, `e->е`, `h->н`, `j->ј`, `k->к`, `m->м`, `o->о`, `p->р`, `t->т`, `x->х`, `y->у`. `b->в` exists so casefold of Latin `B` (the twin of Cyrillic `В`) still matches. It must run after the table-token rewrite, or `tabela` becomes `тавела`. Do not map `i`, `l`, `n`, `r`, `s`, `u`, `v`, or any other letter. Latin `P` folds to `Р`, not `П`. Latin `C` folds to `С`, not `Ц`. Latin `H` folds to `Н`, not `Х`. Latin `V` does not fold to `В`.
4. **Comparison ignores whitespace and `.:;,-` plus hyphen variants, but matching stays a full-prefix substring.** Do not match on the table number alone.
5. **Replace `pattern` regexes with `title` literals.** Delete the ad-hoc `[TТ]`, `[aа]`, and `[ОO]` classes. They become redundant and will rot.
6. **Replace `_TABELA_RE` with `_is_table_title()`.** Same fold. Still require a table number plus at least 10 characters of following text, and still anchor title recognition at the start of the line. Page detection stays an unanchored substring search, which is what `re.search` does today.
7. **Out of scope:** `find_label_cols_count`, pandas dtype handling, camelot extraction, header reconstruction logic beyond the title-skip predicate, footer regexes, full-PDF re-extraction, new dependencies, pushing, and pull requests.
8. **Branch:** `feature/robust-table-detection` from current `main`. This repo has no `dev` branch. Do not use `git worktree`. Do not commit on `main`.

## File structure

- Modify: `extract_tables.py` - title literals, fold helpers, `find_table_pages`, title-line predicate.
- Modify: `tests/test_table_detection.py` - folding, distinctness, and title-capture tests.
- Do not create a new module. One caller, and tests already import `extract_tables`.

---

## Chunk 1: Folded table matching

### Task 1: Create the feature branch

**Files:**
- None

- [ ] **Step 1: Branch from main in the current working tree**

Run from `/mnt/ubuntu-home/sasa/workshop/mfin`:

```bash
git checkout -b feature/robust-table-detection
```

Expected: branch `feature/robust-table-detection` is checked out. Do not use `git worktree`. Do not push.

- [ ] **Step 2: Commit the plan if it is still uncommitted**

If `docs/plans/2026-09-22-robust-table-detection.md` is untracked or unstaged, commit only that file:

```bash
git add docs/plans/2026-09-22-robust-table-detection.md
git commit -m "docs(plans): add robust table detection plan"
```

Do not append a `Co-Authored-By` trailer. If the plan commit already exists on this branch, do not commit it again.

### Task 2: Add failing detection tests

**Files:**
- Modify: `tests/test_table_detection.py`
- Reference: `extract_tables.py:18-91`
- Reference: `extract_tables.py:977-991`

- [ ] **Step 1: Replace the test module with the content below**

Keep the existing RFPIO, Vojvodina, and label-column tests. Add the new cases in the same file.

```python
import unittest

import pandas as pd

from extract_tables import (
    TABLE_PATTERNS,
    _extract_title_from_page,
    _is_table_title,
    _table_match_key,
    consolidate_title_row,
    find_label_cols_count,
    find_table_pages,
)

# Visual twins only. Cyrillic В maps to Latin B, not V. Cyrillic Р maps to Latin P.
_CYRILLIC_LOOKALIKE_TO_LATIN = str.maketrans({
    "а": "a", "в": "b", "с": "c", "е": "e", "н": "h", "ј": "j",
    "к": "k", "м": "m", "о": "o", "р": "p", "т": "t", "х": "x", "у": "y",
    "А": "A", "В": "B", "С": "C", "Е": "E", "Н": "H", "Ј": "J",
    "К": "K", "М": "M", "О": "O", "Р": "P", "Т": "T", "Х": "X", "У": "Y",
})


def _variants(title):
    yield "canonical", title
    yield "upper", title.upper()
    yield "lower", title.lower()
    yield "double-space", title.replace(" ", "  ")
    yield "nbsp", title.replace(" ", "\u00a0")
    yield "newline", title.replace(". ", ".\n")
    yield "glued-punct", title.replace(". ", ".")
    yield "colon", title.replace(".", ":")
    yield "no-punct", title.replace(".", "")
    yield "latin-token", title.replace("Табела", "Tabela")
    yield "upper-latin-token", title.replace("Табела", "TABELA")
    yield "mixed-token", title.replace("Табела", "Taбела")
    yield "glued-latin-token", title.replace("Табела ", "Tabela")
    yield "lookalikes", title.translate(_CYRILLIC_LOOKALIKE_TO_LATIN)


class FindTablePagesTests(unittest.TestCase):
    def test_finds_vojvodina_income_title_with_mixed_script(self):
        key = ("02_budzet_vojvodine", "Табела 1")

        for title in (
            "Табела 1. Примања буџета Војводине у мил. динара",
            "Taбела 1. Примања буџета Војводине у мил. динара  ",
        ):
            with self.subTest(title=title):
                table_pages = find_table_pages({62: title}, log=lambda _message: None)
                self.assertEqual(table_pages.get(key), [62])

    def test_finds_rfpio_income_title_with_cyrillic_or_latin_o(self):
        key = ("04_ooso", "Табела 1")

        for final_o in ("О", "O"):
            with self.subTest(final_o=final_o):
                page_texts = {
                    72: f"Табела 1. Примања РФПИ{final_o} у мил. динара"
                }

                table_pages = find_table_pages(page_texts, log=lambda _message: None)

                self.assertEqual(table_pages.get(key), [72])

    def test_each_title_variant_matches_only_its_own_key(self):
        for xlsx_name, defs in TABLE_PATTERNS.items():
            for tdef in defs:
                key = (xlsx_name, tdef["id"])
                for name, text in _variants(tdef["title"]):
                    with self.subTest(xlsx=xlsx_name, tid=tdef["id"], variant=name):
                        pages = find_table_pages({7: text}, log=lambda _message: None)
                        self.assertEqual(pages, {key: [7]})

    def test_preamble_does_not_hide_a_folded_title(self):
        text = (
            "Министарство финансија\n"
            "Tabela 1. Примања буџета Војводине у мил. динара"
        )
        pages = find_table_pages({62: text}, log=lambda _message: None)
        self.assertEqual(pages, {("02_budzet_vojvodine", "Табела 1"): [62]})

    def test_table_1_does_not_match_table_10_or_11(self):
        pages = find_table_pages(
            {
                1: "Табела 10. Субвенције из буџета",
                2: "ТАБЕЛА 11. Донације и трансфери из буџета",
            },
            log=lambda _message: None,
        )
        self.assertEqual(
            pages,
            {
                ("01_budzet_rs", "Табела 10"): [1],
                ("01_budzet_rs", "Табела 11"): [2],
            },
        )

    def test_neporeski_does_not_match_poreski(self):
        pages = find_table_pages(
            {1: "Табела 6. Непорески приходи"},
            log=lambda _message: None,
        )
        self.assertEqual(pages, {("01_budzet_rs", "Табела 6"): [1]})

    def test_lowercase_latin_b_matches_cyrillic_ve_lookalike(self):
        text = "Табела 1. Примања буџета bојводине"
        pages = find_table_pages({62: text}, log=lambda _message: None)
        self.assertEqual(pages.get(("02_budzet_vojvodine", "Табела 1")), [62])

    def test_latin_v_does_not_match_cyrillic_ve(self):
        text = "Табела 1. Примања буџета vојводине"
        pages = find_table_pages({62: text}, log=lambda _message: None)
        self.assertNotIn(("02_budzet_vojvodine", "Табела 1"), pages)

    def test_full_latin_transliteration_is_not_a_match(self):
        pages = find_table_pages(
            {1: "Tabela 1. Primanja budzeta Vojvodine"},
            log=lambda _message: None,
        )
        self.assertNotIn(("02_budzet_vojvodine", "Табела 1"), pages)

    def test_unrelated_text_matches_nothing(self):
        pages = find_table_pages({1: "Увод у билтен"}, log=lambda _message: None)
        self.assertEqual(pages, {})


class TableMatchKeyTests(unittest.TestCase):
    def test_folds_case_space_token_and_lookalikes_to_one_key(self):
        expected = _table_match_key("Табела 1. Примања РФПИО")
        for text in (
            "ТАБЕЛА 1. ПРИМАЊА РФПИО",
            "Tabela  1: Примања РФПИO",
            "Taбела1.Примања РФПИО",
            "табела 1.\nпримања рфпио",
        ):
            with self.subTest(text=text):
                self.assertEqual(_table_match_key(text), expected)


class TableTitleRecognitionTests(unittest.TestCase):
    def test_recognizes_folded_title_and_returns_original_line(self):
        line = "TABELA  1. Примања РФПИO у мил. динара"
        self.assertTrue(_is_table_title(line))
        self.assertEqual(_extract_title_from_page(line + "\nостало"), line)

    def test_joins_wrapped_mixed_script_title(self):
        text = "Tabela 1.\nПримања буџета Војводине у мил. динара"
        self.assertEqual(
            _extract_title_from_page(text),
            "Tabela 1. Примања буџета Војводине у мил. динара",
        )

    def test_rejects_bare_table_token(self):
        self.assertFalse(_is_table_title("Tabela"))
        self.assertFalse(_is_table_title("Табела 1."))

    def test_strips_folded_title_from_header_cell_without_rewriting_it(self):
        title = "Tabela 1. Примања буџета Војводине у мил. динара"
        df = pd.DataFrame([[title, "x"]])
        out = consolidate_title_row(df, title)
        self.assertEqual(out.iloc[0, 0], title)
        self.assertNotIn(title, out.iloc[1].tolist())


class FindLabelColsCountTests(unittest.TestCase):
    def test_counts_all_missing_leading_column_as_label(self):
        df = pd.DataFrame({0: [None, None], 1: ["1", "2"]})

        self.assertEqual(find_label_cols_count(df), 1)


if __name__ == "__main__":
    unittest.main()
```

- [ ] **Step 2: Run the new tests and confirm they fail**

Run from the repo root:

```bash
if [ -x .venv/bin/python ]; then PY=.venv/bin/python; else PY=python3; fi
$PY -m unittest discover -s tests -p 'test_table_detection.py' -v
```

Expected: collection fails with `ImportError` because `_table_match_key` and `_is_table_title` are not defined yet. That is the failing test. Do not stub those functions just to collect, and do not weaken assertions.

### Task 3: Fold text and match canonical titles

**Files:**
- Modify: `extract_tables.py:1-91`
- Modify: `extract_tables.py:137`
- Modify: `extract_tables.py:977-990`
- Test: `tests/test_table_detection.py`

- [ ] **Step 1: Add `unicodedata` to the imports**

`extract_tables.py` already imports `re`. Add `import unicodedata` next to the other stdlib imports.

- [ ] **Step 2: Replace `TABLE_PATTERNS` with canonical title literals**

Replace the whole `TABLE_PATTERNS` dict. Do not keep `pattern`. Do not keep character classes.

```python
# Canonical Cyrillic title prefixes. Matching folds case, whitespace, the
# Tabela/Табела token, and Latin/Cyrillic lookalikes. Do not put regex here.
TABLE_PATTERNS = {
    "00_fiskalna_kretanja": [
        {"id": "Табела 1", "title": "Табела 1. Консолидовани биланс државе у периоду"},
        {"id": "Табела 2", "title": "Табела 2. Консолидовани биланс државе по нивоима"},
    ],
    "01_budzet_rs": [
        {"id": "Табела 3", "title": "Табела 3. Примања и издаци буџета"},
        {"id": "Табела 4", "title": "Табела 4. Порески приходи"},
        {"id": "Табела 5", "title": "Табела 5. Порез на додату вредност"},
        {"id": "Табела 6", "title": "Табела 6. Непорески приходи"},
        {"id": "Табела 7", "title": "Табела 7. Укупни издаци буџета"},
        {"id": "Табела 8", "title": "Табела 8. Укупни расходи за запослене"},
        {"id": "Табела 9", "title": "Табела 9. Расходи по основу отплате камата"},
        {"id": "Табела 10", "title": "Табела 10. Субвенције из буџета"},
        {"id": "Табела 11", "title": "Табела 11. Донације и трансфери из буџета"},
    ],
    "02_budzet_vojvodine": [
        {"id": "Табела 1", "title": "Табела 1. Примања буџета Војводине"},
        {"id": "Табела 2", "title": "Табела 2. Издаци буџета Војводине"},
    ],
    "03_budzet_opstina": [
        {"id": "Табела 1", "title": "Табела 1. Примања буџета општина"},
        {"id": "Табела 2", "title": "Табела 2. Издаци буџета општина"},
    ],
    "04_ooso": [
        {"id": "Табела 1", "title": "Табела 1. Примања РФПИО"},
        {"id": "Табела 2", "title": "Табела 2. Издаци РФПИО"},
        {"id": "Табела 3", "title": "Табела 3. Примања Републичког фонда за здравствено"},
        {"id": "Табела 4", "title": "Табела 4. Издаци Републичког фонда за здравствено"},
        {"id": "Табела 5", "title": "Табела 5. Примања Националне службе"},
        {"id": "Табела 6", "title": "Табела 6. Издаци Националне службе"},
    ],
}
```

Leave `MULTI_PAGE_TABLES` unchanged.

- [ ] **Step 3: Add the fold helpers immediately after `MULTI_PAGE_TABLES`**

```python
# Table-name token only. b/б and l/л are not lookalikes, but the word
# "table" is edited between scripts as a whole. Run before the lookalike map.
_TABLE_TOKEN_RE = re.compile(
    r"(?<![\w])[tт][aа][bб][eе][lл][aа](?=\d|\s|[:.]|$)"
)

# Lowercase Latin -> Cyrillic visual twins. Applied after casefold, so Latin
# B (twin of Cyrillic В) is already "b". Not an orthographic transliteration.
_LOOKALIKE_TO_CYRILLIC = str.maketrans({
    "a": "а",
    "b": "в",
    "c": "с",
    "e": "е",
    "h": "н",
    "j": "ј",
    "k": "к",
    "m": "м",
    "o": "о",
    "p": "р",
    "t": "т",
    "x": "х",
    "y": "у",
})

_INVISIBLE_RE = re.compile(r"[\u00ad\u200b\u200c\u200d\ufeff]")
_WS_RE = re.compile(r"[\s\u00a0\u202f\u2007\u2008\u2009\u200a\u205f\u3000]+")
_COMPARE_STRIP_RE = re.compile(r"[\s.:;,\-\u2010\u2011\u2013\u2014]+")
_TITLE_RE = re.compile(r"^табела\s*\d+\s*[:.]?\s*\S.{9,}")


def _canonicalize_table_text(text):
    """Fold case, the table-name token, and Latin/Cyrillic lookalikes.

    Punctuation and single spaces are kept. Callers that compare titles use
    `_table_match_key`. Never use this to rewrite a title stored in Excel.
    """
    if not text:
        return ""
    folded = unicodedata.normalize("NFKC", str(text))
    folded = _INVISIBLE_RE.sub("", folded)
    folded = _WS_RE.sub(" ", folded).strip().casefold()
    folded = _TABLE_TOKEN_RE.sub("табела", folded)
    return folded.translate(_LOOKALIKE_TO_CYRILLIC)


def _table_match_key(text):
    """Comparison key with whitespace and light punctuation removed."""
    return _COMPARE_STRIP_RE.sub("", _canonicalize_table_text(text))


def _is_table_title(text):
    """True when text looks like a 'Табела N. ...' title after folding."""
    return _TITLE_RE.match(_canonicalize_table_text(text)) is not None
```

- [ ] **Step 4: Match folded keys in `find_table_pages`**

Replace the body of `find_table_pages` so it no longer calls `re.search` on `tdef["pattern"]` and no longer only swaps newlines. Keep the warning string exact.

```python
def find_table_pages(page_texts, log=None):
    """Find which pages contain which tables. Returns dict: (xlsx_name, table_id) -> [page_numbers]."""
    if log is None:
        log = print
    table_pages = {}
    page_keys = {
        pg_num: _table_match_key(text) for pg_num, text in page_texts.items()
    }

    for xlsx_name, table_defs in TABLE_PATTERNS.items():
        for tdef in table_defs:
            tid = tdef["id"]
            needle = _table_match_key(tdef["title"])
            pages = []
            if needle:
                for pg_num, haystack in sorted(page_keys.items()):
                    if needle in haystack:
                        pages.append(pg_num)
            if pages:
                table_pages[(xlsx_name, tid)] = pages
            else:
                log(f"  WARNING: Could not find {xlsx_name}/{tid}")

    return table_pages
```

- [ ] **Step 5: Replace every `_TABELA_RE` use with `_is_table_title`**

In `_reconstruct_headers`, change:

```python
        if _TABELA_RE.match(text):
```

to:

```python
        if _is_table_title(text):
```

Delete the `_TABELA_RE = re.compile(...)` assignment.

In `_extract_title_from_page`, use `_is_table_title(line)` and `_is_table_title(lines[0] + ' ' + lines[1])`. Still return the original line text, not the folded form.

In `consolidate_title_row`, use `_is_table_title(v.strip())` and `_is_table_title(l.strip())` in place of `_TABELA_RE.match(...)`.

Do not change the rest of header reconstruction, title insertion, or empty-row dropping.

- [ ] **Step 6: Run the focused tests**

```bash
if [ -x .venv/bin/python ]; then PY=.venv/bin/python; else PY=python3; fi
$PY -m unittest discover -s tests -p 'test_table_detection.py' -v
```

Expected: all tests in `test_table_detection.py` pass, including the pre-existing RFPIO, Vojvodina, and empty-column tests.

- [ ] **Step 7: Run complete unit-test discovery**

```bash
if [ -x .venv/bin/python ]; then PY=.venv/bin/python; else PY=python3; fi
$PY -m unittest discover -s tests -v
```

Expected: every discovered test passes. If discovery fails because a dependency is missing, install from `requirements.txt` into the existing venv or a local venv and rerun. Do not pin or upgrade packages as part of this change.

- [ ] **Step 8: Confirm the diff is scoped, then commit**

```bash
git status --short
git diff -- extract_tables.py tests/test_table_detection.py
```

Expected: only `extract_tables.py` and `tests/test_table_detection.py` are in the implementation diff. Do not stage `.claude/`, generated `tabele/` files, or the plan file if it was already committed.

```bash
git add extract_tables.py tests/test_table_detection.py
git commit -m "fix(extraction): fold script, case, and whitespace in table detection"
```

No `Co-Authored-By` trailer. Do not push.
