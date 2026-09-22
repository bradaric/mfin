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
