import unittest

import pandas as pd

from extract_tables import find_label_cols_count, find_table_pages


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


class FindLabelColsCountTests(unittest.TestCase):
    def test_counts_all_missing_leading_column_as_label(self):
        df = pd.DataFrame({0: [None, None], 1: ["1", "2"]})

        self.assertEqual(find_label_cols_count(df), 1)


if __name__ == "__main__":
    unittest.main()
