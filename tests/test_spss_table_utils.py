"""SPSS 표 분리 및 결측치 처리 회귀 테스트."""

import unittest

import numpy as np
import pandas as pd

from spss_table_utils import (
    extract_spss_tables,
    prepare_spss_dataframe,
    replace_missing_stats,
)


class SpssTableUtilsTests(unittest.TestCase):
    def test_numeric_zero_is_preserved(self):
        frame = pd.DataFrame([[0.0, 1.234, -2]])
        result = prepare_spss_dataframe(frame)
        self.assertEqual(result.iloc[0].tolist(), ["0.00", "1.23", "-2.00"])

    def test_empty_cells_remain_missing(self):
        frame = pd.DataFrame([[None, np.nan, "   "]])
        result = prepare_spss_dataframe(frame)
        self.assertTrue(result.iloc[0].isna().all())

    def test_blank_row_separates_tables(self):
        frame = pd.DataFrame([
            ["집단통계량", None],
            ["A", 1.5],
            [None, None],
            ["ANOVA", None],
            ["F", 3.2],
        ])
        result = extract_spss_tables(prepare_spss_dataframe(frame), ["집단통계량", "ANOVA"])
        self.assertEqual(len(result["집단통계량"][0]), 2)
        self.assertEqual(len(result["ANOVA"][0]), 2)

    def test_next_heading_separates_tables_without_blank_row(self):
        frame = pd.DataFrame([["집단통계량"], ["A"], ["ANOVA"], ["F"]])
        result = extract_spss_tables(frame, ["집단통계량", "ANOVA"])
        self.assertEqual(result["집단통계량"][0].iloc[-1, 0], "A")
        self.assertEqual(result["ANOVA"][0].iloc[0, 0], "ANOVA")

    def test_scheffe_stops_at_significance_row(self):
        frame = pd.DataFrame([
            ["Scheffe", None],
            ["집단A", 12],
            ["CTT 유의확률", 0.05],
            ["다른 설명", 7],
        ])
        result = extract_spss_tables(frame, ["Scheffe"])
        self.assertEqual(len(result["Scheffe"][0]), 3)

    def test_missing_statistics_are_not_fabricated_as_zero(self):
        frame = pd.DataFrame([["nan", "nan±nan", "1.23±nan", np.nan, "0.00", ""]])
        result = replace_missing_stats(frame)
        self.assertEqual(result.iloc[0].tolist(), ["—", "—", "—", "—", "0.00", ""])

    def test_no_matching_heading(self):
        frame = pd.DataFrame([["설명", 1]])
        self.assertEqual(extract_spss_tables(frame, ["ANOVA"]), {"ANOVA": []})


if __name__ == "__main__":
    unittest.main()
