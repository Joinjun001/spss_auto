from io import BytesIO

import numpy as np
import pandas as pd
import pytest
from openpyxl import load_workbook
from scipy import stats
from survey_core import (AnalysisError, correlation, descriptives, frequencies,
                         missing_codes, profile, quality, read_survey, to_excel)


def test_csv_cp949_and_whitespace():
    frame = read_survey('이름,점수\n가,3\n나, \n'.encode('cp949'), 'test.csv')
    assert pd.isna(frame.loc[1, '점수'])
    assert frame.loc[0, '이름'] == '가'


def test_xlsx():
    out = BytesIO()
    pd.DataFrame({'a': [1, 2]}).to_excel(out, index=False)
    assert len(read_survey(out.getvalue(), 'test.xlsx')) == 2


def test_invalid_file():
    with pytest.raises(AnalysisError):
        read_survey(b'abc', 'file.sav')


def test_missing_code_no_mutation():
    frame = pd.DataFrame({'x': [0, 99, 2]})
    result = missing_codes(frame, ['99'])
    assert pd.isna(result.loc[1, 'x'])
    assert frame.loc[1, 'x'] == 99
    assert result.loc[0, 'x'] == 0


def test_profile_and_quality():
    frame = pd.DataFrame({'id': [1, 2, 3], 'likert': [1, 2, 3], 'fixed': [0, 0, 0]})
    assert profile(frame).set_index('변수').loc['id', '측정 수준'] == 'ID'
    assert any('fixed' in note for note in quality(frame))


def test_descriptive_ddof():
    result = descriptives(pd.DataFrame({'x': [1, 2, 3, np.nan]}), ['x']).iloc[0]
    assert result['N'] == 3
    assert result['표준편차'] == pytest.approx(1)


def test_frequency_denominators():
    frame = pd.DataFrame({'x': ['a', 'a', 'b', None]})
    result = frequencies(frame, 'x').set_index('값')
    assert result.loc['a', '전체 %'] == pytest.approx(50)
    assert result.loc['a', '유효 %'] == pytest.approx(200 / 3)


def test_pearson_matches_scipy():
    frame = pd.DataFrame({'x': [1, 2, 3, 4, 5, 6], 'y': [3, 2, 5, 6, 4, 8]})
    row = correlation(frame, 'x', 'y', 'pearson').iloc[0]
    expected = stats.pearsonr(frame.x, frame.y)
    assert row['상관계수'] == pytest.approx(expected.statistic)
    assert row['양측 p'] == pytest.approx(expected.pvalue)
    assert row['95% CI 하한'] <= row['상관계수'] <= row['95% CI 상한']


def test_spearman_small_n_does_not_report_p():
    frame = pd.DataFrame({'x': [1, 2, 3, 4, 5], 'y': [3, 2, 5, 1, 4]})
    row = correlation(frame, 'x', 'y', 'spearman').iloc[0]
    assert row['상관계수'] == pytest.approx(stats.spearmanr(frame.x, frame.y).statistic)
    assert pd.isna(row['양측 p'])


def test_correlation_invalid():
    with pytest.raises(AnalysisError):
        correlation(pd.DataFrame({'x': [1, 1, 1], 'y': [1, 2, 3]}), 'x', 'y', 'pearson')


def test_excel_formula_injection():
    data = to_excel({'빈도표': pd.DataFrame({'값': ['=1+1'], '빈도': [1]})})
    sheet = load_workbook(BytesIO(data))['빈도표']
    assert sheet['A2'].data_type != 'f'
    assert sheet['A2'].value == "'=1+1"
