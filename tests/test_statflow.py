from pathlib import Path

import numpy as np
import pandas as pd
import pytest
from scipy import stats

from statflow.data import profile_dataset, read_dataset
from statflow.engine import run_analysis
from statflow.planner import recommend_plan
from statflow.report import summarize
from statflow.schema import AnalysisMethod, AnalysisPlan
from statflow.validation import has_failures, validate_plan


@pytest.fixture
def sample():
    path = Path('sample_data/example_survey.csv')
    return read_dataset(path.read_bytes(), path.name)


def test_sample_dataset_and_profile(sample):
    assert len(sample) == 40
    profile = profile_dataset(sample).set_index('변수')
    assert profile.loc['응답자ID', '추정 측정수준'] == 'ID'
    assert profile.loc['학습시간', '추정 측정수준'] == '연속형'


def test_planner_recommends_regression(sample):
    plan = recommend_plan('학습시간이 스트레스점수에 영향을 미치는가?', sample)
    assert plan.method == AnalysisMethod.LINEAR_REGRESSION
    assert plan.predictors == ('학습시간',)
    assert plan.outcome == '스트레스점수'


def test_planner_recommends_correlation(sample):
    plan = recommend_plan('학습시간과 스트레스점수는 관련이 있는가?', sample)
    assert plan.method == AnalysisMethod.PEARSON
    assert plan.predictors == ('학습시간', '스트레스점수')


def test_planner_recommends_t_test(sample):
    plan = recommend_plan('성별에 따라 스트레스점수 차이가 있는가?', sample)
    assert plan.method == AnalysisMethod.WELCH_T
    assert plan.group == '성별'
    assert plan.outcome == '스트레스점수'


def test_planner_recommends_anova(sample):
    plan = recommend_plan('지역에 따라 스트레스점수 차이가 있는가?', sample)
    assert plan.method == AnalysisMethod.ONE_WAY_ANOVA
    assert plan.group == '지역'


def test_validation_blocks_same_correlation_variable(sample):
    plan = AnalysisPlan(AnalysisMethod.PEARSON, 'q', predictors=('학습시간', '학습시간'))
    assert has_failures(validate_plan(sample, plan))


def test_correlation_engine_matches_scipy(sample):
    plan = AnalysisPlan(AnalysisMethod.PEARSON, 'q', predictors=('학습시간', '스트레스점수'))
    assert not has_failures(validate_plan(sample, plan))
    result = run_analysis(sample, plan)
    expected = stats.pearsonr(sample['학습시간'], sample['스트레스점수'])
    assert result.metrics['r'] == pytest.approx(expected.statistic)
    assert result.metrics['p'] == pytest.approx(expected.pvalue)


def test_welch_engine_matches_scipy(sample):
    plan = AnalysisPlan(AnalysisMethod.WELCH_T, 'q', outcome='스트레스점수', group='성별')
    result = run_analysis(sample, plan)
    groups = [g['스트레스점수'].to_numpy() for _, g in sample.groupby('성별')]
    expected = stats.ttest_ind(groups[0], groups[1], equal_var=False)
    assert result.metrics['t'] == pytest.approx(expected.statistic)
    assert result.metrics['p'] == pytest.approx(expected.pvalue)


def test_anova_engine_matches_scipy(sample):
    plan = AnalysisPlan(AnalysisMethod.ONE_WAY_ANOVA, 'q', outcome='스트레스점수', group='지역')
    result = run_analysis(sample, plan)
    groups = [g['스트레스점수'].to_numpy() for _, g in sample.groupby('지역')]
    expected = stats.f_oneway(*groups)
    assert result.metrics['f'] == pytest.approx(expected.statistic)
    assert result.metrics['p'] == pytest.approx(expected.pvalue)


def test_regression_engine_known_relation():
    x = np.arange(1, 31, dtype=float)
    frame = pd.DataFrame({'x': x, 'y': 2.0 + 3.0 * x})
    plan = AnalysisPlan(AnalysisMethod.LINEAR_REGRESSION, 'q', outcome='y', predictors=('x',))
    result = run_analysis(frame, plan)
    coefficients = result.tables['회귀계수'].set_index('변수')
    assert coefficients.loc['const', '계수'] == pytest.approx(2.0)
    assert coefficients.loc['x', '계수'] == pytest.approx(3.0)
    assert result.metrics['r_squared'] == pytest.approx(1.0)


def test_report_uses_computed_values(sample):
    plan = AnalysisPlan(AnalysisMethod.PEARSON, 'q', predictors=('학습시간', '스트레스점수'))
    result = run_analysis(sample, plan)
    text = ' '.join(summarize(plan, result))
    assert 'r=' in text
    assert '인과관계' in text
