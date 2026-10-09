from pathlib import Path

import pytest

from statflow.data import read_dataset
from statflow.llm_planner import (
    PlannerOutputError,
    build_planner_payload,
    parse_llm_plan,
    plan_with_provider,
    planner_json_schema,
)
from statflow.schema import AnalysisMethod, METHOD_ASSUMPTIONS


def sample_frame():
    path = Path('sample_data/example_survey.csv')
    return read_dataset(path.read_bytes(), path.name)


def test_payload_exposes_metadata_not_raw_rows():
    payload = build_planner_payload('학습시간과 스트레스점수의 관계는?', sample_frame())
    assert payload['research_question']
    assert 'variables' in payload
    assert 'rows' not in payload
    assert all('name' in variable for variable in payload['variables'])
    assert all('values' not in variable for variable in payload['variables'])


def test_json_schema_forbids_extra_fields():
    schema = planner_json_schema()
    assert schema['additionalProperties'] is False
    assert 'method' in schema['properties']
    assert 'p_value' not in schema['properties']


def test_parse_valid_regression_plan_uses_canonical_assumptions():
    frame = sample_frame()
    result = parse_llm_plan({
        'method': 'linear_regression',
        'outcome': '스트레스점수',
        'predictors': ['학습시간'],
        'group': None,
        'rationale': ['학습시간과 스트레스점수의 예측 관계를 검토합니다.'],
        'confidence': 'high',
        'follow_up_questions': ['연구 설계는 횡단 연구인가요?'],
    }, question='학습시간이 스트레스점수에 영향을 미치는가?', frame=frame)

    assert result.plan.method == AnalysisMethod.LINEAR_REGRESSION
    assert result.plan.outcome == '스트레스점수'
    assert result.plan.predictors == ('학습시간',)
    assert result.plan.assumptions == tuple(METHOD_ASSUMPTIONS[AnalysisMethod.LINEAR_REGRESSION])
    assert result.follow_up_questions == ('연구 설계는 횡단 연구인가요?',)


def test_rejects_unknown_variable():
    with pytest.raises(PlannerOutputError, match='없는 변수'):
        parse_llm_plan({
            'method': 'pearson_correlation',
            'outcome': None,
            'predictors': ['학습시간', '환각변수'],
            'group': None,
            'rationale': ['관계를 확인합니다.'],
            'confidence': 'medium',
            'follow_up_questions': [],
        }, question='q', frame=sample_frame())


def test_rejects_statistics_or_other_extra_fields():
    with pytest.raises(PlannerOutputError, match='schema'):
        parse_llm_plan({
            'method': 'pearson_correlation',
            'outcome': None,
            'predictors': ['학습시간', '스트레스점수'],
            'group': None,
            'rationale': ['관계를 확인합니다.'],
            'confidence': 'high',
            'follow_up_questions': [],
            'p_value': 0.001,
        }, question='q', frame=sample_frame())


def test_rejects_invalid_method_specific_roles():
    with pytest.raises(PlannerOutputError, match='Pearson'):
        parse_llm_plan({
            'method': 'pearson_correlation',
            'outcome': '스트레스점수',
            'predictors': ['학습시간'],
            'group': None,
            'rationale': ['잘못된 역할입니다.'],
            'confidence': 'low',
            'follow_up_questions': [],
        }, question='q', frame=sample_frame())


def test_provider_adapter_receives_schema_and_returns_validated_plan():
    class FakeProvider:
        def __init__(self):
            self.call = None

        def generate_structured(self, *, system_prompt, input_payload, json_schema):
            self.call = (system_prompt, input_payload, json_schema)
            return {
                'method': 'welch_t_test',
                'outcome': '스트레스점수',
                'predictors': [],
                'group': '성별',
                'rationale': ['두 독립 집단의 평균 차이를 검토합니다.'],
                'confidence': 'high',
                'follow_up_questions': [],
            }

    provider = FakeProvider()
    result = plan_with_provider(provider, question='성별에 따라 스트레스점수 차이가 있는가?', frame=sample_frame())
    assert result.plan.method == AnalysisMethod.WELCH_T
    assert provider.call is not None
    assert provider.call[1]['constraints']['do_not_compute_statistics'] is True
    assert provider.call[2]['additionalProperties'] is False


def test_json_string_is_supported_but_non_object_is_rejected():
    frame = sample_frame()
    raw = '{"method":"pearson_correlation","outcome":null,"predictors":["학습시간","스트레스점수"],"group":null,"rationale":["관계를 확인합니다."],"confidence":"high","follow_up_questions":[]}'
    result = parse_llm_plan(raw, question='q', frame=frame)
    assert result.plan.method == AnalysisMethod.PEARSON

    with pytest.raises(PlannerOutputError, match='JSON 객체'):
        parse_llm_plan('[]', question='q', frame=frame)
