"""Provider 독립적인 LLM structured-output planner 경계.

외부 모델은 통계값을 계산하지 않고 분석 계획 후보만 반환한다. 이 모듈은
외부 출력을 신뢰하지 않고 JSON schema, 허용 변수 목록, 방법별 역할 규칙으로
검증한 뒤 기존 ``AnalysisPlan`` 도메인 모델로 변환한다.
"""
from __future__ import annotations

import json
from dataclasses import dataclass
from typing import Any, Literal, Mapping, Protocol

import pandas as pd
from pydantic import BaseModel, ConfigDict, Field, ValidationError, field_validator

from .data import profile_dataset
from .schema import AnalysisMethod, AnalysisPlan, METHOD_ASSUMPTIONS


class PlannerOutputError(ValueError):
    """LLM planner 출력이 StatFlow의 안전 경계를 만족하지 못함."""


class StructuredOutputProvider(Protocol):
    """특정 LLM SDK에 종속되지 않는 최소 provider 계약."""

    def generate_structured(
        self,
        *,
        system_prompt: str,
        input_payload: Mapping[str, Any],
        json_schema: Mapping[str, Any],
    ) -> Mapping[str, Any] | str:
        """JSON 객체 또는 JSON 문자열 형태의 structured output을 반환한다."""


class LLMPlanOutput(BaseModel):
    """외부 모델이 반환할 수 있는 필드의 전체 목록.

    통계 수치 필드는 의도적으로 존재하지 않으며, 추가 필드는 모두 거부한다.
    """

    model_config = ConfigDict(extra='forbid', str_strip_whitespace=True)

    method: AnalysisMethod
    outcome: str | None = None
    predictors: list[str] = Field(default_factory=list, max_length=10)
    group: str | None = None
    rationale: list[str] = Field(min_length=1, max_length=4)
    confidence: Literal['low', 'medium', 'high'] = 'medium'
    follow_up_questions: list[str] = Field(default_factory=list, max_length=3)

    @field_validator('predictors')
    @classmethod
    def predictors_are_unique(cls, value: list[str]) -> list[str]:
        if len(value) != len(set(value)):
            raise ValueError('predictors에는 중복 변수를 사용할 수 없습니다.')
        return value


@dataclass(frozen=True)
class LLMPlanningResult:
    plan: AnalysisPlan
    follow_up_questions: tuple[str, ...] = ()
    source: str = 'llm'


SYSTEM_PROMPT = """당신은 연구 통계 분석 계획을 제안하는 planner입니다.
반드시 제공된 JSON schema에 맞는 분석 계획만 반환하세요.
허용된 변수명과 분석 방법만 사용하세요.
p-value, 효과크기, 신뢰구간, 회귀계수 또는 임의의 통계 수치를 계산하거나 생성하지 마세요.
연구 질문만으로 인과관계를 확정하지 말고, 불확실하면 confidence를 낮추고 follow_up_questions를 사용하세요.
실제 통계 계산은 별도의 검증 가능한 엔진이 수행합니다."""


def planner_json_schema() -> dict[str, Any]:
    """Provider의 structured-output 기능에 전달할 JSON schema."""
    return LLMPlanOutput.model_json_schema()


def build_planner_payload(question: str, frame: pd.DataFrame) -> dict[str, Any]:
    """원자료 행은 보내지 않고 분석 계획에 필요한 변수 메타데이터만 만든다."""
    profile = profile_dataset(frame)
    variables = []
    for row in profile.to_dict(orient='records'):
        variables.append({
            'name': row['변수'],
            'measurement_level': row['추정 측정수준'],
            'valid_n': int(row['유효 N']),
            'missing_n': int(row['결측 N']),
            'unique_count': int(row['고유값']),
        })
    return {
        'research_question': question,
        'variables': variables,
        'allowed_methods': [method.value for method in AnalysisMethod],
        'constraints': {
            'do_not_compute_statistics': True,
            'use_only_listed_variables': True,
            'causal_claims_require_design_evidence': True,
        },
    }


def _coerce_output(raw: Mapping[str, Any] | str) -> Mapping[str, Any]:
    if isinstance(raw, str):
        try:
            decoded = json.loads(raw)
        except json.JSONDecodeError as exc:
            raise PlannerOutputError('LLM planner가 유효한 JSON을 반환하지 않았습니다.') from exc
        if not isinstance(decoded, dict):
            raise PlannerOutputError('LLM planner 출력은 JSON 객체여야 합니다.')
        return decoded
    if isinstance(raw, Mapping):
        return raw
    raise PlannerOutputError('LLM planner 출력 형식을 해석할 수 없습니다.')


def _validate_variables(output: LLMPlanOutput, allowed: set[str]) -> None:
    referenced = [*output.predictors]
    if output.outcome:
        referenced.append(output.outcome)
    if output.group:
        referenced.append(output.group)
    unknown = sorted(set(referenced) - allowed)
    if unknown:
        raise PlannerOutputError('데이터에 없는 변수를 사용했습니다: ' + ', '.join(unknown))


def _validate_roles(output: LLMPlanOutput) -> None:
    method = output.method
    predictors = output.predictors

    if method == AnalysisMethod.PEARSON:
        if len(predictors) != 2 or output.outcome is not None or output.group is not None:
            raise PlannerOutputError('Pearson 상관분석은 predictors 두 개만 사용해야 합니다.')
        if predictors[0] == predictors[1]:
            raise PlannerOutputError('상관분석 변수는 서로 달라야 합니다.')
        return

    if method == AnalysisMethod.LINEAR_REGRESSION:
        if not output.outcome or not predictors or output.group is not None:
            raise PlannerOutputError('선형 회귀는 outcome과 하나 이상의 predictors가 필요합니다.')
        if output.outcome in predictors:
            raise PlannerOutputError('종속변수를 독립변수로 동시에 사용할 수 없습니다.')
        return

    if method in {AnalysisMethod.WELCH_T, AnalysisMethod.ONE_WAY_ANOVA}:
        if not output.outcome or not output.group or predictors:
            raise PlannerOutputError('집단 비교 분석은 outcome과 group만 사용해야 합니다.')
        if output.outcome == output.group:
            raise PlannerOutputError('종속변수와 집단변수는 서로 달라야 합니다.')
        return

    raise PlannerOutputError('허용되지 않은 분석 방법입니다.')


def parse_llm_plan(
    raw: Mapping[str, Any] | str,
    *,
    question: str,
    frame: pd.DataFrame,
) -> LLMPlanningResult:
    """외부 structured output을 신뢰 경계 안의 AnalysisPlan으로 변환한다."""
    try:
        output = LLMPlanOutput.model_validate(_coerce_output(raw))
    except ValidationError as exc:
        raise PlannerOutputError('LLM planner 출력 schema 검증에 실패했습니다.') from exc

    _validate_variables(output, {str(column) for column in frame.columns})
    _validate_roles(output)

    # 가정 목록은 LLM이 자유롭게 만들지 않고 StatFlow registry의 canonical 값을 사용한다.
    plan = AnalysisPlan(
        method=output.method,
        research_question=question,
        outcome=output.outcome,
        predictors=tuple(output.predictors),
        group=output.group,
        rationale=tuple(output.rationale),
        assumptions=tuple(METHOD_ASSUMPTIONS[output.method]),
        confidence=output.confidence,
    )
    return LLMPlanningResult(
        plan=plan,
        follow_up_questions=tuple(output.follow_up_questions),
    )


def plan_with_provider(
    provider: StructuredOutputProvider,
    *,
    question: str,
    frame: pd.DataFrame,
) -> LLMPlanningResult:
    """Provider 호출부터 schema/semantic validation까지 한 경계에서 처리한다."""
    raw = provider.generate_structured(
        system_prompt=SYSTEM_PROMPT,
        input_payload=build_planner_payload(question, frame),
        json_schema=planner_json_schema(),
    )
    return parse_llm_plan(raw, question=question, frame=frame)
