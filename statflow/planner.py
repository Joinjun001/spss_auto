"""연구 질문을 구조화된 통계 분석 계획으로 바꾸는 교체 가능한 planning layer.

현재 기본 구현은 API 키 없이 동작하는 설명 가능한 규칙 기반 planner다.
향후 LLM provider는 같은 AnalysisPlan 스키마만 반환하면 교체할 수 있다.
"""
import re

import pandas as pd

from .data import level_map
from .schema import AnalysisMethod, AnalysisPlan, METHOD_ASSUMPTIONS


def _normalize(text: str) -> str:
    return re.sub(r'[\s_\-]+', '', text).lower()


def mentioned_columns(question: str, columns: list[str]) -> list[str]:
    normalized_question = _normalize(question)
    hits = []
    for col in columns:
        token = _normalize(str(col))
        if token and token in normalized_question:
            hits.append(str(col))
    return hits


def recommend_plan(question: str, frame: pd.DataFrame) -> AnalysisPlan:
    levels = level_map(frame)
    mentions = mentioned_columns(question, [str(c) for c in frame.columns])
    numeric = [c for c, level in levels.items() if level in {'연속형', '서열형'}]
    categorical = [c for c, level in levels.items() if level == '명목형']
    mentioned_numeric = [c for c in mentions if c in numeric]
    mentioned_categorical = [c for c in mentions if c in categorical]
    q = _normalize(question)

    difference_words = ('차이', '비교', '다른가', '높은가', '낮은가')
    influence_words = ('영향', '예측', '회귀', '통제', '설명')
    relation_words = ('상관', '관련', '관계', '연관')

    if any(word in q for word in difference_words) and (mentioned_categorical or categorical):
        group = (mentioned_categorical or categorical)[0]
        outcome_candidates = mentioned_numeric or [c for c in numeric if c != group]
        outcome = outcome_candidates[0] if outcome_candidates else None
        group_count = int(frame[group].nunique(dropna=True))
        method = AnalysisMethod.WELCH_T if group_count == 2 else AnalysisMethod.ONE_WAY_ANOVA
        reason = (
            f'질문이 집단 간 차이를 묻고 있고, {group}의 유효 집단 수가 {group_count}개입니다.',
            f'{outcome or "종속변수"}의 평균 차이를 검정하는 계획입니다.',
        )
        return AnalysisPlan(method, question, outcome=outcome, group=group, rationale=reason,
                            assumptions=tuple(METHOD_ASSUMPTIONS[method]), confidence='high' if mentions else 'medium')

    if any(word in q for word in influence_words):
        candidates = mentioned_numeric or numeric
        predictor = candidates[0] if candidates else None
        outcome = candidates[1] if len(candidates) > 1 else (numeric[1] if len(numeric) > 1 else None)
        predictors = (predictor,) if predictor and predictor != outcome else ()
        reason = (
            '질문이 영향·예측 관계를 묻고 있어 회귀모형을 우선 제안합니다.',
            '추천은 인과관계를 보장하지 않으며 연구 설계와 변수 역할을 사용자가 확인해야 합니다.',
        )
        return AnalysisPlan(AnalysisMethod.LINEAR_REGRESSION, question, outcome=outcome,
                            predictors=predictors, rationale=reason,
                            assumptions=tuple(METHOD_ASSUMPTIONS[AnalysisMethod.LINEAR_REGRESSION]),
                            confidence='high' if len(mentioned_numeric) >= 2 else 'medium')

    if any(word in q for word in relation_words) or len(mentioned_numeric) >= 2:
        candidates = mentioned_numeric or numeric
        predictors = tuple(candidates[:2])
        reason = (
            '질문이 두 수치형 변수의 관계를 묻고 있어 Pearson 상관분석을 우선 제안합니다.',
            '선형 관계와 극단치를 확인한 뒤 결과를 해석해야 합니다.',
        )
        return AnalysisPlan(AnalysisMethod.PEARSON, question, predictors=predictors, rationale=reason,
                            assumptions=tuple(METHOD_ASSUMPTIONS[AnalysisMethod.PEARSON]),
                            confidence='high' if len(mentioned_numeric) >= 2 else 'low')

    candidates = mentioned_numeric or numeric
    return AnalysisPlan(
        AnalysisMethod.PEARSON,
        question,
        predictors=tuple(candidates[:2]),
        rationale=('질문의 통계적 의도가 명확하지 않아 탐색적 상관분석을 임시 후보로 제안합니다.',
                   '분석 방법과 변수 역할을 반드시 직접 확인하세요.'),
        assumptions=tuple(METHOD_ASSUMPTIONS[AnalysisMethod.PEARSON]),
        confidence='low',
    )
