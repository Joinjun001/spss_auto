"""StatFlow의 분석 계획과 검증 결과를 표현하는 작은 도메인 모델."""
from dataclasses import dataclass, field
from enum import Enum


class AnalysisMethod(str, Enum):
    PEARSON = 'pearson_correlation'
    WELCH_T = 'welch_t_test'
    ONE_WAY_ANOVA = 'one_way_anova'
    LINEAR_REGRESSION = 'linear_regression'


METHOD_LABELS = {
    AnalysisMethod.PEARSON: 'Pearson 상관분석',
    AnalysisMethod.WELCH_T: 'Welch 독립표본 t-test',
    AnalysisMethod.ONE_WAY_ANOVA: '일원분산분석(One-way ANOVA)',
    AnalysisMethod.LINEAR_REGRESSION: '선형 회귀분석',
}

METHOD_ASSUMPTIONS = {
    AnalysisMethod.PEARSON: ['두 변수의 수치형 측정', '관측치 독립성', '선형 관계', '극단치 점검'],
    AnalysisMethod.WELCH_T: ['수치형 종속변수', '서로 독립인 두 집단', '관측치 독립성', '집단별 극단치·분포 점검'],
    AnalysisMethod.ONE_WAY_ANOVA: ['수치형 종속변수', '서로 독립인 3개 이상 집단', '관측치 독립성', '집단별 극단치·분포 점검'],
    AnalysisMethod.LINEAR_REGRESSION: ['수치형 종속·독립변수', '선형성', '독립성', '잔차 등분산성', '다중공선성 점검'],
}


@dataclass(frozen=True)
class AnalysisPlan:
    method: AnalysisMethod
    research_question: str
    outcome: str | None = None
    predictors: tuple[str, ...] = ()
    group: str | None = None
    rationale: tuple[str, ...] = ()
    assumptions: tuple[str, ...] = ()
    confidence: str = 'medium'


@dataclass(frozen=True)
class ValidationCheck:
    name: str
    status: str  # pass | warning | fail
    message: str


@dataclass
class AnalysisResult:
    method: AnalysisMethod
    metrics: dict[str, object] = field(default_factory=dict)
    tables: dict[str, object] = field(default_factory=dict)
    diagnostics: list[ValidationCheck] = field(default_factory=list)
