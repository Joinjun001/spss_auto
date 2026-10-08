"""실제 통계 결과만 사용해 재현 가능한 설명을 만든다."""
import math

from .schema import AnalysisMethod, AnalysisPlan, AnalysisResult


def _p_text(p: float) -> str:
    if not math.isfinite(p):
        return '계산되지 않음'
    return '< .001' if p < 0.001 else f'= {p:.3f}'


def summarize(plan: AnalysisPlan, result: AnalysisResult) -> list[str]:
    m = result.metrics
    if result.method == AnalysisMethod.PEARSON:
        significance = '통계적으로 유의한' if m['p'] < 0.05 else '통계적으로 유의하다고 보기 어려운'
        return [
            f"{m['x']}와 {m['y']} 사이에는 r={m['r']:.3f}, 95% CI [{m['ci_low']:.3f}, {m['ci_high']:.3f}], p {_p_text(m['p'])}의 {significance} 선형 상관이 관찰되었습니다.",
            '상관관계만으로 인과관계를 결론내릴 수 없습니다. 연구 설계와 잠재적 교란변수를 함께 검토하세요.',
        ]
    if result.method == AnalysisMethod.WELCH_T:
        significance = '유의한 평균 차이가 관찰되었습니다.' if m['p'] < 0.05 else '유의한 평균 차이라고 보기 어렵습니다.'
        return [
            f"{m['group_a']}와 {m['group_b']}의 {m['outcome']} 평균 차이는 {m['mean_diff']:.3f}이었고, Welch t({m['df']:.1f})={m['t']:.3f}, p {_p_text(m['p'])}, Hedges' g={m['hedges_g']:.3f}로 {significance}",
            '집단 배정이 무작위가 아니라면 이 차이를 집단 특성의 인과효과로 해석하면 안 됩니다.',
        ]
    if result.method == AnalysisMethod.ONE_WAY_ANOVA:
        significance = '집단 간 평균 차이가 관찰되었습니다.' if m['p'] < 0.05 else '집단 간 평균 차이가 유의하다고 보기 어렵습니다.'
        return [
            f"{m['outcome']}에 대한 일원분산분석 결과 F({m['df_between']}, {m['df_within']})={m['f']:.3f}, p {_p_text(m['p'])}, η²={m['eta_squared']:.3f}로 {significance}",
            '유의한 경우 어느 집단 간 차이인지 확인하려면 다중비교를 별도로 수행해야 합니다.',
        ]
    if result.method == AnalysisMethod.LINEAR_REGRESSION:
        significance = '모형 전체가 유의했습니다.' if m['f_p'] < 0.05 else '모형 전체가 유의하다고 보기 어렵습니다.'
        return [
            f"{m['outcome']}을 종속변수로 한 선형 회귀모형의 R²={m['r_squared']:.3f}, 조정 R²={m['adj_r_squared']:.3f}, F={m['f']:.3f}, p {_p_text(m['f_p'])}로 {significance}",
            '계수의 통계적 유의성과 별개로 선형성·등분산성·잔차·다중공선성 진단을 함께 확인해야 하며, 관찰자료 회귀계수는 자동으로 인과효과를 의미하지 않습니다.',
        ]
    return ['결과 요약을 지원하지 않는 분석입니다.']
