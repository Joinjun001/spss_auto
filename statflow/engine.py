"""AI가 아닌 검증 가능한 통계 라이브러리가 실제 수치를 계산한다."""
import numpy as np
import pandas as pd
import statsmodels.api as sm
from scipy import stats
from statsmodels.stats.diagnostic import het_breuschpagan
from statsmodels.stats.outliers_influence import variance_inflation_factor

from .schema import AnalysisMethod, AnalysisPlan, AnalysisResult, ValidationCheck


class EngineError(ValueError):
    pass


def _numeric(series: pd.Series) -> pd.Series:
    return pd.to_numeric(series, errors='coerce').replace([np.inf, -np.inf], np.nan)


def run_analysis(frame: pd.DataFrame, plan: AnalysisPlan) -> AnalysisResult:
    if plan.method == AnalysisMethod.PEARSON:
        x, y = plan.predictors[:2]
        data = pd.DataFrame({x: _numeric(frame[x]), y: _numeric(frame[y])}).dropna()
        result = stats.pearsonr(data[x], data[y])
        ci = result.confidence_interval(confidence_level=0.95)
        return AnalysisResult(plan.method, metrics={
            'x': x, 'y': y, 'n': len(data), 'r': float(result.statistic), 'p': float(result.pvalue),
            'ci_low': float(ci.low), 'ci_high': float(ci.high),
        })

    if plan.method == AnalysisMethod.WELCH_T:
        outcome, group = plan.outcome, plan.group
        data = pd.DataFrame({'outcome': _numeric(frame[outcome]), 'group': frame[group]}).dropna()
        labels = list(pd.unique(data['group']))
        if len(labels) != 2:
            raise EngineError('Welch t-test에는 두 집단이 필요합니다.')
        a = data.loc[data['group'] == labels[0], 'outcome'].to_numpy(float)
        b = data.loc[data['group'] == labels[1], 'outcome'].to_numpy(float)
        test = stats.ttest_ind(a, b, equal_var=False)
        diff = float(np.mean(a) - np.mean(b))
        va, vb = np.var(a, ddof=1), np.var(b, ddof=1)
        se2 = va / len(a) + vb / len(b)
        df = se2 ** 2 / ((va / len(a)) ** 2 / (len(a) - 1) + (vb / len(b)) ** 2 / (len(b) - 1))
        critical = stats.t.ppf(0.975, df)
        pooled = np.sqrt(((len(a)-1)*va + (len(b)-1)*vb) / (len(a)+len(b)-2))
        d = diff / pooled if pooled else np.nan
        correction = 1 - 3 / (4 * (len(a) + len(b)) - 9)
        g = d * correction if np.isfinite(d) else np.nan
        table = pd.DataFrame({
            '집단': labels,
            'N': [len(a), len(b)],
            '평균': [np.mean(a), np.mean(b)],
            '표준편차': [np.std(a, ddof=1), np.std(b, ddof=1)],
        })
        return AnalysisResult(plan.method, metrics={
            'outcome': outcome, 'group': group, 'group_a': str(labels[0]), 'group_b': str(labels[1]),
            'mean_diff': diff, 't': float(test.statistic), 'df': float(df), 'p': float(test.pvalue),
            'ci_low': float(diff - critical * np.sqrt(se2)), 'ci_high': float(diff + critical * np.sqrt(se2)),
            'hedges_g': float(g),
        }, tables={'집단기술통계': table})

    if plan.method == AnalysisMethod.ONE_WAY_ANOVA:
        outcome, group = plan.outcome, plan.group
        data = pd.DataFrame({'outcome': _numeric(frame[outcome]), 'group': frame[group]}).dropna()
        grouped = [(label, values['outcome'].to_numpy(float)) for label, values in data.groupby('group', observed=True)]
        if len(grouped) < 3:
            raise EngineError('ANOVA에는 세 집단 이상이 필요합니다.')
        test = stats.f_oneway(*(values for _, values in grouped))
        grand = data['outcome'].mean()
        ss_between = sum(len(values) * (np.mean(values) - grand) ** 2 for _, values in grouped)
        ss_total = float(((data['outcome'] - grand) ** 2).sum())
        eta2 = ss_between / ss_total if ss_total else np.nan
        table = pd.DataFrame([
            {'집단': str(label), 'N': len(values), '평균': np.mean(values), '표준편차': np.std(values, ddof=1)}
            for label, values in grouped
        ])
        return AnalysisResult(plan.method, metrics={
            'outcome': outcome, 'group': group, 'groups': len(grouped),
            'df_between': len(grouped) - 1, 'df_within': len(data) - len(grouped),
            'f': float(test.statistic), 'p': float(test.pvalue), 'eta_squared': float(eta2),
        }, tables={'집단기술통계': table})

    if plan.method == AnalysisMethod.LINEAR_REGRESSION:
        outcome = plan.outcome
        predictors = list(plan.predictors)
        data = pd.DataFrame({outcome: _numeric(frame[outcome]), **{p: _numeric(frame[p]) for p in predictors}}).dropna()
        y = data[outcome].astype(float)
        x = sm.add_constant(data[predictors].astype(float), has_constant='add')
        model = sm.OLS(y, x).fit()
        conf = model.conf_int(alpha=0.05)
        coeff = pd.DataFrame({
            '변수': model.params.index,
            '계수': model.params.values,
            '표준오차': model.bse.values,
            't': model.tvalues.values,
            'p': model.pvalues.values,
            '95% CI 하한': conf[0].values,
            '95% CI 상한': conf[1].values,
        })
        diagnostics: list[ValidationCheck] = []
        if len(model.resid) >= 3:
            if len(model.resid) <= 5000:
                shapiro = stats.shapiro(model.resid)
                diagnostics.append(ValidationCheck('잔차 정규성(Shapiro-Wilk)',
                    'pass' if shapiro.pvalue >= 0.05 else 'warning', f'p={shapiro.pvalue:.4g}'))
            bp = het_breuschpagan(model.resid, model.model.exog)
            diagnostics.append(ValidationCheck('등분산성(Breusch-Pagan)',
                'pass' if bp[1] >= 0.05 else 'warning', f'p={bp[1]:.4g}'))
        if len(predictors) >= 2:
            vif_rows = []
            for index, name in enumerate(x.columns):
                if name == 'const':
                    continue
                vif_rows.append({'변수': name, 'VIF': variance_inflation_factor(x.values, index)})
            vif_table = pd.DataFrame(vif_rows)
        else:
            vif_table = pd.DataFrame(columns=['변수', 'VIF'])
        return AnalysisResult(plan.method, metrics={
            'outcome': outcome, 'predictors': predictors, 'n': int(model.nobs), 'r_squared': float(model.rsquared),
            'adj_r_squared': float(model.rsquared_adj), 'f': float(model.fvalue) if model.fvalue is not None else np.nan,
            'f_p': float(model.f_pvalue) if model.f_pvalue is not None else np.nan,
        }, tables={'회귀계수': coeff, 'VIF': vif_table}, diagnostics=diagnostics)

    raise EngineError('지원하지 않는 분석 방법입니다.')
