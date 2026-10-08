"""계획을 실행하기 전에 데이터가 최소 조건을 만족하는지 검증한다."""
import numpy as np
import pandas as pd

from .schema import AnalysisMethod, AnalysisPlan, ValidationCheck


def _numeric(frame: pd.DataFrame, columns: list[str]) -> tuple[pd.DataFrame | None, list[ValidationCheck]]:
    checks: list[ValidationCheck] = []
    converted = pd.DataFrame(index=frame.index)
    for col in columns:
        if col not in frame:
            checks.append(ValidationCheck('변수 존재', 'fail', f'변수 {col}이 데이터에 없습니다.'))
            return None, checks
        values = pd.to_numeric(frame[col], errors='coerce')
        bad = frame[col].notna() & values.isna()
        if bad.any():
            checks.append(ValidationCheck('수치형 변환', 'fail', f'{col}에 숫자가 아닌 값 {int(bad.sum())}개가 있습니다.'))
        converted[col] = values.replace([np.inf, -np.inf], np.nan)
    return converted, checks


def validate_plan(frame: pd.DataFrame, plan: AnalysisPlan) -> list[ValidationCheck]:
    checks: list[ValidationCheck] = []

    if plan.method == AnalysisMethod.PEARSON:
        cols = list(plan.predictors[:2])
        if len(cols) != 2 or cols[0] == cols[1]:
            return [ValidationCheck('변수 역할', 'fail', '서로 다른 수치형 변수 두 개가 필요합니다.')]
        data, numeric_checks = _numeric(frame, cols)
        checks.extend(numeric_checks)
        if data is None or any(c.status == 'fail' for c in numeric_checks):
            return checks
        complete = data.dropna()
        n = len(complete)
        checks.append(ValidationCheck('완전 사례 수', 'pass' if n >= 10 else 'fail', f'유효 쌍 N={n}'))
        for col in cols:
            status = 'pass' if complete[col].nunique() >= 2 else 'fail'
            checks.append(ValidationCheck('변동성', status, f'{col} 고유값 {complete[col].nunique()}개'))
        if 10 <= n < 30:
            checks.append(ValidationCheck('표본 크기', 'warning', '표본이 작아 극단치와 분포의 영향을 크게 받을 수 있습니다.'))
        return checks

    if plan.method == AnalysisMethod.LINEAR_REGRESSION:
        cols = ([plan.outcome] if plan.outcome else []) + list(plan.predictors)
        if not plan.outcome or not plan.predictors or plan.outcome in plan.predictors:
            return [ValidationCheck('변수 역할', 'fail', '종속변수와 하나 이상의 서로 다른 독립변수가 필요합니다.')]
        data, numeric_checks = _numeric(frame, [c for c in cols if c])
        checks.extend(numeric_checks)
        if data is None or any(c.status == 'fail' for c in numeric_checks):
            return checks
        complete = data.dropna()
        n, p = len(complete), len(plan.predictors)
        minimum = max(20, 10 * (p + 1))
        checks.append(ValidationCheck('완전 사례 수', 'pass' if n >= minimum else 'warning', f'N={n}, 권장 최소 기준≈{minimum}'))
        if complete[plan.outcome].nunique() < 2:
            checks.append(ValidationCheck('종속변수 변동성', 'fail', '종속변수에 변동이 없습니다.'))
        matrix = complete[list(plan.predictors)].to_numpy(dtype=float)
        rank = np.linalg.matrix_rank(np.column_stack([np.ones(len(matrix)), matrix])) if len(matrix) else 0
        expected_rank = p + 1
        checks.append(ValidationCheck('설계행렬', 'pass' if rank == expected_rank else 'fail',
                                      f'rank={rank}, expected={expected_rank}'))
        return checks

    if plan.method in {AnalysisMethod.WELCH_T, AnalysisMethod.ONE_WAY_ANOVA}:
        if not plan.outcome or not plan.group or plan.outcome == plan.group:
            return [ValidationCheck('변수 역할', 'fail', '수치형 종속변수와 별도의 집단변수가 필요합니다.')]
        data, numeric_checks = _numeric(frame, [plan.outcome])
        checks.extend(numeric_checks)
        if data is None or any(c.status == 'fail' for c in numeric_checks):
            return checks
        group_values = frame[plan.group] if plan.group in frame else None
        if group_values is None:
            return checks + [ValidationCheck('집단변수', 'fail', f'{plan.group}이 데이터에 없습니다.')]
        combined = pd.DataFrame({'outcome': data[plan.outcome], 'group': group_values}).dropna()
        counts = combined.groupby('group', observed=True).size()
        groups = len(counts)
        expected_groups = 2 if plan.method == AnalysisMethod.WELCH_T else 3
        valid_group_count = groups == 2 if plan.method == AnalysisMethod.WELCH_T else groups >= 3
        checks.append(ValidationCheck('집단 수', 'pass' if valid_group_count else 'fail', f'유효 집단 {groups}개'))
        if groups:
            min_group = int(counts.min())
            checks.append(ValidationCheck('집단별 표본', 'pass' if min_group >= 5 else 'warning', f'최소 집단 N={min_group}'))
        if groups < expected_groups:
            checks.append(ValidationCheck('분석 가능성', 'fail', '요구되는 집단 수를 충족하지 못합니다.'))
        return checks

    return [ValidationCheck('분석 방법', 'fail', '지원하지 않는 분석 방법입니다.')]


def has_failures(checks: list[ValidationCheck]) -> bool:
    return any(check.status == 'fail' for check in checks)
