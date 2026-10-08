"""설문 원자료의 안전한 읽기, 품질 점검 및 통계 분석. 입력값은 변경하지 않는다."""
from io import BytesIO, StringIO
from pathlib import Path
from zipfile import BadZipFile, ZipFile

import numpy as np
import pandas as pd
from scipy import stats

MAX_BYTES = 20 * 1024 * 1024


class AnalysisError(ValueError):
    """사용자에게 안내할 수 있는 분석 오류."""


def read_survey(data: bytes, filename: str, sheet: str | int = 0) -> pd.DataFrame:
    """응답자별 CSV(UTF-8/CP949) 또는 XLSX를 메모리에서 읽는다."""
    suffix = Path(filename).suffix.lower()
    if suffix not in {'.csv', '.xlsx'} or not data or len(data) > MAX_BYTES:
        raise AnalysisError('20MB 이하 CSV 또는 XLSX 원자료를 업로드하세요.')
    try:
        if suffix == '.xlsx':
            with ZipFile(BytesIO(data)) as z:
                entries = z.infolist()
                if len(entries) > 2000 or sum(i.file_size for i in entries) > 100 * 1024 * 1024:
                    raise AnalysisError('압축 해제 크기가 지나치게 큰 XLSX입니다.')
            frame = pd.read_excel(BytesIO(data), sheet_name=sheet, engine='openpyxl')
        else:
            try:
                content = data.decode('utf-8-sig')
            except UnicodeDecodeError:
                content = data.decode('cp949')
            frame = pd.read_csv(StringIO(content), sep=None, engine='python')
    except (UnicodeError, BadZipFile, OSError, ValueError, KeyError, pd.errors.ParserError) as exc:
        raise AnalysisError('파일을 읽을 수 없습니다. 인코딩, 시트, 첫 행 변수명을 확인하세요.') from exc
    if not isinstance(frame, pd.DataFrame) or frame.empty or len(frame) > 100_000 or len(frame.columns) > 300:
        raise AnalysisError('데이터가 비었거나 10만 행/300열 제한을 초과했습니다.')
    if frame.columns.duplicated().any() or any(str(c).startswith('Unnamed:') for c in frame.columns):
        raise AnalysisError('변수명이 비어 있거나 중복되어 있습니다.')
    frame = frame.copy()
    for col in frame.select_dtypes(include=['object', 'string']).columns:
        frame[col] = frame[col].replace(r'^\s*$', np.nan, regex=True)
    return frame


def missing_codes(frame: pd.DataFrame, tokens: list[str]) -> pd.DataFrame:
    """명시된 결측 코드만 바꾸고 원본을 보존한다."""
    result = frame.copy()
    tokens = [t.strip() for t in tokens if t.strip()]
    for col in result.columns:
        series = result[col]
        mask = series.notna() & series.astype(str).str.strip().isin(tokens)
        if pd.api.types.is_numeric_dtype(series):
            for token in tokens:
                try:
                    num = float(token)
                    if np.isfinite(num):
                        mask |= series.eq(num)
                except ValueError:
                    pass
        result.loc[mask, col] = np.nan
    return result


def profile(frame: pd.DataFrame) -> pd.DataFrame:
    """변수 측정 수준은 확정이 아니라 초기 제안이다."""
    records = []
    for col in frame:
        s = frame[col].dropna()
        unique = s.nunique()
        if not len(s):
            kind, note = '검토 필요', '전체 결측'
        elif (str(col).lower() == 'id' or str(col).lower().endswith('_id')) and unique == len(s):
            kind, note = 'ID', '식별자 추정'
        elif pd.api.types.is_numeric_dtype(s):
            if unique <= 2:
                kind, note = '명목형', '숫자 범주 코드인지 확인'
            elif unique <= 7 and (s % 1 == 0).all():
                kind, note = '서열형', 'Likert/범주 코드인지 확인'
            else:
                kind, note = '연속형', '연속값/ID/범주 코드인지 확인'
        else:
            kind, note = '명목형', '순서가 있으면 서열형으로 수정'
        records.append({'변수': str(col), '측정 수준': kind, '유효 N': len(s),
                        '결측 N': int(frame[col].isna().sum()), '고유값': unique, '검토': note})
    return pd.DataFrame(records)


def quality(frame: pd.DataFrame) -> list[str]:
    warnings = []
    if frame.isna().sum().sum():
        warnings.append(f'결측 셀 {int(frame.isna().sum().sum())}개: 결측 코드와 응답 패턴을 확인하세요.')
    if frame.duplicated().sum():
        warnings.append(f'동일한 응답 행 {int(frame.duplicated().sum())}개: 자동 삭제하지 않습니다.')
    for col in frame:
        if frame[col].nunique(dropna=True) <= 1:
            warnings.append(f'{col}: 상수 또는 전체 결측 변수입니다.')
    return warnings


def descriptives(frame: pd.DataFrame, columns: list[str]) -> pd.DataFrame:
    rows = []
    for col in columns:
        if col not in frame:
            raise AnalysisError(f'없는 변수: {col}')
        values = pd.to_numeric(frame[col], errors='coerce')
        if (frame[col].notna() & values.isna()).any():
            raise AnalysisError(f'{col}: 숫자가 아닌 응답이 있습니다.')
        v = values.replace([np.inf, -np.inf], np.nan).dropna()
        rows.append({'변수': col, 'N': len(v), '결측 N': len(frame) - len(v),
                     '평균': v.mean() if len(v) else np.nan,
                     '표준편차': v.std(ddof=1) if len(v) > 1 else np.nan,
                     '중앙값': v.median() if len(v) else np.nan})
    return pd.DataFrame(rows, columns=['변수', 'N', '결측 N', '평균', '표준편차', '중앙값'])


def frequencies(frame: pd.DataFrame, col: str) -> pd.DataFrame:
    if col not in frame:
        raise AnalysisError(f'없는 변수: {col}')
    counts = frame[col].value_counts(dropna=False, sort=False)
    valid = int(frame[col].notna().sum())
    return pd.DataFrame([{'값': '(결측)' if pd.isna(k) else str(k), '빈도': int(n),
                          '전체 %': n / len(frame) * 100,
                          '유효 %': n / valid * 100 if valid and not pd.isna(k) else np.nan}
                         for k, n in counts.items()])


def correlation(frame: pd.DataFrame, x: str, y: str, method: str) -> pd.DataFrame:
    """쌍별 결측 제외. Pearson 양측 p 및 Fisher z 근사 CI, Spearman 근사 p."""
    if method not in {'pearson', 'spearman'} or x == y or x not in frame or y not in frame:
        raise AnalysisError('상관분석 방법과 서로 다른 두 변수를 확인하세요.')
    pair = frame[[x, y]].copy()
    for col in (x, y):
        values = pd.to_numeric(pair[col], errors='coerce')
        if (pair[col].notna() & values.isna()).any():
            raise AnalysisError(f'{col}: 숫자가 아닌 응답이 있습니다.')
        pair[col] = values
    pair = pair.replace([np.inf, -np.inf], np.nan).dropna()
    n = len(pair)
    if n < 3 or pair[x].nunique() < 2 or pair[y].nunique() < 2:
        raise AnalysisError('유효 쌍 3개 이상과 각 변수의 변동성이 필요합니다.')
    result = stats.pearsonr(pair[x], pair[y]) if method == 'pearson' else stats.spearmanr(pair[x], pair[y])
    r, p = float(result.statistic), float(result.pvalue)
    low = high = np.nan
    if method == 'pearson' and n > 3:
        if abs(r) >= 1:
            low = high = np.sign(r)
        else:
            center = np.arctanh(r)
            half = 1.959963984540054 / np.sqrt(n - 3)
            low, high = np.tanh([center - half, center + half])
    if method == 'spearman' and n < 10:
        p = np.nan  # SciPy의 소표본 근사 p는 보고하지 않음
    return pd.DataFrame([{'변수 X': x, '변수 Y': y, '방법': method, '유효 쌍 N': n,
                          '제외 N': len(frame) - n, '상관계수': r, '양측 p': p,
                          '95% CI 하한': low, '95% CI 상한': high}])


def to_excel(tables: dict[str, pd.DataFrame]) -> bytes:
    """원자료가 아닌 결과표만 XLSX로 출력하고 수식 주입을 차단한다."""
    from openpyxl import Workbook
    wb = Workbook()
    wb.active.title = '안내'
    wb.active.append(['설문 분석 결과', '자동 추정 척도와 분석 가정을 확인하세요.'])
    for name, table in tables.items():
        sheet = wb.create_sheet(name[:31])
        sheet.append(list(table.columns))
        for row in table.itertuples(index=False, name=None):
            safe = []
            for value in row:
                if pd.isna(value):
                    value = None
                elif isinstance(value, str) and value.startswith(('=', '+', '-', '@')):
                    value = "'" + value
                elif isinstance(value, np.generic):
                    value = value.item()
                safe.append(value)
            sheet.append(safe)
    out = BytesIO()
    wb.save(out)
    return out.getvalue()
