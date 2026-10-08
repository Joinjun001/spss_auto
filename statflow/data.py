"""연구 원자료 읽기와 변수 프로파일링."""
from io import BytesIO, StringIO
from pathlib import Path
from zipfile import BadZipFile, ZipFile

import numpy as np
import pandas as pd

MAX_BYTES = 20 * 1024 * 1024
MAX_ROWS = 100_000
MAX_COLUMNS = 300


class DataError(ValueError):
    pass


def read_dataset(data: bytes, filename: str, sheet: str | int = 0) -> pd.DataFrame:
    suffix = Path(filename).suffix.lower()
    if suffix not in {'.csv', '.xlsx'} or not data or len(data) > MAX_BYTES:
        raise DataError('20MB 이하 CSV 또는 XLSX 원자료를 업로드하세요.')
    try:
        if suffix == '.xlsx':
            with ZipFile(BytesIO(data)) as archive:
                entries = archive.infolist()
                if len(entries) > 2000 or sum(item.file_size for item in entries) > 100 * 1024 * 1024:
                    raise DataError('압축 해제 크기가 지나치게 큰 XLSX입니다.')
            frame = pd.read_excel(BytesIO(data), sheet_name=sheet, engine='openpyxl')
        else:
            try:
                text = data.decode('utf-8-sig')
            except UnicodeDecodeError:
                text = data.decode('cp949')
            frame = pd.read_csv(StringIO(text), sep=None, engine='python')
    except (UnicodeError, BadZipFile, OSError, ValueError, KeyError, pd.errors.ParserError) as exc:
        raise DataError('파일을 읽을 수 없습니다. 인코딩, 시트, 첫 행 변수명을 확인하세요.') from exc

    if not isinstance(frame, pd.DataFrame) or frame.empty:
        raise DataError('비어 있는 데이터입니다.')
    if len(frame) > MAX_ROWS or len(frame.columns) > MAX_COLUMNS:
        raise DataError('10만 행 또는 300열 제한을 초과했습니다.')
    if frame.columns.duplicated().any() or any(str(c).startswith('Unnamed:') for c in frame.columns):
        raise DataError('변수명이 비어 있거나 중복되어 있습니다.')

    frame = frame.copy()
    for col in frame.select_dtypes(include=['object', 'string']).columns:
        frame[col] = frame[col].replace(r'^\s*$', np.nan, regex=True)
    return frame


def infer_level(series: pd.Series, name: str) -> tuple[str, str]:
    values = series.dropna()
    unique = values.nunique()
    normalized = name.lower().replace(' ', '').replace('-', '_')
    if not len(values):
        return '검토 필요', '전체 결측'
    if (normalized in {'id', '응답자id', 'respondent_id'} or normalized.endswith('_id')) and unique == len(values):
        return 'ID', '식별자 후보'
    if pd.api.types.is_numeric_dtype(values):
        if unique <= 2:
            return '명목형', '숫자 범주 코드인지 확인'
        if unique <= 7 and (values % 1 == 0).all():
            return '서열형', 'Likert 또는 순서형 코드인지 확인'
        return '연속형', '연속값인지 범주 코드인지 확인'
    return '명목형', '순서가 있다면 서열형으로 수정'


def profile_dataset(frame: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for col in frame.columns:
        level, note = infer_level(frame[col], str(col))
        rows.append({
            '변수': str(col),
            '추정 측정수준': level,
            '유효 N': int(frame[col].notna().sum()),
            '결측 N': int(frame[col].isna().sum()),
            '고유값': int(frame[col].nunique(dropna=True)),
            '검토': note,
        })
    return pd.DataFrame(rows)


def level_map(frame: pd.DataFrame) -> dict[str, str]:
    profile = profile_dataset(frame)
    return dict(zip(profile['변수'], profile['추정 측정수준']))


def quality_messages(frame: pd.DataFrame) -> list[str]:
    messages: list[str] = []
    missing = int(frame.isna().sum().sum())
    if missing:
        messages.append(f'결측 셀 {missing}개가 있습니다. 분석별 완전 사례 수를 확인합니다.')
    duplicates = int(frame.duplicated().sum())
    if duplicates:
        messages.append(f'완전히 동일한 응답 행 {duplicates}개가 있습니다. 자동 삭제하지 않습니다.')
    constant = [str(col) for col in frame.columns if frame[col].nunique(dropna=True) <= 1]
    if constant:
        messages.append('상수 또는 전체 결측 변수: ' + ', '.join(constant))
    return messages
