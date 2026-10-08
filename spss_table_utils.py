"""SPSS Excel 표 추출 및 결측치 표시 공통 함수."""

from numbers import Real
import re

import numpy as np
import pandas as pd


def prepare_spss_dataframe(frame):
    """결측 셀을 보존하면서 숫자만 소수점 두 자리 문자열로 변환한다."""
    def convert(value):
        if pd.isna(value):
            return np.nan
        if isinstance(value, str) and not value.strip():
            return np.nan
        if isinstance(value, Real) and not isinstance(value, (bool, np.bool_)):
            return f"{value:.2f}"
        return value

    return frame.map(convert) if hasattr(frame, "map") else frame.applymap(convert)


def extract_spss_tables(frame, keywords):
    """빈 행 또는 다음 표 제목을 경계로 SPSS 표를 분리한다."""
    result = {keyword: [] for keyword in keywords}
    if frame.shape[1] == 0:
        return result

    first_column = frame.iloc[:, 0].fillna("").astype(str)
    positions = {
        keyword: [i for i, value in enumerate(first_column) if keyword in value]
        for keyword in keywords
    }
    all_starts = sorted({i for indices in positions.values() for i in indices})

    for keyword in keywords:
        for start in positions[keyword]:
            limit = next((i for i in all_starts if i > start), len(frame))
            end = limit
            for i in range(start + 1, limit):
                if frame.iloc[i].isna().all():
                    end = i
                    break
                if keyword == "Scheffe" and "CTT 유의확률" in str(frame.iat[i, 0]):
                    end = i + 1
                    break
            table = frame.iloc[start:end].dropna(how="all").reset_index(drop=True)
            if not table.empty:
                result[keyword].append(table)
    return result


def replace_missing_stats(frame):
    """누락된 통계량을 실제 0으로 오인하지 않도록 대시(—)로 표시한다."""
    def convert(value):
        if pd.isna(value):
            return "—"
        if isinstance(value, str) and re.fullmatch(
            r"(?:nan|nan±.*|.*±nan)", value.strip(), flags=re.IGNORECASE
        ):
            return "—"
        return value

    return frame.map(convert) if hasattr(frame, "map") else frame.applymap(convert)
