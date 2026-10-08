"""StatFlow: 설문 원자료를 메모리에서 점검하고 분석하는 Streamlit 앱."""
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from survey_core import (AnalysisError, correlation, descriptives, frequencies,
                         missing_codes, profile, quality, read_survey, to_excel)

SAMPLE_PATH = Path(__file__).parent / 'sample_data' / 'example_survey.csv'

st.set_page_config(page_title='StatFlow', layout='wide')
st.title('StatFlow')
st.caption('설문 원자료를 점검하고 통계 분석 결과를 탐색하는 연구 데이터 분석 도구입니다.')
st.warning('민감한 설문 자료는 공개 서버에 업로드하지 마세요. 개인 PC에서 실행하는 것을 권장합니다.')

st.subheader('0. 데이터 선택')
source = st.radio(
    '어떻게 시작할까요?',
    ['예시 데이터로 시작', '내 파일 업로드'],
    horizontal=True,
)

try:
    sheet = 0
    if source == '예시 데이터로 시작':
        sample_bytes = SAMPLE_PATH.read_bytes()
        frame = read_survey(sample_bytes, SAMPLE_PATH.name)
        st.success('합성 설문 예시 데이터 40명을 불러왔습니다. 실제 개인정보는 포함되어 있지 않습니다.')
        st.download_button(
            '예시 CSV 다운로드',
            sample_bytes,
            file_name='statflow_example_survey.csv',
            mime='text/csv',
        )
    else:
        upload = st.file_uploader('원자료 업로드 (CSV/XLSX, 첫 행: 변수명)', type=['csv', 'xlsx'])
        if upload is None:
            st.info('CSV 또는 XLSX 파일을 선택하면 분석을 시작합니다.')
            st.stop()
        if upload.name.lower().endswith('.xlsx'):
            sheets = pd.ExcelFile(BytesIO(upload.getvalue()), engine='openpyxl').sheet_names
            sheet = st.selectbox('시트', sheets)
        frame = read_survey(upload.getvalue(), upload.name, sheet)

    tokens = st.text_input('결측 코드 (쉼표로 구분, 자동 삭제하지 않음)', placeholder='예: -99, 999')
    frame = missing_codes(frame, tokens.split(','))
except (AnalysisError, ValueError, OSError) as exc:
    st.error(str(exc))
    st.stop()

st.subheader('1. 데이터 품질')
st.caption(f'{len(frame):,}행 × {len(frame.columns):,}열')
warnings = quality(frame)
if warnings:
    for warning in warnings:
        st.warning(warning)
else:
    st.success('기본 품질 점검에서 즉시 확인할 경고가 없습니다.')
with st.expander('원자료 미리보기 (30행)', expanded=(source == '예시 데이터로 시작')):
    st.dataframe(frame.head(30), use_container_width=True)

st.subheader('2. 변수·코딩 검토')
st.caption('측정 수준은 추정입니다. 코드북을 확인하고 직접 수정하세요.')
variables = st.data_editor(
    profile(frame),
    hide_index=True,
    use_container_width=True,
    disabled=['변수', '유효 N', '결측 N', '고유값', '검토'],
    column_config={
        '측정 수준': st.column_config.SelectboxColumn(
            '측정 수준',
            options=['명목형', '서열형', '연속형', 'ID', '검토 필요'],
            required=True,
        )
    },
)
measures = dict(zip(variables['변수'], variables['측정 수준']))
numeric = [c for c in frame if measures[str(c)] in {'서열형', '연속형'}]
categories = [c for c in frame if measures[str(c)] in {'명목형', '서열형'}]

st.subheader('3. 분석')
tables = {'변수검토': variables}
t1, t2, t3 = st.tabs(['기술통계', '빈도분석', '상관분석'])
with t1:
    default_descriptive = [c for c in ['age', 'study_hours', 'stress_score'] if c in numeric]
    if not default_descriptive:
        default_descriptive = numeric[:5]
    chosen = st.multiselect('분석 변수', numeric, default=default_descriptive)
    try:
        table = descriptives(frame, chosen)
        st.dataframe(table, hide_index=True, use_container_width=True)
        tables['기술통계'] = table
        st.caption('표본 표준편차(ddof=1). 서열형 응답의 평균 해석에는 주의하세요.')
    except AnalysisError as exc:
        st.error(str(exc))
with t2:
    default_category = 'region' if 'region' in categories else ''
    options = [''] + categories
    selected = st.selectbox('범주 변수', options, index=options.index(default_category) if default_category else 0)
    if selected:
        table = frequencies(frame, selected)
        st.dataframe(table, hide_index=True, use_container_width=True)
        st.bar_chart(table[table['값'] != '(결측)'].set_index('값')['빈도'])
        tables['빈도분석'] = table
with t3:
    method = st.radio('방법', ['pearson', 'spearman'], horizontal=True)
    eligible = [c for c in frame if measures[str(c)] == '연속형'] if method == 'pearson' else numeric
    if len(eligible) < 2:
        st.info('분석 가능한 서로 다른 변수가 두 개 이상 필요합니다.')
    else:
        preferred_x = 'study_hours' if 'study_hours' in eligible else eligible[0]
        preferred_y = 'stress_score' if 'stress_score' in eligible and 'stress_score' != preferred_x else eligible[1]
        x = st.selectbox('X 변수', eligible, index=eligible.index(preferred_x))
        y = st.selectbox('Y 변수', eligible, index=eligible.index(preferred_y))
        try:
            table = correlation(frame, x, y, method)
            st.dataframe(table, hide_index=True, use_container_width=True)
            tables['상관분석'] = table
            if method == 'pearson':
                st.scatter_chart(frame[[x, y]].apply(pd.to_numeric, errors='coerce').dropna(), x=x, y=y)
            st.caption('결측 쌍 제외. Pearson CI는 Fisher z 근사, Spearman p는 근사값(N<10 미제공). 상관은 인과가 아닙니다.')
        except AnalysisError as exc:
            st.error(str(exc))

st.subheader('4. 결과 내보내기')
st.download_button(
    'Excel 결과 다운로드',
    to_excel(tables),
    file_name='statflow_results.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
)
st.caption('결과 파일에는 원자료가 포함되지 않습니다. 앱은 업로드 데이터를 별도 디스크에 저장하지 않습니다.')
