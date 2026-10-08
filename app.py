"""독립형 설문 분석 앱. 원자료는 별도 파일로 저장하지 않는다."""
from io import BytesIO

import pandas as pd
import streamlit as st
from survey_core import (AnalysisError, correlation, descriptives, frequencies,
                         missing_codes, profile, quality, read_survey, to_excel)

st.set_page_config(page_title='StatFlow', layout='wide')
st.title('StatFlow')
st.caption('설문 원자료를 점검하고 통계 분석 결과를 탐색하는 연구 데이터 분석 도구입니다.')
st.warning('민감한 설문 자료는 공개 서버에 업로드하지 마세요. 개인 PC에서 실행하는 것을 권장합니다.')
upload = st.file_uploader('원자료 업로드 (CSV/XLSX, 첫 행: 변수명)', type=['csv', 'xlsx'])
if upload is None:
    st.info('샘플: sample_data/example_survey.csv')
    st.stop()
try:
    sheet = 0
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
for warning in quality(frame):
    st.warning(warning)
with st.expander('원자료 미리보기 (30행)'):
    st.dataframe(frame.head(30), use_container_width=True)

st.subheader('2. 변수·코딩 검토')
st.caption('측정 수준은 추정입니다. 코드북을 확인하고 직접 수정하세요.')
variables = st.data_editor(profile(frame), hide_index=True, use_container_width=True,
    disabled=['변수', '유효 N', '결측 N', '고유값', '검토'],
    column_config={'측정 수준': st.column_config.SelectboxColumn('측정 수준',
        options=['명목형', '서열형', '연속형', 'ID', '검토 필요'], required=True)})
measures = dict(zip(variables['변수'], variables['측정 수준']))
numeric = [c for c in frame if measures[str(c)] in {'서열형', '연속형'}]
categories = [c for c in frame if measures[str(c)] in {'명목형', '서열형'}]

st.subheader('3. 분석')
tables = {'변수검토': variables}
t1, t2, t3 = st.tabs(['기술통계', '빈도분석', '상관분석'])
with t1:
    chosen = st.multiselect('분석 변수', numeric, default=numeric[:5])
    try:
        table = descriptives(frame, chosen)
        st.dataframe(table, hide_index=True, use_container_width=True)
        tables['기술통계'] = table
        st.caption('표본 표준편차(ddof=1). 서열형 응답의 평균 해석에는 주의하세요.')
    except AnalysisError as exc:
        st.error(str(exc))
with t2:
    selected = st.selectbox('범주 변수', [''] + categories)
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
        x = st.selectbox('X 변수', eligible)
        y = st.selectbox('Y 변수', eligible, index=1)
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
st.download_button('Excel 결과 다운로드', to_excel(tables), file_name='survey_results.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
st.caption('결과 파일에는 원자료가 포함되지 않습니다. 앱은 업로드 데이터를 별도 디스크에 저장하지 않습니다.')
