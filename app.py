"""StatFlow Streamlit UI: 연구 질문을 검증 가능한 분석 workflow로 변환한다."""
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st

from statflow.data import DataError, level_map, profile_dataset, quality_messages, read_dataset
from statflow.engine import EngineError, run_analysis
from statflow.planner import recommend_plan
from statflow.report import summarize
from statflow.schema import AnalysisMethod, AnalysisPlan, METHOD_ASSUMPTIONS, METHOD_LABELS
from statflow.validation import has_failures, validate_plan

SAMPLE_PATH = Path(__file__).parent / 'sample_data' / 'example_survey.csv'
METHODS = list(AnalysisMethod)

st.set_page_config(page_title='StatFlow', layout='wide')
st.title('StatFlow')
st.caption('연구 질문을 통계 분석 계획으로 바꾸고, 데이터 조건을 검증한 뒤 실제 통계 엔진으로 실행합니다.')

with st.sidebar:
    st.markdown('### 원칙')
    st.markdown('- AI/Planner는 **분석 계획만 제안**합니다.\n- 통계 수치는 SciPy/statsmodels가 계산합니다.\n- 추천은 강제가 아니며 사용자가 수정할 수 있습니다.\n- 결과는 연구 설계와 함께 해석해야 합니다.')

st.subheader('1. 연구 질문과 데이터')
question = st.text_area('연구 질문', value='학습시간이 스트레스점수에 영향을 미치는가?', height=90)
source = st.radio('데이터', ['예시 데이터', '내 파일 업로드'], horizontal=True)

try:
    if source == '예시 데이터':
        raw = SAMPLE_PATH.read_bytes()
        frame = read_dataset(raw, SAMPLE_PATH.name)
        st.caption('합성 설문 40명: 실제 개인정보를 포함하지 않습니다.')
    else:
        upload = st.file_uploader('CSV/XLSX 업로드', type=['csv', 'xlsx'])
        if upload is None:
            st.info('파일을 업로드하면 분석 계획을 만들 수 있습니다.')
            st.stop()
        sheet = 0
        if upload.name.lower().endswith('.xlsx'):
            sheets = pd.ExcelFile(BytesIO(upload.getvalue()), engine='openpyxl').sheet_names
            sheet = st.selectbox('시트', sheets)
        frame = read_dataset(upload.getvalue(), upload.name, sheet)
except (DataError, OSError, ValueError) as exc:
    st.error(str(exc))
    st.stop()

cols = [str(c) for c in frame.columns]
levels = level_map(frame)
numeric = [c for c in cols if levels[c] in {'연속형', '서열형'}]
categorical = [c for c in cols if levels[c] == '명목형']

if not numeric:
    st.error('분석 가능한 수치형 변수를 찾지 못했습니다. 데이터와 변수 코딩을 확인하세요.')
    st.stop()

for message in quality_messages(frame):
    st.warning(message)
with st.expander('데이터/변수 프로파일', expanded=False):
    st.dataframe(profile_dataset(frame), hide_index=True, use_container_width=True)
    st.dataframe(frame.head(20), use_container_width=True)

st.subheader('2. 분석 계획')
proposal = recommend_plan(question, frame)
method_index = METHODS.index(proposal.method)
method = st.selectbox('추천 분석 방법', METHODS, index=method_index, format_func=lambda m: METHOD_LABELS[m])

if proposal.confidence == 'low':
    st.warning('질문에서 변수와 분석 의도를 충분히 확정하지 못했습니다. 아래 변수 역할을 직접 확인하세요.')
else:
    st.info('추천 근거: ' + ' '.join(proposal.rationale))

outcome = None
group = None
predictors: tuple[str, ...] = ()

if method == AnalysisMethod.PEARSON:
    if len(numeric) < 2:
        st.error('상관분석에는 서로 다른 수치형 변수 두 개가 필요합니다.')
        st.stop()
    defaults = list(proposal.predictors[:2]) if len(proposal.predictors) >= 2 else numeric[:2]
    x = st.selectbox('변수 X', numeric, index=numeric.index(defaults[0]) if defaults and defaults[0] in numeric else 0)
    y_candidates = [c for c in numeric if c != x]
    y_default = defaults[1] if len(defaults) > 1 and defaults[1] in y_candidates else y_candidates[0]
    y = st.selectbox('변수 Y', y_candidates, index=y_candidates.index(y_default))
    predictors = (x, y)
elif method == AnalysisMethod.LINEAR_REGRESSION:
    if len(numeric) < 2:
        st.error('회귀분석에는 종속변수와 독립변수로 사용할 수치형 변수가 두 개 이상 필요합니다.')
        st.stop()
    outcome_default = proposal.outcome if proposal.outcome in numeric else (numeric[1] if len(numeric) > 1 else numeric[0])
    outcome = st.selectbox('종속변수', numeric, index=numeric.index(outcome_default))
    predictor_options = [c for c in numeric if c != outcome]
    defaults = [p for p in proposal.predictors if p in predictor_options] or predictor_options[:1]
    predictors = tuple(st.multiselect('독립변수', predictor_options, default=defaults))
else:
    outcome_default = proposal.outcome if proposal.outcome in numeric else numeric[0]
    outcome = st.selectbox('종속변수', numeric, index=numeric.index(outcome_default))
    group_candidates = [c for c in categorical if c != outcome]
    if not group_candidates:
        st.error('명목형 집단변수가 없습니다. 변수 측정수준 추론을 확인하세요.')
        st.stop()
    group_default = proposal.group if proposal.group in group_candidates else group_candidates[0]
    group = st.selectbox('집단변수', group_candidates, index=group_candidates.index(group_default))

plan = AnalysisPlan(
    method=method,
    research_question=question,
    outcome=outcome,
    predictors=predictors,
    group=group,
    rationale=proposal.rationale if method == proposal.method else ('사용자가 추천 분석을 수정했습니다.',),
    assumptions=tuple(METHOD_ASSUMPTIONS[method]),
    confidence=proposal.confidence if method == proposal.method else 'user-confirmed',
)

with st.expander('필요한 가정', expanded=True):
    for assumption in plan.assumptions:
        st.markdown(f'- {assumption}')

st.subheader('3. 실행 전 검증')
checks = validate_plan(frame, plan)
icons = {'pass': '✅', 'warning': '⚠️', 'fail': '❌'}
for check in checks:
    st.write(f"{icons[check.status]} **{check.name}** — {check.message}")

st.subheader('4. 분석 실행')
if has_failures(checks):
    st.error('실패한 검증 항목이 있어 분석을 실행하지 않습니다. 변수 역할이나 데이터를 수정하세요.')
else:
    try:
        result = run_analysis(frame, plan)
    except (EngineError, ValueError, KeyError) as exc:
        st.error(f'분석 실행 실패: {exc}')
        st.stop()

    for name, table in result.tables.items():
        if isinstance(table, pd.DataFrame) and not table.empty:
            st.markdown(f'**{name}**')
            st.dataframe(table, hide_index=True, use_container_width=True)

    if result.diagnostics:
        st.markdown('**실행 후 진단**')
        for diagnostic in result.diagnostics:
            st.write(f"{icons[diagnostic.status]} **{diagnostic.name}** — {diagnostic.message}")

    st.markdown('**결과 해석**')
    for line in summarize(plan, result):
        st.write(line)

    with st.expander('계산된 원시 지표'):
        st.json(result.metrics)

st.caption('현재 MVP의 planner는 API 키 없이 동작하는 설명 가능한 규칙 기반 구현입니다. 향후 LLM은 동일한 AnalysisPlan 스키마를 반환하도록 연결합니다.')
