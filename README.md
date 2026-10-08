# StatFlow

StatFlow는 **연구 질문을 검증 가능한 통계 분석 workflow로 변환하는 연구 분석 오케스트레이션 도구**입니다.

단순히 통계 메뉴를 많이 제공하는 프로그램을 목표로 하지 않습니다. 연구 질문과 원자료를 바탕으로 분석 계획을 제안하고, 사용자가 변수 역할과 방법을 확인한 뒤 데이터 조건을 검증하고, 실제 통계 수치는 SciPy/statsmodels가 계산하도록 역할을 분리합니다.

## 핵심 흐름

```text
연구 질문 + 원자료
        ↓
분석 계획 제안
        ↓
사용자 확인/수정
        ↓
실행 전 데이터·가정 검증
        ↓
SciPy / statsmodels 실행
        ↓
효과크기·신뢰구간·진단
        ↓
근거가 추적 가능한 결과 해석
```

## 현재 MVP

- CSV/XLSX 원자료 읽기와 변수 측정수준 추정
- 연구 질문에서 변수명을 찾아 분석 목적을 설명 가능하게 추천
- 지원 계획: Pearson 상관, Welch t-test, One-way ANOVA, 선형 회귀
- 추천 분석/변수 역할을 사용자가 직접 수정 가능
- 실행 전 표본 수·집단 수·변동성·설계행렬 등 최소 조건 검증
- 실제 계산은 SciPy/statsmodels에 위임
- 회귀 잔차 정규성·Breusch-Pagan·VIF 진단
- 효과크기와 95% 신뢰구간 등 재현 가능한 결과 제공
- 실제 통계값만 사용한 보수적 결과 설명
- 합성 예시 데이터로 즉시 체험

현재 planner는 API 키 없이 동작하는 규칙 기반 구현입니다. 이는 최종 제품이 아니라 **LLM을 연결하기 전에 분석 schema와 안전 경계를 먼저 고정하기 위한 MVP**입니다. 향후 LLM도 자유롭게 수치를 생성하지 않고 동일한 `AnalysisPlan` 구조만 반환하게 합니다.

## 실행

```bash
python -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
streamlit run app.py
```

## 테스트

```bash
python -m pytest -q
python -m compileall -q app.py statflow tests
```

## 설계 원칙

1. AI/Planner는 분석 방법과 변수 역할을 **제안**한다.
2. 통계 수치는 검증 가능한 통계 라이브러리가 계산한다.
3. 추천은 강제가 아니며 근거·가정·한계를 함께 보여준다.
4. 연구 설계가 허용하지 않는 인과 해석을 자동으로 만들지 않는다.
5. 결과는 재현 가능한 구조화 데이터와 함께 제공한다.

자세한 구조는 `docs/ARCHITECTURE.md`, 다음 개발 순서는 `docs/ROADMAP.md`를 참고하세요.
