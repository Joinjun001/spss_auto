# 아키텍처

StatFlow의 중심은 통계 기능의 개수가 아니라 **Planning → Validation → Execution → Reporting** 경계를 명확히 하는 것입니다.

```text
app.py
  │
  ├─ statflow/data.py        원자료 읽기·프로파일링
  ├─ statflow/planner.py     규칙 기반 fallback planner
  ├─ statflow/llm_planner.py LLM structured-output 신뢰 경계
  ├─ statflow/validation.py  실행 전 조건 검증
  ├─ statflow/engine.py      SciPy/statsmodels 실제 계산
  └─ statflow/report.py      계산 결과 기반 설명
              │
              └─ statflow/schema.py  공통 도메인 모델
```

## Planner 경계

현재 `planner.py`는 API 키 없이 작동하는 규칙 기반 fallback이고, `llm_planner.py`는 특정 SDK에 종속되지 않는 `StructuredOutputProvider` 계약을 제공합니다. 외부 모델 출력은 Pydantic JSON schema로 검증하며 추가 필드를 금지하고, 데이터에 존재하는 변수만 허용하며, 분석 방법별 변수 역할까지 다시 검증한 뒤 `AnalysisPlan`으로 변환합니다. 외부 provider에는 원자료 행이 아니라 변수명·측정수준·유효/결측 N·고유값 수 같은 메타데이터만 기본 전달합니다.

LLM이 담당할 수 있는 영역:
- 연구 질문에서 변수 역할 후보 추론
- 분석 방법 후보와 근거 제시
- 필요한 가정과 추가 질문 제안

LLM이 직접 담당하지 않는 영역:
- p-value, 신뢰구간, 회귀계수 등 통계 수치 생성
- 검증 실패를 무시한 분석 실행
- 관찰자료에서 근거 없는 인과 결론 생성

## 실행 계층

통계 수치는 SciPy와 statsmodels에서 계산합니다. 새 분석 방법을 추가할 때는 다음 순서를 지킵니다.

1. `AnalysisMethod`에 방법 등록
2. planner가 해당 방법을 제안할 수 있도록 규칙/LLM schema 확장
3. validation에 최소 실행 조건 추가
4. engine에 검증된 구현 추가
5. 알려진 값 또는 신뢰 가능한 라이브러리와 대조하는 테스트 추가
6. report에 계산된 값만 사용하는 결과 설명 추가
