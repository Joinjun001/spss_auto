# 개발 로드맵

## Phase 1 — 분석 오케스트레이션 기반 (현재)

- 연구 질문 → 구조화된 `AnalysisPlan`
- 사용자 확인/수정
- 실행 전 검증
- Pearson / Welch t-test / One-way ANOVA / 선형 회귀
- 효과크기·신뢰구간·회귀 진단
- 재현 가능한 보수적 결과 설명

## Phase 2 — LLM planner 연결

- [x] provider 독립적인 structured-output adapter와 schema/semantic validation
- 변수명/코드북/연구가설을 함께 전달
- 복수 분석 후보와 추천 점수
- 불확실할 때 사용자에게 필요한 추가 질문 생성
- [x] planner 출력 schema validation 및 허용 분석 registry 경계

## Phase 3 — 연구 설계 검증 강화

- 연구 설계 유형(횡단/종단/실험/반복측정) 입력
- 대응표본, 비모수, 카이제곱, 다중회귀 확장
- 결측 처리 전략과 민감도 경고
- 잔차/극단치/영향점 진단 강화
- 사후검정과 다중검정 보정

## Phase 4 — 재현 가능한 연구 보고서

- 분석 계획과 사용 변수 snapshot
- 실행된 통계 코드/환경 정보
- 표·그래프·가정 검증 결과
- Markdown/HTML/PDF 보고서
- 분석 이력과 재실행

## 제품 원칙

기능 개수보다 **왜 이 분석을 선택했는지, 실행 조건이 충족되는지, 실제 계산이 재현 가능한지**를 우선합니다.
