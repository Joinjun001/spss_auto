# 배포

현재 StatFlow는 Streamlit 애플리케이션입니다. Streamlit Community Cloud에서 `Joinjun001/StatFlow`의 `main` 브랜치와 `app.py`를 연결하는 구성이 가장 단순합니다.

GitHub `main`에 새 커밋이 push되면 연결된 앱이 변경을 감지해 자동으로 갱신됩니다. 배포 환경에는 별도 AI API 키가 필요하지 않습니다. 향후 LLM provider를 추가할 경우 API 키는 저장소에 커밋하지 않고 배포 플랫폼의 secret 관리 기능으로 주입합니다.
