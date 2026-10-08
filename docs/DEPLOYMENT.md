# 배포

StatFlow는 현재 Streamlit 애플리케이션이므로 Streamlit Community Cloud를 기본 배포 대상으로 사용합니다.

## Community Cloud

1. Streamlit Community Cloud에 GitHub 계정으로 로그인합니다.
2. `Joinjun001/StatFlow` 저장소를 선택합니다.
3. branch는 `main`, entrypoint는 `app.py`로 지정합니다.
4. Python 3.12를 선택해 배포합니다.

배포 이후 GitHub `main`에 새 커밋이 push되면 Community Cloud가 저장소 변경을 감지해 앱을 자동 갱신합니다. 의존성 파일이 변경되면 전체 재배포가 수행될 수 있습니다.

현재 앱은 첫 접속 시 합성 예시 데이터를 기본으로 불러오므로 별도 파일 없이 기능을 시험할 수 있습니다.
