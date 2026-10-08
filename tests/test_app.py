from pathlib import Path

from streamlit.testing.v1 import AppTest


def test_default_app_renders_example_workflow():
    app_path = Path(__file__).resolve().parents[1] / 'app.py'
    app = AppTest.from_file(app_path, default_timeout=10).run()
    assert not app.exception
    assert app.title[0].value == 'StatFlow'
    assert any('연구 질문' in item.value for item in app.subheader)
    assert any('분석 실행' in item.value for item in app.subheader)
