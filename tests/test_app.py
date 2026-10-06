"""화면을 실제 네트워크 호출 없이 끝까지 돌려 본다."""
from pathlib import Path

import pytest
import streamlit as st
from streamlit.testing.v1 import AppTest

from conftest import CORP, fake_fetch
from databook import dart, market, profile

APP = str(Path(__file__).resolve().parent.parent / "app.py")
STOCK = {"stock_code": "000000", "name": "테스트전자", "market": "코스피", "cap": 25000.0,
         "price": 51000.0, "change": -1.2, "logo": ""}
OVERVIEW = {"text": "첫 문단입니다.\n\n둘째 문단입니다.\n\n셋째 문단입니다.", "report": "사업보고서 (2025.12)",
            "rcept_no": "1"}


@pytest.fixture(autouse=True)
def fake_sources(monkeypatch):
    st.cache_data.clear()
    monkeypatch.setattr(dart, "load_corp_codes", lambda force=False: [CORP])
    monkeypatch.setattr(dart, "fetch_statements", fake_fetch)
    monkeypatch.setattr(market, "load_ranking", lambda force=False: [STOCK])
    monkeypatch.setattr(profile, "company_info", lambda code: {"ceo_nm": "홍길동", "est_dt": "19990101"})
    monkeypatch.setattr(profile, "business_overview", lambda code: OVERVIEW)


def html_text(at: AppTest) -> str:
    return " ".join(str(h.proto.body) for h in at.get("html"))


def test_explore_then_open_databook():
    at = AppTest.from_file(APP, default_timeout=30)
    at.query_params["code"] = "000000"  # 표에서 행을 고른 것과 같은 효과
    at.run()
    assert not at.exception
    text = html_text(at)
    assert "테스트전자" in text and "홍길동" in text and "첫 문단입니다." in text
    assert "시가총액 2.5조" in text

    at.selectbox(key="year_explore").select(2025).run()
    next(b for b in at.button if b.label == "데이터북 만들기").click().run()
    assert not at.exception and not at.error
    assert at.session_state["view"] == "book"
    text = html_text(at)
    assert "EBITDA" in text and ">7<span" in text  # 7억원
    assert any("미분류 계정이 2개" in md.value for md in at.markdown)

    next(b for b in at.button if b.label == "← 기업 목록").click().run()
    assert not at.exception
    assert "기업 찾기" in html_text(at)


def test_ranking_failure_is_reported(monkeypatch):
    def fail(force=False):
        raise market.MarketError("시가총액 순위를 받지 못했습니다 (Timeout).")

    monkeypatch.setattr(market, "load_ranking", fail)
    at = AppTest.from_file(APP, default_timeout=30).run()
    assert not at.exception
    assert any("시가총액 순위를 받지 못했습니다" in e.value for e in at.error)


def test_missing_key_message(monkeypatch, tmp_path):
    monkeypatch.undo()
    st.cache_data.clear()
    monkeypatch.setattr(dart, "ROOT", tmp_path)
    monkeypatch.setattr(dart, "CACHE_DIR", tmp_path / "cache")
    monkeypatch.delenv("DART_API_KEY", raising=False)

    at = AppTest.from_file(APP, default_timeout=30).run()
    assert not at.exception
    assert any("인증키가 없습니다" in e.value for e in at.error)


def test_refresh_button_refetches(monkeypatch):
    calls = []
    monkeypatch.setattr(market, "load_ranking", lambda force=False: calls.append(force) or [STOCK])
    at = AppTest.from_file(APP, default_timeout=30).run()
    next(b for b in at.button if b.label == "시세 새로고침").click().run()
    assert not at.exception
    assert True in calls  # 캐시를 무시하고 다시 받았다


def click(at: AppTest, label: str) -> AppTest:
    return next(b for b in at.button if b.label == label).click().run()


def test_excel_lab_lessons_and_practice():
    at = AppTest.from_file(APP, default_timeout=60).run()
    click(at, "엑셀 학습하기")
    assert not at.exception
    assert at.session_state["view"] == "excel"
    text = html_text(at)
    assert "엑셀 학습" in text and "Alt → H → O → I" in text and "SUMIFS" in text
    assert any("Sub HighlightHardcodes()" in c.value for c in at.code)

    at.segmented_control(key="excel_mode").set_value("실전 예제에 적용하기").run()
    assert not at.exception
    at.selectbox(key="year_practice").select(2025).run()
    click(at, "과제 만들기")
    assert not at.exception and not at.error
    result = at.session_state["practice"]["result"]
    assert [t.key for t in result.tasks] == ["sum", "sumifs", "lookup", "text"]  # 기본 선택: 매일 쓰는 기본함수
    assert result.macro_code is None
    assert "과제 4개" in html_text(at)

    click(at, "← 기업 찾기")
    assert "기업 찾기" in html_text(at)
