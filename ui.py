"""화면 공통 디자인: 색, 글꼴, 기본 스타일, 상단 내비게이션.

종이 같은 아이보리 바탕에 짙은 네이비 글씨, 샴페인 골드를 가는 선과 작은 글씨에만 쓴다.
제목과 큰 숫자는 명조(세리프), 본문은 Pretendard.
"""
from __future__ import annotations

import streamlit as st

INK, MUTED, LINE = "#0F1B2D", "#69717F", "#E5DED0"
PAPER, CARD = "#F6F3EC", "#FFFFFF"
NAVY, NAVY_DEEP = "#14284B", "#0E1D38"
GOLD, GOLD_SOFT = "#8C6D3A", "#C9A86A"  # 글씨용(대비 확보) / 장식선용
SERIF = "'Noto Serif KR', 'AppleMyungjo', 'Nanum Myeongjo', serif"
SHADOW = "0 1px 2px rgba(15,27,45,.04), 0 16px 36px -18px rgba(15,27,45,.16)"

BASE_CSS = f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Serif+KR:wght@500;600;700;900&display=swap');

.stApp {{ background: radial-gradient(1100px 460px at 50% -8%, #FFFDF8 0%, rgba(255,253,248,0) 62%), {PAPER}; }}
[data-testid="stHeader"] {{ background: transparent; }}
.block-container {{ max-width: 1300px; padding-top: 1.8rem; padding-bottom: 4.5rem; }}
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] {{ display: none; }}

/* 상단 내비게이션 */
.wordmark {{ display: flex; align-items: center; gap: .6rem; font-family: {SERIF}; font-weight: 700;
            font-size: 1.14rem; color: {INK}; letter-spacing: -.01em; }}
.wordmark .mark {{ width: 9px; height: 9px; background: {GOLD_SOFT}; transform: rotate(45deg); }}
.wordmark .sub {{ font-family: Pretendard, sans-serif; font-size: .66rem; font-weight: 600; letter-spacing: .2em;
                 color: {MUTED}; text-transform: uppercase; margin-left: .35rem; }}
.nav-rule {{ height: 1px; background: linear-gradient(90deg, {LINE} 0%, {LINE} 72%, rgba(229,222,208,0) 100%);
            margin: .15rem 0 2rem; }}
div[class*="st-key-nav_"] button p {{ font-size: .92rem; color: {MUTED}; font-weight: 500; }}
div[class*="st-key-nav_"] button:hover p {{ color: {INK}; }}
div[class*="st-key-nav_"][class*="_on"] button p {{ color: {INK}; font-weight: 700;
    box-shadow: inset 0 -2px 0 {GOLD_SOFT}; padding-bottom: .2rem; }}

/* 제목 */
.brand {{ font-size: .7rem; font-weight: 700; letter-spacing: .2em; color: {GOLD}; text-transform: uppercase; }}
.page-title {{ font-family: {SERIF}; font-size: 2.2rem; font-weight: 700; color: {INK}; margin: .3rem 0 .35rem;
              letter-spacing: -.02em; }}
.page-sub {{ color: {MUTED}; font-size: .95rem; margin-bottom: .9rem; }}
.section-label {{ font-size: .7rem; font-weight: 700; letter-spacing: .16em; color: {GOLD}; margin: 1.1rem 0 .5rem;
                 text-transform: uppercase; }}

/* 카드 */
.st-key-panel, .st-key-banner {{ background: {CARD}; border: 1px solid {LINE} !important; border-radius: 18px !important;
    box-shadow: {SHADOW}; }}
.st-key-panel {{ padding: 1.35rem 1.45rem 1.2rem !important; }}
.st-key-banner {{ padding: 1.1rem 1.45rem !important; }}
[data-testid="stDataFrame"] {{ border-radius: 14px; box-shadow: {SHADOW}; }}
[data-testid="stExpander"] details {{ border-radius: 14px; border-color: {LINE}; background: {CARD}; }}

/* 회사 머리 */
.corp-head {{ display: flex; align-items: center; gap: .9rem; }}
.corp-name {{ font-family: {SERIF}; font-size: 1.32rem; font-weight: 700; color: {INK}; line-height: 1.25; }}
.corp-head.large .corp-name {{ font-size: 2.2rem; letter-spacing: -.02em; }}
.corp-meta {{ color: {MUTED}; font-size: .85rem; margin-top: .15rem; }}
.logo {{ position: relative; flex: none; width: 46px; height: 46px; border-radius: 13px; overflow: hidden;
        background: {NAVY}; color: #fff; display: flex; align-items: center; justify-content: center;
        font-weight: 700; font-size: 1.05rem; box-shadow: 0 0 0 1px {LINE}; }}
.corp-head.large .logo {{ width: 60px; height: 60px; border-radius: 16px; font-size: 1.3rem; }}
.logo img {{ position: absolute; inset: 0; width: 100%; height: 100%; object-fit: cover; background: #fff; }}
.facts {{ display: grid; grid-template-columns: 5.2rem 1fr; row-gap: .5rem; font-size: .88rem; margin: .2rem 0; padding: 0; }}
.facts dt {{ color: {MUTED}; margin: 0; padding: 0; }}
.facts dd {{ margin: 0; color: {INK}; overflow-wrap: anywhere; }}
.overview p {{ font-size: .93rem; line-height: 1.8; color: {INK}; margin: 0 0 .85rem; }}
.source {{ color: {MUTED}; font-size: .78rem; }}

/* 핵심 지표 */
.kpis {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: 1rem; margin: 1rem 0 .4rem; }}
.kpi {{ position: relative; background: {CARD}; border: 1px solid {LINE}; border-radius: 18px;
       padding: 1.25rem 1.3rem 1.1rem; box-shadow: {SHADOW}; }}
.kpi::before {{ content: ""; position: absolute; top: 0; left: 1.3rem; width: 30px; height: 2px; background: {GOLD_SOFT}; }}
.kpi .label {{ color: {MUTED}; font-size: .8rem; letter-spacing: .01em; }}
.kpi .value {{ font-family: {SERIF}; color: {INK}; font-size: 2rem; font-weight: 700; letter-spacing: -.02em;
              font-variant-numeric: lining-nums tabular-nums; margin: .25rem 0 .2rem; }}
.kpi .unit {{ font-family: Pretendard, sans-serif; font-size: .85rem; font-weight: 500; color: {MUTED}; margin-left: .3rem; }}
.kpi .sub {{ color: {MUTED}; font-size: .8rem; }}
.empty {{ border: 1px dashed #D8CFBD; border-radius: 18px; padding: 3.4rem 1.5rem; text-align: center;
         color: {MUTED}; font-size: .93rem; line-height: 1.8; background: rgba(255,255,255,.45); }}

/* 버튼과 입력 */
[data-testid="stBaseButton-primary"] {{ background: linear-gradient(180deg, #1C3763 0%, {NAVY_DEEP} 100%) !important;
    border: 1px solid {NAVY_DEEP} !important; border-radius: 999px !important;
    box-shadow: 0 10px 20px -12px rgba(14,29,56,.7); }}
[data-testid="stBaseButton-primary"] p {{ font-weight: 600; letter-spacing: .02em; }}
[data-testid="stBaseButton-primary"]:hover {{ filter: brightness(1.12); }}
[data-testid="stBaseButton-secondary"] {{ border-radius: 999px !important; background: {CARD} !important;
    border-color: {LINE} !important; }}
[data-testid="stBaseButton-secondary"]:hover {{ border-color: {GOLD_SOFT} !important; color: {INK} !important; }}
[data-testid="stTextInputRootElement"], [data-testid="stSelectbox"] [role="group"] {{
    background: {CARD} !important; border-color: {LINE} !important; border-radius: 12px !important; }}
button[role="radio"][aria-checked="true"], button[aria-pressed="true"] {{
    background: {NAVY} !important; border-color: {NAVY} !important; }}
button[role="radio"][aria-checked="true"] p, button[aria-pressed="true"] p {{ color: #FFFFFF !important; }}
[data-testid="stTab"] p {{ font-size: .95rem; }}
[data-testid="stTab"][aria-selected="true"] p {{ color: {INK} !important; font-weight: 700; }}
[data-testid="stTab"][aria-selected="true"] > div:last-child {{ background-color: {GOLD_SOFT} !important; }}

/* 시작 화면 */
.hero {{ text-align: center; padding-top: 5vh; }}
.hero img {{ width: min(520px, 82vw); height: auto; }}
.hero .brand {{ margin-top: 1.7rem; }}
.hero .name {{ font-family: {SERIF}; font-size: 3.2rem; font-weight: 900; color: {INK}; letter-spacing: -.03em;
              margin: .5rem 0 .6rem; }}
.hero .tagline {{ color: {MUTED}; font-size: 1.04rem; margin-bottom: 1.7rem; }}
.hero .rule {{ width: 44px; height: 2px; background: {GOLD_SOFT}; margin: 0 auto 1.1rem; }}
.hero-foot {{ text-align: center; color: #A2A8B0; font-size: .76rem; margin-top: 2.4rem; letter-spacing: .02em; }}
@media (max-width: 900px) {{ .kpis {{ grid-template-columns: repeat(2, 1fr); }} }}
</style>
"""

NAV_ITEMS = (("explore", "기업 찾기"), ("excel", "엑셀 학습"))


def inject_base() -> None:
    st.html(BASE_CSS)


def top_nav(current: str, highlight: str) -> None:
    """상단 내비게이션. current는 지금 화면, highlight는 밑줄을 칠 메뉴."""
    left, *cols = st.columns([5.2, 0.85, 0.85], vertical_alignment="center")
    left.html('<div class="wordmark"><span class="mark"></span>DART 미니 데이터북'
              '<span class="sub">Disclosure · Excel</span></div>')
    for col, (view, label) in zip(cols, NAV_ITEMS):
        key = f"nav_{view}" + ("_on" if view == highlight else "")
        if col.button(label, key=key, type="tertiary", width="stretch") and view != current:
            st.session_state["view"] = view
            st.rerun()
    st.html('<div class="nav-rule"></div>')
