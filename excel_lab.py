"""엑셀 학습 화면: 기능 학습(단축키·함수·매크로)과 실전 예제(과제·정답 파일 생성)."""
from __future__ import annotations

import random
from html import escape
from typing import Callable

import streamlit as st

import ui
from databook import dart, lessons, practice
from databook.model import build_databook

GREEN, GREEN_DARK, GREEN_TINT, GREEN_LINE = "#107C41", "#0B5A2E", "#E9F3EC", "#CFE3D6"
INK, MUTED, CARD = ui.INK, ui.MUTED, ui.CARD
MODE_LEARN, MODE_PRACTICE = "기능 학습", "실전 예제에 적용하기"
DEFAULT_GROUP = practice.G_DAILY
XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

CSS = f"""
<style>
.stApp {{ background: radial-gradient(1100px 460px at 50% -8%, #FFFFFF 0%, rgba(255,255,255,0) 62%), #F3F7F2; }}
.xl-brand {{ font-size: .7rem; font-weight: 700; letter-spacing: .2em; color: {GREEN}; text-transform: uppercase; }}
.xl-title {{ font-family: {ui.SERIF}; font-size: 2.2rem; font-weight: 700; color: {INK}; margin: .3rem 0 .35rem; letter-spacing: -.02em; }}
.xl-sub {{ color: {MUTED}; font-size: .92rem; margin-bottom: .4rem; }}
.xl-label {{ font-size: .76rem; font-weight: 700; letter-spacing: .08em; color: {GREEN}; margin: 1.3rem 0 .5rem; }}
.xl-note {{ background: {GREEN_TINT}; border: 1px solid {GREEN_LINE}; border-radius: 12px; padding: .8rem 1rem;
           font-size: .88rem; line-height: 1.7; color: {INK}; }}
.xl-note p {{ margin: 0 0 .3rem; }}
table.keys {{ width: 100%; border-collapse: separate; border-spacing: 0; background: {CARD};
             border: 1px solid {GREEN_LINE}; border-radius: 14px; overflow: hidden; font-size: .88rem; box-shadow: {ui.SHADOW}; }}
table.keys th {{ text-align: left; background: {GREEN_TINT}; color: {GREEN_DARK}; font-weight: 600;
                padding: .55rem .9rem; font-size: .8rem; }}
table.keys td {{ padding: .55rem .9rem; border-top: 1px solid #EAF1EC; color: {INK}; vertical-align: middle; }}
table.keys td.memo {{ color: {MUTED}; font-size: .82rem; }}
.kbd {{ display: inline-block; font-family: 'SF Mono', Menlo, Consolas, monospace; font-size: .8rem;
       background: #F4F7F4; border: 1px solid #D5E2D9; border-bottom-width: 2px; border-radius: 7px;
       padding: .12rem .5rem; color: {INK}; white-space: nowrap; }}
.kbd.none {{ background: transparent; border-color: transparent; color: {MUTED}; font-family: inherit; }}
.fn {{ background: {CARD}; border: 1px solid {GREEN_LINE}; border-radius: 16px; padding: 1.1rem 1.3rem; margin-bottom: .9rem; box-shadow: {ui.SHADOW}; }}
.fn .name {{ font-family: 'SF Mono', Menlo, Consolas, monospace; font-weight: 700; color: {GREEN_DARK}; font-size: 1.02rem; }}
.fn .summary {{ color: {INK}; font-size: .93rem; margin: .15rem 0 .7rem; }}
.fn .row {{ display: grid; grid-template-columns: 3.4rem 1fr; gap: .6rem; margin-top: .4rem; font-size: .86rem; align-items: baseline; }}
.fn .row .k {{ color: {MUTED}; font-size: .78rem; }}
.fn code {{ font-family: 'SF Mono', Menlo, Consolas, monospace; font-size: .82rem; background: #F4F7F4;
           border-radius: 6px; padding: .15rem .4rem; color: {INK}; overflow-wrap: anywhere; }}
.fn .meaning {{ color: {MUTED}; margin-top: .2rem; }}
.fn .tip {{ color: {INK}; line-height: 1.65; }}
.task {{ background: {CARD}; border: 1px solid {GREEN_LINE}; border-radius: 16px; padding: 1rem 1.25rem; margin-bottom: .8rem; box-shadow: {ui.SHADOW}; }}
.task .head {{ display: flex; gap: .6rem; align-items: baseline; flex-wrap: wrap; }}
.task .no {{ font-weight: 700; color: {GREEN}; font-size: .82rem; }}
.task .title {{ font-weight: 700; color: {INK}; font-size: 1rem; }}
.task .chip {{ background: {GREEN_TINT}; color: {GREEN_DARK}; border-radius: 999px; padding: .1rem .6rem; font-size: .76rem; }}
.task ul {{ margin: .5rem 0 0; padding-left: 1.1rem; color: {INK}; font-size: .88rem; line-height: 1.7; }}
.step {{ display: flex; align-items: center; gap: .6rem; margin: 1.4rem 0 .5rem; }}
.step .n {{ width: 1.6rem; height: 1.6rem; border-radius: 50%; background: {GREEN}; color: #fff; font-weight: 700;
           font-size: .85rem; display: flex; align-items: center; justify-content: center; }}
.step .t {{ font-weight: 700; color: {INK}; font-size: 1.02rem; }}
/* 이 화면에서는 강조색을 엑셀 초록으로. Streamlit 자체 스타일보다 우선하도록 !important를 쓴다. */
[data-testid="stBaseButton-primary"] {{ background-color: {GREEN} !important; border-color: {GREEN} !important; }}
[data-testid="stBaseButton-primary"]:hover {{ background-color: {GREEN_DARK} !important; border-color: {GREEN_DARK} !important; }}
button[role="radio"][aria-checked="true"], button[aria-pressed="true"] {{
    background: {GREEN} !important; border-color: {GREEN} !important; }}
button[role="radio"][aria-checked="true"] p, button[aria-pressed="true"] p {{ color: #FFFFFF !important; }}
[data-testid="stBaseButton-primary"] {{ background: linear-gradient(180deg, #17924F 0%, {GREEN_DARK} 100%) !important;
    box-shadow: 0 10px 20px -12px rgba(11,90,46,.7) !important; }}
div[class*="st-key-nav_"][class*="_on"] button p {{ box-shadow: inset 0 -2px 0 {GREEN} !important; }}
[data-testid="stTab"][aria-selected="true"] p {{ color: {GREEN_DARK} !important; }}
[data-testid="stTab"][aria-selected="true"] > div:last-child {{ background-color: {GREEN} !important; }}
[data-testid="stTextInputRootElement"], [data-testid="stSelectbox"] [role="group"] {{
    background-color: {CARD} !important; border-color: {GREEN_LINE} !important; }}
table.keys {{ table-layout: fixed; }}
table.keys th:first-child {{ width: 28%; }}
table.keys th:last-child {{ width: 26%; }}
</style>
"""


def _keys(text: str) -> str:
    if text == lessons.NONE:
        return '<span class="kbd none">단축키 없음</span>'
    return f'<span class="kbd">{escape(text)}</span>'


def _shortcuts() -> None:
    c1, c2 = st.columns([2.2, 1], vertical_alignment="bottom")
    query = c1.text_input("단축키 검색", placeholder="동작이나 키로 검색 (예: 붙여넣기, 테두리, Alt)",
                          label_visibility="collapsed").strip().lower()
    system = c2.segmented_control("운영체제", ["둘 다", "Windows", "Mac"], default="둘 다",
                                  label_visibility="collapsed") or "둘 다"
    st.html('<div class="xl-note">' + "".join(f"<p>{_md_bold(n)}</p>" for n in lessons.SHORTCUT_NOTES) + "</div>")

    shown = 0
    for group, rows in lessons.SHORTCUT_GROUPS:
        rows = [r for r in rows if not query or query in " ".join(r).lower()]
        if not rows:
            continue
        shown += len(rows)
        head = "<th>동작</th>"
        head += "<th>Windows</th>" if system != "Mac" else ""
        head += "<th>Mac</th>" if system != "Windows" else ""
        body = ""
        for action, win, mac, memo in rows:
            body += f"<tr><td>{escape(action)}</td>"
            body += f"<td>{_keys(win)}</td>" if system != "Mac" else ""
            body += f"<td>{_keys(mac)}</td>" if system != "Windows" else ""
            body += f'<td class="memo">{escape(memo)}</td></tr>'
        st.html(f'<div class="xl-label">{escape(group)}</div>'
                f'<table class="keys"><thead><tr>{head}<th>메모</th></tr></thead><tbody>{body}</tbody></table>')
    if not shown:
        st.caption("검색 결과가 없습니다.")


def _md_bold(text: str) -> str:
    """**굵게** 만 지원하는 아주 작은 변환."""
    parts = escape(text).split("**")
    return "".join(f"<b>{p}</b>" if i % 2 else p for i, p in enumerate(parts))


def _functions() -> None:
    tiers = [lessons.TIER_DAILY, lessons.TIER_SOMETIMES, lessons.TIER_EXTRA]
    tier = st.segmented_control("단계", tiers, default=tiers[0], label_visibility="collapsed") or tiers[0]
    st.html(f'<div class="xl-note">{escape(lessons.TIER_NOTES[tier])}</div>')
    cards = ""
    for fn in lessons.FUNCTIONS:
        if fn["tier"] != tier:
            continue
        cards += (
            '<div class="fn">'
            f'<div class="name">{escape(fn["name"])}</div>'
            f'<div class="summary">{escape(fn["summary"])}</div>'
            f'<div class="row"><div class="k">구문</div><div><code>{escape(fn["syntax"])}</code></div></div>'
            f'<div class="row"><div class="k">예시</div><div><code>{escape(fn["example"])}</code>'
            f'<div class="meaning">{escape(fn["meaning"])}</div></div></div>'
            f'<div class="row"><div class="k">주의</div><div class="tip">{escape(fn["tip"])}</div></div>'
            "</div>"
        )
    st.html(f'<div style="margin-top:.9rem">{cards}</div>')
    st.caption("읽은 함수는 '실전 예제에 적용하기'에서 실제 회사 재무제표로 바로 연습할 수 있습니다.")


def _macros() -> None:
    st.html('<div class="xl-note">매크로는 순서대로 읽는 것이 좋습니다. 3번까지만 따라 해도 첫 매크로를 만들 수 있고, '
            "코드는 그대로 복사해 VBA 편집기에 붙여넣으면 실행됩니다.</div>")
    for i, (title, body, code) in enumerate(lessons.MACRO_SECTIONS):
        with st.expander(title, expanded=i == 0):
            st.markdown(body)
            if code:
                st.code(code, language="vbnet")


def _step(n: int, title: str) -> None:
    st.html(f'<div class="step"><div class="n">{n}</div><div class="t">{escape(title)}</div></div>')


def _practice(companies: list[dict], year_and_fs: Callable[[str], tuple[int, str | None]]) -> None:
    st.html('<div class="xl-note">실제 회사의 재무제표를 받아, 고른 기능을 연습하는 <b>과제 파일</b>과 '
            "<b>정답 파일</b>을 만들어 드립니다. 시간 제한은 없습니다.</div>")

    _step(1, "어느 회사의 재무제표로 연습할까요?")
    if not companies:
        st.error("회사 목록을 불러오지 못했습니다. '기업 찾기' 화면에서 연결 상태를 확인하세요.")
        return
    current = st.session_state.get("book_corp", {}).get("corp_code")
    index = next((i for i, c in enumerate(companies) if c["corp_code"] == current), 0)
    c1, c2 = st.columns([1.3, 1])
    with c1:
        corp = st.selectbox(
            "회사 (이름을 입력하면 검색됩니다)", companies, index=index,
            format_func=lambda c: f"{c['corp_name']} ({c['stock_code']})",
        )
    with c2:
        base_year, fs_div = year_and_fs("practice")

    _step(2, "어떤 기능을 익히고 싶나요? (여러 개 선택)")
    chosen: list[str] = []
    for group in practice.SKILL_GROUPS:
        skills = [s for s in practice.SKILLS if s.group == group]
        labels = st.pills(
            group, [s.label for s in skills], selection_mode="multi", key=f"skills_{group}",
            default=[s.label for s in skills] if group == DEFAULT_GROUP else None,
        )
        chosen += [s.key for s in skills if s.label in (labels or [])]

    _step(3, "과제 만들기")
    c1, c2, _ = st.columns([1.2, 1.4, 2.4])
    make = c1.button("과제 만들기", type="primary", width="stretch", disabled=not chosen)
    again = c2.button("다른 문제로 다시 만들기", width="stretch", disabled=not chosen,
                      help="찾기 과제의 계정 등이 바뀝니다.")
    if not chosen:
        st.caption("기능을 하나 이상 고르면 만들 수 있습니다.")
    if make or again:
        seed = random.randrange(1, 10**6) if again else 0
        try:
            with st.spinner("재무제표를 받아 과제를 만드는 중입니다…"):
                db = build_databook(corp, base_year, fs_div, dart.fetch_statements)
                st.session_state["practice"] = {
                    "result": practice.build_practice(db, chosen, seed),
                    "name": corp["corp_name"], "years": db.years,
                }
        except (dart.DartError, ValueError) as exc:
            st.error(str(exc))

    made = st.session_state.get("practice")
    if not made:
        return
    result: practice.Practice = made["result"]
    name, years = made["name"], made["years"]
    st.html(f'<div class="xl-label">{escape(name)} · FY{years[0]}~FY{years[-1]} · 과제 {len(result.tasks)}개</div>')
    cols = st.columns(3)
    cols[0].download_button("과제 파일 받기", result.task_xlsx, f"{name}_엑셀과제.xlsx", XLSX,
                            type="primary", width="stretch", icon=":material/download:")
    cols[1].download_button("정답 파일 받기", result.answer_xlsx, f"{name}_엑셀과제_정답.xlsx", XLSX,
                            width="stretch", icon=":material/task_alt:")
    if result.macro_code:
        cols[2].download_button("매크로 정답 코드 받기", result.macro_code.encode("utf-8"), f"{name}_매크로_정답코드.txt",
                                "text/plain", width="stretch", icon=":material/code:")
    st.caption("과제 파일에는 재무제표(BS·IS·CF)와 과제 시트가 들어 있습니다. 노란 칸을 채운 뒤 정답 파일과 비교하세요.")

    cards = ""
    for n, task in enumerate(result.tasks, 1):
        items = "".join(f"<li>{escape(line)}</li>" for line in task.brief)
        cards += (
            f'<div class="task"><div class="head"><span class="no">과제 {n}</span>'
            f'<span class="title">{escape(task.title)}</span><span class="chip">{escape(task.functions)}</span></div>'
            f"<ul>{items}</ul></div>"
        )
    st.html(cards)


def render(companies: list[dict], year_and_fs: Callable[[str], tuple[int, str | None]]) -> None:
    st.html(CSS)
    ui.top_nav("excel", "excel")
    st.html(
        '<div class="xl-brand">Excel Lab</div>'
        '<div class="xl-title">엑셀 학습</div>'
        '<div class="xl-sub">단축키 · 함수 · 매크로를 익히고, 실제 회사의 재무제표로 바로 연습합니다.</div>'
    )
    mode = st.segmented_control("모드", [MODE_LEARN, MODE_PRACTICE], default=MODE_LEARN, key="excel_mode",
                                label_visibility="collapsed") or MODE_LEARN
    if mode == MODE_PRACTICE:
        _practice(companies, year_and_fs)
        return
    tab_keys, tab_fn, tab_macro = st.tabs(["단축키", "함수", "매크로"])
    with tab_keys:
        _shortcuts()
    with tab_fn:
        _functions()
    with tab_macro:
        _macros()
