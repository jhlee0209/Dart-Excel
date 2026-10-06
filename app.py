"""DART 미니 데이터북 — Streamlit 화면.

화면은 셋이다. '기업 찾기'(시가총액 순위에서 회사를 고르고 소개를 읽는다),
'데이터북'(고른 회사의 5개년 분석을 보고 엑셀로 내려받는다),
'엑셀 학습'(excel_lab.py: 단축키·함수·매크로를 익히고 실제 재무제표로 과제를 푼다).
"""
from __future__ import annotations

from html import escape

import altair as alt
import pandas as pd
import streamlit as st

import excel_lab
from databook import dart, market, profile
from databook import mapping as m
from databook.analysis import CHECK_KEYS, compute
from databook.excel import build_workbook
from databook.model import FIRST_XBRL_YEAR, FS_LABEL, Databook, build_databook, default_base_year

st.set_page_config(page_title="DART 미니 데이터북", page_icon="📒", layout="wide",
                   initial_sidebar_state="collapsed")

INK, MUTED, LINE, CARD = "#18222F", "#6B7480", "#E3DDD0", "#FFFFFF"
NAVY, BRONZE = "#1F3A5F", "#9A7B4F"
SERIES_REVENUE, SERIES_PROFIT = "#2F5E9E", "#B8802A"  # 차트 계열색(색각 이상 구분 검증을 통과한 조합)
UP, DOWN = "#C0392B", "#2C5FA8"  # 국내 관례: 상승 빨강, 하락 파랑
FS_OPTIONS = {"자동 (연결 우선)": None, "연결": "CFS", "별도": "OFS"}
OTHER_MARKET = "비상장·기타"

st.html(f"""
<style>
.block-container {{ max-width: 1320px; padding-top: 2.2rem; padding-bottom: 4rem; }}
[data-testid="stSidebar"], [data-testid="stSidebarCollapsedControl"] {{ display: none; }}
.brand {{ font-size: .72rem; font-weight: 700; letter-spacing: .16em; color: {BRONZE}; text-transform: uppercase; }}
.page-title {{ font-size: 1.9rem; font-weight: 700; color: {INK}; margin: .15rem 0 .2rem; letter-spacing: -.02em; }}
.page-sub {{ color: {MUTED}; font-size: .92rem; margin-bottom: .4rem; }}
.corp-head {{ display: flex; align-items: center; gap: .85rem; }}
.corp-name {{ font-size: 1.25rem; font-weight: 700; color: {INK}; line-height: 1.25; }}
.corp-head.large .corp-name {{ font-size: 1.9rem; letter-spacing: -.02em; }}
.corp-meta {{ color: {MUTED}; font-size: .85rem; margin-top: .1rem; }}
.logo {{ position: relative; flex: none; width: 44px; height: 44px; border-radius: 12px; overflow: hidden;
        background: {NAVY}; color: #fff; display: flex; align-items: center; justify-content: center;
        font-weight: 700; font-size: 1.05rem; }}
.corp-head.large .logo {{ width: 56px; height: 56px; border-radius: 14px; font-size: 1.3rem; }}
.logo img {{ position: absolute; inset: 0; width: 100%; height: 100%; object-fit: cover; background: #fff; }}
.facts {{ display: grid; grid-template-columns: 5.2rem 1fr; row-gap: .45rem; font-size: .88rem; margin: .2rem 0; }}
.facts dt {{ color: {MUTED}; }}
.facts dd {{ margin: 0; color: {INK}; overflow-wrap: anywhere; }}
.section-label {{ font-size: .74rem; font-weight: 700; letter-spacing: .1em; color: {BRONZE}; margin: .9rem 0 .45rem; }}
.overview p {{ font-size: .92rem; line-height: 1.75; color: {INK}; margin: 0 0 .8rem; }}
.source {{ color: {MUTED}; font-size: .78rem; }}
.kpis {{ display: grid; grid-template-columns: repeat(4, 1fr); gap: .9rem; margin: .6rem 0 .2rem; }}
.kpi {{ background: {CARD}; border: 1px solid {LINE}; border-radius: 14px; padding: 1rem 1.15rem; }}
.kpi .label {{ color: {MUTED}; font-size: .8rem; }}
.kpi .value {{ color: {INK}; font-size: 1.65rem; font-weight: 700; letter-spacing: -.02em;
              font-variant-numeric: tabular-nums; margin: .15rem 0; }}
.kpi .unit {{ font-size: .85rem; font-weight: 500; color: {MUTED}; margin-left: .2rem; }}
.kpi .sub {{ color: {MUTED}; font-size: .8rem; }}
.empty {{ border: 1px dashed {LINE}; border-radius: 14px; padding: 3rem 1.5rem; text-align: center;
         color: {MUTED}; font-size: .92rem; line-height: 1.7; }}
@media (max-width: 900px) {{ .kpis {{ grid-template-columns: repeat(2, 1fr); }} }}
</style>
""")


# ── 데이터 ──────────────────────────────────────────────────────────────
@st.cache_data(show_spinner="회사 목록을 받는 중입니다…")
def corp_codes() -> list[dict]:
    return dart.load_corp_codes()


@st.cache_data(ttl=market.CACHE_MINUTES * 60, show_spinner="시가총액 순위를 받는 중입니다…")
def listed_companies() -> list[dict]:
    """시가총액 순 상장사. DART 고유번호가 없는 종목(우선주 등)은 뺀다."""
    by_stock = {c["stock_code"]: c for c in corp_codes() if c["stock_code"]}
    return [
        {**s, "corp_code": by_stock[s["stock_code"]]["corp_code"], "corp_name": s["name"]}
        for s in market.load_ranking()
        if s["stock_code"] in by_stock
    ]


# ── 서식 ────────────────────────────────────────────────────────────────
def money(v) -> str:
    """원 → 백만원 문자열. 음수는 괄호."""
    if v is None:
        return ""
    v = round(v / 1_000_000)
    return f"({abs(v):,})" if v < 0 else f"{v:,}"


def big_money(v) -> tuple[str, str]:
    """원 → (숫자, 단위). 1조 이상이면 조원, 아니면 억원."""
    if abs(v) >= 1e12:
        return f"{v / 1e12:,.1f}", "조원"
    return f"{v / 1e8:,.0f}", "억원"


def pct(v) -> str:
    return "n/a" if v is None else f"{v * 100:.1f}%"


def mult(v) -> str:
    return "n/a" if v is None else f"{v:.1f}x"


def days(v) -> str:
    return "n/a" if v is None else f"{v:.0f}일"


# ── 공통 조각 ───────────────────────────────────────────────────────────
def corp_head(corp: dict, meta: str, large: bool = False) -> None:
    """로고 + 회사명. 로고가 없거나 못 불러오면 뒤에 깔린 머리글자가 보인다."""
    name = corp["corp_name"]
    img = f'<img src="{escape(corp["logo"])}" alt="">' if corp.get("logo") else ""
    st.html(
        f'<div class="corp-head{" large" if large else ""}">'
        f'<div class="logo">{escape(name[:1])}{img}</div>'
        f'<div><div class="corp-name">{escape(name)}</div><div class="corp-meta">{escape(meta)}</div></div>'
        "</div>"
    )


def corp_meta(corp: dict) -> str:
    parts = [corp.get("stock_code") or "", corp.get("market") or ""]
    if corp.get("cap"):
        parts.append(f"시가총액 {market.format_cap(corp['cap'])}")
    return " · ".join(p for p in parts if p)


def company_profile(corp: dict) -> None:
    """기업개황과 사업의 개요."""
    try:
        with st.spinner("기업 소개를 불러오는 중입니다…"):
            info = profile.company_info(corp["corp_code"])
            overview = profile.business_overview(corp["corp_code"])
    except dart.DartError as exc:
        st.error(str(exc))
        return

    rows = "".join(f"<dt>{escape(k)}</dt><dd>{escape(v)}</dd>" for k, v in profile.facts(info))
    st.html(f'<div class="section-label">기업 개황</div><dl class="facts">{rows}</dl>')

    st.html('<div class="section-label">무엇을 하는 회사인가</div>')
    paragraphs = overview["text"].split("\n\n") if overview["text"] else []
    if not paragraphs:
        st.caption("최근 사업보고서에서 '사업의 개요'를 찾지 못했습니다.")
        return
    lead, rest = paragraphs[:2], paragraphs[2:]
    st.html('<div class="overview">' + "".join(f"<p>{escape(p)}</p>" for p in lead) + "</div>")
    if rest:
        with st.expander("이어서 읽기"):
            st.html('<div class="overview">' + "".join(f"<p>{escape(p)}</p>" for p in rest) + "</div>")
    st.html(f'<div class="source">출처: DART {escape(overview["report"])} · II. 사업의 내용 › 1. 사업의 개요</div>')


def table(db: Databook, rows: list[tuple[str, list, callable]]) -> None:
    """(이름, 연도별 값, 서식 함수) 목록을 표로 보여준다."""
    df = pd.DataFrame(
        [[label] + [fmt(v) for v in values] for label, values, fmt in rows],
        columns=["항목"] + [f"FY{y}" for y in db.years],
    )
    st.dataframe(df, hide_index=True, width="stretch", height=38 + 35 * len(rows))


def statement_df(db: Databook, lines, first_col: str) -> pd.DataFrame:
    data = {first_col: [ln.cls for ln in lines], "계정명": [ln.name for ln in lines]}
    for y in db.years:
        data[f"FY{y}"] = [money(ln.values.get(y)) for ln in lines]
    return pd.DataFrame(data)


def bar_chart(db: Databook, values: list, title: str, color: str) -> None:
    """한 계열짜리 막대 차트. 계열 이름은 제목이 대신한다."""
    unit, scale = ("조원", 1e12) if max(abs(v) for v in values) >= 1e12 else ("억원", 1e8)
    df = pd.DataFrame({"연도": [f"FY{y}" for y in db.years], "값": [v / scale for v in values]})
    chart = (
        alt.Chart(df, title=alt.Title(f"{title} ({unit})", anchor="start", fontSize=13, color=INK, dy=-6))
        .mark_bar(color=color, cornerRadiusEnd=4, size=30)
        .encode(
            x=alt.X("연도:N", sort=None, axis=alt.Axis(title=None, labelAngle=0, ticks=False, domainColor=LINE,
                                                       labelColor=MUTED, labelPadding=8)),
            y=alt.Y("값:Q", axis=alt.Axis(title=None, grid=True, gridColor="#ECE7DC", domain=False, ticks=False,
                                          labelColor=MUTED, tickCount=4, format=",.0f")),
            tooltip=[alt.Tooltip("연도:N"), alt.Tooltip("값:Q", title=f"{title} ({unit})", format=",.1f")],
        )
        .properties(height=230, background="transparent")
        .configure_view(stroke=None)
        .configure(font="Pretendard, 'Apple SD Gothic Neo', sans-serif")
    )
    st.altair_chart(chart, theme=None, width="stretch")


def year_and_fs(key: str) -> tuple[int, str | None]:
    c1, c2 = st.columns(2)
    latest = default_base_year()
    year = c1.selectbox("기준 연도", list(range(latest, FIRST_XBRL_YEAR + 3, -1)), key=f"year_{key}",
                        help="이 해를 마지막으로 최근 5개년을 받습니다.")
    fs = c2.selectbox("재무제표", list(FS_OPTIONS), key=f"fs_{key}")
    return year, FS_OPTIONS[fs]


def open_databook(corp: dict, base_year: int, fs_div: str | None) -> None:
    try:
        with st.spinner("DART에서 재무제표를 받는 중입니다…"):
            db = build_databook(corp, base_year, fs_div, dart.fetch_statements)
    except (dart.DartError, ValueError) as exc:
        st.error(str(exc))
        return
    st.session_state.update(
        db=db, book_corp=corp, auto_cls=[ln.cls for ln in db.bs],
        run=st.session_state.get("run", 0) + 1, view="book",
    )
    st.rerun()


def excel_banner() -> None:
    """화면 맨 아래의 엑셀 학습 입구."""
    st.html('<div style="height:1.2rem"></div>')
    with st.container(border=True):
        left, right = st.columns([4.2, 1.2], vertical_alignment="center")
        left.html(
            '<div class="section-label" style="margin:.1rem 0 .2rem">EXCEL LAB</div>'
            f'<div style="font-weight:700;color:{INK};font-size:1.05rem">엑셀도 같이 익히기</div>'
            f'<div style="color:{MUTED};font-size:.88rem;margin-top:.15rem">단축키 · 함수 · 매크로를 배우고, '
            "원하는 회사의 재무제표로 과제와 정답 파일을 받아 연습합니다.</div>"
        )
        if right.button("엑셀 학습하기", icon=":material/table_view:", width="stretch"):
            st.session_state["view"] = "excel"
            st.rerun()


# ── 화면 1: 기업 찾기 ───────────────────────────────────────────────────
def explore() -> None:
    st.html(
        '<div class="brand">DART Mini Databook</div>'
        '<div class="page-title">기업 찾기</div>'
        '<div class="page-sub">시가총액 순위에서 회사를 고르면 소개를 읽고, 5개년 데이터북을 만들 수 있습니다.</div>'
    )
    try:
        listed = listed_companies()
    except (dart.DartError, market.MarketError) as exc:
        st.error(str(exc))
        listed = []

    c1, c2, c3 = st.columns([3, 1.3, 0.9], vertical_alignment="bottom")
    query = c1.text_input("검색", placeholder="회사명 또는 종목코드로 검색 (예: 삼성전자, 005930)",
                          label_visibility="collapsed")
    scope = c2.segmented_control("시장", ["전체", "코스피", "코스닥"], default="전체",
                                 label_visibility="collapsed") or "전체"
    if c3.button("시세 새로고침", icon=":material/refresh:", width="stretch"):
        try:
            market.load_ranking(force=True)
        except market.MarketError as exc:
            st.error(str(exc))
        listed_companies.clear()
        st.rerun()

    rows = [dict(c) for c in listed if scope == "전체" or c["market"] == scope]
    for rank, row in enumerate(rows, 1):
        row["rank"] = rank
    q = query.strip().lower().replace(" ", "")
    if q:
        rows = [r for r in rows if q in r["corp_name"].lower().replace(" ", "") or q == r["stock_code"]]
        if scope == "전체":
            # 순위에 없는 회사(비상장 사업보고서 제출 법인 등)도 DART 목록에서 찾아 붙인다.
            try:
                known = {r["corp_code"] for r in rows}
                rows += [
                    {**c, "market": OTHER_MARKET, "rank": None, "cap": None, "price": None, "change": None, "logo": ""}
                    for c in dart.search_corps(query, corp_codes())
                    if c["corp_code"] not in known
                ]
            except dart.DartError as exc:
                st.error(str(exc))

    left, right = st.columns([1.55, 1], gap="large")
    with left:
        if not rows:
            st.html('<div class="empty">검색 결과가 없습니다.</div>')
            selected = None
        else:
            df = pd.DataFrame({
                "순위": [r["rank"] for r in rows],
                "로고": [r["logo"] or None for r in rows],
                "기업명": [r["corp_name"] for r in rows],
                "시장": [r["market"] for r in rows],
                "시가총액": [market.format_cap(r["cap"]) for r in rows],
                "현재가": [r["price"] for r in rows],
                "등락률": [r["change"] for r in rows],
            })
            styled = (
                df.style
                .format({"현재가": "{:,.0f}", "등락률": "{:+.2f}%", "순위": "{:.0f}"}, na_rep="")
                .map(lambda v: f"color: {UP if v > 0 else DOWN}" if pd.notna(v) and v else "", subset=["등락률"])
            )
            event = st.dataframe(
                styled, hide_index=True, width="stretch", height=680, row_height=44,
                on_select="rerun", selection_mode="single-row", key=f"ranking_{scope}_{q}",
                column_config={
                    "순위": st.column_config.Column(width=56),
                    "로고": st.column_config.ImageColumn(" ", width=48),
                    "기업명": st.column_config.Column(width="medium"),
                    "시장": st.column_config.Column(width=80),
                },
            )
            as_of = market.fetched_at()
            st.caption(
                f"{len(rows):,}개 회사 · 행을 누르면 오른쪽에 기업 소개가 나옵니다 · "
                f"네이버 증권 시세{f', {as_of:%m-%d %H:%M} 기준' if as_of else ''} "
                f"({market.CACHE_MINUTES}분이 지나면 다시 받습니다)"
            )
            picked = event.selection.rows
            selected = rows[picked[0]] if picked else None
            if selected is None and (code := st.query_params.get("code")):
                selected = next((r for r in rows if r["stock_code"] == code), None)

    with right:
        if selected is None:
            st.html('<div class="empty">왼쪽 목록에서 회사를 선택하세요.<br>무엇을 하는 회사인지와 기본 정보를 볼 수 있습니다.</div>')
        else:
            with st.container(border=True):
                corp_head(selected, corp_meta(selected))
                base_year, fs_div = year_and_fs("explore")
                if st.button("데이터북 만들기", type="primary", width="stretch"):
                    open_databook(selected, base_year, fs_div)
                company_profile(selected)
    excel_banner()


# ── 화면 2: 데이터북 ────────────────────────────────────────────────────
def book(db: Databook) -> None:
    corp = st.session_state["book_corp"]

    # 분류 편집 결과가 요약·다운로드에 반영되도록 편집 표의 현재 값을 먼저 읽는다.
    editor_key = f"cls_editor_{st.session_state['run']}"
    classes = list(st.session_state["auto_cls"])
    for row, change in st.session_state.get(editor_key, {}).get("edited_rows", {}).items():
        classes[int(row)] = change.get("분류", classes[int(row)])
    for line, cls in zip(db.bs, classes):
        line.cls = cls
    a = compute(db)
    last = len(db.years) - 1

    back, _, download = st.columns([1.1, 4, 1.6], vertical_alignment="center")
    if back.button("← 기업 목록", width="stretch"):
        st.session_state["view"] = "explore"
        st.rerun()
    download.download_button(
        "엑셀 데이터북 내려받기", data=build_workbook(db), type="primary", width="stretch",
        file_name=f"{db.corp['corp_name']}_데이터북_FY{db.years[-1]}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.html('<div class="brand" style="margin-top:.6rem">DART Mini Databook</div>')
    corp_head(
        corp,
        f"{corp_meta(corp)} · {FS_LABEL[db.fs_div]}재무제표 · FY{db.years[0]}~FY{db.years[-1]}".lstrip(" ·"),
        large=True,
    )

    has_da = any(a["감가상각비"])
    kpis = [
        (f"매출액 · FY{db.years[-1]}", a[m.REV][last],
         f"전년 대비 {a['매출 성장률'][last] * 100:+.1f}%" if last and a["매출 성장률"][last] is not None else ""),
        ("EBITDA" if has_da else "EBITDA (감가상각비 미반영)", a["EBITDA"][last], f"마진 {pct(a['EBITDA 마진'][last])}"),
        ("순운전자본", a["순운전자본"][last], f"매출 대비 {pct(a['NWC/매출액'][last])}"),
        ("순차입금" if a["순차입금"][last] >= 0 else "순차입금 (순현금)", a["순차입금"][last],
         f"EBITDA 대비 {mult(a['순차입금/EBITDA'][last])}"),
    ]
    cards = ""
    for label, value, sub in kpis:
        number, unit = big_money(value)
        cards += (
            f'<div class="kpi"><div class="label">{escape(label)}</div>'
            f'<div class="value">{number}<span class="unit">{unit}</span></div>'
            f'<div class="sub">{escape(sub)}</div></div>'
        )
    st.html(f'<div class="kpis">{cards}</div>')

    notices = [w for w in db.warnings if "자동 분류하지 못한" not in w]
    n_unclassified = sum(1 for ln in db.bs if ln.cls == m.UNCLASSIFIED)
    if n_unclassified:
        notices.append(f"미분류 계정이 {n_unclassified}개 있습니다. '분류 확인' 탭에서 성격을 판단해 정해 주세요.")
    failed = [k.removeprefix("체크: ") for k in CHECK_KEYS if any(abs(v) > 1 for v in a[k])]
    if failed:
        st.error("검증 실패: " + ", ".join(failed) + " — 소계로 표시할 줄이 더 있는지 확인하세요.")
    if notices:
        with st.expander(f"확인할 사항 {len(notices)}건", icon=":material/info:"):
            for notice in notices:
                st.markdown(f"- {notice}")

    tab_sum, tab_fs, tab_nwc, tab_nd, tab_cls, tab_about = st.tabs(
        ["요약", "재무제표", "순운전자본", "순차입금", "분류 확인", "기업 소개"]
    )

    with tab_sum:
        left, right = st.columns([1.25, 1], gap="large")
        with left:
            st.caption("단위: 백만원")
            table(db, [
                ("매출액", a[m.REV], money),
                ("  매출 성장률", a["매출 성장률"], lambda v: "" if v is None else pct(v)),
                ("영업이익", a[m.EBIT], money),
                ("  영업이익률", a["영업이익률"], pct),
                ("EBITDA", a["EBITDA"], money),
                ("  EBITDA 마진", a["EBITDA 마진"], pct),
                ("당기순이익", a[m.NI], money),
                ("순운전자본 (NWC)", a["순운전자본"], money),
                ("순차입금", a["순차입금"], money),
                ("  순차입금 / EBITDA", a["순차입금/EBITDA"], mult),
                ("FCF (영업현금흐름 − CAPEX)", a["FCF"], money),
                ("부채비율", a["부채비율"], pct),
            ])
        with right:
            bar_chart(db, a[m.REV], "매출액", SERIES_REVENUE)
            bar_chart(db, a[m.EBIT], "영업이익", SERIES_PROFIT)

    with tab_fs:
        choice = st.segmented_control("재무제표", ["재무상태표", "손익계산서", "현금흐름표"], default="재무상태표",
                                      label_visibility="collapsed") or "재무상태표"
        lines, first_col = {
            "재무상태표": (db.bs, "분류"), "손익계산서": (db.is_, "태그"), "현금흐름표": (db.cf, "태그"),
        }[choice]
        st.caption("단위: 백만원 · 각 연도 사업보고서의 당기 금액")
        st.dataframe(statement_df(db, lines, first_col), hide_index=True, width="stretch", height=620)

    with tab_nwc:
        st.caption("단위: 백만원 · 회전일수는 기말 잔액 기준")
        table(db, [
            (m.AR, a[m.AR], money), (m.INV, a[m.INV], money), (m.OTHER_NWC_A, a[m.OTHER_NWC_A], money),
            ("운전자본 자산 합계", a["운전자본자산"], money),
            (m.AP, a[m.AP], money), (m.OTHER_NWC_L, a[m.OTHER_NWC_L], money),
            ("운전자본 부채 합계", a["운전자본부채"], money),
            ("순운전자본 (NWC)", a["순운전자본"], money),
            ("NWC / 매출액", a["NWC/매출액"], pct),
            ("매출채권 회전일수 (DSO)", a["매출채권 회전일수"], days),
            ("재고자산 회전일수 (DIO)", a["재고자산 회전일수"], days),
            ("매입채무 회전일수 (DPO)", a["매입채무 회전일수"], days),
        ])

    with tab_nd:
        st.caption("단위: 백만원 · 순차입금이 음수면 순현금")
        table(db, [
            (m.DEBT, a[m.DEBT], money),
            (m.CASH, a[m.CASH], money),
            ("순차입금", a["순차입금"], money),
            (m.DEBT_LIKE, a[m.DEBT_LIKE], money),
            ("순차입금 (유사차입금 포함)", a["순차입금(유사차입금 포함)"], money),
            ("순차입금 / EBITDA", a["순차입금/EBITDA"], mult),
        ])
        members = [ln for ln in db.bs if ln.cls in (m.DEBT, m.CASH, m.DEBT_LIKE)]
        st.html('<div class="section-label">포함된 계정</div>')
        st.dataframe(statement_df(db, members, "분류"), hide_index=True, width="stretch")

    with tab_cls:
        st.markdown(
            "재무상태표 각 계정의 **분류**를 바꾸면 순운전자본·순차입금과 엑셀 파일에 바로 반영됩니다. "
            f"`{m.UNCLASSIFIED}`은 계정명만으로 판단하기 어려워 남겨 둔 계정입니다."
        )
        base = statement_df(db, db.bs, "분류")
        base["분류"] = st.session_state["auto_cls"]
        base.insert(0, "구분", [ln.section for ln in db.bs])
        st.data_editor(
            base, key=editor_key, hide_index=True, width="stretch", height=620,
            disabled=[c for c in base.columns if c != "분류"],
            column_config={"분류": st.column_config.SelectboxColumn(options=m.BS_CLASSES, required=True)},
        )

    with tab_about:
        left, _ = st.columns([1.6, 1])
        with left:
            company_profile(corp)
    excel_banner()


db: Databook | None = st.session_state.get("db")
view = st.session_state.get("view")
if view == "excel":
    try:
        companies = listed_companies()
    except (dart.DartError, market.MarketError) as exc:
        st.error(str(exc))
        companies = []
    excel_lab.render(companies, year_and_fs)
elif view == "book" and db is not None:
    book(db)
else:
    explore()
