"""파이썬으로 계산한 분석 지표. 화면 표시와 엑셀 수식 검증에 쓴다."""
from __future__ import annotations

from . import mapping as m
from .model import Databook, Line


def _div(a: float, b: float) -> float | None:
    return a / b if b else None


def class_sum(lines: list[Line], cls: str, year: int) -> int:
    return sum(ln.value(year) for ln in lines if ln.cls == cls)


def role_value(db: Databook, role: str, year: int) -> int:
    line = db.role_line(role)
    return line.value(year) if line else 0


def section_sum(db: Databook, suffix: str, year: int) -> int:
    """합계 줄을 뺀 구분별 합. suffix는 '자산'·'부채'·'자본'."""
    return sum(
        ln.value(year) for ln in db.bs
        if ln.section.endswith(suffix) and ln.cls != m.SUBTOTAL
    )


def compute(db: Databook) -> dict[str, list]:
    """지표 이름 → 연도 순서대로의 값 목록."""
    out: dict[str, list] = {}

    def put(name: str, fn) -> None:
        out[name] = [fn(i, y) for i, y in enumerate(db.years)]

    for tag in m.IS_TAGS:
        put(tag, lambda i, y, t=tag: class_sum(db.is_, t, y))
    put("매출총이익", lambda i, y: out[m.REV][i] - out[m.COGS][i])
    put("감가상각비", lambda i, y: class_sum(db.cf, m.DA, y))
    put("EBITDA", lambda i, y: out[m.EBIT][i] + out["감가상각비"][i])
    put("영업활동현금흐름", lambda i, y: class_sum(db.cf, m.CFO, y))
    put("CAPEX", lambda i, y: abs(class_sum(db.cf, m.CAPEX_PPE, y)) + abs(class_sum(db.cf, m.CAPEX_INT, y)))
    put("FCF", lambda i, y: out["영업활동현금흐름"][i] - out["CAPEX"][i])

    for cls in m.BS_CLASSES:
        put(cls, lambda i, y, c=cls: class_sum(db.bs, c, y))
    put("운전자본자산", lambda i, y: sum(out[c][i] for c in m.NWC_ASSET_CLASSES))
    put("운전자본부채", lambda i, y: sum(out[c][i] for c in m.NWC_LIAB_CLASSES))
    put("순운전자본", lambda i, y: out["운전자본자산"][i] - out["운전자본부채"][i])
    put("순차입금", lambda i, y: out[m.DEBT][i] - out[m.CASH][i])
    put("순차입금(유사차입금 포함)", lambda i, y: out["순차입금"][i] + out[m.DEBT_LIKE][i])

    for role in ("자산총계", "부채총계", "자본총계", m.CUR_A, m.CUR_L):
        put(role, lambda i, y, r=role: role_value(db, r, y))

    put("매출 성장률", lambda i, y: _div(out[m.REV][i] - out[m.REV][i - 1], out[m.REV][i - 1]) if i else None)
    put("매출총이익률", lambda i, y: _div(out["매출총이익"][i], out[m.REV][i]))
    put("영업이익률", lambda i, y: _div(out[m.EBIT][i], out[m.REV][i]))
    put("EBITDA 마진", lambda i, y: _div(out["EBITDA"][i], out[m.REV][i]))
    put("순이익률", lambda i, y: _div(out[m.NI][i], out[m.REV][i]))
    put("ROE", lambda i, y: _div(out[m.NI][i], out["자본총계"][i]))
    put("부채비율", lambda i, y: _div(out["부채총계"][i], out["자본총계"][i]))
    put("유동비율", lambda i, y: _div(out[m.CUR_A][i], out[m.CUR_L][i]))
    put("순차입금/EBITDA", lambda i, y: _div(out["순차입금"][i], out["EBITDA"][i]))
    put("NWC/매출액", lambda i, y: _div(out["순운전자본"][i], out[m.REV][i]))
    put("CAPEX/매출액", lambda i, y: _div(out["CAPEX"][i], out[m.REV][i]))
    put("매출채권 회전일수", lambda i, y: _div(out[m.AR][i] * 365, out[m.REV][i]))
    put("재고자산 회전일수", lambda i, y: _div(out[m.INV][i] * 365, out[m.COGS][i]))
    put("매입채무 회전일수", lambda i, y: _div(out[m.AP][i] * 365, out[m.COGS][i]))

    put("체크: 자산-부채-자본", lambda i, y: out["자산총계"][i] - out["부채총계"][i] - out["자본총계"][i])
    put("체크: 자산 분류 합계", lambda i, y: section_sum(db, "자산", y) - out["자산총계"][i])
    put("체크: 부채 분류 합계", lambda i, y: section_sum(db, "부채", y) - out["부채총계"][i])
    put("체크: 자본 분류 합계", lambda i, y: section_sum(db, "자본", y) - out["자본총계"][i])
    return out


CHECK_KEYS = [
    "체크: 자산-부채-자본", "체크: 자산 분류 합계", "체크: 부채 분류 합계", "체크: 자본 분류 합계",
]
