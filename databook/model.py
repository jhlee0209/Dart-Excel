"""DART 응답을 연도별로 합쳐 데이터북 모델로 만든다."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Callable

from . import mapping as m

FIRST_XBRL_YEAR = 2015
FS_LABEL = {"CFS": "연결", "OFS": "별도"}
TOLERANCE = 1  # 원


@dataclass(eq=False)
class Line:
    key: str
    account_id: str
    name: str
    values: dict[int, int | None] = field(default_factory=dict)
    section: str = ""  # BS 구분
    cls: str = ""  # BS 분류 또는 IS·CF 태그
    role: str | None = None  # BS 합계 줄의 역할(자산총계 등)

    def value(self, year: int) -> int:
        return self.values.get(year) or 0


@dataclass
class Databook:
    corp: dict
    fs_div: str
    years: list[int]
    bs: list[Line]
    is_: list[Line]
    cf: list[Line]
    is_source: str  # IS(손익계산서) 또는 CIS(포괄손익계산서)
    warnings: list[str] = field(default_factory=list)

    def role_line(self, role: str) -> Line | None:
        return next((ln for ln in self.bs if ln.role == role), None)


def default_base_year(today: date | None = None) -> int:
    """사업보고서는 보통 3월 말에 나오므로 4월부터 직전 연도를 기본값으로 쓴다."""
    today = today or date.today()
    return today.year - 1 if today.month >= 4 else today.year - 2


def parse_amount(text: str | None) -> int | None:
    text = (text or "").replace(",", "").strip()
    if text in ("", "-"):
        return None
    try:
        return int(text)
    except ValueError:
        try:
            return int(float(text))
        except ValueError:
            return None


def _ordered(rows: list[dict], sj_div: str) -> list[dict]:
    picked = [r for r in rows if r.get("sj_div") == sj_div]
    return sorted(picked, key=lambda r: int(r.get("ord") or 0))


def is_tree_order(rows: list[dict]) -> bool:
    """2023년 이후 보고서는 계정이 표시 순서가 아니라 '합계 → 하위 계정(코드순)'으로 온다.

    재무상태표 첫 줄이 자산총계면 그 형식으로 본다.
    """
    bs = _ordered(rows, "BS")
    return bool(bs) and m.bs_role(bs[0].get("account_id", ""), bs[0].get("account_nm", "")) == "자산총계"


def _bs_sections(rows: list[dict], tree: bool) -> list[str]:
    """재무상태표 각 줄의 구분. 합계 줄이 앞에 오는지(tree) 뒤에 오는지에 따라 읽는 법이 다르다."""
    section, out = m.SEC_A, []
    for row in rows:
        name = row.get("account_nm") or ""
        role = m.bs_role(row.get("account_id", ""), name)
        header = m.HEADER_NAMES.get(m.norm(name))
        if role in (m.CUR_A, m.NONCUR_A, m.CUR_L, m.NONCUR_L):
            section = role
            out.append(role)
        elif role == "자산총계":
            out.append(m.SEC_A)
            section = m.SEC_A if tree else m.SEC_L
        elif role == "부채총계":
            out.append(m.SEC_L)
            section = m.SEC_L if tree else m.SEC_E
        elif role == "자본총계":
            out.append(m.SEC_E)
            section = m.SEC_E
        elif role:  # 부채와자본총계
            out.append(m.SEC_TOTAL)
            section = m.SEC_L if tree else section
        elif header:
            section = header
            out.append(header)
        else:
            out.append(section)
    return out


def _occurrence(seen: dict[str, int], base: str) -> str:
    seen[base] = seen.get(base, 0) + 1
    return f"{base}#{seen[base]}"


def merge_years(rows_by_year: dict[int, list[dict]], sj_div: str, year_order: list[int]) -> list[Line]:
    """연도별 계정 목록을 한 표로 합친다. 줄 순서는 year_order의 첫 연도를 따른다.

    같은 계정이라도 연도에 따라 표준계정ID가 바뀌므로 계정명을 먼저 맞추고, 안 맞으면 ID로 맞춘다.
    재무상태표는 유동/비유동에 같은 이름이 있을 수 있어 구분까지 같아야 같은 줄로 본다.
    """
    lines: list[Line] = []
    by_name: dict[str, Line] = {}
    by_id: dict[str, Line] = {}
    for year in year_order:
        rows = _ordered(rows_by_year[year], sj_div)
        if sj_div == "BS":
            sections = _bs_sections(rows, is_tree_order(rows_by_year[year]))
        else:
            sections = [""] * len(rows)
        seen: dict[str, int] = {}
        prev: Line | None = None
        for row, section in zip(rows, sections):
            account_id = row.get("account_id", "")
            name = (row.get("account_nm") or "").strip()
            name_key = _occurrence(seen, f"n|{section}|{m.norm(name)}")
            id_key = _occurrence(seen, f"i|{section}|{account_id}") if m.is_standard_id(account_id) else None
            line = by_name.get(name_key) or (by_id.get(id_key) if id_key else None)
            if line is None or year in line.values:
                line = Line(key=name_key, account_id=account_id, name=name, section=section)
                # 다른 연도에만 있는 줄은 그 해에 바로 앞에 있던 줄 뒤에 끼워 넣는다.
                pos = next(i for i, ln in enumerate(lines) if ln is prev) + 1 if prev else 0
                lines.insert(pos if lines else 0, line)
            by_name.setdefault(name_key, line)
            if id_key:
                by_id.setdefault(id_key, line)
            line.values[year] = parse_amount(row.get("thstrm_amount"))
            prev = line
    return lines


_SECTION_RANK = {
    m.CUR_A: 0, m.NONCUR_A: 1, m.SEC_A: 2, m.CUR_L: 3, m.NONCUR_L: 4, m.SEC_L: 5, m.SEC_E: 6, m.SEC_TOTAL: 7,
}
_SECTION_HEADS = (m.CUR_A, m.NONCUR_A, m.CUR_L, m.NONCUR_L)
_CLASS_RANK = {
    m.CASH: 0, m.AR: 1, m.INV: 2, m.OTHER_NWC_A: 3, m.FIXED: 4,
    m.AP: 0, m.OTHER_NWC_L: 1, m.DEBT: 2, m.DEBT_LIKE: 3,
    m.UNCLASSIFIED: 5, m.OTHER_A: 6, m.OTHER_L: 6, m.EQUITY: 0,
}


def _bs_rank(line: Line) -> tuple[int, int]:
    """구분 순서, 그 안에서는 유동자산 같은 머리 줄이 먼저고 총계가 마지막."""
    within = 0 if line.role in _SECTION_HEADS else 2 if line.role else 1
    return _SECTION_RANK[line.section], within


def _mark_nested_subtotals(lines: list[Line], years: list[int]) -> dict[int, int]:
    """바로 아래 줄들의 합과 모든 연도에서 같은 줄을 소계로 표시한다(예: 자본금 = 보통주 + 우선주).

    DART 응답에는 계정의 상하 관계가 없어서 숫자로 찾는다. 반환값은 소계 줄 위치 → 하위 줄 끝 위치.
    """
    span: dict[int, int] = {}
    for i in range(len(lines) - 1, -1, -1):
        parent = lines[i]
        nonzero = sum(1 for y in years if parent.value(y))
        if parent.role or not nonzero:
            continue
        sums = dict.fromkeys(years, 0)
        j, children = i + 1, 0
        while j < len(lines) and lines[j].section == parent.section and not lines[j].role:
            for y in years:
                sums[y] += lines[j].value(y)
            children += 1
            j = span.get(j, j + 1)
            if all(abs(sums[y] - parent.value(y)) <= TOLERANCE for y in years):
                # 하위 줄이 하나뿐이면 우연히 같은 값일 수 있으니 두 해 이상 맞을 때만 인정한다.
                if children >= 2 or nonzero >= 2:
                    span[i] = j
                    parent.cls = m.SUBTOTAL
                break
    return span


def _build_bs(rows_by_year: dict[int, list[dict]], year_order: list[int], canonical: bool) -> list[Line]:
    years = sorted(rows_by_year)
    lines = merge_years(rows_by_year, "BS", year_order)
    # 금액 없이 제목으로만 나오는 '자산'·'부채'·'자본' 줄은 뺀다.
    lines = [
        ln for ln in lines
        if not (m.norm(ln.name) in m.HEADER_NAMES and not any(ln.value(y) for y in years))
    ]
    for line in lines:
        line.role = m.bs_role(line.account_id, line.name)
    lines.sort(key=_bs_rank)

    span = _mark_nested_subtotals(lines, years)
    for line in lines:
        if line.role:
            line.cls = m.SUBTOTAL
        elif not line.cls:
            line.cls = m.classify_bs(line.section, line.account_id, line.name)

    if canonical:
        # 표시 순서를 알 수 없는 형식이라 구분 안에서 성격(분류)별로 모은다. 소계와 그 하위 줄은 함께 움직인다.
        blocks, i = [], 0
        while i < len(lines):
            end = span.get(i, i + 1)
            blocks.append(lines[i:end])
            i = end

        def block_rank(block: list[Line]):
            head = block[0]
            cls = block[1].cls if head.cls == m.SUBTOTAL and len(block) > 1 else head.cls
            return (*_bs_rank(head), _CLASS_RANK.get(cls, 9))

        blocks.sort(key=block_rank)
        lines = [ln for block in blocks for ln in block]
    return lines


_IS_ORDER = [  # 표시 순서를 알 수 없을 때 쓰는 손익계산서 줄 순서(정규화한 계정명의 시작 문자열)
    ("매출액", "수익매출액", "영업수익"), ("매출원가",), ("매출총이익",), ("판매비", "판매관리비"),
    ("영업이익", "영업손익", "영업손실"), ("기타수익", "기타영업외수익"), ("기타비용", "기타영업외비용"),
    ("지분법",), ("금융수익",), ("금융비용", "금융원가"), ("법인세비용차감전", "법인세차감전"),
    ("법인세비용", "법인세수익"), ("계속영업",), ("중단영업",), ("당기순이익", "당기순손익", "당기순손실"),
    ("지배기업",), ("비지배",), ("기본주당",), ("희석주당",),
]


def _is_rank(line: Line) -> float:
    n = m.norm(line.name)
    for rank, prefixes in enumerate(_IS_ORDER):
        if n == "매출" or n.startswith(prefixes):
            return 0 if n == "매출" else rank
    return 14.5 if "포괄" in n else 9.5


_CF_HEADS = {
    "ifrs-full_CashFlowsFromUsedInOperatingActivities": 0,
    "ifrs-full_CashFlowsFromUsedInInvestingActivities": 1,
    "ifrs-full_CashFlowsFromUsedInFinancingActivities": 2,
}


def _cf_single_rank(line: Line) -> int | None:
    """활동별 묶음에 속하지 않는 현금흐름표 하단 줄의 순서."""
    n = m.norm(line.name)
    if "외화환산" in n or "환율변동" in n:
        return 3
    if n.startswith("현금및현금성자산의") and ("증가" in n or "증감" in n):
        return 4
    if n.startswith("기초"):
        return 5
    if n.startswith("기말"):
        return 6
    return None


def _reorder_cf(lines: list[Line]) -> list[Line]:
    """영업 → 투자 → 재무 → 환율효과 → 증감 → 기초 → 기말 순으로 묶음을 옮긴다."""
    blocks: list[tuple[float, list[Line]]] = []
    for line in lines:
        head = _CF_HEADS.get(line.account_id)
        single = _cf_single_rank(line)
        if head is not None:
            blocks.append((head, [line]))
        elif single is not None:
            blocks.append((single, [line]))
        elif blocks and blocks[-1][0] in (0, 1, 2):
            blocks[-1][1].append(line)
        else:
            blocks.append((7, [line]))
    blocks.sort(key=lambda b: b[0])
    return [ln for _, block in blocks for ln in block]


def _fix_parent_equity(lines: list[Line], latest: int) -> None:
    """지배기업 소유주지분이 세부 항목 뒤에 합계로 나오는 경우도 소계로 잡는다."""
    total = next((ln for ln in lines if ln.role == "자본총계"), None)
    parents = [
        ln for ln in lines
        if ln.section == m.SEC_E and ln.cls == m.EQUITY and m.is_parent_equity(ln.account_id, ln.name)
    ]
    if total and parents:
        others = sum(
            ln.value(latest) for ln in lines
            if ln.section == m.SEC_E and ln.cls == m.EQUITY and ln not in parents
        )
        if abs(others - total.value(latest)) <= TOLERANCE:
            for ln in parents:
                ln.cls = m.SUBTOTAL


def _tag(lines: list[Line], tagger: Callable[[str, str], str], repeatable: set[str]) -> None:
    used: set[str] = set()
    for line in lines:
        tag = tagger(line.account_id, line.name)
        # 포괄손익계산서에는 당기순이익이 두 번 나오므로 첫 줄에만 태그를 붙인다.
        if tag and (tag in repeatable or tag not in used):
            line.cls = tag
            used.add(tag)


def build_databook(
    corp: dict,
    base_year: int,
    fs_div: str | None,
    fetch: Callable[[str, int, str], list[dict]],
    n_years: int = 5,
) -> Databook:
    """fs_div가 None이면 연결을 먼저 시도하고, 없으면 별도로 받는다."""
    years = [y for y in range(base_year - n_years + 1, base_year + 1) if y >= FIRST_XBRL_YEAR]
    warnings: list[str] = []

    rows_by_year: dict[int, list[dict]] = {}
    used_div = fs_div or "CFS"
    for div in [fs_div] if fs_div else ["CFS", "OFS"]:
        rows_by_year = {}
        for year in years:
            rows = fetch(corp["corp_code"], year, div)
            if rows:
                rows_by_year[year] = rows
        used_div = div
        if rows_by_year:
            break
    if not rows_by_year:
        raise ValueError(
            "이 회사의 재무제표 데이터가 없습니다. 사업보고서를 제출하는 회사(주로 상장사)만 조회됩니다."
        )
    if not fs_div and used_div == "OFS":
        warnings.append("연결재무제표가 없어 별도재무제표로 조회했습니다.")
    missing = [y for y in years if y not in rows_by_year]
    if missing:
        warnings.append(f"데이터가 없는 연도는 제외했습니다: {', '.join(map(str, missing))}")
    years = sorted(rows_by_year)
    latest = years[-1]

    # 표시 순서대로 온 가장 최근 연도를 줄 순서의 기준으로 삼는다. 없으면 직접 정렬한다.
    ordered_years = [y for y in years if not is_tree_order(rows_by_year[y])]
    canonical = not ordered_years
    base = ordered_years[-1] if ordered_years else latest
    year_order = [base] + [y for y in reversed(years) if y != base]

    bs = _build_bs(rows_by_year, year_order, canonical)
    _fix_parent_equity(bs, latest)
    is_source = "IS" if any(r.get("sj_div") == "IS" for r in rows_by_year[latest]) else "CIS"
    is_ = merge_years(rows_by_year, is_source, year_order)
    cf = merge_years(rows_by_year, "CF", year_order)
    if canonical:
        is_.sort(key=_is_rank)
        cf = _reorder_cf(cf)

    _tag(is_, m.tag_is, repeatable=set())
    _tag(cf, m.tag_cf, repeatable={m.DA})

    db = Databook(corp, used_div, years, bs, is_, cf, is_source, warnings)
    if not db.role_line(m.CUR_A):
        warnings.append(
            "유동/비유동 구분이 없는 재무상태표입니다(금융업 등). 운전자본·순차입금 분석이 맞지 않을 수 있습니다."
        )
    if not any(ln.cls == m.DA for ln in cf):
        warnings.append(
            "현금흐름표에서 감가상각비를 찾지 못했습니다. 엑셀 EBITDA 시트의 노란 칸에 주석의 금액을 직접 입력하세요."
        )
    unclassified = sum(1 for ln in bs if ln.cls == m.UNCLASSIFIED)
    if unclassified:
        warnings.append(f"자동 분류하지 못한 재무상태표 계정이 {unclassified}개 있습니다. '분류 확인'에서 정해 주세요.")
    return db
