"""계정 → 분류 규칙.

재무상태표(BS)의 각 줄에는 '분류'를, 손익계산서(IS)·현금흐름표(CF)의 주요 줄에는 '태그'를 붙인다.
엑셀의 분석 시트는 이 분류·태그를 SUMIFS로 집계하므로, 엑셀에서 값을 바꾸면 분석도 따라 바뀐다.
"""
from __future__ import annotations

import re

# ── BS 분류 ──────────────────────────────────────────────────────────────
AR, INV, OTHER_NWC_A = "매출채권", "재고자산", "기타운전자본자산"
AP, OTHER_NWC_L = "매입채무", "기타운전자본부채"
CASH, DEBT, DEBT_LIKE = "현금성자산", "차입금", "유사차입금"
FIXED, OTHER_A, OTHER_L, EQUITY = "고정자산", "기타자산", "기타부채", "자본"
SUBTOTAL, UNCLASSIFIED = "소계", "미분류"

BS_CLASSES = [
    AR, INV, OTHER_NWC_A, AP, OTHER_NWC_L, CASH, DEBT, DEBT_LIKE,
    FIXED, OTHER_A, OTHER_L, EQUITY, SUBTOTAL, UNCLASSIFIED,
]
NWC_ASSET_CLASSES = [AR, INV, OTHER_NWC_A]
NWC_LIAB_CLASSES = [AP, OTHER_NWC_L]

# ── BS 구분(섹션)과 합계 줄 ───────────────────────────────────────────────
CUR_A, NONCUR_A, CUR_L, NONCUR_L = "유동자산", "비유동자산", "유동부채", "비유동부채"
SEC_A, SEC_L, SEC_E, SEC_TOTAL = "자산", "부채", "자본", "합계"

ROLE_IDS = {
    "ifrs-full_CurrentAssets": CUR_A,
    "ifrs-full_NoncurrentAssets": NONCUR_A,
    "ifrs-full_Assets": "자산총계",
    "ifrs-full_CurrentLiabilities": CUR_L,
    "ifrs-full_NoncurrentLiabilities": NONCUR_L,
    "ifrs-full_Liabilities": "부채총계",
    "ifrs-full_Equity": "자본총계",
    "ifrs-full_EquityAndLiabilities": "부채와자본총계",
}
ROLE_NAMES = {
    "유동자산": CUR_A, "비유동자산": NONCUR_A, "자산총계": "자산총계",
    "유동부채": CUR_L, "비유동부채": NONCUR_L, "부채총계": "부채총계",
    "자본총계": "자본총계",
    "부채와자본총계": "부채와자본총계", "자본과부채총계": "부채와자본총계",
    "부채및자본총계": "부채와자본총계", "자본및부채총계": "부채와자본총계",
}
# 금액 없이 제목으로만 나오는 줄
HEADER_NAMES = {"자산": SEC_A, "부채": SEC_L, "자본": SEC_E}
PARENT_EQUITY_ID = "ifrs-full_EquityAttributableToOwnersOfParent"

BS_ID_CLASSES = {
    "ifrs-full_CashAndCashEquivalents": CASH,
    "ifrs-full_Inventories": INV,
    "ifrs-full_PropertyPlantAndEquipment": FIXED,
    "ifrs-full_IntangibleAssetsOtherThanGoodwill": FIXED,
    "ifrs-full_IntangibleAssetsAndGoodwill": FIXED,
    "ifrs-full_Goodwill": FIXED,
    "ifrs-full_RightofuseAssets": FIXED,
    "ifrs-full_InvestmentProperty": FIXED,
    "ifrs-full_ShorttermBorrowings": DEBT,
    "ifrs-full_LongtermBorrowings": DEBT,
    "ifrs-full_CurrentPortionOfLongtermBorrowings": DEBT,
}

# ── IS·CF 태그 ───────────────────────────────────────────────────────────
REV, COGS, SGA, EBIT, PBT, TAX, NI = (
    "매출액", "매출원가", "판관비", "영업이익", "세전이익", "법인세비용", "당기순이익",
)
IS_TAGS = [REV, COGS, SGA, EBIT, PBT, TAX, NI]
CFO, DA, CAPEX_PPE, CAPEX_INT = "영업활동현금흐름", "감가상각비", "유형자산취득", "무형자산취득"
CF_TAGS = [CFO, DA, CAPEX_PPE, CAPEX_INT]

IS_ID_TAGS = {
    "ifrs-full_Revenue": REV,
    "ifrs-full_CostOfSales": COGS,
    "dart_TotalSellingGeneralAdministrativeExpenses": SGA,
    "dart_OperatingIncomeLoss": EBIT,
    "ifrs-full_ProfitLossBeforeTax": PBT,
    "ifrs-full_IncomeTaxExpenseContinuingOperations": TAX,
    "ifrs-full_ProfitLoss": NI,
}
IS_NAME_TAGS = [  # (태그, 정규화한 계정명의 시작 문자열)
    (REV, ("매출액", "수익매출액", "영업수익", "매출")),
    (COGS, ("매출원가",)),
    (SGA, ("판매비와관리비", "판매비및관리비", "판매관리비")),
    (EBIT, ("영업이익", "영업손익", "영업손실")),
    (PBT, ("법인세비용차감전", "법인세차감전")),
    (TAX, ("법인세비용", "법인세수익")),
    (NI, ("당기순이익", "당기순손익", "당기순손실", "연결당기순이익")),
]
IS_EXACT_ONLY = {"매출"}  # '매출총이익', '매출원가' 등과 구분하려고 완전 일치만 인정

CF_ID_TAGS = {
    "ifrs-full_CashFlowsFromUsedInOperatingActivities": CFO,
    "ifrs-full_PurchaseOfPropertyPlantAndEquipmentClassifiedAsInvestingActivities": CAPEX_PPE,
    "ifrs-full_PurchaseOfIntangibleAssetsClassifiedAsInvestingActivities": CAPEX_INT,
}

_STRIP = re.compile(r"[\s\.,·ㆍ\(\)\[\]ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]")
_NUMBERING = re.compile(r"^\s*(?:[IVX]+|\d+)\s*\.\s*")


def norm(name: str) -> str:
    """비교용 계정명: 앞 번호(Ⅰ., 1.), 공백, 괄호 기호를 없앤다."""
    return _STRIP.sub("", _NUMBERING.sub("", name or ""))


def is_standard_id(account_id: str) -> bool:
    return bool(account_id) and not account_id.startswith("-")


def bs_role(account_id: str, name: str) -> str | None:
    """유동자산·자산총계 같은 합계 줄이면 그 역할을, 아니면 None."""
    return ROLE_IDS.get(account_id) or ROLE_NAMES.get(norm(name))


def is_parent_equity(account_id: str, name: str) -> bool:
    n = norm(name)
    return account_id == PARENT_EQUITY_ID or ("지배기업" in n and "비지배" not in n)


def _has(n: str, *words: str) -> bool:
    return any(w in n for w in words)


def classify_bs(section: str, account_id: str, name: str) -> str:
    """합계가 아닌 BS 한 줄의 분류. 판단이 갈리는 계정은 '미분류'로 남겨 사용자가 정하게 한다."""
    n = norm(name)
    if section == SEC_E:
        return EQUITY
    if account_id in BS_ID_CLASSES:
        return BS_ID_CLASSES[account_id]

    if section in (CUR_A, NONCUR_A, SEC_A):
        if _has(n, "유형자산", "무형자산", "사용권자산", "투자부동산", "영업권"):
            return FIXED
        if section == NONCUR_A:
            return OTHER_A
        if _has(n, "현금및현금성자산", "현금및예치금", "단기금융상품", "단기투자자산"):
            return CASH
        if _has(n, "매출채권"):
            return AR
        if _has(n, "재고자산"):
            return INV
        if _has(n, "미수금", "선급금", "선급비용", "계약자산", "기타유동자산", "반품"):
            return OTHER_NWC_A
        if _has(n, "법인세자산", "매각예정"):
            return OTHER_A
        if _has(n, "금융자산", "금융상품", "예금"):
            # '기타금융자산'에는 대여금·미수금이 섞이는 경우가 많아 자동 분류하지 않는다.
            return UNCLASSIFIED if "기타" in n else CASH
        return UNCLASSIFIED

    if _has(n, "차입금", "사채", "유동성장기부채", "리스부채", "차입부채"):
        return DEBT
    if section == NONCUR_L:
        return OTHER_L
    if _has(n, "매입채무"):
        return AP
    if _has(n, "미지급금", "미지급비용", "선수금", "예수금", "계약부채", "선수수익",
            "기타유동부채", "환불부채", "반품"):
        return OTHER_NWC_L
    if _has(n, "법인세부채", "충당부채", "매각예정"):
        return OTHER_L
    return UNCLASSIFIED


def tag_is(account_id: str, name: str) -> str:
    if account_id in IS_ID_TAGS:
        return IS_ID_TAGS[account_id]
    n = norm(name)
    for tag, prefixes in IS_NAME_TAGS:
        for p in prefixes:
            if n == p or (p not in IS_EXACT_ONLY and n.startswith(p)):
                return tag
    return ""


def tag_cf(account_id: str, name: str) -> str:
    if account_id in CF_ID_TAGS:
        return CF_ID_TAGS[account_id]
    n = norm(name)
    if n in ("영업활동현금흐름", "영업활동으로인한현금흐름"):
        return CFO
    if n.startswith("유형자산의취득"):
        return CAPEX_PPE
    if n.startswith("무형자산의취득"):
        return CAPEX_INT
    if "감가상각" in n or n == "상각비":
        return DA
    if "상각" in n and _has(n, "무형자산", "사용권자산", "투자부동산"):
        return DA
    return ""
