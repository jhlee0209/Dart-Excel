"""가상의 회사 '테스트전자'의 DART 응답을 흉내 낸 fixture."""
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from databook.model import build_databook  # noqa: E402

NONSTD = "-표준계정코드 미사용-"
M = 1_000_000

BS = [
    ("ifrs-full_CurrentAssets", "유동자산", 1150),
    ("ifrs-full_CashAndCashEquivalents", "현금및현금성자산", 300),
    (NONSTD, "단기금융상품", 100),
    ("dart_ShortTermTradeReceivable", "매출채권", 400),
    ("ifrs-full_Inventories", "재고자산", 250),
    (NONSTD, "기타유동금융자산", 40),
    (NONSTD, "기타유동자산", 60),
    ("ifrs-full_NoncurrentAssets", "비유동자산", 1850),
    ("ifrs-full_PropertyPlantAndEquipment", "유형자산", 1500),
    (NONSTD, "무형자산", 200),
    (NONSTD, "이연법인세자산", 150),
    ("ifrs-full_Assets", "자산총계", 3000),
    ("ifrs-full_CurrentLiabilities", "유동부채", 900),
    (NONSTD, "매입채무", 350),
    ("ifrs-full_ShorttermBorrowings", "단기차입금", 300),
    (NONSTD, "미지급금", 200),
    (NONSTD, "기타유동금융부채", 50),
    ("ifrs-full_NoncurrentLiabilities", "비유동부채", 600),
    (NONSTD, "장기차입금", 500),
    (NONSTD, "순확정급여부채", 100),
    ("ifrs-full_Liabilities", "부채총계", 1500),
    ("ifrs-full_EquityAttributableToOwnersOfParent", "지배기업 소유주지분", 1400),
    (NONSTD, "자본금", 100),
    (NONSTD, "이익잉여금", 1300),
    (NONSTD, "비지배지분", 100),
    ("ifrs-full_Equity", "자본총계", 1500),
    ("ifrs-full_EquityAndLiabilities", "부채와자본총계", 3000),
]
CIS = [
    ("ifrs-full_Revenue", "매출액", 4000),
    ("ifrs-full_CostOfSales", "매출원가", 2800),
    ("ifrs-full_GrossProfit", "매출총이익", 1200),
    (NONSTD, "판매비와관리비", 700),
    (NONSTD, "영업이익(손실)", 500),
    (NONSTD, "법인세비용차감전순이익", 480),
    (NONSTD, "법인세비용", 110),
    ("ifrs-full_ProfitLoss", "당기순이익", 370),
    (NONSTD, "기타포괄손익", 10),
    ("ifrs-full_ProfitLoss", "당기순이익", 370),
    (NONSTD, "총포괄손익", 380),
]
CF = [
    ("ifrs-full_CashFlowsFromUsedInOperatingActivities", "영업활동현금흐름", 650),
    (NONSTD, "감가상각비", 180),
    (NONSTD, "무형자산상각비", 20),
    (NONSTD, "대손상각비", 5),
    (NONSTD, "투자활동현금흐름", -300),
    (NONSTD, "유형자산의 취득", -250),
    (NONSTD, "무형자산의 취득", -30),
]


def make_rows(scale: float) -> list[dict]:
    rows = []
    for sj_div, table in (("BS", BS), ("CIS", CIS), ("CF", CF)):
        for ord_, (account_id, name, amount) in enumerate(table, 1):
            rows.append({
                "sj_div": sj_div, "account_id": account_id, "account_nm": name,
                "thstrm_amount": str(int(amount * M * scale)), "ord": str(ord_),
            })
    return rows


SCALES = {2023: 0.8, 2024: 0.9, 2025: 1.0}
CORP = {"corp_code": "00000000", "corp_name": "테스트전자", "stock_code": "000000"}


def fake_fetch(corp_code: str, year: int, fs_div: str) -> list[dict]:
    if fs_div != "CFS" or year not in SCALES:
        return []
    rows = make_rows(SCALES[year])
    if year == 2023:  # 과거에만 있던 계정
        rows.append({"sj_div": "BS", "account_id": NONSTD, "account_nm": "매각예정자산",
                     "thstrm_amount": "0", "ord": "7"})
    return rows


@pytest.fixture
def db():
    return build_databook(CORP, 2025, None, fake_fetch)


# 2023년 이후 형식: 합계가 먼저 오고 하위 계정이 코드순으로 따라온다. 표준계정ID도 예전과 다르다.
TREE_BS_ORDER = [
    "자산총계", "유동자산", "기타유동금융자산", "기타유동자산", "현금및현금성자산", "매출채권", "재고자산",
    "단기금융상품", "비유동자산", "이연법인세자산", "무형자산", "유형자산", "자본총계",
    "지배기업 소유주지분", "자본금", "이익잉여금", "비지배지분", "부채와자본총계", "유동부채",
    "기타유동금융부채", "매입채무", "미지급금", "단기차입금", "부채총계", "비유동부채", "장기차입금",
    "순확정급여부채",
]
TREE_CF_ORDER = ["투자활동현금흐름", "무형자산의 취득", "유형자산의 취득", "영업활동현금흐름",
                 "대손상각비", "감가상각비", "무형자산상각비"]


def make_tree_rows(scale: float) -> list[dict]:
    by_name = {(r["sj_div"], r["account_nm"]): r for r in make_rows(scale)}
    rows = [r for r in make_rows(scale) if r["sj_div"] == "CIS"]
    for sj_div, order in (("BS", TREE_BS_ORDER), ("CF", TREE_CF_ORDER)):
        for ord_, name in enumerate(order, 1):
            row = dict(by_name[sj_div, name], ord=str(ord_))
            if row["account_id"] in ("dart_ShortTermTradeReceivable", "ifrs-full_ShorttermBorrowings"):
                row["account_id"] = row["account_id"] + "New"
            rows.append(row)
    return rows


def fake_fetch_mixed(corp_code: str, year: int, fs_div: str) -> list[dict]:
    """2023은 예전 형식, 2024·2025는 새 형식."""
    if fs_div != "CFS" or year not in SCALES:
        return []
    return make_rows(SCALES[year]) if year == 2023 else make_tree_rows(SCALES[year])


def fake_fetch_tree(corp_code: str, year: int, fs_div: str) -> list[dict]:
    """모든 연도가 새 형식."""
    return make_tree_rows(SCALES[year]) if fs_div == "CFS" and year in SCALES else []
