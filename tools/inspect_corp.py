"""회사 하나를 조회해 분류 결과와 주요 지표를 터미널에 찍는다. 분류 규칙을 손볼 때 쓴다.

사용: .venv.nosync/bin/python tools/inspect_corp.py 삼성전자 [기준연도] [--tree]
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from databook import dart, mapping as m  # noqa: E402
from databook.analysis import CHECK_KEYS, compute  # noqa: E402
from databook.model import build_databook, default_base_year, is_tree_order  # noqa: E402

args = [a for a in sys.argv[1:] if not a.startswith("--")]
name = args[0]
year = int(args[1]) if len(args) > 1 else default_base_year()
fetch = dart.fetch_statements
if "--tree" in sys.argv:  # 표시 순서 연도가 하나도 없는 상황을 흉내 낸다
    def fetch(code, y, div):
        rows = dart.fetch_statements(code, y, div)
        return rows if is_tree_order(rows) else []

corp = dart.search_corps(name, dart.load_corp_codes())[0]
db = build_databook(corp, year, None, fetch)
print(corp, db.fs_div, db.years, db.is_source)
for w in db.warnings:
    print("!", w)
T = 1e8  # 억원
for title, lines in (("BS", db.bs), ("IS", db.is_), ("CF", db.cf)):
    print(f"== {title}")
    for ln in lines:
        vals = " ".join(f"{(ln.values.get(y) or 0) / T:>12,.0f}" for y in db.years)
        print(f"{ln.section:5} {ln.cls:9} {ln.name[:24]:24} {vals}")
a = compute(db)
print("== 지표 (억원)")
for k in [m.REV, m.EBIT, "감가상각비", "EBITDA", "순운전자본", "순차입금", "CAPEX", "FCF", "자산총계"] + CHECK_KEYS:
    print(f"{k:22}", " ".join(f"{v / T:>12,.0f}" for v in a[k]))
