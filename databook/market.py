"""시가총액 순위와 로고 주소. 네이버 증권의 비공식 API를 쓰며 10분간 캐시한다."""
from __future__ import annotations

import json
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor

import requests

from .dart import CACHE_DIR

URL = "https://m.stock.naver.com/api/stocks/marketValue/{market}"
MARKETS = {"KOSPI": "코스피", "KOSDAQ": "코스닥"}
PAGE_SIZE = 100
CACHE_MINUTES = 10
HEADERS = {"User-Agent": "Mozilla/5.0"}


class MarketError(Exception):
    pass


def _number(text: str | None) -> float | None:
    try:
        return float((text or "").replace(",", ""))
    except ValueError:
        return None


def _page(market: str, page: int) -> dict:
    res = requests.get(
        URL.format(market=market), params={"page": page, "pageSize": PAGE_SIZE},
        headers=HEADERS, timeout=15,
    )
    res.raise_for_status()
    return res.json()


def _fetch_market(market: str) -> list[dict]:
    first = _page(market, 1)
    pages = range(2, (first["totalCount"] - 1) // PAGE_SIZE + 2)
    with ThreadPoolExecutor(max_workers=6) as pool:
        rest = list(pool.map(lambda p: _page(market, p), pages))
    stocks = [s for data in [first, *rest] for s in data["stocks"]]
    return [
        {
            "stock_code": s["itemCode"],
            "name": s["stockName"],
            "market": MARKETS[market],
            "cap": _number(s.get("marketValue")),  # 억원
            "price": _number(s.get("closePrice")),
            "change": _number(s.get("fluctuationsRatio")),  # %
            "logo": s.get("itemLogoPngUrl") or "",
        }
        for s in stocks
        if s.get("stockEndType") == "stock"  # ETF·ETN 제외
    ]


def fetched_at() -> datetime | None:
    """지금 캐시에 든 시세를 받아온 시각."""
    path = CACHE_DIR / "market.json"
    return datetime.fromtimestamp(path.stat().st_mtime) if path.exists() else None


def load_ranking(force: bool = False) -> list[dict]:
    """코스피·코스닥 전 종목을 시가총액 순으로. 실패하면 오래된 캐시라도 돌려준다."""
    path = CACHE_DIR / "market.json"
    fresh = path.exists() and time.time() - path.stat().st_mtime < CACHE_MINUTES * 60
    if fresh and not force:
        return json.loads(path.read_text(encoding="utf-8"))
    try:
        stocks = [s for market in MARKETS for s in _fetch_market(market)]
    except (requests.RequestException, KeyError, ValueError) as exc:
        if path.exists():
            return json.loads(path.read_text(encoding="utf-8"))
        raise MarketError(f"시가총액 순위를 받지 못했습니다 ({type(exc).__name__}).") from None
    stocks.sort(key=lambda s: -(s["cap"] or 0))
    CACHE_DIR.mkdir(exist_ok=True)
    path.write_text(json.dumps(stocks, ensure_ascii=False), encoding="utf-8")
    return stocks


def format_cap(cap: float | None) -> str:
    """억원 → '1,596.0조' 또는 '8,234억'."""
    if cap is None:
        return ""
    return f"{cap / 10000:,.1f}조" if cap >= 10000 else f"{cap:,.0f}억"
