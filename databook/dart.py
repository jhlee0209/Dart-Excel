"""OpenDART 호출과 디스크 캐시."""
from __future__ import annotations

import io
import json
import os
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree

import requests
from dotenv import load_dotenv

BASE_URL = "https://opendart.fss.or.kr/api"
ROOT = Path(__file__).resolve().parent.parent
CACHE_DIR = ROOT / "cache"
CORP_CACHE_DAYS = 7
ANNUAL_REPORT = "11011"  # 사업보고서

STATUS_MESSAGES = {
    "010": "등록되지 않은 인증키입니다. .env의 DART_API_KEY를 확인하세요.",
    "011": "사용할 수 없는 인증키입니다. OpenDART에서 키 상태를 확인하세요.",
    "012": "이 IP에서는 접근할 수 없습니다.",
    "020": "OpenDART 일일 요청 한도를 초과했습니다. 내일 다시 시도하세요.",
    "800": "OpenDART가 점검 중입니다. 잠시 후 다시 시도하세요.",
    "901": "인증키 사용 기간이 만료되었습니다.",
}
NO_DATA = "013"


class DartError(Exception):
    pass


def api_key() -> str:
    load_dotenv(ROOT / ".env", override=True)
    key = os.environ.get("DART_API_KEY", "").strip()
    if not key:
        raise DartError(
            "인증키가 없습니다. dart-databook/.env 파일에 DART_API_KEY=발급받은키 를 입력하세요."
        )
    return key


def _get(endpoint: str, params: dict) -> requests.Response:
    # requests 예외 메시지에는 인증키가 든 URL이 포함되므로 그대로 올리지 않는다.
    try:
        res = requests.get(
            f"{BASE_URL}/{endpoint}", params={"crtfc_key": api_key(), **params}, timeout=30
        )
        res.raise_for_status()
    except requests.RequestException as exc:
        raise DartError(f"OpenDART에 연결하지 못했습니다 ({type(exc).__name__}).") from None
    return res


def _raise_for_status(status: str, message: str) -> None:
    raise DartError(STATUS_MESSAGES.get(status, f"OpenDART 오류 {status}: {message}"))


def load_corp_codes(force: bool = False) -> list[dict]:
    """전체 회사 고유번호 목록. 일주일간 캐시한다."""
    path = CACHE_DIR / "corp_codes.json"
    if not force and path.exists():
        if time.time() - path.stat().st_mtime < CORP_CACHE_DAYS * 86400:
            return json.loads(path.read_text(encoding="utf-8"))

    res = _get("corpCode.xml", {})
    if not res.content.startswith(b"PK"):
        # 오류일 때는 zip 대신 XML 또는 JSON으로 상태가 온다.
        try:
            root = ElementTree.fromstring(res.content)
            _raise_for_status(root.findtext("status", "?"), root.findtext("message", ""))
        except ElementTree.ParseError:
            raise DartError("회사 목록을 받지 못했습니다.") from None

    with zipfile.ZipFile(io.BytesIO(res.content)) as zf:
        root = ElementTree.fromstring(zf.read(zf.namelist()[0]))
    corps = [
        {
            "corp_code": node.findtext("corp_code", "").strip(),
            "corp_name": node.findtext("corp_name", "").strip(),
            "stock_code": node.findtext("stock_code", "").strip(),
        }
        for node in root.iter("list")
    ]
    CACHE_DIR.mkdir(exist_ok=True)
    path.write_text(json.dumps(corps, ensure_ascii=False), encoding="utf-8")
    return corps


def search_corps(query: str, corps: list[dict], limit: int = 30) -> list[dict]:
    """회사명(부분 일치) 또는 종목코드로 검색. 상장사와 정확히 일치하는 이름을 앞에 둔다."""
    q = query.strip().lower().replace(" ", "")
    if not q:
        return []
    hits = [
        c
        for c in corps
        if q in c["corp_name"].lower().replace(" ", "") or q == c["stock_code"]
    ]

    def rank(c: dict):
        name = c["corp_name"].lower().replace(" ", "")
        return (name != q, not c["stock_code"], not name.startswith(q), len(name), name)

    return sorted(hits, key=rank)[:limit]


def fetch_statements(corp_code: str, year: int, fs_div: str) -> list[dict]:
    """한 사업연도의 전체 재무제표 계정. 데이터가 없으면 빈 리스트."""
    path = CACHE_DIR / f"{corp_code}_{year}_{fs_div}.json"
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))

    data = _get(
        "fnlttSinglAcntAll.json",
        {
            "corp_code": corp_code,
            "bsns_year": str(year),
            "reprt_code": ANNUAL_REPORT,
            "fs_div": fs_div,
        },
    ).json()
    status = data.get("status", "?")
    if status == NO_DATA:
        return []
    if status != "000":
        _raise_for_status(status, data.get("message", ""))

    rows = data.get("list", [])
    CACHE_DIR.mkdir(exist_ok=True)
    path.write_text(json.dumps(rows, ensure_ascii=False), encoding="utf-8")
    return rows
