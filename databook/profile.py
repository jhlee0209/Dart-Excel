"""기업 소개: DART 기업개황과 최근 사업보고서의 '사업의 개요'."""
from __future__ import annotations

import html
import io
import json
import re
import zipfile
from datetime import date, timedelta

from . import dart

CORP_CLASS = {"Y": "유가증권시장", "K": "코스닥", "N": "코넥스", "E": "기타"}
OVERVIEW_LIMIT = 6000
_SECTION = re.compile(r"<TITLE[^>]*>\s*1\.\s*사업의\s*개요\s*</TITLE>(.*?)<TITLE", re.S)
_TABLE = re.compile(r"<TABLE.*?</TABLE>", re.S)
_BREAK = re.compile(r"</P>|<BR\s*/?>", re.I)
_TAG = re.compile(r"<[^>]+>")


def _cached(name: str, build):
    path = dart.CACHE_DIR / name
    if path.exists():
        return json.loads(path.read_text(encoding="utf-8"))
    data = build()
    dart.CACHE_DIR.mkdir(exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    return data


def company_info(corp_code: str) -> dict:
    def build() -> dict:
        data = dart._get("company.json", {"corp_code": corp_code}).json()
        if data.get("status") != "000":
            dart._raise_for_status(data.get("status", "?"), data.get("message", ""))
        return data

    return _cached(f"company_{corp_code}.json", build)


def _date(text: str) -> str:
    return f"{text[:4]}.{text[4:6]}.{text[6:]}" if len(text) == 8 else text


def facts(info: dict) -> list[tuple[str, str]]:
    """화면에 보여줄 (항목, 값) 목록."""
    rows = [
        ("대표이사", info.get("ceo_nm", "")),
        ("설립일", _date(info.get("est_dt", ""))),
        ("상장시장", CORP_CLASS.get(info.get("corp_cls", ""), "")),
        ("결산월", f"{info['acc_mt']}월" if info.get("acc_mt") else ""),
        ("본사", re.sub(r"\s+", " ", info.get("adres", ""))),
        ("홈페이지", info.get("hm_url", "")),
    ]
    return [(k, v) for k, v in rows if v]


def extract_overview(xml: str) -> str:
    """사업보고서 원문에서 'II. 사업의 내용 > 1. 사업의 개요' 본문만 뽑는다. 표는 뺀다."""
    match = _SECTION.search(xml)
    if not match:
        return ""
    body = _TABLE.sub("", match.group(1))
    body = _TAG.sub(" ", _BREAK.sub("\n", body))
    lines = [re.sub(r"[ \t\xa0]+", " ", html.unescape(line)).strip() for line in body.split("\n")]
    return "\n\n".join(line for line in lines if line)[:OVERVIEW_LIMIT]


def business_overview(corp_code: str) -> dict:
    """{'text': 본문, 'report': 보고서명, 'rcept_no': 접수번호}. 사업보고서가 없으면 text가 빈 문자열."""

    def build() -> dict:
        since = (date.today() - timedelta(days=550)).strftime("%Y%m%d")
        listing = dart._get(
            "list.json",
            {"corp_code": corp_code, "bgn_de": since, "pblntf_detail_ty": "A001", "last_reprt_at": "Y"},
        ).json()
        reports = listing.get("list") or []
        if not reports:
            return {"text": "", "report": "", "rcept_no": ""}
        report = reports[0]
        res = dart._get("document.xml", {"rcept_no": report["rcept_no"]})
        if not res.content.startswith(b"PK"):
            return {"text": "", "report": report["report_nm"], "rcept_no": report["rcept_no"]}
        with zipfile.ZipFile(io.BytesIO(res.content)) as zf:
            main = f"{report['rcept_no']}.xml"
            name = main if main in zf.namelist() else max(zf.infolist(), key=lambda i: i.file_size).filename
            xml = zf.read(name).decode("utf-8", errors="replace")
        return {"text": extract_overview(xml), "report": report["report_nm"].strip(), "rcept_no": report["rcept_no"]}

    return _cached(f"overview_{corp_code}.json", build)
