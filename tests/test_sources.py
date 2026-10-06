from databook import market, profile

XML = """<SECTION-1><TITLE ATOC="Y">II. 사업의 내용</TITLE>
<SECTION-2><TITLE ATOC="Y" AASSOCNOTE="D-0-2-1-0">1. 사업의 개요</TITLE>
<P>당사는 전자제품을&nbsp;만듭니다.</P><P></P>
<TABLE><TR><TD>표 안의 숫자 123</TD></TR></TABLE>
<P>국내외에   공장이 있습니다.<BR/>수출 비중이 높습니다.</P>
</SECTION-2><SECTION-2><TITLE>2. 주요 제품 및 서비스</TITLE><P>다음 절</P>"""


def test_extract_overview():
    assert profile.extract_overview(XML) == (
        "당사는 전자제품을 만듭니다.\n\n국내외에 공장이 있습니다.\n\n수출 비중이 높습니다."
    )
    assert profile.extract_overview("<TITLE>감사보고서</TITLE>") == ""


def test_facts_skips_empty():
    rows = dict(profile.facts({"ceo_nm": "홍길동", "est_dt": "19690113", "corp_cls": "Y", "hm_url": ""}))
    assert rows == {"대표이사": "홍길동", "설립일": "1969.01.13", "상장시장": "유가증권시장"}


def test_format_cap():
    assert market.format_cap(15_960_341) == "1,596.0조"
    assert market.format_cap(8234) == "8,234억"
    assert market.format_cap(None) == ""
