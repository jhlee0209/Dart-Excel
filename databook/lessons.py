"""엑셀 학습 내용: 단축키, 함수, 매크로.

함수와 단축키의 뼈대는 사용자가 준 「회계법인 실무에서 매일 쓰는 기본엑셀」 문서를 따른다.
"""
from __future__ import annotations

NONE = "없음"

# (동작, Windows, Mac, 메모). Windows의 'Alt → H → P'는 키를 하나씩 순서대로 누른다는 뜻이다.
SHORTCUT_GROUPS: list[tuple[str, list[tuple[str, str, str, str]]]] = [
    ("이동과 입력", [
        ("아래 / 위 셀로 이동", "Enter / Shift + Enter", "Return / ⇧ Return", ""),
        ("오른쪽 / 왼쪽 셀로 이동", "Tab / Shift + Tab", "Tab / ⇧ Tab", ""),
        ("데이터 끝으로 이동", "Ctrl + 방향키", "⌘ 방향키", "Shift를 함께 누르면 끝까지 선택"),
        ("셀 편집", "F2", "F2 또는 ⌃ U", "맥북은 fn과 함께"),
        ("절대참조($) 전환", "F4", "⌘ T 또는 F4", "수식 편집 중에"),
        ("여러 셀에 한 번에 입력", "Ctrl + Enter", "⌃ Return", "범위를 먼저 선택하고 입력"),
        ("셀 안에서 줄바꿈", "Alt + Enter", "⌥ Return", ""),
        ("오른쪽으로 채우기 / 아래로 채우기", "Ctrl + R / Ctrl + D", "⌃ R / ⌃ D", ""),
        ("행 전체 / 열 전체 선택", "Shift + Space / Ctrl + Space", "⇧ Space / ⌃ Space", "한글 입력 상태면 안 될 수 있음"),
        ("행·열 삽입 / 삭제", "Ctrl + Shift + + / Ctrl + -", "⌃ ⇧ = / ⌘ -", "행·열 전체를 선택한 뒤"),
    ]),
    ("붙여넣기", [
        ("선택하여 붙여넣기 창", "Ctrl + Alt + V  (또는 Alt → E → S)", "⌃ ⌘ V", ""),
        ("값만 붙여넣기", "Ctrl + C, Alt → E → S → V", "⌃ ⌘ V 후 V", "최신 버전은 Ctrl + Shift + V / ⌘ ⇧ V"),
        ("서식만 / 수식만 붙여넣기", "Alt → E → S → T / F", "⌃ ⌘ V 후 T / F", ""),
        ("열 너비만 붙여넣기", "Alt → E → S → W", "⌃ ⌘ V 후 W", ""),
        ("행열 바꿔 붙여넣기", "Alt → E → S → E", "⌃ ⌘ V 후 E", ""),
        ("일괄 덧셈 / 곱셈", "Alt → E → S → D / M", "⌃ ⌘ V 후 D / M", "예: 1000을 복사해 곱하면 단위 변환"),
    ]),
    ("숫자·글꼴 서식", [
        ("셀 서식 창", "Ctrl + 1", "⌘ 1", ""),
        ("퍼센트로 바꾸기", "Alt → H → P", "⌃ ⇧ 5", "Windows도 Ctrl + Shift + 5 가능"),
        ("천 단위 구분 기호", "Ctrl + Shift + 1", "⌃ ⇧ 1", ""),
        ("소수 자릿수 늘리기 / 줄이기", "Alt → H → 0 / Alt → H → 9", NONE, "맥은 홈 탭의 자릿수 버튼"),
        ("글꼴 크기 / 글꼴 / 글자색", "Alt → H → F → S / F / C", NONE, ""),
        ("셀 채우기 색", "Alt → H → H", NONE, ""),
        ("가로 가운데 / 오른쪽 / 왼쪽 정렬", "Alt → H → A → C / R / L", "⌘ E (가운데) · ⌘ L (왼쪽)", "Alt → H → A → M은 세로 가운데"),
        ("병합하고 가운데 맞춤", "Alt → H → M → C", NONE, "여러 행을 행별로 병합: Alt → H → M → A"),
    ]),
    ("테두리", [
        ("바깥 테두리", "Ctrl + Shift + 7", "⌘ ⌥ 0", ""),
        ("모든 테두리", "Alt → H → B → A", NONE, ""),
        ("합계 테두리 (위 실선 + 아래 이중선)", "Alt → H → B → U", NONE, ""),
        ("위 / 아래 테두리", "Alt → H → B → P / O", NONE, ""),
        ("왼쪽 / 오른쪽 테두리", "Alt → H → B → L / R", NONE, ""),
        ("굵은 바깥 테두리", "Alt → H → B → T", NONE, ""),
        ("테두리 지우기", "Ctrl + Shift + -", "⌘ ⌥ -", ""),
    ]),
    ("행·열·보기", [
        ("열 너비 자동 맞춤", "Alt → H → O → I", NONE, "맥은 열 경계선을 더블클릭"),
        ("행 높이 자동 맞춤", "Alt → H → O → A", NONE, ""),
        ("행 숨기기 / 열 숨기기", "Ctrl + 9 / Ctrl + 0", "⌃ 9 / ⌃ 0", ""),
        ("그룹 묶기 / 풀기", "Alt + Shift + → / ←", "⌘ ⇧ K / ⌘ ⇧ J", "숨기기보다 그룹이 검토하기 좋음"),
        ("틀 고정", "Alt → W → F → F", NONE, "맥은 보기 탭 › 틀 고정"),
        ("눈금선 켜기 / 끄기", "Alt → W → V → G", NONE, ""),
        ("수식 그대로 보기", "Ctrl + `", "⌃ `", "숫자 1 왼쪽 키. 큰 파일에서는 느려질 수 있음"),
    ]),
    ("데이터 다루기", [
        ("필터 켜기 / 끄기", "Ctrl + Shift + L  (또는 Alt → D → F → F)", "⌘ ⇧ F", ""),
        ("보이는 셀만 선택", "Alt + ;", "⌘ ⇧ Z", "숨긴 행을 빼고 복사할 때"),
        ("빈칸만 선택", "Ctrl + G, Alt + S, K", "⌃ G › 옵션 › 빈 셀", "선택 후 값 입력 → Ctrl + Enter로 일괄 채움"),
        ("중복 값 제거", "Alt → A → M", NONE, "맥은 데이터 탭 › 중복된 항목 제거"),
        ("중복 값 강조", "Alt → H → L → H → D", NONE, "조건부 서식 › 셀 강조 규칙"),
        ("텍스트 나누기 (20251124 → 날짜)", "Alt → A → E", NONE, "3단계에서 '날짜'를 고르면 2025-11-24로 바뀜"),
        ("피벗 테이블 만들기", "Alt → N → V → T", NONE, "맥은 삽입 탭 › 피벗 테이블"),
        ("자동 합계", "Alt + =", "⌘ ⇧ T", ""),
        ("찾기 / 바꾸기", "Ctrl + F / Ctrl + H", "⌘ F / ⌃ H", ""),
    ]),
    ("시트", [
        ("시트 간 이동", "Ctrl + Page Up / Page Down", "⌥ ← / ⌥ →", ""),
        ("새 시트", "Shift + F11", "⇧ F11", ""),
        ("시트 이름 바꾸기", "Alt → H → O → R", NONE, "맥은 시트 탭 더블클릭"),
        ("시트 이동·복사", "Alt → H → O → M", NONE, ""),
        ("시트 탭 색", "Alt → H → O → T", NONE, ""),
        ("시트 삭제", "Alt → E → L", NONE, "실행 취소가 안 되니 주의"),
    ]),
    ("매크로", [
        ("매크로 목록 열기", "Alt + F8", "⌥ F8", ""),
        ("VBA 편집기 열기", "Alt + F11", "⌥ F11", ""),
    ]),
]

SHORTCUT_NOTES = [
    "**Alt 순차 키는 Windows 전용입니다.** Alt를 누르면 리본에 글자가 뜨고, 그 글자를 하나씩 누르는 방식이라 맥에는 없습니다. 회사 PC는 Windows이니 입사 전에는 원리만 알고, 맥에서는 대응 단축키로 손에 익히세요.",
    "맥북은 F2·F4·F11 같은 기능 키를 **fn과 함께** 눌러야 할 수 있습니다(시스템 설정에서 바꿀 수 있음).",
    "기호: ⌘ Command, ⌃ Control, ⌥ Option, ⇧ Shift.",
]

TIER_DAILY, TIER_SOMETIMES, TIER_EXTRA = "매일 쓰는 기본함수", "가끔 쓰는 응용함수", "함께 알아두면 좋은 함수"
TIER_NOTES = {
    TIER_DAILY: "첨부하신 문서의 '실무에서 매일 쓰는 기본함수'입니다. 이것만 막힘없이 쓰면 대부분의 작업이 됩니다.",
    TIER_SOMETIMES: "문서의 '실무에서 가끔 쓰는 응용함수'입니다. 남이 만든 파일을 읽을 때 특히 자주 만납니다.",
    TIER_EXTRA: "문서에는 없지만 Deals 업무에서 거의 매번 함께 쓰여 추가했습니다.",
}

# name, 한 줄 설명, 구문, 예시 수식, 예시 설명, 주의할 점
FUNCTIONS: list[dict] = [
    dict(tier=TIER_DAILY, name="SUM", summary="지정한 범위의 숫자를 모두 더합니다.",
         syntax="=SUM(범위1, [범위2], …)",
         example="=SUM(D5:D12)", meaning="유동자산 세부 계정(D5~D12)의 합계",
         tip="합계를 낸 뒤에는 원본의 합계 줄과 차이가 0인지 꼭 확인합니다. 범위 중간에 소계 줄이 끼어 있으면 두 번 더해집니다."),
    dict(tier=TIER_DAILY, name="SUBTOTAL", summary="필터로 숨긴 행을 빼고, 보이는 셀만 계산합니다.",
         syntax="=SUBTOTAL(기능번호, 범위)   · 9 = 합계, 3 = 개수(COUNTA), 1 = 평균",
         example="=SUBTOTAL(9, D5:D200)", meaning="필터를 걸면 화면에 남은 행만 합산",
         tip="거래처별·계정별로 필터를 바꿔 가며 금액을 볼 때 씁니다. 손으로 숨긴 행까지 빼려면 9 대신 109를 씁니다."),
    dict(tier=TIER_DAILY, name="SUMIFS", summary="여러 조건을 모두 만족하는 행의 값만 더합니다.",
         syntax="=SUMIFS(합할 범위, 조건범위1, 조건1, [조건범위2, 조건2], …)",
         example='=SUMIFS(BS!$D$4:$D$60, BS!$B$4:$B$60, "차입금")', meaning="분류가 '차입금'인 계정의 합계",
         tip="합할 범위가 맨 앞입니다(SUMIF와 순서가 다름). 범위들의 행 수가 서로 같아야 하고, 복사해서 쓸 수식은 범위를 $로 고정합니다."),
    dict(tier=TIER_DAILY, name="COUNTIFS", summary="여러 조건을 모두 만족하는 행의 개수를 셉니다.",
         syntax="=COUNTIFS(조건범위1, 조건1, [조건범위2, 조건2], …)",
         example='=COUNTIFS(BS!$A$4:$A$60, "유동자산", BS!$B$4:$B$60, "미분류")', meaning="유동자산 중 미분류 계정의 수",
         tip="중복 확인에도 씁니다. =COUNTIFS(A:A, A2)가 2 이상이면 그 값이 중복입니다."),
    dict(tier=TIER_DAILY, name="COUNTA", summary="비어 있지 않은 셀의 개수를 셉니다(문자 포함).",
         syntax="=COUNTA(범위)",
         example="=COUNTA(IS!C4:C40)", meaning="손익계산서 계정명이 몇 줄인지",
         tip="COUNT는 숫자만, COUNTA는 문자까지 셉니다. 받은 자료의 행 수가 원본과 같은지 맞춰 볼 때 유용합니다."),
    dict(tier=TIER_DAILY, name="VLOOKUP", summary="표의 첫 열에서 값을 찾아, 같은 행의 다른 열 값을 가져옵니다.",
         syntax="=VLOOKUP(찾을 값, 표 범위, 가져올 열 번호, FALSE)",
         example='=VLOOKUP("재고자산", BS!$C$4:$H$60, 6, FALSE)', meaning="계정명이 '재고자산'인 행의 6번째 열 값",
         tip="마지막 인수는 항상 FALSE(정확히 일치). 찾는 열이 표의 맨 왼쪽이어야 하고, 중간에 열을 삽입하면 열 번호가 틀어집니다."),
    dict(tier=TIER_DAILY, name="XLOOKUP", summary="찾을 열과 가져올 열을 따로 지정해, 어느 방향이든 찾아옵니다.",
         syntax="=XLOOKUP(찾을 값, 찾을 범위, 가져올 범위, [없을 때 값])",
         example='=XLOOKUP("재고자산", BS!$C$4:$C$60, BS!$H$4:$H$60, 0)', meaning="VLOOKUP과 같은 결과. 열을 삽입해도 깨지지 않음",
         tip="Microsoft 365·Excel 2021 이상에서만 됩니다. 오래된 버전을 쓰는 상대에게 보낼 파일이면 INDEX + MATCH로 씁니다."),
    dict(tier=TIER_DAILY, name="CONCATENATE", summary="여러 문자를 이어 붙입니다. & 기호와 같습니다.",
         syntax='=CONCATENATE(문자1, 문자2, …)   또는   =문자1 & 문자2',
         example='=CONCATENATE(A2, "_", B2)', meaning="'삼성전자_FY2025'처럼 찾기용 키를 만들 때",
         tip="조건이 둘 이상인 VLOOKUP을 할 때, 두 열을 이어 붙인 키 열을 만들어 찾는 방식으로 자주 씁니다."),
    dict(tier=TIER_DAILY, name="LEFT · RIGHT · MID", summary="문자의 왼쪽, 오른쪽, 가운데에서 원하는 만큼 잘라냅니다.",
         syntax="=LEFT(문자, 글자 수)   =RIGHT(문자, 글자 수)   =MID(문자, 시작 위치, 글자 수)",
         example='=RIGHT("FY2025", 4)', meaning="'2025'를 잘라냄. =MID(\"20251124\", 5, 2) 는 '11'",
         tip="결과는 문자입니다. 숫자로 계산하려면 =VALUE(RIGHT(A1, 4)) 또는 =RIGHT(A1, 4)*1 로 바꿉니다."),

    dict(tier=TIER_SOMETIMES, name="INDEX", summary="범위 안에서 지정한 위치(몇 번째 행, 몇 번째 열)의 값을 꺼냅니다.",
         syntax="=INDEX(범위, 행 번호, [열 번호])",
         example="=INDEX(IS!$D$4:$H$40, 3, 5)", meaning="범위의 3번째 행, 5번째 열 값",
         tip="혼자 쓰기보다 MATCH와 짝으로 씁니다. 위치를 MATCH가 찾아 주고, INDEX가 그 자리의 값을 꺼냅니다."),
    dict(tier=TIER_SOMETIMES, name="MATCH", summary="찾는 값이 범위에서 몇 번째에 있는지 알려 줍니다.",
         syntax="=MATCH(찾을 값, 범위, 0)",
         example='=INDEX(IS!$D$4:$H$40, MATCH($A5, IS!$C$4:$C$40, 0), MATCH(B$4, IS!$D$3:$H$3, 0))',
         meaning="계정명(세로)과 연도(가로)를 동시에 찾는 양방향 찾기",
         tip="세 번째 인수 0은 정확히 일치입니다. 생략하면 엉뚱한 값을 가져올 수 있으니 항상 적습니다."),
    dict(tier=TIER_SOMETIMES, name="TRANSPOSE", summary="가로로 놓인 범위를 세로로(또는 반대로) 바꿔서 돌려줍니다.",
         syntax="=TRANSPOSE(범위)",
         example="=TRANSPOSE(IS!D5:H5)", meaning="가로로 놓인 5개년 매출액을 세로 5칸으로",
         tip="값으로 한 번만 바꾸면 되는 경우에는 '행열 바꿔 붙여넣기'가 더 간단합니다. 원본이 바뀔 때 따라 바뀌어야 하면 함수를 씁니다. 구버전에서는 결과 범위를 먼저 선택하고 Ctrl + Shift + Enter로 입력합니다."),
    dict(tier=TIER_SOMETIMES, name="TOCOL", summary="여러 행·열에 걸친 범위를 한 줄짜리 세로 열로 펼칩니다.",
         syntax="=TOCOL(범위, [무시할 값], [열 방향으로 읽기])",
         example="=TOCOL(D5:H7, 1)", meaning="3행 × 5열 표를 15칸 세로 열로. 1은 빈칸 무시",
         tip="Microsoft 365에서만 됩니다. 가로로 넓은 표를 피벗 테이블에 넣기 좋은 긴 형태로 바꿀 때 씁니다."),
    dict(tier=TIER_SOMETIMES, name="INDIRECT", summary="문자로 적힌 주소를 실제 참조로 바꿔 줍니다.",
         syntax='=INDIRECT(주소 문자)',
         example="=SUM(INDIRECT(\"'\" & A2 & \"'!D4:D60\"))", meaning="A2에 적힌 시트 이름의 D4:D60 합계. A2만 바꾸면 다른 시트를 봄",
         tip="편리하지만 참조 추적(Ctrl + [)이 안 되고, 파일이 크면 느려지며, 시트 이름이 바뀌면 조용히 깨집니다. 꼭 필요할 때만 씁니다."),

    dict(tier=TIER_EXTRA, name="$ 절대·혼합 참조", summary="수식을 복사해도 움직이지 않게 행이나 열을 고정합니다.",
         syntax="$D$5 둘 다 고정 · D$5 행만 고정 · $D5 열만 고정   (F4로 전환)",
         example="=D6/D$5", meaning="5행(매출액)은 고정하고 열만 움직여, 수식 하나로 표 전체의 매출 대비 비율 계산",
         tip="'수식 하나를 쓰고 오른쪽·아래로 복사'가 되는 표가 좋은 표입니다. 칸마다 수식이 다르면 검토자가 전부 확인해야 합니다."),
    dict(tier=TIER_EXTRA, name="IF · AND · OR", summary="조건에 따라 다른 값을 돌려줍니다.",
         syntax="=IF(조건, 참일 때, 거짓일 때)",
         example='=IF(ABS(D20) <= 1, "OK", "확인 필요")', meaning="차이가 1 이하면 OK",
         tip="검증 셀을 만들 때 가장 많이 씁니다. IF를 3겹 이상 겹치게 되면 조회표 + VLOOKUP으로 바꾸는 편이 읽기 쉽습니다."),
    dict(tier=TIER_EXTRA, name="IFERROR", summary="수식이 오류일 때 대신 보여줄 값을 정합니다.",
         syntax="=IFERROR(수식, 오류일 때 값)",
         example='=IFERROR(D10/D5, "n/a")', meaning="분모가 0이면 #DIV/0! 대신 n/a",
         tip="오류를 가리는 함수라서 진짜 실수까지 숨깁니다. 나눗셈·찾기처럼 오류 원인이 분명한 곳에만 씁니다."),
    dict(tier=TIER_EXTRA, name="ROUND", summary="지정한 자릿수로 반올림합니다.",
         syntax="=ROUND(숫자, 자릿수)   · 0 = 정수, -6 = 백만 단위",
         example="=ROUND(D5/1000000, 0)", meaning="원 단위를 백만원 단위 정수로",
         tip="표시형식으로 자릿수를 줄이는 것과 다릅니다. ROUND는 값 자체가 바뀌어, 합계에 단수 차이가 생길 수 있습니다."),
    dict(tier=TIER_EXTRA, name="TEXT", summary="숫자나 날짜를 원하는 모양의 문자로 바꿉니다.",
         syntax='=TEXT(값, "서식")',
         example='="매출액 " & TEXT(D5, "#,##0,,") & "백만원"', meaning="'매출액 333,605,938백만원' 같은 문장 만들기",
         tip='서식의 쉼표 두 개(,,)는 백만 단위로 줄여 보여 줍니다. 날짜는 "yyyy-mm-dd"처럼 씁니다.'),
    dict(tier=TIER_EXTRA, name="EOMONTH · EDATE", summary="몇 달 뒤(전)의 월말 또는 같은 날짜를 구합니다.",
         syntax="=EOMONTH(시작일, 개월 수)   =EDATE(시작일, 개월 수)",
         example="=EOMONTH(B3, 1)", meaning="B3 다음 달의 말일. 월별 추정 모델의 날짜 줄에 사용",
         tip="날짜에 30을 더하는 방식은 월말이 어긋납니다. 월 단위 이동은 이 두 함수로 합니다."),
]

# (제목, 설명 마크다운, 코드 또는 None)
MACRO_SECTIONS: list[tuple[str, str, str | None]] = [
    ("1. 매크로란, 그리고 언제 쓰나", """
매크로는 엑셀에서 하는 조작을 **VBA라는 언어로 적어 두고 버튼 하나로 다시 실행**하는 기능입니다.

쓰기 좋은 일은 "같은 조작을 여러 번 반복하는 일"입니다.
- 시트 30개에 같은 서식 입히기, 시트별로 PDF 저장하기
- 받은 자료를 매번 같은 순서로 정리하기 (빈 행 삭제, 열 순서 바꾸기, 서식 적용)
- 모델 점검 (하드코딩된 숫자 색칠, 시트 목차 만들기)

반대로 **계산 자체는 매크로가 아니라 수식으로** 합니다. 매크로로 계산한 값은 어떻게 나왔는지 셀에 남지 않아 검토할 수 없기 때문입니다.
""", None),
    ("2. 준비: 개발 도구 탭과 저장 형식", """
**개발 도구 탭 켜기**
- Windows: 파일 › 옵션 › 리본 사용자 지정 › 오른쪽 목록에서 '개발 도구' 체크
- Mac: Excel › 환경설정 › 리본 메뉴 및 도구 모음 › '개발 도구' 체크

**저장 형식**: 매크로가 든 파일은 `.xlsx`로 저장하면 매크로가 사라집니다. **Excel 매크로 사용 통합 문서(.xlsm)** 로 저장합니다.

**보안 경고**: 매크로가 든 파일을 열면 노란 경고 줄이 뜹니다. 출처를 아는 파일만 '콘텐츠 사용'을 누릅니다. 회사 보안 정책에 따라 외부에서 받은 매크로 파일은 아예 실행이 막혀 있을 수 있습니다.
""", None),
    ("3. 첫 매크로: 기록해서 만들기", """
코드를 몰라도 **매크로 기록**으로 시작할 수 있습니다.

1. 개발 도구 › **매크로 기록** → 이름을 정하고 확인
2. 평소처럼 조작합니다 (예: 범위 선택 → 천 단위 서식 → 굵게)
3. 개발 도구 › **기록 중지**
4. 개발 도구 › 매크로(Alt + F8 / ⌥ F8) › 실행

기록된 코드는 VBA 편집기(Alt + F11 / ⌥ F11)에서 볼 수 있습니다. **기록된 코드를 읽고 고쳐 보는 것**이 가장 빠른 공부법입니다.

기록할 때 '상대 참조로 기록'을 켜면 "지금 선택한 셀 기준"으로, 끄면 "항상 같은 셀"로 기록됩니다. 아래는 기록된 코드를 다듬은 모습입니다.
""", '''Sub FormatNumbers()
    ' 선택한 범위에 천 단위 구분, 음수 괄호 서식을 적용한다
    With Selection
        .NumberFormat = "#,##0;(#,##0);""-"""
        .HorizontalAlignment = xlRight
    End With
End Sub'''),
    ("4. 꼭 아는 문법 여섯 가지", """
| 문법 | 뜻 |
|---|---|
| `Sub 이름() … End Sub` | 매크로 하나의 시작과 끝 |
| `Dim r As Long` | 변수 선언. 행 번호는 Long, 금액은 Double, 문자는 String |
| `Range("A1")`, `Cells(행, 열)` | 셀 가리키기. 반복문에서는 숫자로 쓰는 Cells가 편함 |
| `For r = 4 To lastRow … Next r` | 정해진 횟수만큼 반복 |
| `For Each ws In Worksheets … Next ws` | 모든 시트(또는 셀)를 하나씩 |
| `If 조건 Then … Else … End If` | 조건에 따라 나누기 |

데이터가 몇 행까지 있는지는 매번 달라지므로, **마지막 행을 찾는 한 줄**을 외워 둡니다.
""", """Sub LastRowExample()
    Dim lastRow As Long
    ' C열의 맨 아래에서 위로 올라오다 처음 만나는 값이 있는 행
    lastRow = Cells(Rows.Count, 3).End(xlUp).Row
    MsgBox "마지막 행: " & lastRow
End Sub"""),
    ("5. 예제: 하드코딩된 숫자 찾기", """
모델을 검토할 때 "수식이어야 할 자리에 숫자가 직접 입력된 셀"을 찾는 일이 많습니다.
현재 시트에서 **직접 입력된 숫자만 골라 파란 글씨로** 바꾸는 매크로입니다.
""", """Sub HighlightHardcodes()
    Dim rng As Range
    On Error Resume Next        ' 해당하는 셀이 하나도 없으면 오류가 나므로 건너뛴다
    Set rng = ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants, xlNumbers)
    On Error GoTo 0
    If Not rng Is Nothing Then rng.Font.Color = RGB(0, 0, 255)
End Sub"""),
    ("6. 예제: 모든 시트의 목차 만들기", """
시트가 많은 파일에서 맨 앞에 **시트 이름 목록과 바로가기 링크**를 만드는 매크로입니다. `For Each`로 시트를 하나씩 도는 전형적인 형태입니다.
'목차'라는 시트가 이미 있으면 이름이 겹쳐 오류가 나니, 다시 실행할 때는 먼저 지웁니다.
""", """Sub MakeIndexSheet()
    Dim ws As Worksheet, idx As Worksheet, r As Long

    Set idx = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    idx.Name = "목차"
    idx.Range("A1").Value = "시트 목록"
    r = 2

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> idx.Name Then
            idx.Hyperlinks.Add Anchor:=idx.Cells(r, 1), Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", TextToDisplay:=ws.Name
            r = r + 1
        End If
    Next ws
End Sub"""),
    ("7. 고치고 확인하는 법", """
- **한 줄씩 실행**: 편집기에서 F8을 누르면 한 줄씩 진행됩니다. 엑셀 창을 옆에 두고 무엇이 바뀌는지 봅니다.
- **중단점**: 멈추고 싶은 줄에서 F9. 실행하면 그 줄에서 멈춥니다.
- **변수 값 보기**: 멈춘 상태에서 변수에 마우스를 올리거나, 직접 실행 창(Ctrl + G)에 `?lastRow`를 입력합니다.
- 모듈 맨 위에 `Option Explicit`을 적어 두면, 변수 이름 오타를 실행 전에 잡아 줍니다.
""", None),
    ("8. 반드시 기억할 주의점", """
- **매크로가 한 일은 실행 취소(Ctrl + Z)가 안 됩니다.** 실행 전에 저장하거나 사본으로 연습합니다.
- 코드에서 `ActiveSheet`·`Selection`은 "지금 보고 있는 곳"입니다. 엉뚱한 시트를 보고 실행하면 그 시트가 바뀝니다.
- 고객 자료 파일에는 회사 정책상 매크로 사용이 제한될 수 있습니다. 입사 후 팀의 방식을 먼저 확인하세요.
- Mac의 VBA는 대부분 같지만 파일 경로 표기와 일부 창(사용자 정의 폼 등)이 다릅니다. 이 페이지의 예제는 양쪽에서 같은 코드로 동작하도록 골랐습니다.
- 반복 정리 작업은 **파워 쿼리**가 더 맞는 경우도 많습니다. 매크로에 익숙해진 뒤 살펴볼 다음 주제입니다.
""", None),
]
