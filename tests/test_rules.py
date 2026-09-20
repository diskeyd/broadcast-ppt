from broadcast_ppt.rules.fields import (
    announcement_items, english_line, mask_names, name_role, page_ref, quoted_title, scripture_ref, verse_chunks,
)


def test_scripture_ref_forms():
    r = scripture_ref("삼상26:21-25")
    assert (r.book, r.chapter, r.verse_start, r.verse_end, r.testament) == ("사무엘상", 26, 21, 25, "old")
    r = scripture_ref("사무엘상 7:7-12")
    assert (r.book, r.verse_end) == ("사무엘상", 12)
    r = scripture_ref("시편 62편")  # no verse -> not a reference
    assert r is None
    r = scripture_ref("시편 73편 1-3절")  # 편 is how a Psalms chapter is usually written
    assert (r.book, r.chapter, r.verse_start, r.verse_end) == ("시편", 73, 1, 3)
    assert scripture_ref("시편 73편 1절").verse_end == 1
    r = scripture_ref("에베소서 4장 11절~16절")
    assert (r.book, r.chapter, r.verse_start, r.verse_end) == ("에베소서", 4, 11, 16)
    r = scripture_ref("[롬3:23-24,27-28] 23 모든 사람이")
    assert (r.book, r.verse_start, r.verse_end, r.testament) == ("로마서", 23, 24, "new")
    assert scripture_ref("마26:6-13").book == "마태복음"
    assert scripture_ref("요한계시록1:8").book == "요한계시록"
    assert scripture_ref("2026 하반기 말씀사경회") is None


def test_page_and_title():
    assert page_ref("구약성경 p.417") == {"testament": "old", "page": 417, "raw": "구약성경 p.417"}
    assert page_ref("p 1050")["page"] == 1050
    assert page_ref("-  구 455쪽  -") == {"testament": "old", "page": 455, "raw": "구 455쪽"}
    assert page_ref("- 신 172쪽 -")["testament"] == "new"
    assert page_ref("사무엘상26:21-25") is None
    assert quoted_title("설 교 “갚으시리니” 삼상26") == "갚으시리니"
    assert english_line("기 도\nPrayer of worship\n") == "Prayer of worship"


def test_name_role():
    assert name_role("기 도 Prayer 홍길동 장로") == [{"name": "홍길동", "role": "장로"}]
    assert name_role("담임목사 위임식") == []


def test_mask_names():
    assert mask_names("기 도 홍길동 장로") == "기 도 [이름] 장로"
    assert mask_names("홍길동 장로님") == "[이름] 장로님"
    assert mask_names("담임목사 말씀") == "담임목사 말씀"
    assert mask_names("하나님은 우리와 함께하십니다") == "하나님은 우리와 함께하십니다"


def test_announcement_items():
    text = (
        "우리는 함께 세우는 교회입니다\n"
        "환영 : 방문해주신 모든 분들을 환영합니다\n"
        "[2026 표어] : “교회는 그의 몸이니”\n"
        "[상반기 모임(5월)] :  1차 모임 주간\n"
        "      추가 보고\n"
        "결혼: 두 가정의 결혼을 축복합니다\n"
        "[말씀사경회] : 다음 주 금요일 저녁\n"
        "강사 : 외부 강사\n"
    )
    items = announcement_items(text)
    assert [i["title"] for i in items] == ["환영", "2026 표어", "상반기 모임(5월)", "결혼", "말씀사경회"]
    assert [i["no"] for i in items] == [1, 2, 3, 4, 5]
    assert items[4]["body"] == "다음 주 금요일 저녁 강사 : 외부 강사"
    assert items[2]["body"] == "1차 모임 주간 추가 보고"
    assert items[1]["bracketed"] and not items[0]["bracketed"]


def test_verse_chunks():
    ref = scripture_ref("삼상26:21-25")
    chunks = verse_chunks("삼상26:21-25 21 사울이 이르되 2026년 22 다윗이 대답하여", ref)
    assert [n for n, _ in chunks] == [None, 21, 22]
