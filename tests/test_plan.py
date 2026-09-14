from broadcast_ppt.mapping.split import pack
from broadcast_ppt.mapping.textfit import WidthModel
from broadcast_ppt.rules import kinds as K
from broadcast_ppt.rules.classify import DeckContext, classify_slide


def test_pack_respects_fit_and_order():
    units = ["21 가나다라마바사", "22 아자차카타파하", "23 가나다", "24 라마바사아자차카타파하가나다"]
    fits = lambda t: len(t) <= 20
    chunks = pack(units, fits)
    assert all(fits(c) for c in chunks)
    assert " ".join(chunks).replace("  ", " ") == " ".join(units)


def test_pack_breaks_an_oversized_unit():
    chunks = pack(["첫 문장입니다. 둘째 문장입니다. 셋째 문장입니다."], lambda t: len(t) <= 10)
    assert all(len(c) <= 10 for c in chunks) and len(chunks) >= 3


def test_width_model_monotonic_and_scale():
    m = WidthModel()
    short, long_ = "가나다", "가나다" * 30
    assert m.estimate_lines(short, 9_000_000, 26) == 1
    assert m.estimate_lines(long_, 9_000_000, 26) > m.estimate_lines(long_, 9_000_000, 20)
    wide = WidthModel(width_scale=1.2)
    assert wide.estimate_lines(long_, 9_000_000, 26) <= m.estimate_lines(long_, 9_000_000, 26)


def _slide(index, text, kinds=None, full_bleed=False):
    boxes = [{"kind": "TEXT_BOX", "text": t, "bbox": [0, 0, 100, 100], "depth": 0} for t in text.split("||") if t]
    return {"index": index, "text": "\n".join(b["text"] for b in boxes), "shapes": boxes,
            "kinds": kinds or {"TEXT_BOX": len(boxes)}, "full_bleed_picture": full_bleed, "keywords": [], "text_len": len(text)}


def test_classifier_core_rules():
    ctx = DeckContext()
    assert classify_slide(_slide(2, "예배로의 부름||meditation"), ctx).kind == K.CARD_TITLE
    d = classify_slide(_slide(3, "기 도||Prayer||홍길동 장로"), ctx)
    assert d.kind == K.PRAYER_TITLE and d.fields["role"] == "장로"
    d = classify_slide(_slide(4, "성경봉독  Scripture Reading||사무엘상26:21-25||-  구 455쪽  -"), ctx)
    assert d.kind == K.SCRIPTURE_TITLE and d.fields["page"]["page"] == 455
    d = classify_slide(_slide(5, "삼상26:21-25||21 사울이 이르되 내가 범죄하였도다 내 아들 다윗아 돌아오라 22 다윗이 대답하여 이르되"), ctx)
    assert d.kind == K.SCRIPTURE_BODY and d.fields["ref"].verse_start == 21
    d = classify_slide(_slide(6, "23 여호와께서 각 사람에게 그의 공의와 신실을 따라 갚으시리니 오늘 여호와께서"), ctx)
    assert d.kind == K.SCRIPTURE_BODY and d.fields.get("inherited_ref")
    d = classify_slide(_slide(7, "환영 : 방문해주신 모든 분들을 환영합니다\n[표어] : 교회는 그의 몸이니\n[모임] : 오늘 오후"), ctx)
    assert d.kind == K.ANNOUNCEMENT_BLOCK and len(d.fields["items"]) == 3
    d = classify_slide(_slide(8, "동행의 기도||하나님께서 갚아주시는 삶의 길을 걷게 하소서!"), ctx)
    assert d.kind == K.CARD_TITLE and d.fields["label"] == "동행의 기도" and d.fields["prayer_text"]
    assert classify_slide(_slide(9, "", {"PICTURE": 1}), ctx).kind == K.PICTURE_ONLY
