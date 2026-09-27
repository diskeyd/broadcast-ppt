from broadcast_ppt.mapping.split import pack
from broadcast_ppt.mapping.textfit import WidthModel
from broadcast_ppt.rules import kinds as K
from broadcast_ppt.rules.classify import DeckContext, classify_slide
from broadcast_ppt.rules.fields import scripture_ref


def test_pack_respects_fit_and_order():
    units = ["21 가나다라마바사", "22 아자차카타파하", "23 가나다", "24 라마바사아자차카타파하가나다"]
    def fits(t):
        return len(t) <= 20
    chunks = pack(units, fits)
    assert all(fits(c) for c in chunks)
    assert " ".join(chunks).replace("  ", " ") == " ".join(units)


def test_pack_breaks_an_oversized_unit():
    chunks = pack(["첫 문장입니다. 둘째 문장입니다. 셋째 문장입니다."], lambda t: len(t) <= 10)
    assert all(len(c) <= 10 for c in chunks) and len(chunks) >= 3


def test_pack_moves_a_whole_unit_to_balance_the_last_chunk():
    # the shape that split a verse across two caption pages: three verses fill a chunk, the fourth
    # is small enough to trigger the rebalance, and the rebalance used to move two *words* back
    units = [f"{n} " + "가" * c for n, c in ((8, 31), (9, 31), (10, 36), (11, 34))]
    chunks = pack(units, lambda t: len(t) <= 110)
    assert [c.split(" ")[0] for c in chunks] == ["8", "10"]   # every page starts on a verse
    assert " ".join(chunks) == " ".join(units)


def test_pack_keeps_each_unit_on_its_own_line():
    units = ["1 가나다 라마바 사아자", "2 차카타 파하가 나다라", "3 짧다"]
    chunks = pack(units, lambda t: len(t) <= 30, sep="\n")
    # the tiny-last-chunk rebalance must move whole units, never two words across a verse boundary
    assert all(line[:1].isdigit() for c in chunks for line in c.split("\n"))
    assert "\n".join(chunks) == "\n".join(units)


def test_width_model_monotonic_and_scale():
    m = WidthModel()
    short, long_ = "가나다", "가나다" * 30
    assert m.estimate_lines(short, 9_000_000, 26) == 1
    assert m.estimate_lines(long_, 9_000_000, 26) > m.estimate_lines(long_, 9_000_000, 20)
    wide = WidthModel(width_scale=1.2)
    assert wide.estimate_lines(long_, 9_000_000, 26) <= m.estimate_lines(long_, 9_000_000, 26)


def _slide(index, text, kinds=None, full_bleed=False, tops=None):
    """`tops` gives the boxes distinct y positions, so a test can put them on screen out of XML order."""
    boxes = [{"kind": "TEXT_BOX", "text": t, "bbox": [0, tops[i] if tops else 0, 100, 100], "depth": 0}
             for i, t in enumerate(t for t in text.split("||") if t)]
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
    d = classify_slide(_slide(61, "[롬3:23-24,27-28] 23 모든 사람이 죄를 범하였으매 하나님의 영광에 이르지 못하더니"), ctx)
    assert d.kind == K.SCRIPTURE_BODY and d.fields["ref"].contains(27) and not d.fields["ref"].contains(25)
    d = classify_slide(_slide(62, "21 사울이 이르되 내가 범죄하였도다 내 아들 다윗아 돌아오라 내가 다시는 너를 해하지 아니하리니"), ctx)
    assert d.rule_id == "R210" and d.fields["ref"].chapter == 26  # outside 롬3 -> the 성경봉독 reference
    d = classify_slide(_slide(63, "99 전혀 다른 본문이 절 번호로 시작하는 경우에는 참조를 비워 두어야 합니다 그렇지 않으면"), ctx)
    assert d.rule_id == "R211" and d.fields["ref"] is None and "review:reference_unknown" in d.flags
    two = _slide(64, "[롬3:23-24] 23 모든 사람이 죄를 범하였으매 하나님의 영광에 이르지 못하더니 24 그리스도 예수 안에 있는 속량으로")
    two["shapes"].append({"kind": "TEXT_BOX", "text": "[고전4:20] 하나님의 나라는 말에 있지 아니하고 오직 능력에 있음이라", "bbox": [0, 0, 1, 1]})
    two["text"] += "\n[고전4:20] 하나님의 나라는 말에 있지 아니하고 오직 능력에 있음이라"
    d = classify_slide(two, ctx)
    assert d.kind == K.SCRIPTURE_BODY and [e["ref"].book for e in d.fields["extra"]] == ["고린도전서"]
    d = classify_slide(_slide(7, "환영 : 방문해주신 모든 분들을 환영합니다\n[표어] : 교회는 그의 몸이니\n[모임] : 오늘 오후"), ctx)
    assert d.kind == K.ANNOUNCEMENT_BLOCK and len(d.fields["items"]) == 3
    d = classify_slide(_slide(8, "동행의 기도||하나님께서 갚아주시는 삶의 길을 걷게 하소서!"), ctx)
    assert d.kind == K.CARD_TITLE and d.fields["label"] == "동행의 기도" and d.fields["prayer_text"]
    assert classify_slide(_slide(9, "", {"PICTURE": 1}), ctx).kind == K.PICTURE_ONLY
    d = classify_slide(_slide(10, "묵 상   Meditate||여호와께서 여기까지 우리를 도우셨다"), ctx)
    assert d.kind == K.MEDITATION and d.fields["quote"] == "여호와께서 여기까지 우리를 도우셨다"


def test_print_plan_handles_every_fit_shape(capsys):
    from broadcast_ppt.cli import _print_plan
    plan = {"slides": [{"src": 1, "kind": "meditation", "rule": "R160"}, {"src": 2, "kind": "scripture_body", "rule": "R200"}],
            "ops": [{"op": "template", "src": 1, "template_slide": 8, "fit": {"chars": 25, "sz": 62.0, "base_sz": 72.0}},
                    {"op": "copy", "src": 2},
                    {"op": "template", "src": 2, "template_slide": 61, "fit": {"chars": 76, "est_lines": 2, "max_lines": 2}, "flags": ["review:x"]}],
            "stats": {}}
    _print_plan(plan)
    out = capsys.readouterr().out
    assert "T8(25자/62pt)" in out and "T61(76자/2줄)" in out


def _ai_deck(text):
    return {"slides": [_slide(1, text)]}


def test_ai_fallback_applies_confident_answer(monkeypatch):
    from broadcast_ppt.rules import ai
    from broadcast_ppt.rules.classify import classify_deck
    seen = {}

    def fake(payload, model):
        seen.update(payload)
        return {"kind": "caption_sentence", "confidence": 0.9, "lines": [2, 3], "reason": "short sermon sentences"}

    monkeypatch.setattr(ai, "ask", fake)
    text = "가나다라마바사아자차카타파하 홍길동 집사||하나님은 우리와 함께하시며 우리를 붙드십니다 그러므로||우리는 두려워하지 않고 오늘도 걸어갑니다 믿음으로||" + "라" * 70
    assert classify_deck(_ai_deck(text))[0].kind == K.UNKNOWN  # no fallback: rules only
    d = classify_deck(_ai_deck(text), fallback=ai.classify_view)[0]
    assert (d.kind, d.rule_id, d.flags) == (K.CAPTION_SENTENCE, "AI", ["review:ai_classified"])
    assert d.fields["body"] == "하나님은 우리와 함께하시며 우리를 붙드십니다 그러므로 우리는 두려워하지 않고 오늘도 걸어갑니다 믿음으로"
    assert "[이름] 집사" in seen["lines"][0] and "홍길동" not in " ".join(seen["lines"])


def test_ai_fallback_keeps_unknown_on_low_confidence_or_bad_answer(monkeypatch):
    from broadcast_ppt.rules import ai
    from broadcast_ppt.rules.classify import classify_deck
    text = "가" * 40 + "||" + "나" * 40 + "||" + "다" * 40
    monkeypatch.setattr(ai, "ask", lambda p, m: {"kind": "graphic", "confidence": 0.5, "lines": [1], "reason": "unsure"})
    d = classify_deck(_ai_deck(text), fallback=ai.classify_view)[0]
    assert (d.kind, d.rule_id) == (K.UNKNOWN, "R900") and d.fields["ai"]["kind"] == "graphic"
    for bad in ({"kind": "prayer_title", "confidence": 0.99, "lines": [1], "reason": ""},
                {"kind": "meditation", "confidence": 0.99, "lines": [1], "reason": ""},  # has a keyword: rules own it
                {"kind": "caption_sentence", "confidence": 0.99, "lines": [99], "reason": ""}):
        monkeypatch.setattr(ai, "ask", lambda p, m, bad=bad: bad)
        d = classify_deck(_ai_deck(text), fallback=ai.classify_view)[0]
        assert (d.kind, d.rule_id) == (K.UNKNOWN, "R900")
    monkeypatch.setattr(ai, "ask", lambda p, m: (_ for _ in ()).throw(RuntimeError("model not found")))
    assert classify_deck(_ai_deck(text), fallback=ai.classify_view)[0].rule_id == "R900"


def test_scripture_reads_every_box_in_reading_order():
    ctx = DeckContext()
    # XML order puts the longest box first; on screen a heading sits above it and a verse box below
    s = _slide(20, "[사55:8-9] 8 가나다라마바사 아자차카타파하 가나다라마바사아자차카타파하 가나다라마||"
                   "[제 삼 권, 아삽의 시]||"
                   "17 하나다라마바사 아자차카타파하 가나다라마바사아자차카타파하 가나다라마바사",
               tops=[200, 100, 300])
    d = classify_slide(s, ctx)
    assert d.kind == K.SCRIPTURE_BODY and d.rule_id == "R205"
    # the heading is neither a reference nor verse-numbered: it fills the tag instead of being dropped
    assert d.fields["heading"] == "제 삼 권, 아삽의 시"
    assert d.fields["ref"].book == "이사야" and d.fields["body"].startswith("8 ")
    # the verse box below used to vanish: only bracketed boxes were kept as extras
    assert [u["body"][:2] for u in d.fields["extra"]] == ["17"]
    assert "review:reference_unknown" in d.flags


def test_scripture_takes_a_reference_printed_under_the_body():
    ctx = DeckContext()
    s = _slide(21, "26 가나다라마바사 아자차카타파하 가나다라마바사아자차카타파하 가나다라마바사||창세기50:22-26",
               tops=[100, 200])
    d = classify_slide(s, ctx)
    assert d.rule_id == "R200" and d.fields["ref"].chapter == 50
    assert d.fields["body"].startswith("26 ") and not d.fields["extra"]


def test_scripture_extra_boxes_survive_without_a_reference_on_the_slide():
    ctx = DeckContext(title_ref=scripture_ref("시편 73편 18-28절"))
    s = _slide(22, "18 가나다라마바사 아자차카타파하 가나다라마바사아자차카타파하 가나다라||"
                   "19 하나다라마바사 아자차카타파하 가나다라마바사아자차카타파하 가나다라",
               tops=[100, 200])
    d = classify_slide(s, ctx)
    # R210/R211 used to build no `extra` key at all, so every sibling box was lost
    assert d.rule_id == "R210" and len(d.fields["extra"]) == 1
    assert d.fields["extra"][0]["ref"].chapter == 73


def test_line_estimate_follows_the_calibrated_width_model():
    from pathlib import Path

    from broadcast_ppt.cli import CATALOG, WIDTHS
    from broadcast_ppt.mapping.catalog import load_catalog
    from broadcast_ppt.mapping.plan import Mapper, PlanOptions
    from broadcast_ppt.mapping.textfit import load_model, strict_model

    model = load_model(Path(WIDTHS))
    m = Mapper(load_catalog(Path(CATALOG)), model, PlanOptions())
    tpl = m.fam["scripture_ref"][-1][0]
    spec = m._body_spec(tpl)
    bp = spec["bodyPr"]
    # the width factor is fitted from PowerPoint's own autofit boxes, and PowerPoint is what the
    # broadcast PC renders with: the planner answers with that, not with the stricter guard
    for n in range(1, 40):
        text = "가나다라마바사 " * n
        assert m.lines(tpl, text) == model.estimate_lines(text, spec["bbox"][2], m._sz(spec, "body"),
                                                          bp["lIns"], bp["rIns"])
    # the strict model still exists, still narrower, and is what the validator warns from
    assert strict_model(model).width_scale < model.width_scale
    # and a chunk is never packed past what the template can hold
    assert all(m.lines(tpl, c) <= spec["max_lines"]
               for c in pack(["가나다라마바사 " * 12] * 4, lambda t: m.fits(tpl, t), sep="\n"))


def test_scripture_caption_puts_each_verse_on_its_own_line_and_uses_the_heading(tmp_path):
    from pathlib import Path

    from pptx import Presentation
    from pptx.util import Inches

    from broadcast_ppt.cli import CATALOG, WIDTHS
    from broadcast_ppt.mapping.plan import PlanOptions, build_plan

    prs = Presentation()
    # the 성경봉독 card writes its chapter with 편, which used to leave the deck with no title reference
    card = prs.slides.add_slide(prs.slide_layouts[6])
    for i, part in enumerate(("성경봉독  Scripture Reading", "시편 73편 18-28절", "-  구 455쪽  -")):
        card.shapes.add_textbox(Inches(1), Inches(1 + i), Inches(8), Inches(1)).text_frame.text = part
    # heading above, verses below, and verses 1-3 fall outside the title reference: the tag stays unknown
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(1)).text_frame.text = "[제 삼 권, 아삽의 시]"
    s.shapes.add_textbox(Inches(1), Inches(3), Inches(8), Inches(3)).text_frame.text = (
        "1 가나다라마바사 아자차카타파하 2 하나다라마바사 아자차카타파하 3 마나다라마바사 아자차카타파하")
    src = tmp_path / "psalm.pptx"
    prs.save(str(src))

    plan = build_plan(src, Path(CATALOG), Path(WIDTHS), PlanOptions())
    assert (plan["slides"][0]["rule"], plan["slides"][1]["rule"]) == ("R120", "R211")
    caps = [o for o in plan["ops"] if o["op"] == "template" and o["src"] == 2]
    assert caps and all(o["fill"]["ref"] == "제 삼 권, 아삽의 시" for o in caps)
    assert all("review:no_reference_tag" not in o.get("flags", []) for o in caps)
    # every line of a caption starts with its verse number, and the line count stays honest
    for o in caps:
        assert all(line.split(" ")[0].isdigit() for line in o["fill"]["body"].split("\n"))
        assert o["fit"]["est_lines"] <= o["fit"]["max_lines"]
    assert sum(o["fill"]["body"].count("\n") + 1 for o in caps) == 3

    # a reading with its own reference gets no line breaks: the tag already carries the range
    plain = prs.slides.add_slide(prs.slide_layouts[6])
    # the hard line break is how the original slide was laid out; the bar is a different width
    plain.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(3)).text_frame.text = (
        "시편 73:18-20 18 가나다라마바사 아자차카타파하\n19 하나다라마바사 아자차카타파하 20 마나다라마바사")
    src2 = tmp_path / "reading.pptx"
    prs.save(str(src2))
    plan = build_plan(src2, Path(CATALOG), Path(WIDTHS), PlanOptions())
    assert plan["slides"][2]["rule"] == "R200"
    reading = [o for o in plan["ops"] if o["op"] == "template" and o["src"] == 3]
    assert reading and all("\n" not in o["fill"]["body"] for o in reading)
    assert all(o["fill"]["ref"].startswith("시편73:") for o in reading)


def test_overrides_ref_and_quote_card(tmp_path):
    from pathlib import Path

    import pytest
    from pptx import Presentation
    from pptx.util import Inches

    from broadcast_ppt.cli import CATALOG, WIDTHS
    from broadcast_ppt.mapping.plan import PlanOptions, build_plan

    prs = Presentation()
    for text in ("[요한복음 3:16] 하나님이 세상을 이처럼 사랑하사 독생자를 주셨으니 이는 그를 믿는 자마다 멸망하지 않고",
                 "1 태초에 말씀이 계시니라 이 말씀이 하나님과 함께 계셨으니 2 그가 태초에 하나님과 함께 계셨고",
                 "가나다라마바사아자차카타파하 가나다라마바사||아자차카타파하 가나다라마바사아자||차카타파하 가나다라마바사아자차카타파하"):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        for i, part in enumerate(text.split("||")):
            s.shapes.add_textbox(Inches(1), Inches(1 + i), Inches(8), Inches(1)).text_frame.text = part
    src = tmp_path / "src.pptx"
    prs.save(str(src))

    plain = build_plan(src, Path(CATALOG), Path(WIDTHS), PlanOptions())
    assert [(r["rule"], "review:reference_unknown" in r["flags"]) for r in plain["slides"]][:2] == [("R205", False), ("R211", True)]

    plan = build_plan(src, Path(CATALOG), Path(WIDTHS), PlanOptions(), {"slides": {2: {"ref": "요 1:1-2"}, 3: {"op": "quote"}}})
    caps = [o for o in plan["ops"] if o["src"] == 2 and o["op"] == "template"]
    assert caps and all(o["fill"]["ref"] == "요한복음1:1-2" and "review:no_reference_tag" not in o["flags"] for o in caps)
    assert "review:reference_unknown" not in plan["slides"][1]["flags"]
    quote = [o for o in plan["ops"] if o["src"] == 3]
    assert len(quote) == 1 and quote[0]["template_slide"] == 8 and quote[0]["fill"]["label"] == ""
    assert quote[0]["fill"]["quote"].startswith("“가나다") and (plan["slides"][2]["kind"], plan["slides"][2]["rule"]) == ("meditation", "OVR")

    with pytest.raises(ValueError):
        build_plan(src, Path(CATALOG), Path(WIDTHS), PlanOptions(), {"slides": {3: {"ref": "요 1:1"}}})


def test_catechism_splits_qa_and_verse(tmp_path):
    from pathlib import Path

    from pptx import Presentation
    from pptx.util import Inches

    from broadcast_ppt.cli import CATALOG, WIDTHS
    from broadcast_ppt.mapping.plan import PlanOptions, build_plan

    parts = ["제1문 사람의 제일 되는 목적은 무엇입니까?", "답: 사람의 제일 되는 목적은\n하나님을 영화롭게 하는 것입니다",
             "[엡1:3] 찬송하리로다 하나님 곧 우리 주 예수 그리스도의 아버지께서\n그리스도 안에서 하늘에 속한 모든 신령한 복으로"]
    d = classify_slide(_slide(15, "||".join(parts)), DeckContext())
    assert (d.kind, d.rule_id) == (K.CATECHISM, "R170") and len(d.fields["verses"]) == 1
    ref = d.fields["verses"][0]["ref"]
    assert (ref.book, ref.chapter, ref.verse_start) == ("에베소서", 1, 3)
    assert d.fields["qa"].startswith("제1문") and "영화롭게" in d.fields["qa"] and "찬송" not in d.fields["qa"]
    assert d.fields["qa"].split("\n")[1].startswith("답")  # question line, answer below

    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    for i, part in enumerate(parts):
        s.shapes.add_textbox(Inches(1), Inches(1 + i * 2), Inches(8), Inches(1.5)).text_frame.text = part
    src = tmp_path / "src.pptx"
    prs.save(str(src))
    ops = build_plan(src, Path(CATALOG), Path(WIDTHS), PlanOptions())["ops"]
    assert [o["note"] for o in ops] == ["catechism full screen", "caption", "scripture caption"]
    assert ops[1]["fill"]["body"].startswith("제1문") and ops[2]["fill"]["ref"] == "에베소서1:3"
    assert ops[1]["para_style"] == [{"color": "FF9300"}, {"sz": 28.0}] and "para_style" not in ops[2]


def test_overrides_caption_and_scripture_ops(tmp_path, monkeypatch):
    from pathlib import Path

    import pytest
    from pptx import Presentation
    from pptx.util import Inches

    from broadcast_ppt.cli import CATALOG, WIDTHS
    from broadcast_ppt.mapping.plan import PlanOptions, build_plan

    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    for i, part in enumerate(("새 양식 첫 줄 가나다라마바사아자차카타파하", "둘째 줄 가나다라마바사아자차카타파하", "셋째 줄 가나다라마바사아자차카타파하 끝")):
        s.shapes.add_textbox(Inches(1), Inches(1 + i), Inches(8), Inches(1)).text_frame.text = part
    prs.slides.add_slide(prs.slide_layouts[6])  # slide 2: no text
    src = tmp_path / "src.pptx"
    prs.save(str(src))
    build = lambda over: build_plan(src, Path(CATALOG), Path(WIDTHS), PlanOptions(), {"slides": {1: over}})  # noqa: E731

    assert build({})["slides"][0]["rule"] == "R900"
    cap = build({"op": "caption"})
    assert (cap["slides"][0]["kind"], cap["slides"][0]["rule"]) == (K.CAPTION_SENTENCE, "OVR")
    assert [o["note"] for o in cap["ops"]][:2] == ["sentence full screen", "caption"] and cap["ops"][1]["fill"]["body"].startswith("새 양식")
    scr = build({"op": "scripture", "ref": "요 1:1-2"})
    caps = [o for o in scr["ops"] if o["note"] == "scripture caption"]
    assert caps and all(o["fill"]["ref"] == "요한복음1:1-2" for o in caps) and scr["slides"][0]["flags"] == []
    assert build({"op": "scripture"})["slides"][0]["flags"] == ["review:reference_unknown"]
    with pytest.raises(ValueError):
        build({"op": "bogus"})
    with pytest.raises(ValueError):  # a typed reference would be dropped silently
        build({"op": "caption", "ref": "요 1:1"})
    with pytest.raises(ValueError):  # a blank quote card must not go on air
        build_plan(src, Path(CATALOG), Path(WIDTHS), PlanOptions(), {"slides": {2: {"op": "quote"}}})

    from broadcast_ppt.rules import ai
    calls = []
    monkeypatch.setattr(ai, "ask", lambda p, m: calls.append(m) or {"kind": "unknown", "confidence": 0.1, "lines": [], "reason": ""})
    build_plan(src, Path(CATALOG), Path(WIDTHS), PlanOptions(ai_model="test-model"), {"slides": {1: {"skip": True}}})
    assert calls == []  # an overridden slide never reaches the model
    build_plan(src, Path(CATALOG), Path(WIDTHS), PlanOptions(ai_model="test-model"))
    assert calls == ["test-model"]


def test_catechism_rule_edges():
    ctx = DeckContext()
    d = classify_slide(_slide(7, "환영 : 오늘 오신 모든 분들을 환영합니다\n[새가족] : 소요리문답 제3문 암송은 다음 주일입니다\n[모임] : 오늘 오후"), ctx)
    assert d.kind == K.ANNOUNCEMENT_BLOCK  # 제3문 inside a line is not a catechism slide
    d = classify_slide(_slide(16, "소요리 제1문 사람의 제일 되는 목적은 무엇입니까?||사람의 제일 되는 목적은 하나님을 영화롭게 하는 것입니다"), ctx)
    assert d.kind == K.CATECHISM and d.fields["qa"].split("\n")[0].startswith("소요리 제1문")  # heading before 제N문 (older decks)
    s = _slide(15, "답: 사람의 제일 되는 목적은 하나님을 영화롭게 하는 것입니다||제1문 사람의 제일 되는 목적은 무엇입니까?"
                   "||[엡1:3] 찬송하리로다 하나님 곧 우리 주 예수 그리스도의 아버지께서||[롬8:1] 그러므로 이제 그리스도 예수 안에 있는 자에게는")
    s["shapes"][0]["bbox"], s["shapes"][1]["bbox"] = [0, 300, 100, 100], [0, 100, 100, 100]  # answer box first in the tree
    d = classify_slide(s, ctx)
    assert d.kind == K.CATECHISM and d.fields["qa"].startswith("제1문") and d.fields["qa"].split("\n")[1].startswith("답")
    assert [q["ref"].book for q in d.fields["verses"]] == ["에베소서", "로마서"]


def test_catechism_long_answer_keeps_styles_on_every_page(tmp_path):
    from pathlib import Path

    from pptx import Presentation
    from pptx.util import Inches

    from broadcast_ppt.cli import CATALOG, WIDTHS
    from broadcast_ppt.mapping.plan import PlanOptions, build_plan

    answer = "답: " + " ".join(["하나님의 작정은 그의 뜻의 계획을 따라 된 영원한 목적이니 이로써 그는 자기의 영광을 위하여 일어날 모든 일을 미리 정하셨습니다"] * 2)
    parts = ["제7문 하나님의 작정은 무엇입니까?", answer, "[엡1:11] 모든 일을 그의 뜻의 결정대로 일하시는 이의 계획을 따라", "[롬8:28] 하나님을 사랑하는 자 곧 그의 뜻대로 부르심을 입은 자들에게는"]
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    for i, part in enumerate(parts):
        s.shapes.add_textbox(Inches(0.5), Inches(0.5 + i * 1.6), Inches(9), Inches(1.2)).text_frame.text = part
    src = tmp_path / "src.pptx"
    prs.save(str(src))
    ops = build_plan(src, Path(CATALOG), Path(WIDTHS), PlanOptions())["ops"]
    caps = [o for o in ops if o["note"] == "caption"]
    assert len(caps) >= 2 and caps[0]["fill"]["body"].startswith("제7문") and caps[0]["para_style"] == [{"color": "FF9300"}, {"sz": 28.0}]
    assert all(o["para_style"] == [{"sz": 28.0}] for o in caps[1:]) and all(o["fit"]["est_lines"] <= o["fit"]["max_lines"] for o in caps)
    assert [o["fill"]["ref"] for o in ops if o["note"] == "scripture caption"] == ["에베소서1:11", "로마서8:28"]


def _shape_slide(index):
    """Two ellipses, each with a short label drawn inside it. The ellipses hold no text of their own."""
    def box(kind, text, bbox):
        return {"kind": kind, "text": text, "bbox": bbox, "depth": 0}
    shapes = [box("AUTO_SHAPE", "\n\n", [222, 219, 276, 89]),
              box("AUTO_SHAPE", "\n\n", [222, 108, 276, 86]),
              box("TEXT_BOX", "어떻게 되는가?", [243, 134, 233, 32]),
              box("TEXT_BOX", "어떻게 사는가?", [243, 247, 244, 32])]
    return {"index": index, "text": "\n".join(s["text"] for s in shapes), "shapes": shapes,
            "kinds": {"TEXT_BOX": 2, "AUTO_SHAPE": 2}, "full_bleed_picture": False, "keywords": [], "text_len": 30}


def test_labelled_shapes_are_laid_out_instead_of_captioned():
    d = classify_slide(_shape_slide(14), DeckContext())
    # this used to fall through to R550, which merged both labels into one line of caption text
    assert (d.kind, d.rule_id) == (K.SHAPES_LAYOUT, "R410") and "review:shapes_layout" in d.flags


def test_shape_layout_flags_text_no_shape_can_carry():
    s = _shape_slide(16)
    # a credit line in the corner: not inside any ellipse, so the layout cannot take it along
    s["shapes"].append({"kind": "TEXT_BOX", "text": "출처 가나다라마", "bbox": [443, 362, 307, 31], "depth": 0})
    s["text"] += "\n출처 가나다라마"
    # the real deck repeats that line on every slide, so it is dropped as boilerplate before the
    # rules see it — the slide still classifies, and the box is still on it
    from broadcast_ppt.rules.fields import compact
    ctx = DeckContext(boilerplate={compact("출처 가나다라마")})
    d = classify_slide(s, ctx)
    assert d.kind == K.SHAPES_LAYOUT and "review:text_dropped" in d.flags
    # without it nothing is left behind
    assert "review:text_dropped" not in classify_slide(_shape_slide(17), DeckContext()).flags


def test_free_floating_short_boxes_are_not_a_shape_layout():
    s = _shape_slide(15)
    s["shapes"][2]["bbox"] = [700, 10, 200, 30]  # a label outside every drawn shape
    assert classify_slide(s, DeckContext()).kind != K.SHAPES_LAYOUT


def test_validator_flags_a_caption_that_opens_mid_verse(tmp_path):
    from pathlib import Path

    from pptx import Presentation

    from broadcast_ppt.cli import CATALOG, WIDTHS
    from broadcast_ppt.validator.check import validate_deck

    deck = tmp_path / "out.pptx"
    prs = Presentation()
    for _ in range(3):
        prs.slides.add_slide(prs.slide_layouts[6])
    prs.save(str(deck))
    # the shape the packer used to produce: page 2 of 2 opening on a word instead of a verse number
    plan = {"template": {"sha256": ""}, "source": {"path": str(deck), "sha256": ""}, "rules_version": "test",
            "ops": [{"op": "copy", "out_index": 1, "chunk": {"i": 1, "n": 2}, "fill": {"body": "8 가나다 라마바"}},
                    {"op": "copy", "out_index": 2, "chunk": {"i": 2, "n": 2}, "fill": {"body": "다 마시며 11 말하기를"}},
                    {"op": "copy", "out_index": 3, "chunk": {"i": 1, "n": 1}, "fill": {"body": "절 번호로 시작하지 않는 첫 장"}}]}

    rep = validate_deck(deck, plan, Path(CATALOG), Path(WIDTHS))
    split = [f for f in rep["findings"] if f["check"] == "verse_split"]
    # only the second page: the first page of a passage may legitimately open on the reference tail
    assert [f["out_slide"] for f in split] == [2] and split[0]["level"] == "warning"


def test_each_announcement_item_is_one_slide():
    from pathlib import Path

    from broadcast_ppt.cli import CATALOG, WIDTHS
    from broadcast_ppt.mapping.catalog import load_catalog
    from broadcast_ppt.mapping.plan import Mapper, PlanOptions
    from broadcast_ppt.mapping.textfit import load_model
    from broadcast_ppt.rules.fields import announcement_items

    m = Mapper(load_catalog(Path(CATALOG)), load_model(Path(WIDTHS)), PlanOptions())
    items = announcement_items("환영 : 오신 분들을 환영합니다\n[행사] : 다음 주 주일 오후,\n함께 모여\n기도합시다!\n[수련회] : 모두 초청합시다!")
    ops = m.announcements(1, items)
    assert [o["fill"]["body"] for o in ops] == ["오신 분들을 환영합니다", "다음 주 주일 오후, 함께 모여 기도합시다!", "모두 초청합시다!"]


def test_prayer_card_with_a_subtitle_instead_of_a_person():
    d = classify_slide(_slide(3, "기 도  Prayer||함께 세우는 기도"), DeckContext())
    assert d.kind == K.PRAYER_TITLE and d.rule_id == "R101" and d.fields["subtitle"] == "함께 세우는 기도"
    # the name/role form still wins
    assert classify_slide(_slide(3, "기 도||Prayer||홍길동 장로"), DeckContext()).rule_id == "R100"
    # and it goes on air as the original page, not as a chroma card
    from pathlib import Path

    from broadcast_ppt.cli import CATALOG, WIDTHS
    from broadcast_ppt.mapping.catalog import load_catalog
    from broadcast_ppt.mapping.plan import Mapper, PlanOptions
    from broadcast_ppt.mapping.textfit import load_model
    m = Mapper(load_catalog(Path(CATALOG)), load_model(Path(WIDTHS)), PlanOptions())
    assert [o["op"] for o in m.map(3, d)] == ["copy"]
