from broadcast_ppt.mapping.split import pack
from broadcast_ppt.mapping.textfit import WidthModel
from broadcast_ppt.rules import kinds as K
from broadcast_ppt.rules.classify import DeckContext, classify_slide


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
