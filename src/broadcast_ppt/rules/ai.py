"""Gemini fallback for slides no rule matched. Runs only when GEMINI_API_KEY is set.

The model picks a kind and which lines carry the text; Python rebuilds the fields from the
original slide text. Nothing the model writes is ever placed on a slide."""

from __future__ import annotations

import json
import os
import sys

from . import kinds as K
from .classify import Decision, SlideView
from .fields import compact, mask_names, norm

MODEL = os.environ.get("GEMINI_MODEL", "gemini-3.5-flash")  # ponytail: id from env; the current flash model as of 2026-05
MIN_CONFIDENCE = 0.8  # ponytail: single gate; the planned 4-step gate if weekly data asks for it
# no prayer/scripture/sermon (names, references), no announcement_block (structured items), no hymn (deferred:hymn flag)
ALLOWED = [K.CAPTION_SENTENCE, K.MEDITATION, K.CARD_TITLE, K.ANNOUNCEMENT_TITLE, K.GRAPHIC, K.PICTURE_ONLY, K.NAME_LIST, K.UNKNOWN]
SCHEMA = {
    "type": "object",
    "properties": {
        "kind": {"type": "string", "enum": ALLOWED},
        "confidence": {"type": "number", "minimum": 0, "maximum": 1},
        "lines": {"type": "array", "items": {"type": "integer"}, "description": "1-based numbers of the lines that carry the text"},
        "reason": {"type": "string", "description": "one short English sentence, do not quote the slide"},
    },
    "required": ["kind", "confidence", "lines", "reason"],
}
PROMPT = """You classify one slide from a Korean church worship PowerPoint. The rule engine could not.
Pick exactly one kind:
- caption_sentence: one or two short sermon sentences that go into a lower-third caption
- meditation: a 묵상 slide, one quoted sentence to meditate on
- card_title: an order-of-worship card (예배로의 부름, 동행의 기도, 결단의 기도, 찬양 및 봉헌 ...): a label, nothing else
- announcement_title: a 교회소식 / announcements heading
- graphic: diagram, outline or discussion questions; too many separate pieces for a caption
- picture_only: decoration, no meaningful text
- name_list: a roster of people
- unknown: none of the above
`lines` = the 1-based line numbers whose text belongs to the caption/quote/label. Never rewrite the text.
Names are masked as [이름]. Reply as JSON.

Shapes: {kinds}, text boxes: {boxes}, full-screen picture: {picture}
Lines:
{lines}"""


def enabled() -> str | None:
    return MODEL if os.environ.get("GEMINI_API_KEY") else None


def ask(payload: dict) -> dict:
    """The one network call; tests replace it."""
    from google import genai  # key read from GEMINI_API_KEY
    lines = "\n".join(f"{i}: {t}" for i, t in enumerate(payload["lines"], 1))
    r = genai.Client().models.generate_content(
        model=MODEL, contents=PROMPT.format(kinds=payload["kinds"], boxes=payload["boxes"], picture=payload["picture"], lines=lines),
        config={"response_mime_type": "application/json", "response_json_schema": SCHEMA})
    return r.parsed if isinstance(r.parsed, dict) else json.loads(r.text or "")


def classify_view(v: SlideView, d: Decision) -> Decision | None:
    """Returns a Decision for a confident answer; otherwise None, leaving the suggestion in d.fields["ai"]."""
    lines = [ln.strip() for ln in norm(v.text).splitlines() if compact(ln)]
    try:
        res = ask({"lines": [mask_names(ln) for ln in lines], "kinds": dict(v.kinds), "boxes": len(v.boxes), "picture": v.full_bleed_picture})
        kind, conf, sel = res["kind"], float(res["confidence"]), [lines[i - 1] for i in res["lines"] if 1 <= i <= len(lines)]
        if kind not in ALLOWED:
            raise ValueError(f"kind {kind!r}")
    except Exception as exc:  # noqa: BLE001 - network, schema, model id: the slide simply stays unclassified
        print(f"ai fallback failed on slide {v.index}: {exc!r}", file=sys.stderr)
        return None
    d.fields["ai"] = {"kind": kind, "confidence": conf, "reason": res.get("reason", "")}
    if conf < MIN_CONFIDENCE or kind == K.UNKNOWN:
        return None
    text = " ".join(sel)
    fields = {K.CAPTION_SENTENCE: {"body": text}, K.MEDITATION: {"quote": text}, K.CARD_TITLE: {"label": sel[0] if sel else ""}}.get(kind, {})
    if any(not val for val in fields.values()):
        return None
    flags = ["review:ai_classified"] + (["passthrough"] if kind in K.PASSTHROUGH else [])
    return Decision(kind, "AI", "gemini fallback", fields, flags, conf)
