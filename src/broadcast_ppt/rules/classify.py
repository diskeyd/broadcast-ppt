"""Rule-based classification of original (live) slides. Ordered specific -> general; first match wins.

Input is one slide dict produced by `analyzer.core.analyze_file` plus a DeckContext that carries
deck-level state. Reasons never quote slide text (it contains names)."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field

from . import kinds as K
from .fields import (
    announcement_items, compact, english_line, name_role, norm, page_ref, quoted_title, scripture_ref,
)
from .version import RULES_VERSION

VERSE_START_RE = re.compile(r"^\s*\d{1,3}\s+\S")
BRACKET_REF_RE = re.compile(r"^\s*[\[［]")
NAME_LIST_HEADER_RE = re.compile(r"[\[［][^\]］]*명단[\]］]")


@dataclass
class Decision:
    kind: str
    rule_id: str
    reason: str
    fields: dict = field(default_factory=dict)
    flags: list[str] = field(default_factory=list)
    confidence: float = 1.0


def _flat(shapes):
    for s in shapes:
        yield s
        yield from _flat(s.get("children", []))


@dataclass
class DeckContext:
    boilerplate: set[str] = field(default_factory=set)
    last_ref: object = None     # most recent reference (scripture title, body or bracketed quote)
    title_ref: object = None    # the service's 성경봉독 reference
    last_kind: str | None = None
    rules_version: str = RULES_VERSION

    @classmethod
    def from_deck(cls, doc: dict) -> "DeckContext":
        """Lines repeated on many slides (church slogan footer etc.) are decoration, not content."""
        slides = [s for s in doc.get("slides", []) if "error" not in s]
        counter: Counter = Counter()
        for s in slides:
            seen = {compact(line) for line in norm(s.get("text", "")).splitlines() if compact(line)}
            counter.update(seen)
        threshold = max(4, int(len(slides) * 0.25))
        return cls(boilerplate={k for k, n in counter.items() if n >= threshold and len(k) >= 6})


@dataclass
class SlideView:
    index: int
    text: str                      # content text with boilerplate lines removed
    boxes: list[dict]              # text boxes (flat, with bbox/text/font_pt) minus boilerplate-only boxes
    kinds: Counter
    full_bleed_picture: bool
    raw: dict

    @property
    def clen(self) -> int:
        return len(compact(self.text))

    @property
    def main_box(self) -> dict | None:
        return max(self.boxes, key=lambda b: len(compact(b["text"])), default=None)


def view_of(slide: dict, ctx: DeckContext) -> SlideView:
    boxes = []
    for sh in _flat(slide.get("shapes", [])):
        t = sh.get("text")
        if not t:
            continue
        kept = "\n".join(line for line in norm(t).splitlines() if compact(line) and compact(line) not in ctx.boilerplate)
        if compact(kept):
            boxes.append({**sh, "text": kept})
    text = "\n".join(b["text"] for b in boxes)
    return SlideView(slide["index"], text, boxes, Counter(slide.get("kinds", {})), slide.get("full_bleed_picture", False), slide)


def _card_label(text: str, word: str) -> str:
    for line in norm(text).splitlines():
        if word in compact(line):
            return re.sub(r"\s+", " ", line).strip()
    return word


def classify_slide(slide: dict, ctx: DeckContext) -> Decision:
    v = view_of(slide, ctx)
    c = compact(v.text)
    pics, groups, media = v.kinds.get("PICTURE", 0), v.kinds.get("GROUP", 0), v.kinds.get("MEDIA", 0)
    d = _classify(v, c, pics, groups, media, ctx)
    ctx.last_kind = d.kind
    if d.kind == K.SCRIPTURE_TITLE and d.fields.get("ref"):
        ctx.last_ref = d.fields["ref"]
        ctx.title_ref = d.fields["ref"]
    if d.kind == K.SCRIPTURE_BODY and d.fields.get("ref") is not None:
        ctx.last_ref = d.fields["ref"]
    return d


def _classify(v: SlideView, c: str, pics: int, groups: int, media: int, ctx: DeckContext) -> Decision:
    if media and "기도로준비" in c:
        return Decision(K.INTRO_MEDIA, "R010", "video + '기도로 준비' notice")
    if v.clen == 0:
        if groups >= 1 or pics >= 1 or media:
            return Decision(K.PICTURE_ONLY, "R600", "no text; pictures/video only", flags=["passthrough"])
        return Decision(K.PICTURE_ONLY, "R601", "empty slide", flags=["passthrough", "review:empty_slide"])

    if v.clen <= 4 and len(v.boxes) <= 3:
        return Decision(K.PICTURE_ONLY, "R602", "decoration only (a few characters)", flags=["passthrough"])

    # 동행의 기도 / 결단의 기도: title plus an optional prayer sentence on the same slide
    if v.clen <= 180:
        for word, label in (("결단의기도", "결단의 기도"), ("동행의기도", "동행의 기도")):
            if word in c:
                body = max((line for line in norm(v.text).splitlines() if word not in compact(line) and not english_line(line)),
                           key=lambda s: len(compact(s)), default="")
                return Decision(K.CARD_TITLE, "R105", f"keyword {label}", {"label": label, "prayer_text": body.strip() if len(compact(body)) >= 8 else None})

    # order-of-worship cards (short text)
    if v.clen <= 60:
        if "봉헌기도" in c and "축도" in c:
            return Decision(K.CARD_TITLE, "R140", "keywords 봉헌기도+축도", {"label": "봉헌기도 및 축도", "card": "offertory"})
        if "예배로의부름" in c or "예배의부름" in c:
            return Decision(K.CARD_TITLE, "R110", "keyword 예배로의 부름", {"label": "예배로의 부름", "card": "call_to_worship"})
        if "성경봉독" in c:
            ref = scripture_ref(v.text)
            page = page_ref(v.text)
            flags = [] if ref else ["review:no_reference"]
            return Decision(K.SCRIPTURE_TITLE, "R120", "keyword 성경봉독" + (" + reference" if ref else ""),
                            {"ref": ref, "page": page}, flags)
        if "설교" in c and "Sermon" in v.text or re.search(r"^설\s*교", norm(v.text), re.M):
            return Decision(K.SERMON_TITLE, "R130", "keyword 설교", {"title": quoted_title(v.text), "ref": scripture_ref(v.text)},
                            [] if quoted_title(v.text) else ["review:no_sermon_title"])
        if ("교회소식" in c or "announcement" in c.lower()) and v.clen <= 40:
            return Decision(K.ANNOUNCEMENT_TITLE, "R150", "keyword 교회소식", {"label": "교회 소식"})
        if re.search(r"(^|\s)기\s*도(\s|$)", norm(v.text), re.M) or c.startswith("기도"):
            people = name_role(v.text)
            if people:
                p = people[0]
                return Decision(K.PRAYER_TITLE, "R100", "keyword 기도 + name/role", {"label": "기도", "name": p["name"], "role": p["role"]})
        if "묵상" in c:
            quote = quoted_title(v.text)
            if not quote:  # quote written without quotation marks: the longest line that is not the label
                lines = [ln.strip() for ln in norm(v.text).splitlines()
                         if ln.strip() and "묵상" not in compact(ln) and not english_line(ln) and "meditat" not in ln.lower()]
                quote = max(lines, key=lambda s: len(compact(s)), default=None)
                if quote and len(compact(quote)) < 4:
                    quote = None
            if not quote:
                return Decision(K.GRAPHIC, "R161", "keyword 묵상 without a quote line", flags=["passthrough", "review:no_quote"])
            return Decision(K.MEDITATION, "R160", "keyword 묵상", {"quote": quote})

    main = v.main_box["text"] if v.main_box else v.text
    if NAME_LIST_HEADER_RE.search(v.text):
        return Decision(K.NAME_LIST, "R500", "bracket header with 명단", flags=["passthrough"])
    items = announcement_items(main)
    if len(items) >= 2 and ("환영" in c or "교회소식" in c or sum(1 for i in items if i["bracketed"]) >= 2):
        return Decision(K.ANNOUNCEMENT_BLOCK, "R300", f"{len(items)} titled announcement lines", {"items": items},
                        [i["flag"] for i in items if i["flag"]])

    # scripture: reference at the start, or verse-numbered continuation of the last reference
    head = norm(main).lstrip()[:40]
    ref = scripture_ref(head)
    if v.clen >= 30 and ref is not None and (head.startswith(("[", "［")) or norm(main).lstrip().find(ref.raw) <= 2):
        body = norm(main).lstrip()
        if body.startswith(("[", "［")):
            close = min([i for i in (body.find("]"), body.find("］")) if i >= 0] or [len(ref.raw)])
            body = body[close + 1:].lstrip()
        else:
            body = body[body.find(ref.raw) + len(ref.raw):].lstrip(" \n")
        return Decision(K.SCRIPTURE_BODY, "R200" if not BRACKET_REF_RE.match(main) else "R205",
                        "reference at start" + (" (bracketed quote)" if BRACKET_REF_RE.match(main) else ""),
                        {"ref": ref, "body": body, "quote": bool(BRACKET_REF_RE.match(main))})
    if v.clen >= 30 and VERSE_START_RE.match(norm(main)) and (ctx.last_ref is not None or ctx.title_ref is not None):
        first = int(re.match(r"\s*(\d{1,3})", norm(main)).group(1))
        # inherit only a reference whose verse range contains the first verse number:
        # a blank tag is a visible gap, a wrong tag goes on air
        for cand, which in ((ctx.last_ref, "last"), (ctx.title_ref, "title")):
            if cand is not None and cand.contains(first):
                return Decision(K.SCRIPTURE_BODY, "R210", f"verse number continues the {which} reference",
                                {"ref": cand, "body": norm(main).strip(), "quote": False, "inherited_ref": True},
                                ["review:inherited_reference"], confidence=0.8)
        return Decision(K.SCRIPTURE_BODY, "R211", "verse-numbered text outside every known reference",
                        {"ref": None, "body": norm(main).strip(), "quote": False},
                        ["review:reference_unknown"], confidence=0.6)

    # hymn before graphic: score pages also carry several drawn shapes
    if groups >= 2 and pics >= 1 and v.clen <= 120:
        return Decision(K.HYMN_IMAGE, "R400", "picture + grouped lyric boxes", flags=["passthrough", "deferred:hymn"])
    drawn = v.kinds.get("AUTO_SHAPE", 0) + v.kinds.get("FREEFORM", 0) + v.kinds.get("CHART", 0) + v.kinds.get("LINE", 0)
    if drawn >= 3 and len(v.boxes) >= 3:
        return Decision(K.GRAPHIC, "R420", "diagram or question sheet built from many shapes", flags=["passthrough", "review:graphic"])

    if len(v.boxes) <= 2 and 6 <= v.clen <= 60 and not v.full_bleed_picture:
        return Decision(K.CAPTION_SENTENCE, "R550", "one short sentence", {"body": re.sub(r"\s*\n\s*", " ", v.text).strip()})
    if len(v.boxes) <= 2 and 6 <= v.clen <= 60:
        return Decision(K.CAPTION_SENTENCE, "R551", "one short sentence over a picture",
                        {"body": re.sub(r"\s*\n\s*", " ", v.text).strip()}, ["review:sentence_over_picture"], 0.8)
    return Decision(K.UNKNOWN, "R900", "no rule matched", flags=["passthrough", "review:unclassified"], confidence=0.0)


def classify_deck(doc: dict) -> list[Decision]:
    ctx = DeckContext.from_deck(doc)
    return [classify_slide(s, ctx) for s in doc.get("slides", []) if "error" not in s]
