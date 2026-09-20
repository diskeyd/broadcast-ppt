"""Rule-based classification of original (live) slides. Ordered specific -> general; first match wins.

Input is one slide dict produced by `analyzer.core.analyze_file` plus a DeckContext that carries
deck-level state. Reasons never quote slide text (it contains names)."""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field

from ..analyzer.core import flatten
from . import kinds as K
from .fields import (
    _BRACKET_TITLE_RE, announcement_items, compact, english_line, name_role, norm, page_ref, quoted_title,
    scripture_ref,
)

VERSE_START_RE = re.compile(r"^\s*\d{1,3}\s+\S")
BRACKET_REF_RE = re.compile(r"^\s*[\[［]")
NAME_LIST_HEADER_RE = re.compile(r"[\[［][^\]］]*명단[\]］]")
# at a line start, optionally after a 소요리/대요리(문답) heading: 제3문 inside a notice line is not a catechism slide
CATECHISM_RE = re.compile(r"^\s*(?:[소대]요리(?:문답)?\s*[:：]?\s*)?제\s*\d{1,3}\s*문(?![가-힣])", re.M)


@dataclass
class Decision:
    kind: str
    rule_id: str
    reason: str
    fields: dict = field(default_factory=dict)
    flags: list[str] = field(default_factory=list)
    confidence: float = 1.0


@dataclass
class DeckContext:
    boilerplate: set[str] = field(default_factory=set)
    last_ref: object = None     # most recent reference (scripture title, body or bracketed quote)
    title_ref: object = None    # the service's 성경봉독 reference

    @classmethod
    def from_deck(cls, doc: dict) -> DeckContext:
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
    for sh in flatten(slide.get("shapes", [])):
        t = sh.get("text")
        if not t:
            continue
        kept = "\n".join(line for line in norm(t).splitlines() if compact(line) and compact(line) not in ctx.boilerplate)
        if compact(kept):
            boxes.append({**sh, "text": kept})
    text = "\n".join(b["text"] for b in boxes)
    return SlideView(slide["index"], text, boxes, Counter(slide.get("kinds", {})), slide.get("full_bleed_picture", False), slide)


def _reading_order(b: dict) -> list:
    """Sort key for boxes as a person reads them: top, then left."""
    return (b.get("bbox") or [0, 0])[1::-1]


def _close_bracket(t: str) -> int:
    return min([i for i in (t.find("]"), t.find("］")) if i >= 0] or [-1])


def _unit(text: str) -> dict | None:
    """One text box read as a scripture unit, or None when it is neither a reference nor verse-numbered.

    A body may be empty: a slide sometimes prints the reference in a box of its own. Leading bracketed
    headings ([제 삼 권, 아삽의 시]) are peeled one by one until a bracket holds a reference, so the
    reference is read from that bracket rather than from a fixed-length slice of the box."""
    t = norm(text).lstrip()
    while t.startswith(("[", "［")):
        close = _close_bracket(t)
        if close < 0:
            break
        ref = scripture_ref(t[1:close])
        if ref is not None:
            return {"ref": ref, "body": t[close + 1:].lstrip(), "quote": True, "verse": None}
        t = t[close + 1:].lstrip()
    line = t.split("\n", 1)[0]
    ref = scripture_ref(line)
    if ref is not None and line.find(ref.raw) <= 2:
        i = t.find(ref.raw)
        return {"ref": ref, "body": t[i + len(ref.raw):].lstrip(" \n"), "quote": False, "verse": None}
    if VERSE_START_RE.match(t):
        return {"ref": None, "body": t.strip(), "quote": False, "verse": int(re.match(r"\s*(\d{1,3})", t).group(1))}
    return None


def _labels_inside_shapes(v: SlideView) -> bool:
    """Every text box sits inside a drawn shape — a few labelled shapes, not free-floating text.

    This is what separates the slide from an announcement card or a meditation quote, which also
    carry two short boxes but draw nothing around them."""
    def inside(a, b) -> bool:
        return b[0] <= a[0] and b[1] <= a[1] and a[0] + a[2] <= b[0] + b[2] and a[1] + a[3] <= b[1] + b[3]

    drawn = [sh for sh in flatten(v.raw.get("shapes", [])) if sh.get("kind") in ("AUTO_SHAPE", "FREEFORM")]
    return bool(drawn) and all(b.get("bbox") and any(sh.get("bbox") and inside(b["bbox"], sh["bbox"]) for sh in drawn)
                               for b in v.boxes)


def _bracketed_verse(u: dict | None) -> bool:
    """A unit the catechism rule counts as a verse: a bracketed reference with text under it."""
    return bool(u and u["quote"] and u["ref"] and compact(u["body"]))


def scripture_units(v: SlideView) -> tuple[list[dict], list[str]]:
    """Every text box in reading order: the scripture units, and the text of the boxes that are not.

    Nothing is dropped here. A box that is not a unit comes back as a string so the caller can put it
    in the reference tag or flag it — silently discarding boxes is what hid missing verses for months.
    Boxes repeating the same text (drop-shadow copies) are kept once."""
    units: list[dict] = []
    other: list[str] = []
    seen: set[str] = set()
    for b in sorted(v.boxes, key=_reading_order):
        key = compact(b["text"])
        if not key or key in seen:
            continue
        seen.add(key)
        u = _unit(b["text"])
        if u is None:
            other.append(norm(b["text"]).strip())
        else:
            units.append(u)
    return units, other


def _resolve_refs(units: list[dict], ctx: DeckContext) -> list[dict]:
    """Give every unit a reference, in reading order. `origin` is where it came from.

    A unit without a printed reference inherits one only when the range contains its first verse
    number: a blank tag is a visible gap, a wrong tag goes on air. The nearest reference already
    resolved on this slide comes first, so a passage continuing in a second box keeps its own; any
    other reference printed on the slide comes next, because some slides print it under the text."""
    printed = [u["ref"] for u in units if u["ref"] is not None]
    out, here = [], None
    for u in units:
        ref, origin = u["ref"], "slide"
        if ref is None:
            for cand, which in [(here, "slide")] + [(p, "slide") for p in printed] + \
                               [(ctx.last_ref, "last"), (ctx.title_ref, "title")]:
                if cand is not None and u["verse"] is not None and cand.contains(u["verse"]):
                    ref, origin = cand, which
                    break
        if ref is None:
            origin = None
        else:
            here = ref
        out.append({**u, "ref": ref, "origin": origin})
    return out


def _scripture_decision(body_units: list[dict], other: list[str]) -> Decision:
    """Rule id and flags from the first unit; the rest ride along as `extra`, as they always have."""
    first = body_units[0]
    heading = next((t for t in other if page_ref(t) is None and compact(t)), None)
    # a bracketed title box is a book division ([제 삼 권, 아삽의 시]); an unbracketed leftover is
    # usually a reference this module failed to parse. Both can fill an empty tag, but only the
    # division changes how the verses are laid out, so the caller needs to tell them apart.
    division = False
    if heading is not None:
        m = _BRACKET_TITLE_RE.match(heading)
        division = m is not None
        heading = m.group("t").strip() if m else heading
    flags, fields = [], {"ref": first["ref"], "body": first["body"], "quote": first["quote"],
                         "extra": [{"ref": u["ref"], "body": u["body"]} for u in body_units[1:]],
                         "heading": heading, "division": division}
    if any(u["origin"] in ("last", "title") for u in body_units):
        flags.append("review:inherited_reference")
    if any(u["ref"] is None for u in body_units):
        flags.append("review:reference_unknown")
    if len(other) > 1 or (other and heading is None):
        flags.append("review:text_dropped")  # a box we could use for neither caption nor tag
    if first["origin"] == "slide":
        rule = "R205" if first["quote"] else "R200"
        reason = "reference at start" + (" (bracketed quote)" if first["quote"] else "")
        return Decision(K.SCRIPTURE_BODY, rule, reason, fields, sorted(flags))
    if first["ref"] is not None:
        fields["inherited_ref"] = True
        return Decision(K.SCRIPTURE_BODY, "R210", f"verse number continues the {first['origin']} reference",
                        fields, sorted(flags), confidence=0.8)
    return Decision(K.SCRIPTURE_BODY, "R211", "verse-numbered text outside every known reference",
                    fields, sorted(flags), confidence=0.6)


def classify_slide(slide: dict, ctx: DeckContext) -> Decision:
    v = view_of(slide, ctx)
    c = compact(v.text)
    pics, groups, media = v.kinds.get("PICTURE", 0), v.kinds.get("GROUP", 0), v.kinds.get("MEDIA", 0)
    d = _classify(v, c, pics, groups, media, ctx)
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
            title = quoted_title(v.text)
            return Decision(K.SERMON_TITLE, "R130", "keyword 설교", {"title": title, "ref": scripture_ref(v.text)},
                            [] if title else ["review:no_sermon_title"])
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

    # catechism (제N문 at a line start): question + answer on caption pages, each bracketed verse on its own.
    # After announcements (a 소식 line may mention 제3문), before scripture (the verse box is usually the longest)
    if CATECHISM_RE.search(v.text):
        quotes = [(b, _unit(b["text"])) for b in sorted(v.boxes, key=_reading_order)]
        # a verse box here is a bracketed reference with text under it, as it was before _unit existed
        verses = [{"ref": u["ref"], "body": u["body"]} for _, u in quotes if u and _bracketed_verse(u)]
        lines = [ln.strip() for b, u in quotes if not _bracketed_verse(u)
                 for ln in norm(b["text"]).splitlines() if compact(ln)]
        # question on its own line, the answer below: a 2-line tail left a blank line on the 3-line bar
        ans = next((i for i, ln in enumerate(lines) if re.match(r"(답(?![가-힣])|A\s*[.:])", ln)), 1)
        qa = "\n".join(p for p in (" ".join(lines[:ans]), " ".join(lines[ans:])) if p)
        return Decision(K.CATECHISM, "R170", "제N문 question" + (f" + {len(verses)} bracketed verse(s)" if verses else ""),
                        {"qa": qa, "verses": verses})

    # scripture: every box that starts with a reference or a verse number is a unit, in reading order.
    # Anchored to the main box so this claims no slide the older, narrower test would have left alone.
    units, other = scripture_units(v)
    main_u = _unit(main)
    if v.clen >= 30 and main_u is not None and (
            any(u["ref"] for u in units) or ctx.last_ref is not None or ctx.title_ref is not None):
        resolved = _resolve_refs(units, ctx)
        body_units = [u for u in resolved if compact(u["body"])]
        if body_units:
            return _scripture_decision(body_units, other)

    # no hymn rule here: a score page carries no text, so it is picture_only (R600) and the run of them
    # is recognised in mapping/plan.py, where the neighbouring slides can be seen
    drawn = v.kinds.get("AUTO_SHAPE", 0) + v.kinds.get("FREEFORM", 0) + v.kinds.get("CHART", 0) + v.kinds.get("LINE", 0)
    # before R420, which would pass a 3-shape slide straight through: a few drawn shapes each holding
    # one short label is a layout to rearrange, not a diagram to copy
    if 2 <= drawn <= 4 and pics == 0 and not v.full_bleed_picture and 2 <= len(v.boxes) <= 4 \
            and all(len(compact(b["text"])) <= 20 for b in v.boxes) and _labels_inside_shapes(v):
        return Decision(K.SHAPES_LAYOUT, "R410", f"{drawn} drawn shapes, each holding a short label",
                        flags=["review:shapes_layout"], confidence=0.8)
    if drawn >= 3 and len(v.boxes) >= 3:
        return Decision(K.GRAPHIC, "R420", "diagram or question sheet built from many shapes", flags=["passthrough", "review:graphic"])

    if len(v.boxes) <= 2 and 6 <= v.clen <= 60 and not v.full_bleed_picture:
        return Decision(K.CAPTION_SENTENCE, "R550", "one short sentence", {"body": re.sub(r"\s*\n\s*", " ", v.text).strip()})
    if len(v.boxes) <= 2 and 6 <= v.clen <= 60:
        return Decision(K.CAPTION_SENTENCE, "R551", "one short sentence over a picture",
                        {"body": re.sub(r"\s*\n\s*", " ", v.text).strip()}, ["review:sentence_over_picture"], 0.8)
    return Decision(K.UNKNOWN, "R900", "no rule matched", flags=["passthrough", "review:unclassified"], confidence=0.0)


def classify_deck(doc: dict, fallback=None) -> list[Decision]:
    """`fallback(view, decision)` is asked about slides no rule matched; None keeps the rule decision."""
    ctx = DeckContext.from_deck(doc)
    out = []
    for s in doc.get("slides", []):
        if "error" in s:
            continue
        d = classify_slide(s, ctx)
        if d.kind == K.UNKNOWN and fallback:
            d = fallback(view_of(s, ctx), d) or d
        out.append(d)
    return out
