"""Build a reviewable conversion plan: original slide -> decision -> output operations."""

from __future__ import annotations

import hashlib
import json
import re
from dataclasses import asdict, dataclass
from datetime import datetime, timezone
from pathlib import Path

from ..analyzer.core import analyze_file
from ..rules import kinds as K
from ..rules.classify import classify_deck
from ..rules.fields import ScriptureRef, verse_chunks
from ..rules.version import RULES_VERSION
from .catalog import families, load_catalog
from .split import pack
from .textfit import WidthModel, load_model

TEMPLATE = {"call_to_worship": 3, "prayer": 4, "scripture": 6, "sermon": 7, "meditation": 8, "generic": 17,
            "announcement_title": 20, "announcement_first": 21, "announcement_item": 23, "offertory": 31}


@dataclass
class PlanOptions:
    slide_size: str = "template"         # template | source
    scripture_both: bool = True          # full-screen copy + lower-third captions
    sermon_material: str = "lower-third"  # lower-third | passthrough


def sha256(path: Path) -> str:
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


class Mapper:
    def __init__(self, catalog: dict, model: WidthModel, opts: PlanOptions):
        self.cat = catalog
        self.model = model
        self.opts = opts
        self.fam = families(catalog)

    # --- fitting helpers -------------------------------------------------------------------------
    def _body_spec(self, tpl: int, slot: str = "body") -> dict:
        for spec in self.cat["slides"][tpl]["shapes"].values():
            if spec.get("slot_map") and slot in spec["slot_map"]:
                return spec
        raise KeyError(f"template {tpl} has no slot {slot}")

    def _sz(self, spec: dict, slot: str) -> float:
        flat = [s for p in spec["paragraphs"] for s in p["segments"] if s["sample_len"] or s["sz"]]
        idx = spec["slot_map"].index(slot) if not spec.get("single_segment") else 0
        seg = flat[min(idx, len(flat) - 1)]
        return seg["sz"] or max(s["sz"] or 0 for s in flat)

    def lines(self, tpl: int, text: str, slot: str = "body") -> int:
        spec = self._body_spec(tpl, slot)
        bp = spec["bodyPr"]
        if bp["wrap"] == "none":
            return 1 if self.model.width_emu(text, self._sz(spec, slot)) <= spec["bbox"][2] * 1.0 else 2
        return self.model.estimate_lines(text, spec["bbox"][2], self._sz(spec, slot), bp["lIns"], bp["rIns"])

    def fits(self, tpl: int, text: str, slot: str = "body", max_lines: int | None = None) -> bool:
        spec = self._body_spec(tpl, slot)
        return self.lines(tpl, text, slot) <= (max_lines or spec.get("max_lines") or 1)

    def smallest_variant(self, family: str, text: str) -> int | None:
        for idx, entry in self.fam[family]:
            if self.fits(idx, text):
                return idx
        return None

    # --- per-kind mapping ------------------------------------------------------------------------
    def copy(self, src: int, note: str) -> dict:
        return {"op": "copy", "src": src, "note": note}

    def card(self, src: int, tpl: int, fill: dict, note: str) -> dict:
        return {"op": "template", "src": src, "template_slide": tpl, "fill": {k: v for k, v in fill.items() if v is not None}, "note": note}

    def scripture_captions(self, src: int, ref: ScriptureRef | None, body: str) -> list[dict]:
        units = [t for _, t in verse_chunks(body, ref)] if ref else [body]
        big = self.fam["scripture_ref"][-1][0]
        chunks = pack(units, lambda t: self.fits(big, t))
        ops = []
        for i, chunk in enumerate(chunks, 1):
            tpl = self.smallest_variant("scripture_ref", chunk) or big
            verses = [int(n) for n in re.findall(r"(?:^|\s)(\d{1,3})\s(?=\S)", chunk)]
            if ref is not None:
                in_range = [v for v in verses if ref.verse_start <= v <= ref.verse_end] or [ref.verse_start]
                v1, v2 = min(in_range), max(in_range)
                tag = f"{ref.book}{ref.chapter}:{v1}" + (f"-{v2}" if v2 != v1 else "")
            else:
                tag = None
            op = self.card(src, tpl, {"ref": tag, "body": chunk}, "scripture caption")
            op["chunk"] = {"i": i, "n": len(chunks)}
            op["fit"] = {"chars": len(chunk), "est_lines": self.lines(tpl, chunk), "max_lines": self._body_spec(tpl)["max_lines"]}
            if tag is None:
                op.setdefault("flags", []).append("review:no_reference_tag")
            ops.append(op)
        return ops

    def caption(self, src: int, text: str) -> list[dict]:
        big = self.fam["caption"][-1][0]
        ops = []
        for chunk in pack([text], lambda t: self.fits(big, t)):
            tpl = self.smallest_variant("caption", chunk) or big
            op = self.card(src, tpl, {"body": chunk}, "caption")
            op["fit"] = {"chars": len(chunk), "est_lines": self.lines(tpl, chunk), "max_lines": self._body_spec(tpl)["max_lines"]}
            ops.append(op)
        return ops

    def announcements(self, src: int, items: list[dict]) -> list[dict]:
        ops = []
        for it in items:
            if it["no"] == 1 and it["title"].replace(" ", "") == "환영":
                label = f"교회 소식 – {it['no']}. {it['title']}"
                tpl = TEMPLATE["announcement_first"]
                for chunk in pack([it["body"]], lambda t: self.fits(tpl, t)):
                    ops.append(self.card(src, tpl, {"label": label, "body": chunk}, "announcement"))
                continue
            title = f"[{it['title']}] : " if it["title"] else ""
            label = f"교회 소식 – {it['no']}. {title}"
            tpl = TEMPLATE["announcement_item"]
            spec = self._body_spec(tpl, "label")
            inner = spec["bbox"][2] - spec["bodyPr"]["lIns"] - spec["bodyPr"]["rIns"]
            label_w = self.model.width_emu(label, self._sz(spec, "label"))
            fits = lambda t: label_w + self.model.width_emu(t, self._sz(spec, "body")) <= inner
            for chunk in pack([it["body"]], fits):
                op = self.card(src, tpl, {"label": label, "body": chunk}, "announcement")
                if it.get("flag"):
                    op.setdefault("flags", []).append(it["flag"])
                ops.append(op)
        return ops

    def map(self, src: int, d) -> list[dict]:
        f = d.fields
        if d.kind in K.PASSTHROUGH:
            return [self.copy(src, d.kind)]
        if d.kind == K.CARD_TITLE:
            tpl = TEMPLATE.get(f.get("card"), TEMPLATE["generic"])
            ops = [self.card(src, tpl, {"label": f.get("label")}, "card")]
            if f.get("prayer_text"):
                ops.append(self.copy(src, "prayer text full screen"))
            return ops
        if d.kind == K.PRAYER_TITLE:
            return [self.card(src, TEMPLATE["prayer"], {"label": "기도 ", "person": f"– {f['name']} {f['role']}"}, "prayer card")]
        if d.kind == K.SCRIPTURE_TITLE:
            ref, page = f.get("ref"), f.get("page")
            testament = (page or {}).get("testament") or (ref.testament if ref else None)
            page_text = f"{'구약' if testament == 'old' else '신약'}성경 p.{page['page']}" if page else ""
            op = self.card(src, TEMPLATE["scripture"], {"ref": ref.long if ref else None, "page": page_text}, "scripture card")
            if not page:
                op["flags"] = ["review:no_page_number"]  # blank rather than the template's sample page
            return [op]
        if d.kind == K.SERMON_TITLE:
            title = f"– “{f['title']}”" if f.get("title") else None
            ref = f.get("ref")
            return [self.card(src, TEMPLATE["sermon"], {"label": "설교 ", "title": title, "ref": ref.long if ref else None}, "sermon card")]
        if d.kind == K.MEDITATION:
            return [self.card(src, TEMPLATE["meditation"], {"quote": f"“{f['quote']}”" if f.get("quote") else None}, "meditation")]
        if d.kind == K.ANNOUNCEMENT_TITLE:
            return [self.card(src, TEMPLATE["announcement_title"], {}, "announcement card")]
        if d.kind == K.ANNOUNCEMENT_BLOCK:
            return self.announcements(src, f["items"])
        if d.kind == K.SCRIPTURE_BODY:
            ops = [self.copy(src, "scripture full screen")] if self.opts.scripture_both else []
            return ops + self.scripture_captions(src, f.get("ref"), f.get("body", ""))
        if d.kind == K.CAPTION_SENTENCE:
            if self.opts.sermon_material == "passthrough":
                return [self.copy(src, "sermon material")]
            ops = [self.copy(src, "sentence full screen")] if self.opts.scripture_both else []
            return ops + self.caption(src, f["body"])
        return [self.copy(src, f"unmapped kind {d.kind}")]


def _jsonable(o):
    if isinstance(o, ScriptureRef):
        return {"book": o.book, "chapter": o.chapter, "verses": [o.verse_start, o.verse_end], "testament": o.testament}
    if isinstance(o, (set, tuple)):
        return list(o)
    raise TypeError(type(o))


def build_plan(src: Path, catalog_path: Path, widths_path: Path, opts: PlanOptions, overrides: dict | None = None) -> dict:
    src = Path(src)
    catalog = load_catalog(catalog_path)
    doc = analyze_file(src, src.parent)
    decisions = classify_deck(doc)
    mapper = Mapper(catalog, load_model(widths_path), opts)
    tw, th = catalog["source"]["slide_size"]
    sw, sh = doc["slide_size_emu"]
    scale = tw / sw if opts.slide_size == "template" else 1.0
    slide_over = ((overrides or {}).get("slides") or {})
    rows, ops = [], []
    slides = [s for s in doc["slides"] if "error" not in s]
    for s, d in zip(slides, decisions):
        o = slide_over.get(s["index"]) or {}
        if o.get("skip"):
            mapped = []
        elif o.get("op") == "copy":
            mapped = [mapper.copy(s["index"], "override: copy")]
        else:
            mapped = mapper.map(s["index"], d)
        for m in mapped:
            m["out_index"] = len(ops) + 1
            m["kind"] = d.kind
            m["rule"] = d.rule_id
            m["flags"] = sorted(set(m.get("flags", []) + [x for x in d.flags if x]))
            ops.append(m)
        rows.append({"src": s["index"], "kind": d.kind, "rule": d.rule_id, "reason": d.reason, "confidence": d.confidence,
                     "flags": d.flags, "out": [m["out_index"] for m in mapped], "override": bool(o)})
    return {
        "version": 1,
        "created": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "rules_version": RULES_VERSION,
        "source": {"name": src.name, "path": str(src), "sha256": sha256(src), "slides": len(slides), "slide_size": [sw, sh]},
        "template": {"file": catalog["source"]["file"], "sha256": catalog["source"]["sha256"], "slide_size": [tw, th]},
        "options": {**asdict(opts), "scale": scale},
        "slides": rows,
        "ops": ops,
        "stats": {"out_slides": len(ops), "copy": sum(o["op"] == "copy" for o in ops),
                  "template": sum(o["op"] == "template" for o in ops),
                  "flagged": sum(1 for o in ops if o["flags"]), "unknown": sum(r["kind"] == K.UNKNOWN for r in rows)},
    }


def save_plan(plan: dict, path: Path) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(plan, ensure_ascii=False, indent=1, default=_jsonable), encoding="utf-8")
    return path


def load_plan(path: Path) -> dict:
    return json.loads(Path(path).read_text(encoding="utf-8"))
