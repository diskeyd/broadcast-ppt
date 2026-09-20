"""Build a reviewable conversion plan: original slide -> decision -> output operations."""

from __future__ import annotations

import functools
import hashlib
import json
import math
import re
from dataclasses import asdict, dataclass, replace
from datetime import datetime, UTC
from pathlib import Path

from ..analyzer.core import analyze_file
from ..rules import ai
from ..rules import kinds as K
from ..rules.classify import DeckContext, classify_deck, view_of
from ..rules.fields import ScriptureRef, compact, norm, scripture_ref, verse_chunks
from ..rules.version import RULES_VERSION
from .catalog import families, load_catalog
from .split import pack
from .textfit import WidthModel, load_model

TEMPLATE = {"call_to_worship": 3, "prayer": 4, "scripture": 6, "sermon": 7, "meditation": 8, "generic": 17,
            "announcement_title": 20, "announcement_first": 21, "announcement_item": 23, "offertory": 31}
REF_FLAGS = {"review:reference_unknown", "review:inherited_reference", "review:no_reference"}
QUESTION_COLOR, ANSWER_PT = "FF9300", 28.0  # catechism caption: template emphasis orange; answer at the scripture caption size
OVERRIDE_OPS = (None, "copy", "quote", "caption", "scripture")


@dataclass
class PlanOptions:
    slide_size: str = "template"         # template | source
    scripture_both: bool = True          # full-screen copy + lower-third captions
    sermon_material: str = "lower-third"  # lower-third | passthrough
    ai_model: str | None = None          # Gemini model asked about unclassified slides; None = rules only


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
        for idx, _ in self.fam[family]:
            if self.fits(idx, text):
                return idx
        return None

    def _fit(self, op: dict, tpl: int, chunk: str) -> None:
        op["fit"] = {"chars": len(chunk), "est_lines": self.lines(tpl, chunk), "max_lines": self._body_spec(tpl)["max_lines"]}
        spare = op["fit"]["max_lines"] - op["fit"]["est_lines"]
        if spare > 0 and self.cat["slides"][tpl].get("shrink"):
            op["shrink_lines"] = spare  # no empty lines at the bottom of the bar

    # --- per-kind mapping ------------------------------------------------------------------------
    def copy(self, src: int, note: str) -> dict:
        return {"op": "copy", "src": src, "note": note}

    def card(self, src: int, tpl: int, fill: dict, note: str) -> dict:
        return {"op": "template", "src": src, "template_slide": tpl, "fill": {k: v for k, v in fill.items() if v is not None}, "note": note}

    def scripture_captions(self, src: int, ref: ScriptureRef | None, body: str, heading: str | None = None) -> list[dict]:
        units = [t for _, t in verse_chunks(body, ref)]
        big = self.fam["scripture_ref"][-1][0]
        # verses are joined with a newline so each one starts its own line in the bar
        chunks = pack(units, lambda t: self.fits(big, t), sep="\n")
        ops = []
        for i, chunk in enumerate(chunks, 1):
            tpl = self.smallest_variant("scripture_ref", chunk) or big
            verses = [int(n) for n in re.findall(r"(?:^|\s)(\d{1,3})\s(?=\S)", chunk)]
            in_range = [v for v in verses if ref is not None and ref.contains(v)]
            if ref is not None and (in_range or not verses):
                v1, v2 = (min(in_range), max(in_range)) if in_range else (ref.verse_start, ref.verse_end)
                tag = f"{ref.book}{ref.chapter}:{v1}" + (f"-{v2}" if v2 != v1 else "")
            else:
                tag = None
            # a Psalms book division ([제 삼 권, 아삽의 시]) fills a tag the slide gives no reference for
            if tag is None and heading and self.fits(tpl, heading, "ref"):
                tag = heading
            # never leave the template's sample reference in place: blank + review flag instead
            op = self.card(src, tpl, {"ref": tag if tag is not None else "", "body": chunk}, "scripture caption")
            op["chunk"] = {"i": i, "n": len(chunks)}
            self._fit(op, tpl, chunk)
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
            self._fit(op, tpl, chunk)
            ops.append(op)
        return ops

    def catechism_captions(self, src: int, question: str, answer: str) -> list[dict]:
        """Question (template size, orange) and the start of the answer (ANSWER_PT) on the 3-line caption; the rest of
        the answer on further pages. pack() would join the paragraphs with spaces, so the first page is filled here."""
        tpl = self.fam["caption"][-1][0]
        spec = self._body_spec(tpl)
        bp, cap = spec["bodyPr"], spec["max_lines"]

        def est(paras: list[str], first: bool) -> int:
            return sum(self.model.estimate_lines(p, spec["bbox"][2], self._sz(spec, "body") if first and i == 0 else ANSWER_PT, bp["lIns"], bp["rIns"])
                       for i, p in enumerate(paras))

        words = answer.split()
        k = len(words)
        while k and est([question, " ".join(words[:k])], True) > cap:
            k -= 1
        pages = [[question] + ([" ".join(words[:k])] if k else [])]
        if words[k:]:
            pages += [[t] for t in pack([" ".join(words[k:])], lambda t: est([t], False) <= cap)]
        ops = []
        for i, paras in enumerate(pages):
            op = self.card(src, tpl, {"body": "\n".join(paras)}, "caption")
            n = est(paras, i == 0)
            op["fit"] = {"chars": sum(map(len, paras)), "est_lines": n, "max_lines": cap}
            # positional styles: _fill_single writes one paragraph per line of the value
            op["para_style"] = ([{"color": QUESTION_COLOR}] if i == 0 else []) + [{"sz": ANSWER_PT}] * (len(paras) - (i == 0))
            if cap > n and self.cat["slides"][tpl].get("shrink"):
                op["shrink_lines"] = cap - n
            ops.append(op)
        return ops

    def meditation(self, src: int, text: str, label: str | None = None) -> dict:
        """The quote stays on one line: when it is wider than the box, its font size goes down.
        label="" blanks the 묵상 label (override quote card); None keeps the template's."""
        tpl = TEMPLATE["meditation"]
        op = self.card(src, tpl, {"quote": text, "label": label}, "meditation")
        spec = self._body_spec(tpl, "quote")
        fit = spec.get("fit_one_line")
        base = self._sz(spec, "quote")
        size = base
        if fit:
            bp = spec["bodyPr"]
            inner = (spec["bbox"][2] - bp["lIns"] - bp["rIns"]) * float(fit.get("margin", 0.96))
            width = self.model.width_emu(text, base) * float(fit.get("font_scale", 1.0))
            if width > inner:
                size = max(float(fit.get("min_pt", 36)), math.floor(base * inner / width * 2) / 2)
                op["font_size"] = {"quote": size}
                if size == float(fit.get("min_pt", 36)):
                    op["flags"] = ["review:quote_too_long"]
        op["fit"] = {"chars": len(text), "sz": size, "base_sz": base}
        return op

    def announcements(self, src: int, items: list[dict]) -> list[dict]:
        ops = []
        for it in items:
            if it["no"] == 1 and it["title"].replace(" ", "") == "환영":
                label = f"교회 소식 – {it['no']}. {it['title']}"
                tpl = TEMPLATE["announcement_first"]
                for chunk in pack([it["body"]], functools.partial(self.fits, tpl)):
                    ops.append(self.card(src, tpl, {"label": label, "body": chunk}, "announcement"))
                continue
            title = f"[{it['title']}] : " if it["title"] else ""
            label = f"교회 소식 – {it['no']}. {title}"
            tpl = TEMPLATE["announcement_item"]
            spec = self._body_spec(tpl, "label")
            inner = spec["bbox"][2] - spec["bodyPr"]["lIns"] - spec["bodyPr"]["rIns"]
            label_w = self.model.width_emu(label, self._sz(spec, "label"))
            def fits(t):
                return label_w + self.model.width_emu(t, self._sz(spec, "body")) <= inner  # noqa: B023
            for chunk in pack([it["body"]], fits):
                ops.append(self.card(src, tpl, {"label": label, "body": chunk}, "announcement"))
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
            quote = f["quote"].strip().strip("“”\"")
            return [self.meditation(src, f"“{quote}”")]
        if d.kind == K.ANNOUNCEMENT_TITLE:
            return [self.card(src, TEMPLATE["announcement_title"], {}, "announcement card")]
        if d.kind == K.ANNOUNCEMENT_BLOCK:
            return self.announcements(src, f["items"])
        if d.kind == K.SCRIPTURE_BODY:
            ops = [self.copy(src, "scripture full screen")] if self.opts.scripture_both else []
            ops += self.scripture_captions(src, f.get("ref"), f.get("body", ""), f.get("heading"))
            for extra in f.get("extra") or []:  # further passages, in the order they are read
                ops += self.scripture_captions(src, extra["ref"], extra["body"], f.get("heading"))
            return ops
        if d.kind == K.CATECHISM:
            ops = [self.copy(src, "catechism full screen")] if self.opts.scripture_both else []
            q, _, a = f["qa"].partition("\n")
            if q:
                ops += self.catechism_captions(src, q, a)
            for verse in f.get("verses") or []:
                ops += self.scripture_captions(src, verse["ref"], verse["body"])
            return ops
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
    if "error" in doc:  # .ppt, a Mac '._' sidecar from a USB stick, or a broken file
        raise ValueError(f"{src.name}: .pptx로 읽을 수 없습니다. PowerPoint에서 .pptx로 다시 저장하거나, '._'로 시작하는 파일이면 진짜 파일을 고르세요 ({doc['error']})")
    slide_over = ((overrides or {}).get("slides") or {})
    # a slide the overrides file decides is never sent to the model
    decisions = classify_deck(doc, fallback=(lambda v, d: None if v.index in slide_over else ai.classify_view(v, d, opts.ai_model))
                              if opts.ai_model else None)
    mapper = Mapper(catalog, load_model(widths_path), opts)
    tw, th = catalog["source"]["slide_size"]
    sw, sh = doc["slide_size_emu"]
    scale = tw / sw if opts.slide_size == "template" else 1.0
    rows, ops = [], []
    slides = [s for s in doc["slides"] if "error" not in s]
    ctx = DeckContext.from_deck(doc)  # override text drops the same boilerplate lines the rules drop
    for s, d in zip(slides, decisions):
        o = slide_over.get(s["index"]) or {}
        op = o.get("op")
        if op not in OVERRIDE_OPS:
            raise ValueError(f"override slide {s['index']}: unknown op {op!r} (use copy, quote, caption or scripture)")
        if "ref" in o:  # a reference the source slide does not show, typed by a person
            if op in ("copy", "quote", "caption"):
                raise ValueError(f"override slide {s['index']}: op {op} takes no ref (use op: scripture)")
            ref = scripture_ref(str(o["ref"]))
            if ref is None or ("ref" not in d.fields and op != "scripture"):
                raise ValueError(f"override slide {s['index']}: ref {o['ref']!r} is not a reference or {d.kind} takes none")
            d = replace(d, fields={**d.fields, "ref": ref}, flags=[f for f in d.flags if f not in REF_FLAGS])
        # new slide forms are handled here, in the overrides file, not with new rules
        text = " ".join(ln.strip() for ln in norm(view_of(s, ctx).text).splitlines() if compact(ln)) if op in OVERRIDE_OPS[2:] else ""
        if op in OVERRIDE_OPS[2:] and not text:
            raise ValueError(f"override slide {s['index']}: op {op} needs slide text")
        if o.get("skip"):
            mapped = []
        elif op == "copy":
            mapped = [mapper.copy(s["index"], "override: copy")]
        elif op == "quote":  # quote card without the 묵상 label
            d = replace(d, kind=K.MEDITATION, rule_id="OVR", reason="override: quote card", flags=[])
            mapped = [mapper.meditation(s["index"], "“" + text.strip("“”\"") + "”", label="")]
        elif op == "caption":
            d = replace(d, kind=K.CAPTION_SENTENCE, rule_id="OVR", reason="override: caption", fields={"body": text}, flags=[])
            mapped = mapper.map(s["index"], d)
        elif op == "scripture":
            ref = d.fields.get("ref")
            d = replace(d, kind=K.SCRIPTURE_BODY, rule_id="OVR", reason="override: scripture", fields={"ref": ref, "body": text, "quote": False},
                        flags=[] if ref is not None else ["review:reference_unknown"])
            mapped = mapper.map(s["index"], d)
        else:
            mapped = mapper.map(s["index"], d)
        for m in mapped:
            m["out_index"] = len(ops) + 1
            m["kind"] = d.kind
            m["rule"] = d.rule_id
            m["flags"] = sorted(set(m.get("flags", []) + [x for x in d.flags if x]))
            ops.append(m)
        rows.append({"src": s["index"], "kind": d.kind, "rule": d.rule_id, "reason": d.reason, "confidence": d.confidence,
                     "flags": d.flags, "out": [m["out_index"] for m in mapped], "override": bool(o),
                     **({"ai": d.fields["ai"]} if "ai" in d.fields else {})})
    return {
        "version": 1,
        "created": datetime.now(UTC).isoformat(timespec="seconds"),
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
