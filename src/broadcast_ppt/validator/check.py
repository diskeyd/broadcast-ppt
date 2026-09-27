"""Validate a built broadcast deck against its plan and the template catalog.

Findings never include slide text. Levels: critical (must fix before broadcast), warning, info."""

from __future__ import annotations

from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path

from pptx import Presentation

from ..analyzer.core import classify_rgb
from ..builder.fill import flat_segments, shape_by_order
from ..mapping.catalog import load_catalog, sha1
from ..mapping.textfit import load_model, strict_model
from ..ooxml import A, P, R


@dataclass
class Finding:
    level: str
    check: str
    out_slide: int | None
    message: str


def _catalog_segments(spec: dict) -> list[dict]:
    return [s for p in spec.get("paragraphs", []) for s in p["segments"] if s["sample_len"] or s["sz"]]


def validate_deck(deck: Path, plan: dict, catalog_path: Path, widths_path: Path | None = None) -> dict:
    findings: list[Finding] = []
    prs = Presentation(str(deck))
    cat = load_catalog(catalog_path)
    # the planner works from the top of the measured width window; this is the bottom of it, so a
    # warning marks a chunk that could overflow anywhere inside what was actually measured on air
    strict = strict_model(load_model(widths_path))
    ops = plan["ops"]
    slides = list(prs.slides)

    if len(slides) != len(ops):
        findings.append(Finding("critical", "slide_count", None, f"deck has {len(slides)} slides, plan has {len(ops)} operations"))
    master_bg = prs.slide_masters[0]._element.find(f"{{{P}}}cSld/{{{P}}}bg//{{{A}}}srgbClr")
    master_green = master_bg is not None and classify_rgb(master_bg.get("val")) == "green"

    for op, slide in zip(ops, slides):
        n = op["out_index"]
        # relationships resolve and internal targets are not empty
        rel_ids = set(slide.part.rels.keys())
        for el in slide._element.iter():
            if not isinstance(el.tag, str):
                continue
            for attr, rid in el.attrib.items():
                if attr.startswith(f"{{{R}}}") and rid not in rel_ids:
                    findings.append(Finding("critical", "rel_resolves", n, f"{attr.split('}')[1]}={rid} has no relationship"))
        for rel in slide.part.rels.values():
            if not rel.is_external and rel.reltype.endswith(("/image", "/media", "/video", "/audio")) and not rel.target_part.blob:
                findings.append(Finding("critical", "rel_resolves", n, f"empty {rel.reltype.rsplit('/', 1)[-1]} part"))
        ids = Counter(el.get("id") for el in slide._element.iter(f"{{{P}}}cNvPr"))
        if any(v > 1 for v in ids.values()):
            findings.append(Finding("warning", "duplicate_shape_id", n, "shape ids repeat on this slide"))

        if op["op"] != "template":
            continue
        entry = cat["slides"][op["template_slide"]]
        fill = op.get("fill", {})
        bg = slide._element.find(f"{{{P}}}cSld/{{{P}}}bg//{{{A}}}srgbClr")
        green = classify_rgb(bg.get("val")) == "green" if bg is not None else master_green
        if not green:
            findings.append(Finding("warning", "green_background", n, "template slide is not on a green background"))
        for spec in entry["shapes"].values():
            slot_map = spec.get("slot_map")
            if not slot_map:
                continue
            sp = shape_by_order(slide, spec["order"])
            txBody = sp.find(f".//{{{P}}}txBody")
            if txBody is None:
                findings.append(Finding("critical", "shape_missing", n, f"template shape order {spec['order']} has no text"))
                continue
            segs = flat_segments(txBody)
            mode = spec.get("mode", "fill")
            if spec.get("single_segment"):
                slot = slot_map[0]
                if mode == "fill" and slot not in fill:
                    findings.append(Finding("critical", "empty_fill", n, f"slot '{slot}' not filled; template sample text remains"))
            elif op.get("emph") and slot_map == ["body"]:
                # emphasis split the body into styled runs: the box must read exactly the filled body
                if "".join(s["text"] for s in segs).strip() != fill.get("body", "").replace("\n", " ").strip():
                    findings.append(Finding("critical", "extra_template_text", n, "emphasised body differs from the planned text"))
            else:
                if mode == "fill" and len(segs) > len(slot_map):
                    findings.append(Finding("critical", "extra_template_text", n,
                                            f"{len(segs)} text segments for {len(slot_map)} slots; template text may remain"))
                for seg, cseg, slot in zip(segs, _catalog_segments(spec), slot_map):
                    if slot is None:
                        if mode == "fill" and cseg["sample_len"] and sha1(seg["text"]) == cseg["sample_sha1"]:
                            findings.append(Finding("critical", "leftover_sample_text", n, "an unmapped segment still holds template text"))
                        continue
                    if sha1(seg["text"]) == cseg["sample_sha1"] and (mode == "fill" or slot in fill):
                        value = fill.get(slot)
                        intended = value is not None and sha1(value.replace("\n", " ")) == cseg["sample_sha1"]
                        if not intended:  # a label that legitimately equals the template's own label is fine
                            findings.append(Finding("critical", "leftover_sample_text", n, f"slot '{slot}' still holds the template sample"))
                    if mode == "fill" and slot not in fill:
                        findings.append(Finding("critical", "empty_fill", n, f"slot '{slot}' has no value"))
            # line fit for wrapping boxes
            bp = spec.get("bodyPr", {})
            if bp.get("wrap") == "square" and spec.get("max_lines"):
                text = "\n".join("".join(r.findtext(f"{{{A}}}t", "") for r in p.findall(f"{{{A}}}r")) for p in txBody.findall(f"{{{A}}}p"))
                sz = max((s["sz"] or 0) for s in _catalog_segments(spec)) or 20
                lines = strict.estimate_lines(text, spec["bbox"][2], sz, bp["lIns"], bp["rIns"])
                # the bar may have been cut to fewer lines than the template holds: compare with what is there
                room = spec["max_lines"] - (op.get("shrink_lines", 0) if spec.get("slot_map") else 0)
                if lines > room:
                    findings.append(Finding("warning", "lines_fit", n, f"estimated {lines} lines in a {room}-line box (strict width)"))
            # single-line boxes: filled text wider than the bar it sits on
            grown = sp.find(f".//{{{A}}}bodyPr").get("wrap") == "square"  # a 교회소식 bar grown to more lines
            if bp.get("wrap") == "none" and spec.get("backing") and not grown:
                width = sum(strict.width_emu(seg["text"], int(seg["key"][0]) / 100 if seg["key"][0] else 28) for seg in segs)
                bar = max(b["bbox"][0] + b["bbox"][2] for b in spec["backing"]) - spec["bbox"][0]
                if width + bp["lIns"] + bp["rIns"] > bar:
                    findings.append(Finding("warning", "text_wider_than_bar", n, f"single-line text ~{int(width / 12700)}pt wide on a {int(bar / 12700)}pt bar"))
        for flag in op.get("flags", []):
            if flag.startswith("review:"):
                findings.append(Finding("info", "review", n, flag))
    for op in ops:
        if op["op"] != "template":  # template flags are collected above, while the slide is open
            for flag in op.get("flags", []):
                if flag.startswith(("review:", "deferred:")):
                    findings.append(Finding("info", "review", op["out_index"], flag))
        # a caption page after the first opens on a verse number. Starting mid-sentence means a verse
        # was cut between two pages, which reads as broken text on air and nothing else here sees it
        chunk = op.get("chunk") or {}
        body = (op.get("fill") or {}).get("body") or ""
        if chunk.get("i", 1) > 1 and body and not body.split(" ", 1)[0].strip().isdigit():
            findings.append(Finding("warning", "verse_split", op["out_index"],
                                    f"caption {chunk['i']} of {chunk.get('n')} starts mid-verse"))

    counts = Counter(f.level for f in findings)
    return {"deck": str(deck), "ok": counts["critical"] == 0,
            "counts": {k: counts.get(k, 0) for k in ("critical", "warning", "info")},
            "findings": [asdict(f) for f in findings]}
