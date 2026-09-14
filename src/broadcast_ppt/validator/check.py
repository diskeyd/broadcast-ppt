"""Validate a built broadcast deck against its plan and the template catalog.

Findings never include slide text. Levels: critical (must fix before broadcast), warning, info."""

from __future__ import annotations

from collections import Counter
from dataclasses import asdict, dataclass
from pathlib import Path

from pptx import Presentation

from ..analyzer.core import classify_rgb
from ..builder.fill import flat_segments, segment_sha1, shape_by_order
from ..mapping.catalog import load_catalog
from ..mapping.textfit import WidthModel, load_model

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


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
    # stricter than the planner, but not blind: the planner uses the calibrated width factor (never
    # under-estimated on the calibration set); the validator removes most of that bonus so warnings mark
    # chunks that are close to overflowing, not every chunk.
    planner_model = load_model(widths_path)
    strict = WidthModel(em=planner_model.em, line_height=planner_model.line_height,
                        width_scale=max(1.0, planner_model.width_scale - 0.06))
    ops = plan["ops"]
    slides = list(prs.slides)

    if len(slides) != len(ops):
        findings.append(Finding("critical", "slide_count", None, f"deck has {len(slides)} slides, plan has {len(ops)} operations"))
    W, H = prs.slide_width, prs.slide_height
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
            else:
                for seg, cseg, slot in zip(segs, _catalog_segments(spec), slot_map):
                    if slot is None:
                        continue
                    if segment_sha1(seg["text"]) == cseg["sample_sha1"] and (mode == "fill" or slot in fill):
                        value = fill.get(slot)
                        intended = value is not None and segment_sha1(value.replace("\n", " ")) == cseg["sample_sha1"]
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
                if lines > spec["max_lines"]:
                    findings.append(Finding("warning", "lines_fit", n, f"estimated {lines} lines in a {spec['max_lines']}-line box (strict width)"))
            # single-line boxes: filled text wider than the bar it sits on
            if bp.get("wrap") == "none" and spec.get("backing"):
                text = "".join(seg["text"] for seg in segs)
                width = sum(strict.width_emu(seg["text"], int(seg["key"][0]) / 100 if seg["key"][0] else 28) for seg in segs)
                bar = max(b["bbox"][0] + b["bbox"][2] for b in spec["backing"]) - spec["bbox"][0]
                if width + bp["lIns"] + bp["rIns"] > bar:
                    findings.append(Finding("warning", "text_wider_than_bar", n, f"single-line text ~{int(width / 12700)}pt wide on a {int(bar / 12700)}pt bar"))
        for flag in op.get("flags", []):
            if flag.startswith("review:"):
                findings.append(Finding("info", "review", n, flag))
    for op in ops:
        if op["op"] == "copy":
            for flag in op.get("flags", []):
                if flag.startswith(("review:", "deferred:")):
                    findings.append(Finding("info", "review", op["out_index"], flag))

    counts = Counter(f.level for f in findings)
    return {"deck": str(deck), "ok": counts["critical"] == 0,
            "counts": {k: counts.get(k, 0) for k in ("critical", "warning", "info")},
            "findings": [asdict(f) for f in findings]}
