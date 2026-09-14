"""Build the broadcast deck from a plan."""

from __future__ import annotations

import hashlib
from pathlib import Path

from pptx import Presentation

from ..mapping.catalog import load_catalog
from ..mapping.textfit import WidthModel, load_model
from .clone import copy_foreign_slide, duplicate_template_slide
from .fill import fill_shape, shape_by_order
from .geometry import in_group, runs_with_size, set_font_size, set_width, shift
from .package import current_slide_ids, drop_section_list, open_base, remove_slides

SINGLE_LINE_SLACK = 1.06  # room for renderers whose glyphs run a little wider than PowerPoint's


def adjust_filled_shape(sp, spec: dict, op: dict, filled: dict, model: WidthModel) -> list[str]:
    flags = []
    if spec.get("set_cx"):  # body box narrowed to stay inside its bar
        set_width(sp, spec["set_cx"], keep_center=False)
    if spec.get("nudge_y"):
        shift(sp, dy=int(spec["nudge_y"]))
    segs = [s for p in spec.get("paragraphs", []) for s in p["segments"] if s["sample_len"] or s["sz"]]
    for slot, pt in (op.get("font_size") or {}).items():
        if slot in (spec.get("slot_map") or []):
            idx = 0 if spec.get("single_segment") else spec["slot_map"].index(slot)
            base = segs[min(idx, len(segs) - 1)]["sz"]
            if base and set_font_size(sp, int(round(base * 100)), int(round(pt * 100))):
                flags.append(f"font_size:{slot}")
    bp = spec.get("bodyPr", {})
    if filled and bp.get("wrap") == "none" and not spec.get("backing") and not in_group(sp):
        # a single-line box keeps the width PowerPoint measured for the sample text; renderers that honour
        # the stored width (Keynote, Quick Look) wrap longer text, so size the box to the new text
        default = max((s["sz"] or 0) for s in segs) if segs else 18
        width = sum(model.width_emu(t, sz) for t, sz in runs_with_size(sp, default))
        new_cx = int(width * SINGLE_LINE_SLACK + bp.get("lIns", 91440) + bp.get("rIns", 91440))
        set_width(sp, new_cx, keep_center=True)
    return flags


def shrink_variant(slide, entry: dict, lines: int) -> None:
    """Take unused body lines off a group-backed caption: bar, tag and text move down, the bottom stays."""
    s = entry["shrink"]
    d = int(lines * s["line_emu"])
    for role in ("group", "bar", "body"):
        shift(shape_by_order(slide, s[role]), dy=d, dcy=-d)
    shift(shape_by_order(slide, s["tag"]), dy=d)


def build_deck(plan: dict, template: Path, catalog_path: Path, out: Path, force: bool = False,
               widths_path: Path | None = Path("config/widths.yaml")) -> dict:
    template, out = Path(template), Path(out)
    catalog = load_catalog(catalog_path)
    model = load_model(widths_path)
    tsha = hashlib.sha256(template.read_bytes()).hexdigest()
    if tsha != plan["template"]["sha256"] and not force:
        raise SystemExit(f"template changed since planning: {template}")
    src_path = Path(plan["source"]["path"])
    if hashlib.sha256(src_path.read_bytes()).hexdigest() != plan["source"]["sha256"] and not force:
        raise SystemExit(f"source deck changed since planning: {src_path}")
    prs = open_base(template)
    base_ids = current_slide_ids(prs)
    tpl_slides = list(prs.slides)
    src = Presentation(str(src_path))
    scale = plan["options"]["scale"]
    results = []
    for op in plan["ops"]:
        rec = {"out_index": op["out_index"], "op": op["op"], "flags": list(op.get("flags", []))}
        if op["op"] == "template":
            entry = catalog["slides"][op["template_slide"]]
            slide, rep = duplicate_template_slide(prs, tpl_slides[op["template_slide"] - 1])
            rec["flags"] += rep.flags
            for spec in entry["shapes"].values():
                if not spec.get("slot_map"):
                    continue
                sp = shape_by_order(slide, spec["order"])
                res = fill_shape(sp, spec, op.get("fill", {}))
                rec["flags"] += res.flags + [f"missing:{m}" for m in res.missing]
                rec["flags"] += adjust_filled_shape(sp, spec, op, res.filled, model)
            if entry.get("shrink") and op.get("shrink_lines"):
                shrink_variant(slide, entry, int(op["shrink_lines"]))
        elif op["op"] == "copy":
            slide, rep = copy_foreign_slide(prs, src.slides[op["src"] - 1], scale=scale, keep_bg=True)
            rec["flags"] += rep.flags
        results.append(rec)
    remove_slides(prs, base_ids)
    drop_section_list(prs)
    prs.core_properties.comments = f"broadcast-ppt rules={plan['rules_version']} source={plan['source']['sha256'][:12]} template={tsha[:12]}"
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    return {"out": str(out), "slides": len(results), "ops": results}
