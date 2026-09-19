"""Build the broadcast deck from a plan."""

from __future__ import annotations

import hashlib
from pathlib import Path

from pptx import Presentation

from ..mapping.catalog import load_catalog
from ..ooxml import REPO_ROOT
from ..mapping.textfit import WidthModel, load_model
from .clone import copy_foreign_slide, duplicate_template_slide
from .fill import fill_shape, shape_by_order
from .geometry import (EMU_PER_PT, cy_of, drop_blank_paragraphs, hide_shape, in_group, insets, one_line_height, plain_line_spacing,
                       runs_with_size, set_font_size, style_paragraphs, set_width, set_y, shift, split_paragraph_boxes, text_align, y_of)
from .package import current_slide_ids, drop_section_list, remove_slides

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
    if op.get("para_style") and "body" in (spec.get("slot_map") or []):
        style_paragraphs(sp, op["para_style"])
    if spec.get("hide_when_empty") and not any((op.get("fill") or {}).get(s) for s in spec.get("slot_map") or [] if s):
        hide_shape(sp)  # unknown reference: no empty tag box on air
    if spec.get("drop_empty_paragraphs"):
        drop_blank_paragraphs(sp)
    bp = spec.get("bodyPr", {})
    if filled and bp.get("wrap") == "none" and not spec.get("backing") and not in_group(sp):
        default = max((s["sz"] or 0) for s in segs) if segs else 18
        fit_single_line_width(sp, default, model)
    return flags


def fit_single_line_width(sp, default_sz: float, model: WidthModel) -> None:
    """A single-line box keeps the width PowerPoint measured for the sample text; renderers that honour the
    stored width (Keynote, Quick Look) wrap longer text, so the box is sized to the new text. A centred box
    keeps its centre; a left-aligned label keeps its left edge so it stays in line with the text under it."""
    ins = insets(sp)
    width = sum(model.width_emu(t, sz) for t, sz in runs_with_size(sp, default_sz))
    set_width(sp, int(width * SINGLE_LINE_SLACK + ins["lIns"] + ins["rIns"]), keep_center=text_align(sp) == "ctr")


def layout_stack(slide, entry: dict, model: WidthModel) -> None:
    """Card rows (yellow label, white lines under it) are placed bottom-up with a fixed visual gap between
    the glyphs of neighbouring rows. The lowest row keeps the template's bottom edge. Glyph offsets per row
    (top and bottom of the drawn text, measured from the box top) are measured renders stored in roles.yaml."""
    st = entry["stack"]
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    next_id = max(int(c.get("id")) for c in slide.shapes._spTree.iter(f"{{{P}}}cNvPr")) + 1
    split: dict[int, tuple[list, int]] = {}
    rows = []
    for row in st["rows"]:
        el = shape_by_order(slide, row["order"])
        if row.get("paragraph") is not None:
            if row["order"] not in split:
                bottom = y_of(el) + cy_of(el)
                boxes = split_paragraph_boxes(el, next_id)
                next_id += len(boxes)
                for bx in boxes:
                    fit_single_line_width(bx, 18, model)
                split[row["order"]] = (boxes, bottom)
            boxes, bottom = split[row["order"]]
            if row["paragraph"] >= len(boxes):
                continue
            el = boxes[row["paragraph"]]
        else:
            # same treatment as split rows: no percent line spacing, one-line height, bottom edge kept
            bottom = y_of(el) + cy_of(el)
            plain_line_spacing(el)
            one_line_height(el)
            set_y(el, bottom - cy_of(el))
        rows.append((el, row))
    if not rows:
        return
    last_el, last_row = rows[-1]
    if last_row.get("paragraph") is not None:
        set_y(last_el, split[last_row["order"]][1] - cy_of(last_el))
    max_bottom = st.get("max_bottom_pt")
    if max_bottom is not None and len(last_row["glyph"]) > 1:
        # the lowest drawn text must not run lower than the approved card bottom (or off the slide)
        over = y_of(last_el) + int(last_row["glyph"][1] * EMU_PER_PT) - int(float(max_bottom) * EMU_PER_PT)
        if over > 0:
            set_y(last_el, y_of(last_el) - over)
    gap = int(float(st.get("gap_pt", 4)) * EMU_PER_PT)
    for i in range(len(rows) - 2, -1, -1):
        el, row = rows[i]
        below, below_row = rows[i + 1]
        glyph_top_below = int(below_row["glyph"][0] * EMU_PER_PT)
        glyph_bottom = int(row["glyph"][1] * EMU_PER_PT)
        set_y(el, y_of(below) + glyph_top_below - gap - glyph_bottom)


def shrink_variant(slide, entry: dict, lines: int) -> None:
    """Take unused body lines off a group-backed caption: bar, tag and text move down, the bottom stays."""
    s = entry["shrink"]
    d = int(lines * s["line_emu"])
    for role in ("group", "bar", "body"):
        if role in s:
            shift(shape_by_order(slide, s[role]), dy=d, dcy=-d)
    if "tag" in s:
        shift(shape_by_order(slide, s["tag"]), dy=d)


def build_deck(plan: dict, template: Path, catalog_path: Path, out: Path, force: bool = False,
               widths_path: Path | None = REPO_ROOT / "config/widths.yaml") -> dict:
    template, out = Path(template), Path(out)
    catalog = load_catalog(catalog_path)
    model = load_model(widths_path)
    tsha = hashlib.sha256(template.read_bytes()).hexdigest()
    if tsha != plan["template"]["sha256"] and not force:
        raise ValueError(f"template changed since planning: {template}")
    src_path = Path(plan["source"]["path"])
    if hashlib.sha256(src_path.read_bytes()).hexdigest() != plan["source"]["sha256"] and not force:
        raise ValueError(f"source deck changed since planning: {src_path}")
    prs = Presentation(str(template))
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
            if entry.get("stack"):
                layout_stack(slide, entry, model)
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
