"""Build the broadcast deck from a plan."""

from __future__ import annotations

import hashlib
from pathlib import Path

from pptx import Presentation

from ..mapping.catalog import load_catalog
from ..ooxml import P, REPO_ROOT
from ..mapping.textfit import WidthModel, load_model
from .clone import copy_foreign_slide, duplicate_template_slide
from .fill import fill_shape, shape_by_order
from .geometry import (EMU_PER_PT, box_of, cy_of, drop_blank_paragraphs, hide_shape, in_group, insets, one_line_height,
                       plain_line_spacing, runs_with_size, set_font_size, style_paragraphs, set_width, set_y, shift,
                       split_paragraph_boxes, text_align, translate, y_of)
from .package import current_slide_ids, drop_section_list, remove_slides
from .scale import scale_tree

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
    if filled and bp.get("wrap") == "none" and not in_group(sp) and (not spec.get("backing") or spec.get("fit_backing")):
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


def fit_boxed_line(slide, entry: dict, spec: dict, model: WidthModel) -> list[str]:
    """Hymn caption: the rounded box behind the lyric is sized to the lyric, so every line keeps the same
    margin around its text. A line too long for the box loses font size instead of pushing the box off
    the slide; if it is still too wide the validator's text_wider_than_bar warning picks it up."""
    cfg = spec["fit_backing"]
    pad = int(float(cfg.get("pad_pt", 12)) * EMU_PER_PT)
    limit = int(float(cfg["max_width_pt"]) * EMU_PER_PT)
    floor = float(cfg.get("min_pt", 0))
    sp = shape_by_order(slide, spec["order"])
    segs = [s for p in spec.get("paragraphs", []) for s in p["segments"] if s["sample_len"] or s["sz"]]
    base = max((s["sz"] or 0) for s in segs) or 44.0
    flags = []
    ins = insets(sp)
    # what the text may measure so that the box (text box + slack + insets + padding) still fits
    room = (limit - 2 * pad - ins["lIns"] - ins["rIns"]) / SINGLE_LINE_SLACK
    width = sum(model.width_emu(t, sz) for t, sz in runs_with_size(sp, base))
    if floor and width and width > room:
        pt = max(floor, base * room / width)
        if set_font_size(sp, int(round(base * 100)), int(round(pt * 100))):
            flags.append("font_size:body")
        base = pt
    fit_single_line_width(sp, base, model)
    box = box_of(sp)
    if box is None:
        return flags
    left = None
    for b in spec.get("backing", []):
        bar = shape_by_order(slide, b["order"])
        set_width(bar, min(box[2] + 2 * pad, limit), keep_center=True)
        bar_box = box_of(bar)
        left = bar_box[0] if left is None else min(left, bar_box[0])
    # a label over the bar keeps the lyric's left margin, so it lines up with the text under it
    for sib in entry["shapes"].values():
        if left is not None and sib.get("align_left_to_backing"):
            el = shape_by_order(slide, sib["order"])
            here = box_of(el)
            if here is not None:
                translate(el, left + pad - here[0], 0)
    return flags


def caption_band(catalog: dict) -> tuple[int, int, int, int]:
    """(x, y, cx, cy) of the widest caption body in the template: where a lower third lives.

    Read from the catalog rather than written down, so these slides follow the template if it moves."""
    bodies = [spec["bbox"] for entry in catalog["slides"].values() if entry.get("family") == "caption"
              for spec in entry["shapes"].values() if spec.get("slot_map")]
    x, y, cx, cy = max(bodies, key=lambda b: b[2] * b[3])
    return int(x), int(y), int(cx), int(cy)


def _clusters(slide) -> list[list]:
    """Top-level shapes grouped by containment: a label drawn inside an ellipse belongs to it.

    A "shape" the owner points at is what the eye sees as one thing, not one XML element, so a box
    whose frame sits inside another's is carried along with it instead of being placed on its own."""
    items = []
    for el in slide.shapes._spTree:
        box = box_of(el)
        if box is not None:
            items.append((el, box))
    items.sort(key=lambda t: t[1][2] * t[1][3], reverse=True)  # containers before what sits in them
    out: list[list] = []
    for el, (x, y, cx, cy) in items:
        host = None
        for c in out:
            if any(bx <= x and by <= y and x + cx <= bx + bcx and y + cy <= by + bcy for _, (bx, by, bcx, bcy) in c):
                host = c
                break
        if host is None:
            host = []
            out.append(host)
        host.append((el, (x, y, cx, cy)))
    return out


def _union(cluster) -> tuple[int, int, int, int]:
    xs = [b[0] for _, b in cluster] + [b[0] + b[2] for _, b in cluster]
    ys = [b[1] for _, b in cluster] + [b[1] + b[3] for _, b in cluster]
    return min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys)


def layout_bottom_row(slide, band: tuple[int, int, int, int]) -> int:
    """Spread the slide's shape clusters along the bottom band, equally spaced, bottom aligned.

    `band` is (x, y, cx, cy) in EMU — the lower-third caption box, so these slides sit where every
    other caption sits. Clusters only ever shrink; the gaps carry whatever width is left over."""
    groups = [c for c in _clusters(slide) if c]
    if not groups:
        return 0
    groups.sort(key=lambda c: _union(c)[0])  # left to right, as they were read on the original
    bx, by, bcx, bcy = band
    boxes = [_union(c) for c in groups]
    k = min([1.0] + [bcy / h for *_, h in boxes if h > 0])
    total = sum(w for *_, w, _ in boxes) * k
    if total > bcx:  # too wide even at the band height: shrink until the row fits with no gap
        k *= bcx / total
        total = bcx
    # the outer clusters sit at the band's edges (bottom left and bottom right) with equal gaps
    # between them, so two shapes land in the corners and three or four space out evenly
    gap = (bcx - total) / (len(groups) - 1) if len(groups) > 1 else 0.0
    x = bx if len(groups) > 1 else bx + (bcx - total) / 2
    for cluster, (ux, uy, ucx, ucy) in zip(groups, boxes, strict=True):
        for el, _ in cluster:
            scale_tree(el, k)
        dy = int(by + bcy - (uy + ucy) * k)          # bottom of the band
        dx = int(x - ux * k)
        for el, _ in cluster:
            translate(el, dx, dy)
        x += ucx * k + gap
    return len(groups)


def layout_stack(slide, entry: dict, model: WidthModel) -> None:
    """Card rows (yellow label, white lines under it) are placed bottom-up with a fixed visual gap between
    the glyphs of neighbouring rows. The lowest row keeps the template's bottom edge. Glyph offsets per row
    (top and bottom of the drawn text, measured from the box top) are measured renders stored in roles.yaml."""
    st = entry["stack"]
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
    # last, so a row moves on its own: inside the chain the rows above would follow it, and before
    # the max_bottom clamp the lowest row's offset would be clamped away (this is why nudge_y_pt fails here)
    for el, row in rows:
        if row.get("offset_pt"):
            set_y(el, y_of(el) + int(float(row["offset_pt"]) * EMU_PER_PT))


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
            for spec in entry["shapes"].values():
                if spec.get("fit_backing"):
                    rec["flags"] += fit_boxed_line(slide, entry, spec, model)
            if entry.get("stack"):
                layout_stack(slide, entry, model)
            if entry.get("shrink") and op.get("shrink_lines"):
                shrink_variant(slide, entry, int(op["shrink_lines"]))
        elif op["op"] in ("copy", "shapes"):
            # "shapes": drop the source background so the template master's green shows through,
            # then move the shape clusters into the band the captions use
            keep_bg = op["op"] == "copy"
            slide, rep = copy_foreign_slide(prs, src.slides[op["src"] - 1], scale=scale, keep_bg=keep_bg)
            rec["flags"] += rep.flags
            if op["op"] == "shapes":
                n = layout_bottom_row(slide, caption_band(catalog))
                rec["flags"] += [] if n else ["missing:shapes"]
        results.append(rec)
    remove_slides(prs, base_ids)
    drop_section_list(prs)
    prs.core_properties.comments = f"broadcast-ppt rules={plan['rules_version']} source={plan['source']['sha256'][:12]} template={tsha[:12]}"
    out.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(out))
    return {"out": str(out), "slides": len(results), "ops": results}
