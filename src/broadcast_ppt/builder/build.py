"""Build the broadcast deck from a plan."""

from __future__ import annotations

import hashlib
from pathlib import Path

from pptx import Presentation

from ..mapping.catalog import load_catalog
from .clone import copy_foreign_slide, duplicate_template_slide
from .fill import fill_shape, shape_by_order
from .package import current_slide_ids, drop_section_list, open_base, remove_slides


def build_deck(plan: dict, template: Path, catalog_path: Path, out: Path, force: bool = False) -> dict:
    template, out = Path(template), Path(out)
    catalog = load_catalog(catalog_path)
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
                res = fill_shape(shape_by_order(slide, spec["order"]), spec, op.get("fill", {}))
                rec["flags"] += res.flags + [f"missing:{m}" for m in res.missing]
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
