"""Template catalog: geometry, text-box formatting segments and line capacity per template slide.

Only lengths and sha1 digests of the template's sample text are stored — the template carries
real names in its samples.
"""

from __future__ import annotations

import hashlib
import unicodedata
from pathlib import Path

import yaml
from lxml import etree
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
}
DEFAULT_INSETS = {"lIns": 91440, "rIns": 91440, "tIns": 45720, "bIns": 45720}
LINE_HEIGHT = 1.212  # measured: spAutoFit box height = lines * 1.212 * font size + insets
EMU_PER_PT = 12700
COLOR_SETS = {"FFFF00": "a", "FFFF99": "b", "FFCCFF": "c"}


def sha1(text: str) -> str:
    return hashlib.sha1(unicodedata.normalize("NFC", text).encode()).hexdigest()


def _run_key(r) -> tuple:
    rpr = r.find("a:rPr", NS)
    if rpr is None:
        return (None, None, None, None, None)
    color = rpr.find("a:solidFill/a:srgbClr", NS)
    latin = rpr.find("a:latin", NS)
    ea = rpr.find("a:ea", NS)
    font = (ea.get("typeface") if ea is not None else None) or (latin.get("typeface") if latin is not None else None)
    return (rpr.get("sz"), color.get("val") if color is not None else None, font, rpr.get("b"), rpr.get("i"))


def segments_of_paragraph(p) -> list[dict]:
    """Collapse consecutive runs with the same size/colour/font/bold/italic into segments."""
    segs: list[dict] = []
    for r in p.findall("a:r", NS):
        t = r.findtext("a:t", default="", namespaces=NS)
        key = _run_key(r)
        blank = not t.strip()
        last = segs[-1] if segs else None
        # whitespace-only runs never start a segment: they join the segment before them, or at the start
        # of a paragraph the first real text after them (must match builder.fill.paragraph_segments)
        if last is not None and last["_lead"]:
            last["_text"] += t
            last["_runs"] += 1
            if not blank:
                last["_key"], last["_lead"] = key, False
        elif last is not None and (blank or last["_key"] == key):
            last["_text"] += t
            last["_runs"] += 1
        else:
            segs.append({"_key": key, "_text": t, "_runs": 1, "_lead": blank})
    return segs


def body_pr(txBody) -> dict:
    bp = txBody.find("a:bodyPr", NS)
    out = {k: int(bp.get(k)) if bp is not None and bp.get(k) else v for k, v in DEFAULT_INSETS.items()}
    out["wrap"] = (bp.get("wrap") if bp is not None else None) or "square"
    out["anchor"] = bp.get("anchor") if bp is not None else None
    out["autofit"] = next((etree.QName(c).localname for c in bp), None) if bp is not None and len(bp) else None
    return out


def text_shapes(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from text_shapes(sh.shapes)
        else:
            yield sh


def describe_shape(sh, slide_w: int, slide_h: int) -> dict:
    d: dict = {
        "id": sh.shape_id,
        "name": sh.name,
        "bbox": [int(sh.left), int(sh.top), int(sh.width), int(sh.height)],
        "bbox_frac": [round(sh.left / slide_w, 4), round(sh.top / slide_h, 4), round(sh.width / slide_w, 4), round(sh.height / slide_h, 4)],
        "in_group": sh._element.getparent().tag.endswith("}grpSp"),
    }
    if not sh.has_text_frame:
        return d
    txBody = sh.text_frame._txBody
    d["bodyPr"] = body_pr(txBody)
    paras = []
    max_sz = 0.0
    for p in txBody.findall("a:p", NS):
        segs = []
        for s in segments_of_paragraph(p):
            sz = int(s["_key"][0]) / 100 if s["_key"][0] else None
            if sz:
                max_sz = max(max_sz, sz)
            segs.append({"sz": sz, "color": s["_key"][1], "font": s["_key"][2], "bold": s["_key"][3] == "1",
                         "sample_len": len(s["_text"].strip()), "sample_sha1": sha1(s["_text"])})
        paras.append({"align": (p.find("a:pPr", NS).get("algn") if p.find("a:pPr", NS) is not None else None), "segments": segs})
    d["paragraphs"] = paras
    if max_sz and d["bodyPr"]["autofit"] == "spAutoFit":
        inner = d["bbox"][3] - d["bodyPr"]["tIns"] - d["bodyPr"]["bIns"]
        # each paragraph contributes lines at its own largest font size; the body paragraph size
        # (largest) sets the capacity of the remaining height
        para_sz = [max([sg["sz"] or 0 for sg in pr["segments"]] or [0]) or max_sz for pr in paras]
        d["line_emu"] = int(LINE_HEIGHT * max_sz * EMU_PER_PT)
        d["lines_measured"] = round(inner / (LINE_HEIGHT * max_sz * EMU_PER_PT), 2)
        fixed = sum(LINE_HEIGHT * sz * EMU_PER_PT for sz in para_sz[1:]) if len(para_sz) > 1 else 0
        d["max_lines"] = max(len(paras), round((inner - fixed) / (LINE_HEIGHT * max_sz * EMU_PER_PT)) + len(paras) - 1)
    return d


def build_catalog(template: Path, roles_path: Path) -> dict:
    template = Path(template)
    roles = yaml.safe_load(Path(roles_path).read_text(encoding="utf-8")) or {}
    prs = Presentation(str(template))
    W, H = int(prs.slide_width), int(prs.slide_height)
    master_bg = prs.slide_masters[0]._element.find("p:cSld/p:bg//a:srgbClr", NS)
    cat: dict = {
        "version": 1,
        "source": {"file": template.name, "sha256": hashlib.sha256(template.read_bytes()).hexdigest(),
                   "slide_size": [W, H], "slides": len(prs.slides),
                   "master_bg": master_bg.get("val") if master_bg is not None else None},
        "slides": {},
    }
    for idx, slide in enumerate(prs.slides, start=1):
        spec = (roles.get("slides") or {}).get(idx)
        if spec is None:
            continue
        shapes = {sh.shape_id: sh for sh in text_shapes(slide.shapes)}
        order = [int(el.get("id")) for el in slide.shapes._spTree.iter("{http://schemas.openxmlformats.org/presentationml/2006/main}cNvPr")]
        entry: dict = {"role": spec["role"], "family": spec["family"], "layout": slide.slide_layout.name,
                       "has_slide_bg": slide._element.find("p:cSld/p:bg", NS) is not None, "shapes": {}}
        colors = set()
        for sid, sspec in (spec.get("shapes") or {}).items():
            sh = shapes.get(int(sid))
            if sh is None:
                raise ValueError(f"template slide {idx}: shape id {sid} not found")
            desc = describe_shape(sh, W, H)
            desc["order"] = order.index(int(sid))  # position among cNvPr in document order (ids are renumbered on copy)
            desc["mode"] = sspec.get("mode", "fill")
            desc["single_segment"] = bool(sspec.get("single_segment"))
            slots = sspec.get("slots") or []
            flat = [seg for p in desc.get("paragraphs", []) for seg in p["segments"] if seg["sample_len"] or seg["sz"]]
            if slots and desc["single_segment"]:
                desc["slot_map"] = [slots[0]]
            elif slots:
                if len(slots) > len(flat):
                    raise ValueError(f"template slide {idx} shape {sid}: {len(slots)} slots but {len(flat)} segments")
                desc["slot_map"] = slots + [None] * (len(flat) - len(slots))
            for seg in flat:
                if seg["color"] in COLOR_SETS:
                    colors.add(COLOR_SETS[seg["color"]])
            clamp = sspec.get("clamp_to")
            if clamp is not None:
                # keep the text box inside its bar with the same margin on both sides
                bar = shapes.get(int(clamp))
                if bar is None:
                    raise ValueError(f"template slide {idx}: clamp shape {clamp} not found")
                margin = desc["bbox"][0] - int(bar.left)
                limit = int(bar.left) + int(bar.width) - margin
                if desc["bbox"][0] + desc["bbox"][2] > limit:
                    desc["bbox"][2] = limit - desc["bbox"][0]
                    desc["bbox_frac"][2] = round(desc["bbox"][2] / W, 4)
                    desc["set_cx"] = desc["bbox"][2]
            if sspec.get("nudge_y_pt"):
                desc["nudge_y"] = int(float(sspec["nudge_y_pt"]) * EMU_PER_PT)
            if sspec.get("drop_empty_paragraphs"):
                desc["drop_empty_paragraphs"] = True
            if sspec.get("fit_one_line"):
                desc["fit_one_line"] = dict(sspec["fit_one_line"])
            backing = (spec.get("backing") or {}).get(sid) or (spec.get("backing") or {}).get(int(sid))
            if backing:
                desc["backing"] = []
                for bid in backing:
                    b = shapes.get(int(bid))
                    if b is None:
                        raise ValueError(f"template slide {idx}: backing shape {bid} not found")
                    bd = describe_shape(b, W, H)
                    desc["backing"].append({"id": bd["id"], "order": order.index(int(bid)), "bbox": bd["bbox"], "in_group": bd["in_group"]})
            entry["shapes"][int(sid)] = desc
        if spec.get("stack"):
            rows = []
            for key in spec["stack"]["rows"]:
                sid, _, para = str(key).partition(".")
                glyph = (spec["stack"].get("glyph") or {}).get(str(key))
                if glyph is None:
                    raise ValueError(f"template slide {idx}: stack row {key} has no glyph offsets")
                rows.append({"order": order.index(int(sid)), "paragraph": int(para) if para else None,
                             "glyph": [float(glyph[0]), float(glyph[1])]})
            entry["stack"] = {"rows": rows, "gap_pt": float(spec["stack"].get("gap_pt", roles.get("card_gap_pt", 4)))}
            bottom = spec["stack"].get("max_bottom_pt", roles.get("card_max_bottom_pt"))
            if bottom is not None:
                entry["stack"]["max_bottom_pt"] = float(bottom)
        if spec.get("shrink"):
            sh = spec["shrink"]
            body = entry["shapes"][int(sh["body"])]
            entry["shrink"] = {k: order.index(int(sh[k])) for k in ("group", "bar", "tag", "body")}
            entry["shrink"]["line_emu"] = body["line_emu"]
        caps = [s.get("max_lines") for s in entry["shapes"].values() if s.get("slot_map") and "body" in s["slot_map"]]
        entry["line_capacity"] = max([c for c in caps if c] or [1])
        entry["color_set"] = sorted(colors)[0] if colors else None
        cat["slides"][idx] = entry
    return cat


def dump_catalog(cat: dict, out: Path) -> Path:
    out = Path(out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(yaml.safe_dump(cat, allow_unicode=True, sort_keys=False, width=140), encoding="utf-8")
    return out


def load_catalog(path: Path) -> dict:
    return yaml.safe_load(Path(path).read_text(encoding="utf-8"))


def families(cat: dict) -> dict[str, list[tuple[int, dict]]]:
    out: dict[str, list[tuple[int, dict]]] = {}
    for idx, entry in cat["slides"].items():
        out.setdefault(entry["family"], []).append((int(idx), entry))
    for v in out.values():
        v.sort(key=lambda t: t[1]["line_capacity"])
    return out
