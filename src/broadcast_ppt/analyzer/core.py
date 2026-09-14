"""Read-only structure inventory of a .pptx file.

One JSON document per file: file facts, then per slide the resolved background,
a recursive shape inventory (groups, tables, pictures), text, keywords and
chroma-key candidates. Nothing here writes to the source file.
"""

from __future__ import annotations

import colorsys
import hashlib
import io
import re
import unicodedata
import zipfile
from collections import Counter
from pathlib import Path

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}
R_EMBED = "{%s}embed" % NS["r"]

# Order matters only for display; every matching label is recorded.
KEYWORDS: list[tuple[str, tuple[str, ...]]] = [
    ("예배의부름", ("예배의부름", "예배로의부름")),
    ("묵도", ("묵도",)),
    ("신앙고백", ("사도신경", "신앙고백")),
    ("주기도문", ("주기도문",)),
    ("찬송", ("찬송", "찬양")),
    ("교독문", ("교독문", "교독")),
    ("대표기도", ("대표기도",)),
    ("기도", ("기도",)),
    ("성경봉독", ("성경봉독", "말씀봉독", "봉독")),
    ("설교", ("설교", "말씀선포")),
    ("헌금", ("헌금", "봉헌")),
    ("광고", ("광고", "교회소식", "교우소식", "알림")),
    ("축도", ("축도",)),
    ("환영", ("환영", "새가족")),
    ("특송", ("특송", "특별찬양", "찬양대", "성가대")),
    ("성찬", ("성찬",)),
    ("세례", ("세례",)),
    ("묵상", ("묵상",)),
]

SCRIPTURE_RE = re.compile(
    r"[가-힣]{1,8}\s?\d{1,3}\s?[:：]\s?\d{1,3}"
    r"|\d{1,3}\s?[장편]\s?\d{1,3}\s?절"
)
HYMN_RE = re.compile(r"(찬송가?|새찬송가)\s?\d{1,3}\s?장")

KNOWN_FAMILIES = (
    "주일오전예배", "주일오후예배", "주일저녁예배", "세퍼드라이프", "성탄감사예배",
    "송구영신예배", "송구예배", "수요오후예배", "수요저녁예배", "새벽기도회",
    "합본", "바이블인카운터", "부활주일", "주일오후기도회",
)

PRESET_COLORS = {"green": "008000", "lime": "00FF00", "blue": "0000FF", "black": "000000", "white": "FFFFFF"}
CLRMAP_DEFAULT = {"bg1": "lt1", "tx1": "dk1", "bg2": "lt2", "tx2": "dk2"}


def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def date_token(name: str) -> str | None:
    m = re.search(r"(?<!\d)20(\d{6})(?!\d)", name)
    if m:
        return m.group(1)
    m = re.search(r"(?<!\d)(\d{6})(?!\d)", name)
    if m:
        return m.group(1)
    m = re.search(r"\((20\d{2})\)", name)
    return m.group(1) if m else None


def family_of(name: str) -> tuple[str, str | None]:
    stem = nfc(Path(name).stem)
    suffix = None
    m = re.search(r"\)\s*-\s*(.+)$", stem)
    if m:
        suffix = m.group(1).strip()
    compact = re.sub(r"\s", "", stem)
    for fam in KNOWN_FAMILIES:
        if fam in compact:
            return fam, suffix
    m = re.search(r"([가-힣]+예배)", compact)
    if m:
        return m.group(1), suffix
    return "기타", suffix


def classify_rgb(hexs: str | None) -> str | None:
    """'green' / 'blue' when the colour is a saturated key colour, else None."""
    if not hexs or len(hexs) != 6:
        return None
    try:
        r, g, b = (int(hexs[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except ValueError:
        return None
    h, s, v = colorsys.rgb_to_hsv(r, g, b)
    hue = h * 360
    if s >= 0.6 and v >= 0.55:
        if 75 <= hue <= 165:
            return "green"
        if 195 <= hue <= 255:
            return "blue"
    return None


def keywords_in(text: str) -> list[str]:
    compact = re.sub(r"\s", "", text)
    return [label for label, pats in KEYWORDS if any(p in compact for p in pats)]


def ln(el) -> str:
    return etree.QName(el).localname


class Context:
    """Per-presentation caches: theme colours per master, image stats per blob."""

    def __init__(self, prs):
        self.prs = prs
        self.slide_w = int(prs.slide_width or 0)
        self.slide_h = int(prs.slide_height or 0)
        self.image_cache: dict[str, dict] = {}
        self.theme_cache: dict[int, tuple[dict, dict]] = {}

    @property
    def slide_area(self) -> int:
        return max(self.slide_w * self.slide_h, 1)

    def theme_for(self, master) -> tuple[dict, dict]:
        key = id(master.part)
        if key in self.theme_cache:
            return self.theme_cache[key]
        colors: dict[str, str] = {}
        for rel in master.part.rels.values():
            if rel.is_external or not rel.reltype.endswith("/theme"):
                continue
            root = etree.fromstring(rel.target_part.blob)
            scheme = root.find(".//a:clrScheme", NS)
            if scheme is None:
                continue
            for child in scheme:
                if len(child) == 0:
                    continue
                c = child[0]
                val = c.get("val") if ln(c) == "srgbClr" else c.get("lastClr")
                if val:
                    colors[ln(child)] = val.upper()
        clrmap = dict(CLRMAP_DEFAULT)
        cm = master._element.find("p:clrMap", NS)
        if cm is not None:
            clrmap.update(cm.attrib)
        self.theme_cache[key] = (colors, clrmap)
        return self.theme_cache[key]

    def image_stats(self, blob: bytes, content_type: str | None) -> dict:
        key = hashlib.sha1(blob).hexdigest()
        if key in self.image_cache:
            return self.image_cache[key]
        out: dict = {"sha1": key[:12], "content_type": content_type, "bytes": len(blob)}
        try:
            im = Image.open(io.BytesIO(blob))
            out["format"] = im.format
            out["px"] = list(im.size)
            out["mode"] = im.mode
            if im.format == "JPEG":
                im.draft("RGB", (128, 128))
            rgba = im.convert("RGBA")
            rgba.thumbnail((64, 64))
            px = list(rgba.getdata())
            n = max(len(px), 1)
            opaque = [c[:3] for c in px if c[3] >= 128]
            out["transparent_frac"] = round(1 - len(opaque) / n, 3)
            m = max(len(opaque), 1)
            out["mean_rgb"] = "%02X%02X%02X" % tuple(sum(c[i] for c in opaque) // m for i in range(3))
            kinds = Counter(classify_rgb("%02X%02X%02X" % c) for c in opaque)
            out["green_frac"] = round(kinds["green"] / n, 3)
            out["blue_frac"] = round(kinds["blue"] / n, 3)
            out["black_frac"] = round(sum(1 for c in opaque if max(c) <= 24) / n, 3)
            buckets = Counter(tuple(min(v // 51 * 51, 255) for v in c) for c in opaque)
            out["top_colors"] = [["%02X%02X%02X" % k, round(v / n, 3)] for k, v in buckets.most_common(3)]
        except Exception as exc:  # EMF/WMF and friends
            out["decode_error"] = type(exc).__name__
        self.image_cache[key] = out
        return out


def color_of(fill_el, ctx: Context, master) -> dict:
    """Resolve an a:solidFill (or a colour-bearing element) to an RGB record."""
    if fill_el is None or len(fill_el) == 0:
        return {"rgb": None, "src": None}
    c = fill_el[0]
    tag = ln(c)
    rgb = None
    if tag == "srgbClr":
        rgb = c.get("val")
    elif tag == "schemeClr":
        colors, clrmap = ctx.theme_for(master)
        name = c.get("val")
        rgb = colors.get(clrmap.get(name, name))
    elif tag == "sysClr":
        rgb = c.get("lastClr")
    elif tag == "prstClr":
        rgb = PRESET_COLORS.get(c.get("val", ""))
    rec = {"rgb": rgb.upper() if rgb else None, "src": tag if tag != "schemeClr" else "scheme:" + c.get("val", "")}
    mods = [ln(m) for m in c]
    if mods:
        rec["mods"] = mods
    rec["key"] = classify_rgb(rec["rgb"])
    return rec


def fill_record(container, part, ctx: Context, master) -> dict | None:
    """Fill inside p:bgPr / p:spPr / p:grpSpPr."""
    if container is None:
        return None
    solid = container.find("a:solidFill", NS)
    if solid is not None:
        return {"type": "solid", **color_of(solid, ctx, master)}
    blip_fill = container.find("a:blipFill", NS)
    if blip_fill is not None:
        blip = blip_fill.find("a:blip", NS)
        rid = blip.get(R_EMBED) if blip is not None else None
        rec: dict = {"type": "picture"}
        if rid:
            try:
                img_part = part.related_part(rid)
                rec["image"] = ctx.image_stats(img_part.blob, img_part.content_type)
            except Exception as exc:
                rec["error"] = type(exc).__name__
        return rec
    grad = container.find("a:gradFill", NS)
    if grad is not None:
        stops = [color_of(gs, ctx, master) for gs in grad.findall(".//a:gs", NS)]
        return {"type": "gradient", "stops": [s["rgb"] for s in stops], "key": next((s["key"] for s in stops if s.get("key")), None)}
    if container.find("a:pattFill", NS) is not None:
        return {"type": "pattern"}
    if container.find("a:noFill", NS) is not None:
        return {"type": "none"}
    return None


def background_of(slide, ctx: Context) -> dict:
    layout = slide.slide_layout
    master = layout.slide_master
    for level, obj in (("slide", slide), ("layout", layout), ("master", master)):
        bg = obj._element.find("p:cSld/p:bg", NS)
        if bg is None:
            continue
        bgpr = bg.find("p:bgPr", NS)
        if bgpr is not None:
            rec = fill_record(bgpr, obj.part, ctx, master) or {"type": "unknown"}
            return {"level": level, **rec}
        ref = bg.find("p:bgRef", NS)
        if ref is not None:
            return {"level": level, "type": "themeRef", "idx": ref.get("idx"), **color_of(ref, ctx, master)}
    return {"level": "default", "type": "none"}


def picture_key(rec: dict | None) -> str | None:
    img = (rec or {}).get("image") or {}
    if img.get("green_frac", 0) >= 0.5:
        return "green"
    if img.get("blue_frac", 0) >= 0.5:
        return "blue"
    return None


def shape_kind(sh) -> str:
    try:
        if sh.is_placeholder:
            return "ph:" + sh.placeholder_format.type.name
    except Exception:
        pass
    try:
        st = sh.shape_type
        if st is not None:
            return st.name
    except Exception:
        pass
    gd = sh._element.find(".//a:graphicData", NS)
    if gd is not None:
        uri = gd.get("uri", "")
        if "diagram" in uri:
            return "SMART_ART"
        if "chart" in uri:
            return "CHART"
        if "table" in uri:
            return "TABLE"
    return "UNKNOWN"


def walk_shapes(shapes, part, ctx: Context, master, depth: int = 0) -> list[dict]:
    out: list[dict] = []
    for sh in shapes:
        d: dict = {"name": sh.name, "kind": shape_kind(sh), "depth": depth}
        try:
            left, top, width, height = sh.left, sh.top, sh.width, sh.height
            if None not in (left, top, width, height):
                d["bbox"] = [int(left), int(top), int(width), int(height)]
                d["area_frac"] = round(int(width) * int(height) / ctx.slide_area, 3)
        except Exception:
            pass
        sppr = sh._element.find("p:spPr", NS)
        if sppr is None:
            sppr = sh._element.find("p:grpSpPr", NS)
        fill = fill_record(sppr, part, ctx, master)
        if fill and fill["type"] != "none":
            d["fill"] = fill

        if getattr(sh, "has_text_frame", False) and sh.has_text_frame:
            paras = []
            fonts = set()
            for p in sh.text_frame.paragraphs:
                paras.append("".join(r.text for r in p.runs))
                for r in p.runs:
                    if r.font.name:
                        fonts.add(r.font.name)
                    if r.font.size:
                        d.setdefault("font_pt", set()).add(round(r.font.size.pt, 1))
            text = "\n".join(paras).strip()
            if text:
                d["text"] = text
            if fonts:
                d["fonts"] = sorted(fonts)
            if "font_pt" in d:
                d["font_pt"] = sorted(d["font_pt"])

        if getattr(sh, "has_table", False) and sh.has_table:
            tbl = sh.table
            cells = [[cell.text for cell in row.cells] for row in tbl.rows]
            d["table"] = {"rows": len(tbl.rows), "cols": len(tbl.columns), "cells": cells}

        if d["kind"] in ("PICTURE", "ph:PICTURE") or hasattr(sh, "image"):
            try:
                img = sh.image
                big = d.get("area_frac", 0) >= 0.25
                d["image"] = ctx.image_stats(img.blob, img.content_type) if big else {
                    "content_type": img.content_type, "bytes": len(img.blob)}
            except Exception as exc:
                if d["kind"] == "PICTURE":
                    d["image_error"] = type(exc).__name__

        if d["kind"] == "GROUP":
            d["children"] = walk_shapes(sh.shapes, part, ctx, master, depth + 1)
        out.append(d)
    return out


def flatten(shapes: list[dict]):
    for s in shapes:
        yield s
        yield from flatten(s.get("children", []))


def all_text(shapes: list[dict]) -> str:
    parts = []
    for s in flatten(shapes):
        if s.get("text"):
            parts.append(s["text"])
        if s.get("table"):
            parts.extend(c for row in s["table"]["cells"] for c in row if c.strip())
    return "\n".join(parts)


def chroma_candidates(bg: dict, shapes: list[dict], level: str) -> list[dict]:
    found = []
    bg_key = bg.get("key") or picture_key(bg)
    if bg_key and level == "slide":
        found.append({"source": "background:" + bg["level"], "fill": bg["type"], "key": bg_key,
                      "rgb": bg.get("rgb") or (bg.get("image") or {}).get("mean_rgb"), "area_frac": 1.0})
    for s in flatten(shapes):
        area = s.get("area_frac", 0)
        fill = s.get("fill") or {}
        key = fill.get("key") or picture_key(fill) or picture_key(s)
        if key and area >= 0.05:
            found.append({"source": level + "-shape", "kind": s["kind"], "key": key, "area_frac": area,
                          "rgb": fill.get("rgb") or (s.get("image") or fill.get("image") or {}).get("mean_rgb"),
                          "bbox": s.get("bbox")})
    return found


def signature(shapes: list[dict], layout_name: str, ctx: Context) -> str:
    items = []
    for s in shapes:
        bb = s.get("bbox")
        if bb and ctx.slide_w and ctx.slide_h:
            g = (round(bb[0] / ctx.slide_w, 1), round(bb[1] / ctx.slide_h, 1),
                 round(bb[2] / ctx.slide_w, 1), round(bb[3] / ctx.slide_h, 1))
        else:
            g = None
        items.append((s["kind"], g))
    raw = layout_name + "|" + repr(sorted(items, key=repr))
    return hashlib.sha1(raw.encode()).hexdigest()[:10]


def ratio_label(w: int, h: int) -> str:
    if not h:
        return "?"
    r = w / h
    for label, val in (("16:9", 16 / 9), ("4:3", 4 / 3), ("16:10", 1.6)):
        if abs(r - val) < 0.02:
            return label
    return f"{r:.3f}"


def zip_facts(path: Path) -> dict:
    facts: dict = {}
    with zipfile.ZipFile(path) as z:
        names = z.namelist()
        media = Counter(Path(n).suffix.lower() for n in names if n.startswith("ppt/media/"))
        facts["media"] = dict(media)
        facts["embedded_fonts"] = sum(1 for n in names if n.startswith("ppt/fonts/"))
        if "docProps/app.xml" in names:
            root = etree.fromstring(z.read("docProps/app.xml"))
            for el in root:
                if ln(el) in ("Application", "AppVersion"):
                    facts[ln(el)] = el.text
    return facts


def _open_skipping_bad_members(path: Path):
    """Re-pack a deck whose zip has unreadable members (e.g. one corrupt image).

    Broken media parts are replaced by empty bytes so the rest of the deck can be
    inventoried. Returns (Presentation, [damaged part names]) or None.
    """
    bad: list[str] = []
    buf = io.BytesIO()
    try:
        with zipfile.ZipFile(path) as src, zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as dst:
            for item in src.infolist():
                try:
                    data = src.read(item.filename)
                except Exception:
                    if not item.filename.startswith("ppt/media/"):
                        return None
                    bad.append(item.filename)
                    data = b""
                dst.writestr(item, data)
        if not bad:
            return None
        buf.seek(0)
        return Presentation(buf), bad
    except Exception:
        return None


def analyze_file(path: Path, root: Path | None = None) -> dict:
    path = Path(path)
    rel = path.relative_to(root) if root else Path(path.name)
    name = nfc(path.name)
    fam, suffix = family_of(name)
    info: dict = {
        "path": nfc(str(rel)),
        "name": name,
        "folder": nfc(str(rel.parent)) if str(rel.parent) != "." else "",
        "family": fam,
        "suffix": suffix,
        "date": date_token(name),
        "bytes": path.stat().st_size,
        "mtime": int(path.stat().st_mtime),
    }
    try:
        info.update(zip_facts(path))
    except Exception as exc:
        info["error"] = f"zip: {type(exc).__name__}: {exc}"
        return info
    try:
        prs = Presentation(str(path))
    except Exception as exc:
        repaired = _open_skipping_bad_members(path)
        if repaired is None:
            info["error"] = f"open: {type(exc).__name__}: {exc}"
            return info
        prs, bad = repaired
        info["damaged_parts"] = bad

    ctx = Context(prs)
    info["slide_size_emu"] = [ctx.slide_w, ctx.slide_h]
    info["ratio"] = ratio_label(ctx.slide_w, ctx.slide_h)
    info["masters"] = len(prs.slide_masters)

    layout_cache: dict[int, dict] = {}
    slides = []
    fonts = Counter()
    for idx, slide in enumerate(prs.slides, start=1):
        layout = slide.slide_layout
        master = layout.slide_master
        try:
            shapes = walk_shapes(slide.shapes, slide.part, ctx, master)
        except Exception as exc:
            slides.append({"index": idx, "error": f"{type(exc).__name__}: {exc}"})
            continue
        bg = background_of(slide, ctx)
        text = all_text(shapes)
        show_master = slide._element.get("showMasterSp", "1") not in ("0", "false")

        lkey = id(layout.part)
        if lkey not in layout_cache:
            lshapes = [s for s in walk_shapes(layout.shapes, layout.part, ctx, master) if not s["kind"].startswith("ph:")]
            mshapes = [s for s in walk_shapes(master.shapes, master.part, ctx, master) if not s["kind"].startswith("ph:")]
            layout_cache[lkey] = {
                "chroma": chroma_candidates({}, lshapes, "layout") + chroma_candidates({}, mshapes, "master"),
                "decor_shapes": len(lshapes) + len(mshapes),
            }
        chroma = chroma_candidates(bg, shapes, "slide")
        if bg.get("level") in ("layout", "master") and (bg.get("key") or picture_key(bg)):
            chroma.append({"source": "background:" + bg["level"], "fill": bg["type"],
                           "key": bg.get("key") or picture_key(bg), "area_frac": 1.0,
                           "rgb": bg.get("rgb") or (bg.get("image") or {}).get("mean_rgb")})
        if show_master:
            chroma.extend(layout_cache[lkey]["chroma"])

        flat = list(flatten(shapes))
        kinds = Counter(s["kind"] for s in flat)
        for s in flat:
            fonts.update(s.get("fonts", []))
        full_pics = [s for s in flat if s["kind"] in ("PICTURE", "ph:PICTURE") and s.get("area_frac", 0) >= 0.9]
        notes = ""
        if slide.has_notes_slide:
            notes = (slide.notes_slide.notes_text_frame.text or "").strip() if slide.notes_slide.notes_text_frame else ""

        slides.append({
            "index": idx,
            "layout": nfc(layout.name or ""),
            "show_master_shapes": show_master,
            "background": bg,
            "kinds": dict(kinds),
            "shapes": shapes,
            "text": text,
            "notes": notes,
            "keywords": keywords_in(text),
            "scripture_ref": bool(SCRIPTURE_RE.search(text)),
            "hymn_no": bool(HYMN_RE.search(text)),
            "chroma": chroma,
            "full_bleed_picture": bool(full_pics),
            "text_len": len(re.sub(r"\s", "", text)),
            "signature": signature(shapes, layout.name or "", ctx),
        })
    info["slide_count"] = len(slides)
    info["layouts_used"] = dict(Counter(s.get("layout", "?") for s in slides))
    info["fonts"] = dict(fonts.most_common())
    info["slides"] = slides
    return info
