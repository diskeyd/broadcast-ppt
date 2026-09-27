"""Geometry adjustments applied to a duplicated template slide after its text is filled."""

from __future__ import annotations

import copy

from lxml import etree

from ..ooxml import A, AUTOFIT_LINE_HEIGHT as AUTOFIT_LINE, DEFAULT_INSETS, EMU_PER_PT, P


def xfrm(el):
    x = el.find(f"{{{P}}}spPr/{{{A}}}xfrm")
    if x is None:
        x = el.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm")
    return x


def shift(el, dy: int = 0, dcy: int = 0) -> None:
    """Move an element down by dy and change its height by dcy (EMU). Groups move their child frame too,
    so children keep a 1:1 mapping and can be shifted in the same units."""
    xf = xfrm(el)
    if xf is None:
        return
    off, ext = xf.find(f"{{{A}}}off"), xf.find(f"{{{A}}}ext")
    off.set("y", str(int(off.get("y")) + dy))
    ext.set("cy", str(max(0, int(ext.get("cy")) + dcy)))
    ch_off, ch_ext = xf.find(f"{{{A}}}chOff"), xf.find(f"{{{A}}}chExt")
    if ch_off is not None and ch_ext is not None:
        ch_off.set("y", str(int(ch_off.get("y")) + dy))
        ch_ext.set("cy", str(max(0, int(ch_ext.get("cy")) + dcy)))


def set_width(el, cx: int, keep_center: bool) -> None:
    xf = xfrm(el)
    if xf is None:
        return
    off, ext = xf.find(f"{{{A}}}off"), xf.find(f"{{{A}}}ext")
    old = int(ext.get("cx"))
    if keep_center:
        off.set("x", str(int(off.get("x")) + (old - int(cx)) // 2))
    ext.set("cx", str(int(cx)))


def set_font_size(sp_el, from_sz: int, to_sz: int) -> int:
    n = 0
    for tag in ("rPr", "endParaRPr", "defRPr"):
        for rpr in sp_el.iter(f"{{{A}}}{tag}"):
            if rpr.get("sz") == str(from_sz):
                rpr.set("sz", str(to_sz))
                n += 1
    return n


FILLS = ("noFill", "solidFill", "gradFill", "blipFill", "pattFill", "grpFill")


def style_paragraphs(sp_el, styles: list[dict]) -> None:
    """Per-paragraph {"sz": pt, "color": "RRGGBB"} on every run of that paragraph (catechism: orange question, smaller answer)."""
    for p, st in zip(sp_el.iter(f"{{{A}}}p"), styles, strict=False):
        for r in p.findall(f"{{{A}}}r"):
            rpr = _rpr(r)
            if "sz" in st:
                rpr.set("sz", str(int(round(st["sz"] * 100))))
            if "color" in st:
                _set_color(rpr, st["color"])


def _rpr(r):
    rpr = r.find(f"{{{A}}}rPr")
    if rpr is None:
        rpr = etree.Element(f"{{{A}}}rPr")
        r.insert(0, rpr)
    return rpr


def _set_color(rpr, color: str) -> None:
    for old in [c for c in rpr if c.tag in {f"{{{A}}}{t}" for t in FILLS}]:  # a run holds one fill
        rpr.remove(old)
    fill = etree.Element(f"{{{A}}}solidFill")
    etree.SubElement(fill, f"{{{A}}}srgbClr", val=color)
    ln = rpr.find(f"{{{A}}}ln")  # schema order: ln, then the fill
    rpr.insert(list(rpr).index(ln) + 1 if ln is not None else 0, fill)


def emphasize_runs(sp_el, phrases: list[dict]) -> None:
    """Colour / bold the source's emphasised phrases inside the filled body: each run holding one is split
    so the phrase has a run of its own, a copy of the run's rPr plus the colour and b="1"."""
    for r in list(sp_el.iter(f"{{{A}}}r")):
        text = r.findtext(f"{{{A}}}t", "")
        cuts = []
        for ph in phrases:
            i = text.find(ph["text"])
            if i >= 0 and not any(i < e and s < i + len(ph["text"]) for s, e, _ in cuts):
                cuts.append((i, i + len(ph["text"]), ph))
        if not cuts:
            continue
        pieces, at = [], 0
        for s, e, ph in sorted(cuts, key=lambda c: c[0]):
            pieces += [(text[at:s], None), (text[s:e], ph)]
            at = e
        pieces.append((text[at:], None))
        anchor = r
        for piece, ph in [pc for pc in pieces if pc[0]]:
            nr = copy.deepcopy(r)
            nr.find(f"{{{A}}}t").text = piece
            if ph is not None:
                rpr = _rpr(nr)
                if ph.get("color"):
                    _set_color(rpr, ph["color"])
                if ph.get("bold"):
                    rpr.set("b", "1")
            anchor.addnext(nr)
            anchor = nr
        r.getparent().remove(r)


def hide_shape(sp_el) -> None:
    """Draw nothing (no fill, no line, no theme effect) but keep the element: slots are found by shape order."""
    sppr = sp_el.find(f"{{{P}}}spPr")
    if sppr is None:
        return
    for el in [c for c in sppr if c.tag in {f"{{{A}}}{t}" for t in FILLS + ("effectLst", "effectDag")}]:
        sppr.remove(el)
    geom = next((c for c in sppr if c.tag in (f"{{{A}}}prstGeom", f"{{{A}}}custGeom")), None)
    anchor = geom if geom is not None else sppr.find(f"{{{A}}}xfrm")  # schema order: xfrm, geometry, fill, ln
    sppr.insert(list(sppr).index(anchor) + 1 if anchor is not None else 0, etree.Element(f"{{{A}}}noFill"))
    ln = sppr.find(f"{{{A}}}ln")
    if ln is None:
        ln = etree.Element(f"{{{A}}}ln")
        sppr.insert(list(sppr).index(sppr.find(f"{{{A}}}noFill")) + 1, ln)
    for el in [c for c in ln if c.tag in {f"{{{A}}}{t}" for t in FILLS}]:
        ln.remove(el)
    ln.insert(0, etree.Element(f"{{{A}}}noFill"))
    sppr.insert(list(sppr).index(ln) + 1, etree.Element(f"{{{A}}}effectLst"))  # overrides the style's effectRef


def runs_with_size(sp_el, default_sz: float) -> list[tuple[str, float]]:
    out = []
    for r in sp_el.iter(f"{{{A}}}r"):
        rpr = r.find(f"{{{A}}}rPr")
        sz = int(rpr.get("sz")) / 100 if rpr is not None and rpr.get("sz") else default_sz
        out.append((r.findtext(f"{{{A}}}t", default=""), sz))
    return out


def text_align(sp_el) -> str | None:
    p = sp_el.find(f".//{{{A}}}p")
    ppr = p.find(f"{{{A}}}pPr") if p is not None else None
    return ppr.get("algn") if ppr is not None else None


def drop_blank_paragraphs(sp_el) -> int:
    """Remove paragraphs that hold no visible text (template spacers), so a centred box centres the text."""
    body = sp_el.find(f".//{{{A}}}p/..")
    if body is None:
        return 0
    paras = body.findall(f"{{{A}}}p")
    if len(paras) < 2:
        return 0
    n = 0
    for p in paras:
        if any(r.findtext(f"{{{A}}}t", default="").strip() for r in p.findall(f"{{{A}}}r")):
            continue
        body.remove(p)
        n += 1
    return n


def in_group(el) -> bool:
    return el.getparent() is not None and el.getparent().tag == f"{{{P}}}grpSp"



def y_of(el) -> int:
    return int(xfrm(el).find(f"{{{A}}}off").get("y"))


def cy_of(el) -> int:
    return int(xfrm(el).find(f"{{{A}}}ext").get("cy"))


def set_y(el, y: int) -> None:
    xfrm(el).find(f"{{{A}}}off").set("y", str(int(y)))


def box_of(el) -> tuple[int, int, int, int] | None:
    """(x, y, cx, cy) in EMU, or None for an element with no frame of its own."""
    xf = xfrm(el)
    if xf is None:
        return None
    off, ext = xf.find(f"{{{A}}}off"), xf.find(f"{{{A}}}ext")
    if off is None or ext is None:
        return None
    return int(off.get("x")), int(off.get("y")), int(ext.get("cx")), int(ext.get("cy"))


def translate(el, dx: int, dy: int) -> None:
    """Move an element and, for a group, its child frame with it."""
    xf = xfrm(el)
    if xf is None:
        return
    for tag in ("off", "chOff"):
        node = xf.find(f"{{{A}}}{tag}")
        if node is not None:
            node.set("x", str(int(node.get("x")) + dx))
            node.set("y", str(int(node.get("y")) + dy))


def insets(sp_el) -> dict:
    bp = sp_el.find(f".//{{{A}}}bodyPr")
    d = dict(DEFAULT_INSETS)
    if bp is not None:
        d.update({k: int(bp.get(k)) for k in d if bp.get(k)})
    return d


def max_run_size(sp_el, default: float) -> float:
    sizes = [int(r.get("sz")) / 100 for r in sp_el.iter(f"{{{A}}}rPr") if r.get("sz")]
    return max(sizes) if sizes else default


def plain_line_spacing(sp_el) -> None:
    """Drop percent line spacing and space before/after. Renderers disagree on them (Keynote draws a
    reduced line height very differently from PowerPoint), so stacked card rows are placed by geometry."""
    for ppr in sp_el.iter(f"{{{A}}}pPr"):
        for tag in ("lnSpc", "spcBef", "spcAft"):
            el = ppr.find(f"{{{A}}}{tag}")
            if el is not None:
                ppr.remove(el)


def one_line_height(sp_el) -> None:
    ins = insets(sp_el)
    xfrm(sp_el).find(f"{{{A}}}ext").set("cy", str(int(AUTOFIT_LINE * max_run_size(sp_el, 18) * EMU_PER_PT + ins["tIns"] + ins["bIns"])))


def split_paragraph_boxes(sp_el, next_id: int) -> list:
    """Give every paragraph of a text box its own box. The first paragraph stays in the original element;
    copies for the others are appended to the end of the slide tree, so the document order of existing
    shapes does not change. Percent line spacing and space-before only made sense for stacked paragraphs
    and are removed; each box gets a one-line auto-fit height."""
    body = sp_el.find(f"{{{P}}}txBody")
    paras = body.findall(f"{{{A}}}p")
    plain_line_spacing(sp_el)
    tree = sp_el.getparent()
    boxes = [sp_el]
    for i in range(1, len(paras)):
        c = copy.deepcopy(sp_el)
        cbody = c.find(f"{{{P}}}txBody")
        for j, cp in enumerate(cbody.findall(f"{{{A}}}p")):
            if j != i:
                cbody.remove(cp)
        cnv = c.find(f".//{{{P}}}cNvPr")
        cnv.set("id", str(next_id + i - 1))
        cnv.set("name", f"{cnv.get('name')} ({i + 1})")
        tree.append(c)
        boxes.append(c)
    for p in paras[1:]:
        body.remove(p)
    for b in boxes:
        one_line_height(b)
    return boxes
