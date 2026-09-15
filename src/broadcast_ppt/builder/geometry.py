"""Geometry adjustments applied to a duplicated template slide after its text is filled."""

from __future__ import annotations

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


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


EMU_PER_PT = 12700
AUTOFIT_LINE = 1.212  # PowerPoint spAutoFit height per line, in font sizes


def y_of(el) -> int:
    return int(xfrm(el).find(f"{{{A}}}off").get("y"))


def cy_of(el) -> int:
    return int(xfrm(el).find(f"{{{A}}}ext").get("cy"))


def set_y(el, y: int) -> None:
    xfrm(el).find(f"{{{A}}}off").set("y", str(int(y)))


def insets(sp_el) -> dict:
    bp = sp_el.find(f".//{{{A}}}bodyPr")
    d = {"lIns": 91440, "rIns": 91440, "tIns": 45720, "bIns": 45720}
    if bp is not None:
        d.update({k: int(bp.get(k)) for k in d if bp.get(k)})
    return d


def max_run_size(sp_el, default: float) -> float:
    sizes = [int(r.get("sz")) / 100 for r in sp_el.iter(f"{{{A}}}rPr") if r.get("sz")]
    return max(sizes) if sizes else default


def split_paragraph_boxes(sp_el, next_id: int) -> list:
    """Give every paragraph of a text box its own box. The first paragraph stays in the original element;
    copies for the others are appended to the end of the slide tree, so the document order of existing
    shapes does not change. Percent line spacing and space-before only made sense for stacked paragraphs
    and are removed; each box gets a one-line auto-fit height."""
    import copy
    body = sp_el.find(f"{{{P}}}txBody")
    paras = body.findall(f"{{{A}}}p")
    for p in paras:
        ppr = p.find(f"{{{A}}}pPr")
        if ppr is not None:
            for tag in ("lnSpc", "spcBef", "spcAft"):
                el = ppr.find(f"{{{A}}}{tag}")
                if el is not None:
                    ppr.remove(el)
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
        ins = insets(b)
        xfrm(b).find(f"{{{A}}}ext").set("cy", str(int(AUTOFIT_LINE * max_run_size(b, 18) * EMU_PER_PT + ins["tIns"] + ins["bIns"])))
    return boxes
