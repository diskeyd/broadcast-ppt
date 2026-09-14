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


def in_group(el) -> bool:
    return el.getparent() is not None and el.getparent().tag == f"{{{P}}}grpSp"
