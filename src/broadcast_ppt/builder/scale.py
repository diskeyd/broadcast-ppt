"""Scale DrawingML geometry and text sizes when moving shapes between slide sizes.

The originals are 10in wide (9144000 EMU) and the template 13.333in (12192000 EMU); PowerPoint's own
paste between them multiplies positions, sizes, font sizes and insets by 4/3 — including insets that
were only implicit defaults, which it writes out explicitly (the template's lIns="60959").
"""

from __future__ import annotations

from lxml import etree

from ..ooxml import A, DEFAULT_INSETS

# local element name -> attributes scaled by k
SCALED_ATTRS: dict[str, tuple[str, ...]] = {
    "off": ("x", "y"),
    "ext": ("cx", "cy"),
    "chOff": ("x", "y"),
    "chExt": ("cx", "cy"),
    "rPr": ("sz", "spc", "kern"),
    "defRPr": ("sz", "spc", "kern"),
    "endParaRPr": ("sz", "spc", "kern"),
    "pPr": ("marL", "marR", "indent", "defTabSz"),
    "spcPts": ("val",),
    "buSzPts": ("val",),
    "ln": ("w",),
    "outerShdw": ("blurRad", "dist"),
    "innerShdw": ("blurRad", "dist"),
    "glow": ("rad",),
    "softEdge": ("rad",),
    "gridCol": ("w",),
    "tr": ("h",),
    "tab": ("pos",),
}
LVL_PPR = {f"lvl{i}pPr" for i in range(1, 10)}


def _scale_attr(el, name: str, k: float) -> None:
    v = el.get(name)
    if v is None:
        return
    try:
        el.set(name, str(int(round(int(v) * k))))
    except ValueError:
        pass  # percentages etc.


def scale_tree(root, k: float) -> None:
    """Multiply every length and font size in `root` (in place) by k."""
    if abs(k - 1.0) < 1e-9:
        return
    for el in root.iter():
        if not isinstance(el.tag, str):
            continue
        name = etree.QName(el).localname
        if name == "bodyPr" and etree.QName(el).namespace == A:
            for attr, default in DEFAULT_INSETS.items():
                el.set(attr, str(int(round(int(el.get(attr, default)) * k))))
            continue
        attrs = SCALED_ATTRS.get(name) or (SCALED_ATTRS["pPr"] if name in LVL_PPR else None)
        if attrs:
            for attr in attrs:
                _scale_attr(el, attr, k)
