"""Shared DrawingML units, defaults and run grouping used by the catalog, the builder and the validator."""

from __future__ import annotations

from pathlib import Path

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS = {"a": A, "p": P}

EMU_PER_PT = 12700
AUTOFIT_LINE_HEIGHT = 1.212  # measured: spAutoFit box height = lines * 1.212 * font size + insets
DEFAULT_INSETS = {"lIns": 91440, "rIns": 91440, "tIns": 45720, "bIns": 45720}

REPO_ROOT = Path(__file__).resolve().parents[2]


def run_key(r) -> tuple:
    """Formatting identity of a run: size, colour, font, bold, italic."""
    rpr = r.find("a:rPr", NS)
    if rpr is None:
        return (None, None, None, None, None)
    color = rpr.find("a:solidFill/a:srgbClr", NS)
    latin = rpr.find("a:latin", NS)
    ea = rpr.find("a:ea", NS)
    font = (ea.get("typeface") if ea is not None else None) or (latin.get("typeface") if latin is not None else None)
    return (rpr.get("sz"), color.get("val") if color is not None else None, font, rpr.get("b"), rpr.get("i"))


def run_segments(p) -> list[dict]:
    """Group a paragraph's runs into formatting segments ({key, runs, text}).

    Consecutive runs with the same formatting form one segment. Whitespace-only runs never start a segment:
    they join the segment before them, or — at the start of a paragraph — the first real text after them
    (a padding space in another size is not a slot)."""
    segs: list[dict] = []
    for r in p.findall("a:r", NS):
        t = r.findtext("a:t", default="", namespaces=NS)
        key = run_key(r)
        blank = not t.strip()
        last = segs[-1] if segs else None
        if last is not None and last["lead"]:
            last["runs"].append(r)
            last["text"] += t
            if not blank:
                last["key"], last["lead"] = key, False
        elif last is not None and (blank or last["key"] == key):
            last["runs"].append(r)
            last["text"] += t
        else:
            segs.append({"key": key, "runs": [r], "text": t, "lead": blank})
    return segs
