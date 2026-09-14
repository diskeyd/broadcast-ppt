"""Text width and line-count estimation without a renderer.

Character widths (in em) per character class are fitted from PowerPoint's own measurements:
single-line auto-fit boxes (wrap="none" + spAutoFit) store the width PowerPoint measured, and
wrapping auto-fit boxes store a height that is an exact multiple of the line height.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import yaml

EMU_PER_PT = 12700
LINE_HEIGHT = 1.212
CLASSES = ("hangul", "latin_upper", "latin_lower", "digit", "space", "punct", "wide")
DEFAULT_EM = {"hangul": 0.95, "latin_upper": 0.65, "latin_lower": 0.52, "digit": 0.55, "space": 0.28, "punct": 0.35, "wide": 0.9}


def char_class(ch: str) -> str:
    if ch == " " or ch == " ":
        return "space"
    o = ord(ch)
    if 0xAC00 <= o <= 0xD7A3 or 0x3131 <= o <= 0x318E or 0x4E00 <= o <= 0x9FFF:
        return "hangul"
    if "A" <= ch <= "Z":
        return "latin_upper"
    if "a" <= ch <= "z":
        return "latin_lower"
    if ch.isdigit():
        return "digit"
    if unicodedata.east_asian_width(ch) in ("W", "F"):
        return "wide"
    return "punct"


def class_counts(text: str) -> dict[str, int]:
    c = {k: 0 for k in CLASSES}
    for ch in unicodedata.normalize("NFC", text):
        if ch in "\r\n\x0b":
            continue
        c[char_class(ch)] += 1
    return c


@dataclass
class WidthModel:
    em: dict[str, float] = field(default_factory=lambda: dict(DEFAULT_EM))
    line_height: float = LINE_HEIGHT

    def width_emu(self, text: str, sz_pt: float) -> float:
        counts = class_counts(text)
        return sum(self.em[k] * n for k, n in counts.items()) * sz_pt * EMU_PER_PT

    def wrap_lines(self, text: str, box_inner_emu: float, sz_pt: float) -> list[str]:
        """Greedy word wrap per paragraph; words longer than a line break per character (Korean does)."""
        out: list[str] = []
        for para in unicodedata.normalize("NFC", text).split("\n"):
            words = re.split(r"(\s+)", para)
            line = ""
            for w in words:
                if not w:
                    continue
                cand = line + w
                if self.width_emu(cand.rstrip(), sz_pt) <= box_inner_emu or not line.strip():
                    if self.width_emu(cand.rstrip(), sz_pt) > box_inner_emu and not w.isspace():
                        # a single word wider than the line: break by characters
                        for ch in w:
                            if self.width_emu((line + ch).rstrip(), sz_pt) > box_inner_emu and line.strip():
                                out.append(line.rstrip())
                                line = ch
                            else:
                                line += ch
                    else:
                        line = cand
                else:
                    out.append(line.rstrip())
                    line = "" if w.isspace() else w
            out.append(line.rstrip())
        return out

    def estimate_lines(self, text: str, box_width_emu: int, sz_pt: float, lIns: int = 91440, rIns: int = 91440,
                       safety: float = 1.0) -> int:
        inner = (box_width_emu - lIns - rIns) * safety
        return len(self.wrap_lines(text, inner, sz_pt))

    def height_emu(self, lines: int, sz_pt: float, tIns: int = 45720, bIns: int = 45720) -> int:
        return int(lines * self.line_height * sz_pt * EMU_PER_PT + tIns + bIns)


def load_model(path: Path | None) -> WidthModel:
    if path and Path(path).exists():
        data = yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}
        em = dict(DEFAULT_EM)
        em.update({k: float(v) for k, v in (data.get("em") or {}).items()})
        return WidthModel(em=em, line_height=float(data.get("line_height", LINE_HEIGHT)))
    return WidthModel()
