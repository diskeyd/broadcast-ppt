"""Greedy, boundary-aware packing of text units into chunks that fit a template box."""

from __future__ import annotations

import re
from typing import Callable

SENTENCE_RE = re.compile(r"(?<=[.?!…다요라」”'\"])\s+")


def _pieces(unit: str, fits: Callable[[str], bool]) -> list[str]:
    """Break one unit that does not fit: sentences, then words, then characters."""
    if fits(unit):
        return [unit]
    for splitter in (SENTENCE_RE, re.compile(r"(?<=[,，;])\s+"), re.compile(r"\s+")):
        parts = [p for p in splitter.split(unit) if p.strip()]
        if len(parts) > 1:
            return pack(parts, fits)
    out, cur = [], ""
    for ch in unit:
        if cur and not fits(cur + ch):
            out.append(cur)
            cur = ch
        else:
            cur += ch
    return out + ([cur] if cur else [])


def pack(units: list[str], fits: Callable[[str], bool], joiner: str = " ") -> list[str]:
    chunks: list[str] = []
    cur = ""
    for unit in (u.strip() for u in units):
        if not unit:
            continue
        cand = f"{cur}{joiner}{unit}" if cur else unit
        if fits(cand):
            cur = cand
            continue
        if cur:
            chunks.append(cur)
            cur = ""
        pieces = _pieces(unit, fits)
        chunks.extend(pieces[:-1])
        cur = pieces[-1] if pieces else ""
    if cur:
        chunks.append(cur)
    # avoid a tiny last chunk: move the last unit of the previous chunk forward when both still fit
    if len(chunks) >= 2 and len(chunks[-1]) < 0.35 * max(len(c) for c in chunks[:-1]):
        prev_words = chunks[-2].split(" ")
        if len(prev_words) > 3:
            moved = " ".join(prev_words[-2:])
            new_prev, new_last = " ".join(prev_words[:-2]), f"{moved} {chunks[-1]}"
            if fits(new_last) and fits(new_prev):
                chunks[-2], chunks[-1] = new_prev, new_last
    return chunks
