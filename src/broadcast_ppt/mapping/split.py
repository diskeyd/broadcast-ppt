"""Greedy, boundary-aware packing of text units into chunks that fit a template box."""

from __future__ import annotations

import re
from collections.abc import Callable

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


def pack(units: list[str], fits: Callable[[str], bool], sep: str = " ") -> list[str]:
    """Group units into chunks that fit, keeping every unit whole and in order.

    Chunks are carried as lists of units, not as joined strings, so nothing downstream can cut
    through a unit: a verse split across two caption pages reads as a broken sentence on air.
    `sep` joins them at the end — "\\n" puts each unit on its own line."""
    chunks: list[list[str]] = []
    cur: list[str] = []
    for unit in (u.strip() for u in units):
        if not unit:
            continue
        if fits(sep.join(cur + [unit])):
            cur.append(unit)
            continue
        if cur:
            chunks.append(cur)
            cur = []
        pieces = _pieces(unit, fits)  # one unit too long for a chunk of its own: the only place a
        chunks.extend([p] for p in pieces[:-1])  # unit is broken, and then at a sentence or a word
        cur = [pieces[-1]] if pieces else []
    if cur:
        chunks.append(cur)
    # avoid a tiny last chunk by moving one whole unit forward, when both still fit
    if len(chunks) >= 2 and len(sep.join(chunks[-1])) < 0.35 * max(len(sep.join(c)) for c in chunks[:-1]):
        if len(chunks[-2]) > 1:
            new_prev, new_last = chunks[-2][:-1], chunks[-2][-1:] + chunks[-1]
            if fits(sep.join(new_last)) and fits(sep.join(new_prev)):
                chunks[-2], chunks[-1] = new_prev, new_last
    return [sep.join(c) for c in chunks]
