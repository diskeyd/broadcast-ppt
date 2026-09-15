"""Replace text inside template text boxes while keeping their formatting.

A box's runs are grouped into formatting segments exactly like the catalog does (same size, colour,
font, bold, italic; whitespace-only runs join the segment before them). Each slot value goes into the
first run of its segment; the other runs of that segment are removed, so the first run's rPr — the
template's formatting — is what the text is drawn with.
"""

from __future__ import annotations

import copy
import hashlib
import unicodedata
from dataclasses import dataclass, field

from lxml import etree

from ..ooxml import A, run_segments

NS = {"a": A}


@dataclass
class FillResult:
    filled: dict[str, str] = field(default_factory=dict)
    missing: list[str] = field(default_factory=list)
    flags: list[str] = field(default_factory=list)


paragraph_segments = run_segments


def flat_segments(txBody) -> list[dict]:
    """Segments in reading order, filtered the same way as catalog slot maps (text or explicit size)."""
    out = []
    for pi, p in enumerate(txBody.findall("a:p", NS)):
        for seg in paragraph_segments(p):
            if seg["text"].strip() or seg["key"][0]:
                seg["paragraph"] = pi
                out.append(seg)
    return out


def _set_run_text(r, text: str) -> None:
    t = r.find("a:t", NS)
    if t is None:
        t = etree.SubElement(r, f"{{{A}}}t")
    t.text = text


def segment_sha1(text: str) -> str:
    return hashlib.sha1(unicodedata.normalize("NFC", text).encode()).hexdigest()


def _fill_single(txBody, value: str, result: FillResult) -> None:
    paras = txBody.findall("a:p", NS)
    runs = [r for p in paras for r in p.findall("a:r", NS)]
    if not runs:
        result.flags.append("no_runs")
        return
    base = next((r for r in runs if r.find("a:rPr/a:solidFill", NS) is None and r.findtext("a:t", "", NS).strip()), runs[0])
    first_p = base.getparent()
    lines = value.split("\n")
    for p in paras:
        if p is not first_p:
            txBody.remove(p)
    for r in first_p.findall("a:r", NS):
        if r is not base:
            first_p.remove(r)
    for br in first_p.findall("a:br", NS):
        first_p.remove(br)
    _set_run_text(base, lines[0])
    anchor = first_p
    for line in lines[1:]:
        np_ = copy.deepcopy(first_p)
        _set_run_text(np_.find("a:r", NS), line)
        anchor.addnext(np_)
        anchor = np_


def fill_shape(sp_el, spec: dict, values: dict[str, str]) -> FillResult:
    """Fill one shape element according to its catalog spec (slot_map, mode, single_segment)."""
    result = FillResult()
    mode = spec.get("mode", "fill")
    slot_map = spec.get("slot_map") or []
    if mode == "fixed" or not slot_map:
        return result
    txBody = sp_el.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}txBody")
    if txBody is None:
        result.flags.append("no_text_body")
        return result
    if spec.get("single_segment"):
        slot = slot_map[0]
        value = values.get(slot)
        if value is None:
            if mode == "fill":
                result.missing.append(slot)
            return result
        _fill_single(txBody, value, result)
        result.filled[slot] = value
        return result
    segs = flat_segments(txBody)
    if len(segs) != len(slot_map):
        result.flags.append(f"segment_mismatch:{len(segs)}!={len(slot_map)}")
    for seg, slot in zip(segs, slot_map):
        if slot is None:
            continue
        value = values.get(slot)
        if value is None:
            if mode == "fill":
                result.missing.append(slot)
            continue
        # write into the first run that carries text (its rPr is the segment's formatting)
        target = next((r for r in seg["runs"] if r.findtext("a:t", "", NS).strip()), seg["runs"][0])
        _set_run_text(target, value.replace("\n", " "))
        for r in seg["runs"]:
            if r is not target:
                r.getparent().remove(r)
        result.filled[slot] = value
    return result


def shape_by_order(slide, order: int):
    """The shape element (p:sp / p:pic / p:grpSp …) owning the cNvPr at document-order position `order`."""
    P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    cnv = list(slide.shapes._spTree.iter(f"{{{P}}}cNvPr"))
    el = cnv[order]
    # cNvPr -> nvSpPr/nvPicPr/nvGrpSpPr -> shape element
    return el.getparent().getparent()
