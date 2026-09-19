"""Slide cloning at the XML level (python-pptx has no slide-copy API).

- duplicate_template_slide: copy a slide of the output package itself (the opened template) onto a new
  blank slide; relationships can point at the same parts.
- copy_foreign_slide: copy a slide from another presentation; image/media parts are re-created in the
  output package under fresh part names, geometry and fonts are scaled to the output slide size,
  placeholders are turned into plain shapes.
"""

from __future__ import annotations

import copy
import io
import posixpath
from dataclasses import dataclass, field

from lxml import etree
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI

from .package import add_blank_slide
from .scale import scale_tree

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
MEDIA_RELTYPES = {RT.MEDIA, RT.VIDEO, RT.AUDIO, "http://schemas.microsoft.com/office/2007/relationships/media"}


@dataclass
class CloneReport:
    flags: list[str] = field(default_factory=list)

    def flag(self, f: str) -> None:
        if f not in self.flags:
            self.flags.append(f)


def _sp_tree_children(slide):
    tree = slide.shapes._spTree
    for child in tree:
        name = etree.QName(child).localname
        if name in ("nvGrpSpPr", "grpSpPr"):
            continue
        yield child


def _renumber_ids(slide) -> None:
    next_id = 2
    for el in slide.shapes._spTree.iter(f"{{{P}}}cNvPr"):
        el.set("id", str(next_id))
        next_id += 1


def _copy_part_into(dst_package, src_part) -> Part:
    ext = posixpath.splitext(str(src_part.partname))[1] or ".bin"
    folder = posixpath.dirname(str(src_part.partname))
    base = posixpath.basename(str(src_part.partname)).rstrip("0123456789." + ext.lstrip(".")) or "media"
    tmpl = f"{folder}/{base}%d{ext}"
    partname = dst_package.next_partname(tmpl)
    return Part(PackURI(partname), src_part.content_type, dst_package, src_part.blob)


def _rewrite_rels(elements, src_slide, dst_slide, same_package: bool, report: CloneReport) -> None:
    src_rels = src_slide.part.rels
    dst_part = dst_slide.part
    cache: dict[str, str] = {}
    for root in elements:
        for el in list(root.iter()):
            if not isinstance(el.tag, str):
                continue
            for attr, rid in list(el.attrib.items()):
                if not attr.startswith(f"{{{R}}}"):
                    continue
                if rid in cache:
                    el.set(attr, cache[rid])
                    continue
                rel = src_rels.get(rid) if hasattr(src_rels, "get") else None
                if rel is None:
                    del el.attrib[attr]
                    report.flag("missing_relationship")
                    continue
                if rel.is_external:
                    new = dst_part.relate_to(rel.target_ref, rel.reltype, is_external=True)
                elif rel.reltype == RT.SLIDE:
                    parent = el.getparent()  # internal jump to another slide of the source deck
                    if parent is not None:
                        parent.remove(el)
                    report.flag("dropped_slide_link")
                    continue
                elif same_package:
                    new = dst_part.relate_to(rel.target_part, rel.reltype)
                elif rel.reltype == RT.IMAGE:
                    _, new = dst_part.get_or_add_image_part(io.BytesIO(rel.target_part.blob))
                elif rel.reltype in MEDIA_RELTYPES:
                    new_part = _copy_part_into(dst_part.package, rel.target_part)
                    new = dst_part.relate_to(new_part, rel.reltype)
                else:
                    report.flag("unsupported_relationship")
                    del el.attrib[attr]
                    continue
                cache[rid] = new
                el.set(attr, new)


def _materialise_placeholder(sp, src_slide, report: CloneReport) -> None:
    ph = sp.find(f"{{{P}}}nvSpPr/{{{P}}}nvPr/{{{P}}}ph")
    if ph is None:
        return
    ph_type, ph_idx = ph.get("type", "body"), ph.get("idx")
    sp_pr = sp.find(f"{{{P}}}spPr")
    if sp_pr is not None and sp_pr.find(f"{{{A}}}xfrm") is None:
        for holder in (src_slide.slide_layout, src_slide.slide_layout.slide_master):
            for cand in holder.placeholders:
                f = cand.placeholder_format
                if (ph_idx is not None and str(f.idx) == ph_idx) or (f.type is not None and f.type.name.lower().startswith(ph_type.lower())):
                    xfrm = cand._element.find(f"{{{P}}}spPr/{{{A}}}xfrm")
                    if xfrm is not None:
                        sp_pr.insert(0, copy.deepcopy(xfrm))
                    break
            if sp_pr.find(f"{{{A}}}xfrm") is not None:
                break
    # inherited run size from the source master's title/body style, for runs without an explicit size
    master = src_slide.slide_layout.slide_master._element
    style = master.find(f"{{{P}}}txStyles/{{{P}}}{'titleStyle' if ph_type in ('title', 'ctrTitle') else 'bodyStyle'}")
    sz = None
    if style is not None:
        d = style.find(f"{{{A}}}lvl1pPr/{{{A}}}defRPr")
        sz = d.get("sz") if d is not None else None
    if sz:
        for r in sp.iter(f"{{{A}}}r"):
            rpr = r.find(f"{{{A}}}rPr")
            if rpr is None:
                rpr = etree.SubElement(r, f"{{{A}}}rPr")
                r.insert(0, rpr)
            if rpr.get("sz") is None:
                rpr.set("sz", sz)
    ph.getparent().remove(ph)
    report.flag("placeholder_materialised")


def _resolved_background(slide):
    for holder in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        bg = holder._element.find(f"{{{P}}}cSld/{{{P}}}bg")
        if bg is not None:
            return bg, holder
    return None, None


def _set_background(dst_slide, bg) -> None:
    csld = dst_slide._element.find(f"{{{P}}}cSld")
    old = csld.find(f"{{{P}}}bg")
    if old is not None:
        csld.remove(old)
    csld.insert(0, bg)


def duplicate_template_slide(prs, tpl_slide) -> tuple[object, CloneReport]:
    report = CloneReport()
    dst = add_blank_slide(prs)
    tree = dst.shapes._spTree
    copies = [copy.deepcopy(ch) for ch in _sp_tree_children(tpl_slide)]
    for c in copies:
        tree.append(c)
    bg = tpl_slide._element.find(f"{{{P}}}cSld/{{{P}}}bg")
    bg_copy = copy.deepcopy(bg) if bg is not None else None
    _rewrite_rels(copies + ([bg_copy] if bg_copy is not None else []), tpl_slide, dst, same_package=True, report=report)
    if bg_copy is not None:
        _set_background(dst, bg_copy)
    _renumber_ids(dst)
    return dst, report


def copy_foreign_slide(prs, src_slide, *, scale: float, keep_bg: bool = True) -> tuple[object, CloneReport]:
    report = CloneReport()
    dst = add_blank_slide(prs)
    tree = dst.shapes._spTree
    copies = []
    for ch in _sp_tree_children(src_slide):
        c = copy.deepcopy(ch)
        for sp in [c] + list(c.iter(f"{{{P}}}sp")):
            if etree.QName(sp).localname == "sp":
                _materialise_placeholder(sp, src_slide, report)
        scale_tree(c, scale)
        copies.append(c)
        tree.append(c)
    elements = list(copies)
    if keep_bg:
        bg, holder = _resolved_background(src_slide)
        if bg is not None:
            bg_copy = copy.deepcopy(bg)
            scale_tree(bg_copy, scale)
            # background image relationships live on the part that owns the bg element
            _rewrite_rels([bg_copy], holder, dst, same_package=False, report=report)
            _set_background(dst, bg_copy)
    _rewrite_rels(elements, src_slide, dst, same_package=False, report=report)
    if src_slide._element.find(f"{{{P}}}timing") is not None:
        report.flag("animation_dropped")
    _renumber_ids(dst)
    return dst, report
