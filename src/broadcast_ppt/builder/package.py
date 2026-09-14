"""Output package handling: open the template as the base document, add blank slides,
remove the template's own slides and the section list that would otherwise dangle."""

from __future__ import annotations

from pathlib import Path

from pptx import Presentation

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
SECTION_EXT_URI = "{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"
BLANK_LAYOUT_NAMES = ("빈 화면", "Blank")


def open_base(template: Path) -> Presentation:
    return Presentation(str(template))


def blank_layout(prs):
    for layout in prs.slide_layouts:
        if layout.name in BLANK_LAYOUT_NAMES:
            return layout
    return prs.slide_layouts[6]


def add_blank_slide(prs):
    """A slide on the blank layout; python-pptx does not clone layout placeholders, so it has no shapes."""
    slide = prs.slides.add_slide(blank_layout(prs))
    for sh in list(slide.shapes):  # defensive: some layouts do give placeholders
        sh._element.getparent().remove(sh._element)
    return slide


def slide_id_list(prs):
    # python-pptx has no public API to remove slides; all direct sldIdLst access lives in this module.
    return prs.slides._sldIdLst


def current_slide_ids(prs) -> list[int]:
    return [int(s.get("id")) for s in slide_id_list(prs)]


def remove_slides(prs, slide_ids: list[int]) -> int:
    """Drop slides by sldId (relationship + list entry). Orphan parts are not written on save."""
    wanted = set(slide_ids)
    lst = slide_id_list(prs)
    removed = 0
    for sld in list(lst):
        if int(sld.get("id")) in wanted:
            prs.part.drop_rel(sld.get(f"{{{NS_R}}}id"))
            lst.remove(sld)
            removed += 1
    return removed


def drop_section_list(prs) -> bool:
    """Remove the PowerPoint 2010 section list; its slide references break when slides are removed."""
    root = prs.part._element
    ext_lst = root.find(f"{{{NS_P}}}extLst")
    if ext_lst is None:
        return False
    dropped = False
    for ext in list(ext_lst):
        if ext.get("uri") == SECTION_EXT_URI:
            ext_lst.remove(ext)
            dropped = True
    if len(ext_lst) == 0:
        root.remove(ext_lst)
    return dropped
