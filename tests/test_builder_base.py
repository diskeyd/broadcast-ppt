import io
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.util import Inches

from broadcast_ppt.builder.package import add_blank_slide, current_slide_ids, drop_section_list, remove_slides
from broadcast_ppt.builder.scale import scale_tree

NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"


def _deck_with_sections(n=3):
    prs = Presentation()
    for _ in range(n):
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes.add_textbox(Inches(1), Inches(1), Inches(3), Inches(1)).text_frame.text = "x"
    ids = current_slide_ids(prs)
    root = prs.part._element
    ext_xml = (
        f'<p:extLst xmlns:p="{NS_P}"><p:ext uri="{{521415D9-36F7-43E2-AB2F-B90AF26B5E84}}">'
        '<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
        '<p14:section name="s" id="{00000000-0000-0000-0000-000000000001}"><p14:sldIdLst>'
        + "".join(f'<p14:sldId id="{i}"/>' for i in ids)
        + "</p14:sldIdLst></p14:section></p14:sectionLst></p:ext></p:extLst>"
    )
    root.append(etree.fromstring(ext_xml))
    return prs, ids


def test_remove_template_slides_and_sections():
    prs, ids = _deck_with_sections(3)
    new = add_blank_slide(prs)
    assert len(new.shapes) == 0
    assert remove_slides(prs, ids) == 3
    assert drop_section_list(prs) is True
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    again = Presentation(buf)
    assert len(again.slides) == 1
    names = zipfile.ZipFile(io.BytesIO(buf.getvalue())).namelist()
    assert sum(1 for n in names if n.startswith("ppt/slides/slide") and n.endswith(".xml")) == 1
    xml = zipfile.ZipFile(io.BytesIO(buf.getvalue())).read("ppt/presentation.xml").decode()
    assert "sectionLst" not in xml


def test_scale_tree_positions_fonts_insets():
    xml = (
        '<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
        '<p:spPr><a:xfrm><a:off x="300" y="600"/><a:ext cx="9000" cy="3000"/></a:xfrm><a:ln w="12700"/></p:spPr>'
        '<p:txBody><a:bodyPr wrap="square"/><a:p><a:pPr marL="342900"><a:lnSpc><a:spcPts val="3200"/></a:lnSpc></a:pPr>'
        '<a:r><a:rPr lang="ko-KR" sz="3200"/><a:t>x</a:t></a:r></a:p>'
        '<a:extLst><a:ext uri="{X}"/></a:extLst></p:txBody></p:sp>'
    )
    el = etree.fromstring(xml)
    scale_tree(el, 4 / 3)
    a = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    assert el.find(f".//{a}off").get("x") == "400"
    assert el.find(f".//{a}xfrm/{a}ext").get("cx") == "12000"
    assert el.find(f".//{a}rPr").get("sz") == "4267"  # 32pt -> 42.67pt, as in the 220326 broadcast deck
    bp = el.find(f".//{a}bodyPr")
    assert bp.get("lIns") == "121920" and bp.get("tIns") == "60960"  # defaults materialised then scaled
    assert el.find(f".//{a}ln").get("w") == "16933"
    assert el.find(f".//{a}spcPts").get("val") == "4267"
    assert el.find(f".//{a}extLst/{a}ext").get("cx") is None  # extension entries untouched


def test_style_paragraphs_sets_size_and_colour():
    from pptx.dml.color import RGBColor

    from broadcast_ppt.builder.geometry import style_paragraphs
    from broadcast_ppt.ooxml import A

    tb = Presentation().slides.add_slide(Presentation().slide_layouts[6]).shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(2))
    tf = tb.text_frame
    tf.text = "제1문 질문?"
    tf.paragraphs[0].runs[0].font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    tf.add_paragraph().add_run().text = "답: 대답"
    style_paragraphs(tb._element, [{"color": "FF9300"}, {"sz": 28}])
    q, a = tf.paragraphs
    assert str(q.runs[0].font.color.rgb) == "FF9300" and len(q.runs[0]._r.find(f"{{{A}}}rPr").findall(f"{{{A}}}solidFill")) == 1
    assert a.runs[0].font.size.pt == 28 and q.runs[0].font.size is None


def test_hide_shape_keeps_the_element_but_draws_nothing():
    from pptx.dml.color import RGBColor
    from pptx.enum.shapes import MSO_SHAPE

    from broadcast_ppt.builder.geometry import hide_shape

    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    tag = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(1), Inches(2), Inches(0.5))
    tag.fill.solid()
    tag.fill.fore_color.rgb = RGBColor(0xFF, 0xD4, 0x79)
    n = len(slide.shapes)
    hide_shape(tag._element)
    tags = [c.tag.split("}")[1] for c in tag._element.spPr]
    assert "solidFill" not in tags and tags.index("prstGeom") < tags.index("noFill") < tags.index("ln") < tags.index("effectLst")
    assert tag._element.spPr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}ln")[0].tag.endswith("noFill")
    assert len(slide.shapes) == n


def _card_rows(offset, max_bottom):
    """A two-row card: a label with an English subtitle under it, laid out bottom-up."""
    from pptx.util import Emu

    from broadcast_ppt.builder.build import layout_stack
    from broadcast_ppt.builder.fill import shape_by_order
    from broadcast_ppt.builder.geometry import y_of
    from broadcast_ppt.mapping.textfit import WidthModel

    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[6])
    for top in (Inches(4), Inches(5)):
        s.shapes.add_textbox(Inches(1), top, Inches(6), Inches(1)).text_frame.text = "x"
    rows = [{"order": 1, "paragraph": None, "glyph": [10.0, 65.0], "offset_pt": 0.0},
            {"order": 2, "paragraph": None, "glyph": [20.0, 60.0], "offset_pt": offset}]
    entry = {"stack": {"rows": rows, "gap_pt": 0.5, "max_bottom_pt": max_bottom}}
    layout_stack(s, entry, WidthModel())
    return [Emu(y_of(shape_by_order(s, r["order"]))).pt for r in rows]


def test_stack_offset_moves_one_row_after_the_bottom_clamp():
    # the clamp binds here, which is what swallowed nudge_y_pt: the offset has to be applied after it
    base = _card_rows(0.0, 380)
    moved = _card_rows(3.0, 380)
    assert round(moved[1] - base[1], 2) == 3.0
    assert round(moved[0] - base[0], 2) == 0.0  # the Korean row above keeps its place
