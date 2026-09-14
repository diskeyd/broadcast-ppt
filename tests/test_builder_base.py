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
