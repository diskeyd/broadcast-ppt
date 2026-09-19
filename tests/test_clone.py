import io

from lxml import etree
from PIL import Image
from pptx import Presentation
from pptx.opc.constants import RELATIONSHIP_TYPE as RT
from pptx.opc.package import Part
from pptx.opc.packuri import PackURI
from pptx.util import Inches, Pt

from broadcast_ppt.builder.clone import _copy_part_into, copy_foreign_slide, duplicate_template_slide
from broadcast_ppt.builder.package import current_slide_ids, remove_slides
from broadcast_ppt.ooxml import R


def _png(color):
    buf = io.BytesIO()
    Image.new("RGB", (8, 8), color).save(buf, "PNG")
    buf.seek(0)
    return buf


def _deck(color, width=10, height=5.625):
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(width), Inches(height)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(_png(color), Inches(1), Inches(1), Inches(2), Inches(2))
    tb = s.shapes.add_textbox(Inches(4), Inches(1), Inches(4), Inches(1))
    tb.text_frame.text = "본문"
    tb.text_frame.paragraphs[0].runs[0].font.size = Pt(32)
    return prs


def test_duplicate_within_package_keeps_images():
    prs = _deck((255, 0, 0))
    src = prs.slides[0]
    dup, rep = duplicate_template_slide(prs, src)
    assert len(dup.shapes) == len(src.shapes)
    assert dup.shapes[0].image.blob == src.shapes[0].image.blob
    buf = io.BytesIO()
    prs.save(buf)
    buf.seek(0)
    assert len(Presentation(buf).slides) == 2


def test_copy_foreign_scales_and_does_not_collide_images():
    out = _deck((0, 255, 0), width=13.333, height=7.5)
    base_ids = current_slide_ids(out)
    a, b = _deck((255, 0, 0)), _deck((0, 0, 255))  # both packages name their image image1.png
    k = out.slide_width / a.slide_width
    s1, _ = copy_foreign_slide(out, a.slides[0], scale=k)
    s2, _ = copy_foreign_slide(out, b.slides[0], scale=k)
    remove_slides(out, base_ids)
    buf = io.BytesIO()
    out.save(buf)
    buf.seek(0)
    again = Presentation(buf)
    assert len(again.slides) == 2
    px = [Image.open(io.BytesIO(sl.shapes[0].image.blob)).getpixel((0, 0)) for sl in again.slides]
    assert px == [(255, 0, 0), (0, 0, 255)]
    tb = again.slides[0].shapes[1]
    assert abs(tb.left / again.slide_width - Inches(4) / Inches(10)) < 0.001
    assert tb.text_frame.paragraphs[0].runs[0].font.size.pt == 42.67


def test_copied_media_part_keeps_the_media_base_name():
    pkg = Presentation().part.package
    m4a = Part(PackURI("/ppt/media/media3.m4a"), "audio/mp4", pkg, b"x")  # 'a' and 'm' of the extension once ate into the base name
    mp4 = Part(PackURI("/ppt/media/media3.mp4"), "video/mp4", pkg, b"x")
    assert _copy_part_into(pkg, m4a).partname == "/ppt/media/media1.m4a"
    assert _copy_part_into(pkg, mp4).partname == "/ppt/media/media1.mp4"


def test_copy_foreign_keeps_svg_images():
    out = _deck((0, 255, 0), width=13.333, height=7.5)
    slide = _deck((255, 0, 0)).slides[0]
    svg = Part(PackURI("/ppt/media/image9.svg"), "image/svg+xml", slide.part.package, b"<svg xmlns='http://www.w3.org/2000/svg'/>")
    blip = slide.shapes[0]._element.xpath(".//a:blip")[0]  # PowerPoint icons: PNG blip plus an svgBlip extension, which PIL cannot read
    etree.SubElement(blip, "{http://schemas.microsoft.com/office/drawing/2016/SVG/main}svgBlip").set(f"{{{R}}}embed", slide.part.relate_to(svg, RT.IMAGE))
    copy_foreign_slide(out, slide, scale=4 / 3)
    buf = io.BytesIO()
    out.save(buf)
    buf.seek(0)
    svgs = [p for p in Presentation(buf).part.package.iter_parts() if p.content_type == "image/svg+xml"]
    assert [p.blob for p in svgs] == [svg.blob]
