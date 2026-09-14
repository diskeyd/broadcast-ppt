import io

from PIL import Image
from pptx import Presentation
from pptx.util import Inches, Pt

from broadcast_ppt.builder.clone import copy_foreign_slide, duplicate_template_slide
from broadcast_ppt.builder.package import current_slide_ids, remove_slides


def _png(color):
    buf = io.BytesIO()
    Image.new("RGB", (8, 8), color).save(buf, "PNG")
    buf.seek(0)
    return buf


def _deck(color, width=Inches(10), height=Inches(5.625)):
    prs = Presentation()
    prs.slide_width, prs.slide_height = width, height
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
    buf = io.BytesIO(); prs.save(buf); buf.seek(0)
    assert len(Presentation(buf).slides) == 2


def test_copy_foreign_scales_and_does_not_collide_images():
    out = _deck((0, 255, 0), width=Inches(13.333), height=Inches(7.5))
    base_ids = current_slide_ids(out)
    a, b = _deck((255, 0, 0)), _deck((0, 0, 255))  # both packages name their image image1.png
    k = out.slide_width / a.slide_width
    s1, _ = copy_foreign_slide(out, a.slides[0], scale=k)
    s2, _ = copy_foreign_slide(out, b.slides[0], scale=k)
    remove_slides(out, base_ids)
    buf = io.BytesIO(); out.save(buf); buf.seek(0)
    again = Presentation(buf)
    assert len(again.slides) == 2
    px = [Image.open(io.BytesIO(sl.shapes[0].image.blob)).getpixel((0, 0)) for sl in again.slides]
    assert px == [(255, 0, 0), (0, 0, 255)]
    tb = again.slides[0].shapes[1]
    assert abs(tb.left / again.slide_width - Inches(4) / Inches(10)) < 0.001
    assert tb.text_frame.paragraphs[0].runs[0].font.size.pt == 42.67
