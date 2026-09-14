from lxml import etree

from broadcast_ppt.builder.fill import fill_shape, flat_segments

P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"


def _sp(paragraphs_xml: str):
    return etree.fromstring(
        f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:txBody><a:bodyPr wrap="none"/>{paragraphs_xml}</p:txBody></p:sp>'
    )


YELLOW = '<a:rPr sz="5400"><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></a:rPr>'
WHITE = '<a:rPr sz="3600"/>'


def test_fill_two_segments_keeps_first_run_formatting():
    sp = _sp(f'<a:p><a:r>{YELLOW}<a:t>기도 </a:t></a:r><a:r>{WHITE}<a:t>– </a:t></a:r>'
             f'<a:r>{WHITE}<a:t>샘플 </a:t></a:r><a:r><a:rPr sz="3600"><a:latin typeface="X"/></a:rPr><a:t> </a:t></a:r>'
             f'<a:r>{WHITE}<a:t>직분</a:t></a:r></a:p>')
    before = etree.tostring(sp.findall(f".//{{{A}}}r")[1].find(f"{{{A}}}rPr"))
    res = fill_shape(sp, {"slot_map": ["label", "person"]}, {"label": "기도 ", "person": "– 새이름 집사"})
    runs = sp.findall(f".//{{{A}}}r")
    assert [r.findtext(f"{{{A}}}t") for r in runs] == ["기도 ", "– 새이름 집사"]
    assert etree.tostring(runs[1].find(f"{{{A}}}rPr")) == before
    assert res.filled == {"label": "기도 ", "person": "– 새이름 집사"} and not res.flags


def test_leading_space_run_joins_the_next_segment():
    pad = '<a:rPr sz="1867"><a:latin typeface="R"/></a:rPr>'
    tag = '<a:rPr sz="1700"><a:latin typeface="R"/></a:rPr>'
    sp = _sp(f'<a:p><a:r>{pad}<a:t> </a:t></a:r><a:r>{tag}<a:t>샘플 </a:t></a:r><a:r>{tag}<a:t>18:21</a:t></a:r></a:p>')
    assert len(flat_segments(sp.find(f".//{{{P}}}txBody"))) == 1
    fill_shape(sp, {"slot_map": ["ref"]}, {"ref": "사무엘상26:24"})
    runs = sp.findall(f".//{{{A}}}r")
    assert [r.findtext(f"{{{A}}}t") for r in runs] == ["사무엘상26:24"]
    assert runs[0].find(f"{{{A}}}rPr").get("sz") == "1700"


def test_group_shift_keeps_child_frame_and_width_keeps_center():
    from broadcast_ppt.builder.geometry import set_width, shift
    grp = etree.fromstring(
        f'<p:grpSp xmlns:p="{P}" xmlns:a="{A}"><p:grpSpPr><a:xfrm><a:off x="10" y="100"/><a:ext cx="50" cy="40"/>'
        f'<a:chOff x="10" y="100"/><a:chExt cx="50" cy="40"/></a:xfrm></p:grpSpPr></p:grpSp>')
    shift(grp, dy=5, dcy=-5)
    xf = grp.find(f"{{{P}}}grpSpPr/{{{A}}}xfrm")
    assert [xf.find(f"{{{A}}}{t}").get(k) for t, k in (("off", "y"), ("ext", "cy"), ("chOff", "y"), ("chExt", "cy"))] == ["105", "35", "105", "35"]
    sp = etree.fromstring(f'<p:sp xmlns:p="{P}" xmlns:a="{A}"><p:spPr><a:xfrm><a:off x="100" y="0"/><a:ext cx="200" cy="10"/></a:xfrm></p:spPr></p:sp>')
    set_width(sp, 300, keep_center=True)
    assert sp.find(f".//{{{A}}}off").get("x") == "50" and sp.find(f".//{{{A}}}ext").get("cx") == "300"


def test_optional_slot_keeps_template_text_and_fill_reports_missing():
    sp = _sp(f'<a:p><a:r>{WHITE}<a:t>Prayer of worship</a:t></a:r></a:p>')
    res = fill_shape(sp, {"slot_map": ["subtitle_en"], "mode": "optional"}, {})
    assert sp.findtext(f".//{{{A}}}t") == "Prayer of worship" and res.missing == []
    res = fill_shape(sp, {"slot_map": ["subtitle_en"], "mode": "fill"}, {})
    assert res.missing == ["subtitle_en"]


def test_single_segment_body_drops_emphasis_and_extra_paragraphs():
    orange = '<a:rPr sz="2600"><a:solidFill><a:srgbClr val="FF9300"/></a:solidFill></a:rPr>'
    body = '<a:rPr sz="2600"><a:latin typeface="나눔스퀘어OTF Bold"/></a:rPr>'
    sp = _sp(f'<a:p><a:r>{body}<a:t>20 </a:t></a:r><a:r>{orange}<a:t>강조</a:t></a:r><a:r>{body}<a:t>나머지</a:t></a:r></a:p>'
             f'<a:p><a:r>{body}<a:t>둘째 문단</a:t></a:r></a:p>')
    fill_shape(sp, {"slot_map": ["body"], "single_segment": True}, {"body": "1 새 본문\n2 다음 절"})
    paras = sp.findall(f".//{{{A}}}p")
    assert [p.findtext(f"{{{A}}}r/{{{A}}}t") for p in paras] == ["1 새 본문", "2 다음 절"]
    assert all(p.find(f"{{{A}}}r/{{{A}}}rPr/{{{A}}}solidFill") is None for p in paras)
    assert len(flat_segments(sp.find(f".//{{{P}}}txBody"))) == 2
