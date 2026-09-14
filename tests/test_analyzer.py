import json
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.util import Inches

from broadcast_ppt.analyzer.core import analyze_file, classify_rgb, date_token, family_of, keywords_in
from broadcast_ppt.analyzer.summary import summarize
from broadcast_ppt.cli import main


def build_deck(path: Path) -> Path:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    blank = prs.slide_layouts[6]

    s1 = prs.slides.add_slide(blank)
    s1.background.fill.solid()
    s1.background.fill.fore_color.rgb = RGBColor(0x00, 0xFF, 0x00)
    box = s1.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1))
    box.text_frame.text = "성경봉독"
    grp = s1.shapes.add_group_shape()
    inner = grp.shapes.add_textbox(Inches(1), Inches(3), Inches(4), Inches(1))
    inner.text_frame.text = "시편 62:1-8"

    s2 = prs.slides.add_slide(blank)
    rect = s2.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, prs.slide_width, prs.slide_height)
    rect.fill.solid()
    rect.fill.fore_color.rgb = RGBColor(0x00, 0x00, 0xFF)
    tbl = s2.shapes.add_table(2, 2, Inches(1), Inches(1), Inches(6), Inches(2)).table
    tbl.cell(0, 0).text = "대표기도"
    tbl.cell(1, 1).text = "찬송 305장"

    s3 = prs.slides.add_slide(blank)
    s3.shapes.add_textbox(Inches(1), Inches(1), Inches(4), Inches(1)).text_frame.text = "광고"
    prs.save(path)
    return path


def test_helpers():
    assert classify_rgb("00FF00") == "green"
    assert classify_rgb("00B140") == "green"
    assert classify_rgb("0000FF") == "blue"
    assert classify_rgb("FFFFFF") is None
    assert classify_rgb("333333") is None
    assert classify_rgb("002060") is None  # dark navy text box, not a key colour
    assert date_token("주일오전예배ppt(230702)-맥추감사주일.pptx") == "230702"
    assert date_token("송구영신예배ppt(2025).pptx") == "2025"
    assert date_token("합본 20260517.pptx") == "260517"
    assert family_of("합본 20260517.pptx")[0] == "합본"
    assert family_of("(260805) 바이블인카운터 1 ppt.pptx")[0] == "바이블인카운터"
    assert family_of("주일오전예배ppt(230702)-맥추감사주일.pptx") == ("주일오전예배", "맥추감사주일")
    assert family_of("(231129) 세퍼드라이프2권(2B) ppt.pptx")[0] == "세퍼드라이프"
    assert family_of("24.상반기.pptx")[0] == "기타"
    assert keywords_in("대표 기도") == ["대표기도", "기도"]


def test_analyze_synthetic(tmp_path):
    deck = build_deck(tmp_path / "주일오전예배ppt(240101).pptx")
    doc = analyze_file(deck, tmp_path)
    assert doc["ratio"] == "16:9"
    assert doc["slide_count"] == 3
    s1, s2, s3 = doc["slides"]

    assert s1["background"]["type"] == "solid" and s1["background"]["rgb"] == "00FF00"
    assert any(c["key"] == "green" and c["area_frac"] == 1.0 for c in s1["chroma"])
    assert "시편 62:1-8" in s1["text"]  # text inside a group is reached
    assert s1["scripture_ref"] and "성경봉독" in s1["keywords"]
    grp = next(s for s in s1["shapes"] if s["kind"] == "GROUP")
    assert grp["children"][0]["depth"] == 1

    assert any(c["key"] == "blue" and c["area_frac"] >= 0.99 for c in s2["chroma"])
    assert "대표기도" in s2["keywords"] and s2["hymn_no"]  # table cells are read

    assert s3["chroma"] == [] and s3["keywords"] == ["광고"]
    json.dumps(doc, ensure_ascii=False)


def test_cli_and_summary(tmp_path):
    src = tmp_path / "in"
    (src / "미편집").mkdir(parents=True)
    build_deck(src / "주일오전예배ppt(240101).pptx")
    build_deck(src / "미편집" / "주일오전예배ppt(240101).pptx")
    (src / "broken.pptx").write_bytes(b"not a zip")
    out = tmp_path / "out"
    assert main(["analyze", str(src), "--out", str(out), "--workers", "1", "--scope", ""]) == 0
    md = (out / "clusters.md").read_text(encoding="utf-8")
    assert "paired: 1" in md
    assert "broken.pptx" in md
    assert "시편" not in md  # no slide text leaks into the shareable report
    rows = (out / "summary.csv").read_text(encoding="utf-8-sig").splitlines()
    assert len(rows) == 4


def test_deck_with_corrupt_media_is_still_analyzed(tmp_path):
    import zipfile
    good = build_deck(tmp_path / "good.pptx")
    bad = tmp_path / "bad.pptx"
    with zipfile.ZipFile(good) as src, zipfile.ZipFile(bad, "w", zipfile.ZIP_DEFLATED) as dst:
        for item in src.infolist():
            dst.writestr(item, src.read(item.filename))
        dst.writestr("ppt/media/broken.jpeg", b"x" * 4096)
    raw = bytearray(bad.read_bytes())
    info = next(i for i in zipfile.ZipFile(bad).infolist() if i.filename == "ppt/media/broken.jpeg")
    start = info.header_offset + 30 + len(info.filename.encode()) + len(info.extra)
    raw[start:start + 8] = b"\xff" * 8  # corrupt the deflate stream of that member only
    bad.write_bytes(bytes(raw))
    doc = analyze_file(bad, tmp_path)
    assert "error" not in doc, doc.get("error")
    assert doc["damaged_parts"] == ["ppt/media/broken.jpeg"]
    assert doc["slide_count"] == 3


def test_scope_filter():
    from broadcast_ppt.analyzer.summary import in_scope
    scope = {"families": ["주일오전예배"], "name_contains": ["부활주일"],
             "exclude": {"folders": ["새 폴더"], "families": ["합본"], "name_startswith": ["24.상반기"]}}
    assert in_scope({"name": "주일오전예배ppt(240101).pptx", "family": "주일오전예배", "folder": ""}, scope)
    assert in_scope({"name": "부활주일 ppt(220417).pptx", "family": "기타", "folder": ""}, scope)
    assert not in_scope({"name": "합본 20260517.pptx", "family": "합본", "folder": "미편집"}, scope)
    assert not in_scope({"name": "21.상반기.pptx", "family": "기타", "folder": "새 폴더"}, scope)
    assert not in_scope({"name": "24.상반기.pptx", "family": "기타", "folder": ""}, scope)
    assert not in_scope({"name": "수요오후예배ppt(230125).pptx", "family": "수요오후예배", "folder": ""}, scope)
    assert in_scope({"name": "x.pptx", "family": "기타", "folder": ""}, None)
