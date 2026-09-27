"""Hymn library index, song matching and run grouping."""

from __future__ import annotations

from types import SimpleNamespace

import pytest
from pptx import Presentation
from pptx.util import Emu, Pt

from broadcast_ppt.mapping import hymns
from broadcast_ppt.mapping.plan import _hymn_runs
from broadcast_ppt.rules import kinds as K


def _song_deck(path, songs):
    """A library deck: each song opens with a slide carrying a 20pt label under the 44pt lyric."""
    prs = Presentation()
    prs.slide_width, prs.slide_height = Emu(12192000), Emu(6858000)
    blank = prs.slide_layouts[6]
    for title, lines in songs:
        for i, line in enumerate(lines):
            slide = prs.slides.add_slide(blank)
            box = slide.shapes.add_textbox(Emu(1000000), Emu(5900000), Emu(10000000), Emu(700000))
            box.text_frame.paragraphs[0].add_run().text = line
            box.text_frame.paragraphs[0].runs[0].font.size = Pt(44)
            if i == 0:
                tag = slide.shapes.add_textbox(Emu(1000000), Emu(5500000), Emu(10000000), Emu(400000))
                tag.text_frame.paragraphs[0].add_run().text = title
                tag.text_frame.paragraphs[0].runs[0].font.size = Pt(20)
    prs.save(str(path))
    return path


SONG = ("108장 그 어린 주 예수", ["그 어린 주 예수 눕히실 자리 없어", "그 귀하신 몸이 구유에 있네", "저 하늘 별들이 반짝이는데"])
OTHER = ("새 노래 한 곡", ["한 줄씩 불러 봅니다", "함께 걸어갑니다 매일", "온 맘 다해 노래합니다"])


def test_library_reads_songs_and_skips_legacy_ppt(tmp_path):
    _song_deck(tmp_path / "108_그 어린 주 예수.pptx", [SONG])
    (tmp_path / "옛날파일.ppt").write_bytes(b"not a pptx")
    songs, skipped = hymns.build_library(tmp_path)
    assert [s["title"] for s in songs] == ["그 어린 주 예수"]      # the filename names the song, not the label box
    assert songs[0]["number"] == 108 and songs[0]["label"] == "108장 그 어린 주 예수"
    assert songs[0]["lines"] == SONG[1]
    assert len(skipped) == 1 and "옛날파일.ppt" in skipped[0]


def test_weekly_deck_splits_at_each_label(tmp_path):
    _song_deck(tmp_path / "0115 두 곡.pptx", [SONG, OTHER])
    songs, _ = hymns.build_library(tmp_path)
    assert [s["title"] for s in songs] == [SONG[0], OTHER[0]]     # several songs: each label box names its own
    assert [len(s["lines"]) for s in songs] == [3, 3]


def test_a_name_after_the_song_is_not_a_lyric(tmp_path):
    """A weekly deck runs songs together with the service: the slide naming who prays follows the song
    in the same 44pt box. It must end the song, never be indexed — a name must not reach the screen."""
    name_slide = ("찬양", ["주를 찬양합니다 온 맘 다해", "높이 들어 올립니다", "영광 받으소서 주님", "홍길동 권사"])
    _song_deck(tmp_path / "0924 한 곡.pptx", [name_slide, OTHER])
    songs, _ = hymns.build_library(tmp_path)
    assert [s["title"] for s in songs] == ["찬양", OTHER[0]]
    assert songs[0]["lines"] == name_slide[1][:3]           # the song stops at the name
    assert not any("홍길동" in line for s in songs for line in s["lines"])


def test_match_needs_two_lines_or_a_title(tmp_path):
    _song_deck(tmp_path / "108_그 어린 주 예수.pptx", [SONG])
    songs, _ = hymns.build_library(tmp_path)
    assert hymns.match(songs, ["그 어린 주 예수 눕히실 자리 없어", "그 귀하신 몸이 구유에 있네"]) == (songs[0]["id"], 0, 1)
    assert hymns.match(songs, ["그 귀하신 몸이 구유에 있네"]) is None   # one line on its own is not enough
    # a page that only shows the title line still resolves, because the title reads the same
    assert hymns.match(songs, ["그 어린 주 예수", "저 하늘 별들이 반짝이는데"]) is not None
    assert hymns.match(songs, ["전혀 다른 가사입니다 여기", "이것도 다른 노래"]) is None


def test_match_run_uses_the_cache_without_ocr(tmp_path, monkeypatch):
    _song_deck(tmp_path / "108_그 어린 주 예수.pptx", [SONG])
    songs, _ = hymns.build_library(tmp_path)
    deck = tmp_path / "source.pptx"
    prs = Presentation()
    png = tmp_path / "score.png"
    png.write_bytes(bytes.fromhex(  # 1x1 PNG
        "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4890000000a49444154"
        "789c63000100000500010d0a2db40000000049454e44ae426082"))
    for _ in range(3):
        slide = prs.slides.add_slide(prs.slide_layouts[6])
        slide.shapes.add_picture(str(png), Emu(0), Emu(0), Emu(1000000), Emu(1000000))
    prs.save(str(deck))
    import hashlib
    sha = hashlib.sha1(png.read_bytes()).hexdigest()
    index = {"songs": songs, "images": {sha: {"song": songs[0]["id"], "line": 0, "last": 1}}}
    monkeypatch.setattr(hymns, "ocr", lambda path: pytest.fail("cached pages must not be read again"))
    res = hymns.match_run(deck, [1, 2, 3], index, tmp_path / "work")
    assert res.match is not None
    assert (res.match.start, res.match.end, res.match.label) == (0, 1, "108장 그 어린 주 예수")


def _slide(index, pictures):
    return {"index": index, "kinds": {"PICTURE": pictures}}


def _pic(n):
    return SimpleNamespace(kind=K.PICTURE_ONLY), _slide(n, 1)


def test_runs_need_three_score_slides_in_a_row():
    slides, decisions = [], []
    for n in range(1, 10):
        d, s = _pic(n)
        if n in (4, 8):  # a caption between the runs
            d, s = SimpleNamespace(kind=K.CAPTION_SENTENCE), _slide(n, 0)
        slides.append(s)
        decisions.append(d)
    # 1-3 is a run, 5-7 is a run, the single slide 9 is not
    assert _hymn_runs(slides, decisions, set()) == [[1, 2, 3], [5, 6, 7]]


def test_a_slide_with_an_override_keeps_its_run_out():
    slides, decisions = zip(*[(s, d) for d, s in (_pic(n) for n in range(1, 4))])
    assert _hymn_runs(list(slides), list(decisions), {2}) == []


def test_clean_lyric_drops_repeat_signs_and_held_notes():
    assert hymns.clean_lyric(": 새 노래로 함께 부르며") == "새 노래로 함께 부르며"
    assert hymns.clean_lyric("우리가 노-래하는 오늘") == "우리가 노래하는 오늘"
    assert hymns.clean_lyric("반가워요 - 그길을 비춰준 아침에-") == "반가워요 그길을 비춰준 아침에"
    assert hymns.clean_lyric("우리 모두가- 함께 걷는 길이라는걸--") == "우리 모두가 함께 걷는 길이라는걸"


def test_ocr_run_captions_a_song_the_library_lacks(tmp_path, monkeypatch):
    blobs = {3: b"page-a", 4: b"page-b"}
    monkeypatch.setattr(hymns, "_images", lambda src, idx: blobs)
    pages = {b"page-a": [(0.04, ": 새 노래로 함께 부르며"), (0.13, "F Major"), (0.27, ": 기뻐해요- 알지 못했던 날에도-")],
             b"page-b": [(0.05, "우리가 노-래하는 오늘"), (0.5, "2")]}
    monkeypatch.setattr(hymns, "read_lines", lambda path: pages[path.read_bytes()])
    m = hymns.ocr_run(tmp_path / "deck.pptx", [3, 4], {"songs": [], "images": {}}, tmp_path)
    # the title band of the first page names the song; later pages have no title band
    assert m.ocr and m.label == "새 노래로 함께 부르며"
    assert m.lines == ["기뻐해요 알지 못했던 날에도", "우리가 노래하는 오늘"] and (m.start, m.end) == (0, 1)


@pytest.mark.skipif(__import__("sys").platform != "win32", reason="Windows OCR engine")
def test_windows_ocr_reads_a_rendered_word(tmp_path):
    from PIL import Image, ImageDraw, ImageFont
    assert hymns.ocr_available(), "no Windows OCR engine for Korean or the user's languages"
    img = Image.new("RGB", (900, 200), "white")
    ImageDraw.Draw(img).text((40, 50), "HELLO WORLD", fill="black", font=ImageFont.load_default(size=72))
    path = tmp_path / "word.png"
    img.save(path)
    assert "HELLO" in " ".join(hymns.ocr(path)).upper()


def test_lettering_copies_keep_each_style_solid(tmp_path):
    from PIL import Image
    page = Image.new("RGBA", (400, 200), (0, 0, 0, 0))  # a transparent score page
    page.paste((17, 34, 66, 255), (20, 20, 180, 60))      # navy lyric
    page.paste((255, 255, 255, 255), (20, 120, 180, 160))  # white lyric
    path = tmp_path / "p.png"
    page.save(path)
    dark, light = (Image.open(p).convert("L") for p in hymns._lettering(path))
    assert dark.size == (800, 400)
    assert dark.getpixel((100, 80)) == 0 and dark.getpixel((100, 280)) == 255
    assert light.getpixel((100, 280)) == 0 and light.getpixel((100, 80)) == 255


def test_merge_reads_keeps_the_better_reading_per_line():
    dark = [(0.27, "기뻐해요- 알지 못했던 날에도-"), (0.54, "&....??...")]
    light = [(0.28, "V."), (0.545, "우리 모두가- 함께 걷는 길이라는걸--")]
    assert [t for _, t in hymns.merge_reads([dark, light])] == ["기뻐해요- 알지 못했던 날에도-", "우리 모두가- 함께 걷는 길이라는걸--"]


def test_line_text_joins_syllables_the_engine_split():
    def word(text, x, w=40, h=40):
        return SimpleNamespace(text=text, bounding_rect=SimpleNamespace(x=x, width=w, height=h))
    # 이 제 read as syllables 4px apart, a real word gap of 30px before 나도
    line = SimpleNamespace(text="이 제 나도", words=[word("이", 0), word("제", 44), word("나도", 114, w=80)])
    assert hymns._line_text(line) == "이제 나도"


def test_a_no_song_entry_without_a_reading_is_read_again(tmp_path, monkeypatch):
    reads = []
    monkeypatch.setattr(hymns, "ocr", lambda path: reads.append(path) or ["기뻐해요 알지 못했던 날에도"])
    index = {"songs": [], "images": {"abc": {"song": None, "ocr": 0}}}  # written where nothing could read
    hymns._page(index, "abc", b"png", tmp_path)
    assert len(reads) == 1 and index["images"]["abc"]["read"] == ["기뻐해요 알지 못했던 날에도"]
    hymns._page(index, "abc", b"png", tmp_path)  # now it carries its reading: trusted
    assert len(reads) == 1


def test_a_repeated_chorus_takes_the_title_of_its_song(tmp_path, monkeypatch):
    from broadcast_ppt.mapping import plan
    first = hymns.HymnMatch("ocr", "t", "새 노래로 함께 부르며", ["기뻐해요", "오늘도 노래하기 위해", "그리고 그 노래 나누기 위해"],
                            0, 2, False, ocr=True)
    again = hymns.HymnMatch("ocr", "t", "", ["오늘도 노래하기 위해", "그리고 그 노래 나누기 위해"], 0, 1, False, ocr=True)
    runs = {19: first, 26: again}
    monkeypatch.setattr(plan, "_hymn_runs", lambda *a: [[19, 20, 21], [26, 27, 28]])
    monkeypatch.setattr(hymns, "match_run", lambda src, run, index, work: hymns.RunResult(None, readable=True))
    monkeypatch.setattr(hymns, "ocr_run", lambda src, run, index, work: runs[run[0]])
    monkeypatch.setattr(hymns, "load_index", lambda: {"songs": [{"id": "x"}], "images": {}})
    monkeypatch.setattr(hymns, "save_index", lambda index: None)
    found, _, _ = plan._hymn_plan(tmp_path / "d.pptx", [], [], SimpleNamespace(fam={"hymn": []}), set())
    assert found[26][0].label == "새 노래로 함께 부르며"
