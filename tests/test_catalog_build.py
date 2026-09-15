"""Catalog -> build -> validate on a tiny synthetic template (no real decks, no names)."""

from pathlib import Path

import yaml
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

from broadcast_ppt.builder.build import build_deck
from broadcast_ppt.mapping.catalog import build_catalog, dump_catalog
from broadcast_ppt.ooxml import AUTOFIT_LINE_HEIGHT, EMU_PER_PT, REPO_ROOT
from broadcast_ppt.validator.check import validate_deck

SAMPLE_LABEL = "샘플라벨"
SAMPLE_SUB = "sample subtitle"


def _template(path: Path) -> tuple[int, int]:
    prs = Presentation()
    prs.slide_width, prs.slide_height = Inches(13.333), Inches(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    label = s.shapes.add_textbox(Inches(4), Inches(5.8), Inches(4), Inches(1))
    run = label.text_frame.paragraphs[0].add_run()
    run.text, run.font.size = SAMPLE_LABEL, Pt(54)
    run.font.color.rgb = RGBColor(0xFF, 0xFF, 0x00)
    sub = s.shapes.add_textbox(Inches(4), Inches(6.6), Inches(4), Inches(0.8))
    r2 = sub.text_frame.paragraphs[0].add_run()
    r2.text, r2.font.size = SAMPLE_SUB, Pt(40)
    prs.save(str(path))
    return label.shape_id, sub.shape_id


def _roles(path: Path, label_id: int, sub_id: int) -> None:
    path.write_text(yaml.safe_dump({
        "version": 1, "card_gap_pt": 2,
        "slides": {1: {"role": "card_test", "family": "card.test",
                       "shapes": {label_id: {"slots": ["label"]}, sub_id: {"slots": ["subtitle_en"], "mode": "optional"}},
                       "stack": {"rows": [label_id, sub_id], "glyph": {str(label_id): [9.0, 64.0], str(sub_id): [20.0, 60.0]}}}},
    }, allow_unicode=True), encoding="utf-8")


def test_catalog_build_validate_card_stack(tmp_path):
    tpl, roles, cat_path = tmp_path / "tpl.pptx", tmp_path / "roles.yaml", tmp_path / "templates.yaml"
    label_id, sub_id = _template(tpl)
    _roles(roles, label_id, sub_id)

    cat = build_catalog(tpl, roles)
    dump_catalog(cat, cat_path)
    entry = cat["slides"][1]
    assert entry["shapes"][label_id]["slot_map"] == ["label"]
    assert [r["order"] for r in entry["stack"]["rows"]] == [entry["shapes"][label_id]["order"], entry["shapes"][sub_id]["order"]]
    assert entry["stack"]["gap_pt"] == 2.0
    assert SAMPLE_LABEL not in cat_path.read_text(encoding="utf-8")  # sample text is never stored

    src = tmp_path / "src.pptx"
    Presentation().save(str(src))
    plan = {"template": {"sha256": ""}, "source": {"path": str(src), "sha256": ""}, "options": {"scale": 1.0},
            "rules_version": "test",
            "ops": [{"op": "template", "out_index": 1, "template_slide": 1, "fill": {"label": "새 라벨"}, "flags": []}]}
    out = tmp_path / "out.pptx"
    build_deck(plan, tpl, cat_path, out, force=True, widths_path=REPO_ROOT / "config/widths.yaml")

    label, sub = list(Presentation(str(out)).slides[0].shapes)  # ids are renumbered on copy; order is kept
    assert label.text_frame.text == "새 라벨" and sub.text_frame.text == SAMPLE_SUB
    # lowest row keeps its bottom edge with a one-line height; the label sits `gap` above its drawn text
    one_line = int(AUTOFIT_LINE_HEIGHT * 40 * EMU_PER_PT + 2 * 45720)
    assert sub.height == one_line
    expected_label_top = sub.top + int(20.0 * EMU_PER_PT) - int(2.0 * EMU_PER_PT) - int(64.0 * EMU_PER_PT)
    assert abs(label.top - expected_label_top) <= 2

    rep = validate_deck(out, plan, cat_path, REPO_ROOT / "config/widths.yaml")
    assert rep["counts"]["critical"] == 0, rep["findings"]
