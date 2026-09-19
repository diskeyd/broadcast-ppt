"""Render slides through Keynote (macOS) and measure where text is drawn.

Keynote is sandboxed: it cannot open files from the project folder, but it can always read its own iCloud
documents folder. Decks are copied to a temporary folder there, exported as PNGs, copied back into
local/work/keynote/, and the temporary folder is deleted. Never render slides that carry names — use a
deck built with placeholder text (see `calibrate-cards`).

Usage:
  uv run python scripts/keynote_check.py render <deck.pptx> [--slides 2,4,12] [--out local/work/keynote/check]
  uv run python scripts/keynote_check.py bands <image.png> ...
  uv run python scripts/keynote_check.py calibrate-cards [--gap 30]
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import time
import unicodedata
from pathlib import Path

from PIL import Image
from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from broadcast_ppt.builder.package import current_slide_ids, drop_section_list, remove_slides  # noqa: E402
from broadcast_ppt.ooxml import EMU_PER_PT  # noqa: E402

KEYNOTE_DOCS = Path.home() / "Library/Mobile Documents/com~apple~Keynote/Documents"
APP = "Keynote Creator Studio"


def bands(path: Path, h_pt: float = 540.0, tol: int = 60) -> list[tuple]:
    """Vertical bands of non-background pixels in slide points; background = top-left pixel colour."""
    im = Image.open(path).convert("RGB")
    w, h = im.size
    px = im.load()
    bg = px[2, 2]
    out, start = [], None
    for y in range(h):
        n = sum(1 for x in range(0, w, 2) if sum(abs(a - b) for a, b in zip(px[x, y], bg)) > tol)
        if n > 1 and start is None:
            start = y
        elif n <= 1 and start is not None:
            if y - start > 1:
                out.append((round(start * h_pt / h, 1), round(y * h_pt / h, 1)))
            start = None
    if start is not None:
        out.append((round(start * h_pt / h, 1), None))  # runs off the slide bottom
    return out


def ensure_keynote(timeout: float = 60) -> None:
    """Start Keynote in the background and wait until it answers Apple events (a cold start answers -609)."""
    cold = subprocess.run(["pgrep", "-x", "Keynote"], capture_output=True).returncode != 0
    subprocess.run(["open", "-g", "-a", APP], check=False)
    deadline = time.time() + timeout
    while time.time() < deadline:
        r = subprocess.run(["osascript", "-e", f'tell application "{APP}" to get name'], capture_output=True, text=True, timeout=30)
        if r.returncode == 0:
            # the scripting bridge answers before Keynote can export; a cold start needs longer
            time.sleep(20 if cold else 3)
            return
        time.sleep(2)
    raise RuntimeError("Keynote did not respond to Apple events")


def keynote_export(deck: Path, out_dir: Path) -> list[Path]:
    ensure_keynote()
    tmp = KEYNOTE_DOCS / "bppt-tmp"
    shutil.rmtree(tmp, ignore_errors=True)
    (tmp / "out").mkdir(parents=True)
    try:
        shutil.copy(deck, tmp / deck.name)
        script = f'''with timeout of 150 seconds
tell application "{APP}"
set d to open POSIX file "{tmp / deck.name}"
delay 3
export d to POSIX file "{tmp / 'out'}" as slide images with properties {{image format:PNG, skipped slides:false}}
close d saving no
end tell
end timeout'''
        for _ in range(3):
            r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=200)
            if r.returncode == 0 or "-609" not in r.stderr:
                break
            time.sleep(5)
        if r.returncode != 0:
            raise RuntimeError(f"Keynote export failed: {r.stderr.strip()}")
        shutil.rmtree(out_dir, ignore_errors=True)
        out_dir.mkdir(parents=True)
        pngs = []
        for p in sorted((tmp / "out").glob("*.png")):
            shutil.copy(p, out_dir / p.name)
            pngs.append(out_dir / p.name)
        return pngs
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


def text_box_tops(slide) -> list[float]:
    return sorted(round(sh.top / EMU_PER_PT, 1) for sh in slide.shapes if sh.has_text_frame and sh.text_frame.text.strip())


def cmd_render(args) -> int:
    deck = Path(args.deck)
    prs = Presentation(str(deck))
    keep = [int(x) for x in args.slides.split(",")] if args.slides else list(range(1, len(prs.slides) + 1))
    ids = current_slide_ids(prs)
    remove_slides(prs, [i for k, i in enumerate(ids) if k + 1 not in keep])
    drop_section_list(prs)
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    subset = out / "render.pptx"
    prs.save(str(subset))
    pngs = keynote_export(subset, out / "png")
    shown = Presentation(str(subset))
    for n, slide, png in zip(keep, shown.slides, pngs):
        print(f"slide {n}: text boxes at {text_box_tops(slide)} pt, drawn text bands {bands(png)} -> {png}")
    return 0


def cmd_bands(args) -> int:
    for p in args.images:
        print(Path(p).name, bands(Path(p)))
    return 0


def cmd_calibrate_cards(args) -> int:
    """Build every card template with placeholder text and rows spread `gap` pt apart, render in Keynote and
    print each row's [top, bottom] drawn-text offset from its box top (the `glyph` values in roles.yaml)."""
    import broadcast_ppt.builder.build as B
    from broadcast_ppt.mapping.catalog import load_catalog

    cat = load_catalog(ROOT / "config/templates.yaml")
    for e in cat["slides"].values():
        if e.get("stack"):
            e["stack"]["gap_pt"] = args.gap
            e["stack"].pop("max_bottom_pt", None)
    B.load_catalog = lambda _p: cat
    src = next(p for p in (ROOT / "local/input/미편집").iterdir() if unicodedata.normalize("NFC", p.name).startswith("주일오전예배ppt("))
    fills = {3: {"label": "예배로의 부름"}, 4: {"label": "기도 ", "person": "– 홍길동 장로"},
             6: {"ref": "사무엘상 26:21-25", "page": "구약성경 p.455"},
             7: {"label": "설교 ", "title": "– “제목”", "ref": "사무엘상 26:21-25"},
             17: {"label": "동행의 기도"}, 20: {}, 31: {}}
    plan = {"template": {"sha256": ""}, "source": {"path": str(src), "sha256": ""}, "options": {"scale": 4 / 3},
            "rules_version": "calibration",
            "ops": [{"op": "template", "out_index": i + 1, "template_slide": t, "fill": f, "flags": []}
                    for i, (t, f) in enumerate(fills.items())]}
    out = ROOT / "local/work/keynote/calibrate"
    out.mkdir(parents=True, exist_ok=True)
    deck = out / "cards.pptx"
    B.build_deck(plan, ROOT / "local/templates/24.상반기.pptx", ROOT / "config/templates.yaml", deck, force=True)
    # lift everything so the lowest rows are not cut off at the slide bottom
    prs = Presentation(str(deck))
    for slide in prs.slides:
        for sh in slide.shapes:
            if sh.has_text_frame and sh.text_frame.text.strip():
                sh.top = sh.top - 60 * EMU_PER_PT
    prs.save(str(deck))
    pngs = keynote_export(deck, out / "png")
    for t, slide, png in zip(fills, Presentation(str(deck)).slides, pngs):
        tops = text_box_tops(slide)
        found = bands(png)
        offsets = [(round(b[0] - y, 1), round(b[1] - y, 1) if b[1] is not None else None) for y, b in zip(tops, found)]
        print(f"template {t}: rows top->bottom glyph offsets {offsets}" + ("" if len(found) == len(tops) else "  (bands merged: raise --gap)"))
    return 0


def main() -> int:
    ap = argparse.ArgumentParser(prog="keynote_check")
    sub = ap.add_subparsers(dest="cmd", required=True)
    r = sub.add_parser("render")
    r.add_argument("deck")
    r.add_argument("--slides")
    r.add_argument("--out", default=str(ROOT / "local/work/keynote/check"))
    r.set_defaults(func=cmd_render)
    b = sub.add_parser("bands")
    b.add_argument("images", nargs="+")
    b.set_defaults(func=cmd_bands)
    c = sub.add_parser("calibrate-cards")
    c.add_argument("--gap", type=float, default=30.0)
    c.set_defaults(func=cmd_calibrate_cards)
    args = ap.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
