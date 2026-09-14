"""Fit per-class character widths from auto-fit text boxes in real decks (no renderer needed)."""

from __future__ import annotations

import statistics
import unicodedata
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

from ..mapping.textfit import CLASSES, DEFAULT_EM, EMU_PER_PT, LINE_HEIGHT, WidthModel, class_counts

A = "http://schemas.openxmlformats.org/drawingml/2006/main"
DEF_INS = {"lIns": 91440, "rIns": 91440, "tIns": 45720, "bIns": 45720}


def _walk(shapes):
    for sh in shapes:
        if sh.shape_type == MSO_SHAPE_TYPE.GROUP:
            yield from _walk(sh.shapes)
        else:
            yield sh


def collect_samples(paths: list[Path]) -> tuple[list[dict], list[dict]]:
    """A: single-line wrap=none autofit boxes (text, size, inner width). B: wrapping autofit boxes (text, size, width, lines)."""
    single, wrapping = [], []
    for path in paths:
        try:
            prs = Presentation(str(path))
        except Exception:
            continue
        for slide in prs.slides:
            for sh in _walk(slide.shapes):
                if not sh.has_text_frame or sh.width is None or sh.height is None:
                    continue
                body = sh.text_frame._txBody
                bp = body.find(f"{{{A}}}bodyPr")
                if bp is None or bp.find(f"{{{A}}}spAutoFit") is None:
                    continue
                if bp.get("vert") not in (None, "horz") or sh.rotation:
                    continue
                paras = sh.text_frame.paragraphs
                sizes = {r.font.size.pt for p in paras for r in p.runs if r.font.size and r.text.strip()}
                missing = any(r.font.size is None and r.text.strip() for p in paras for r in p.runs)
                if len(sizes) != 1 or missing:
                    continue
                sz = sizes.pop()
                texts = ["".join(r.text for r in p.runs) for p in paras]
                if any("\x0b" in t for t in texts) or not any(t.strip() for t in texts):
                    continue
                ins = {k: int(bp.get(k)) if bp.get(k) else v for k, v in DEF_INS.items()}
                if bp.get("wrap") == "none" and len([t for t in texts if t.strip()]) == 1:
                    text = next(t for t in texts if t.strip())
                    single.append({"text": text, "sz": sz, "inner": sh.width - ins["lIns"] - ins["rIns"]})
                elif bp.get("wrap") in (None, "square") and len(paras) == 1:
                    lines = (sh.height - ins["tIns"] - ins["bIns"]) / (LINE_HEIGHT * sz * EMU_PER_PT)
                    if abs(lines - round(lines)) < 0.08 and round(lines) >= 1:
                        wrapping.append({"text": texts[0], "sz": sz, "width": int(sh.width), "lIns": ins["lIns"],
                                         "rIns": ins["rIns"], "lines": int(round(lines))})
    return single, wrapping


def _solve(rows: list[list[float]], ys: list[float], ridge: float = 1e-3) -> list[float]:
    n = len(rows[0])
    ata = [[sum(r[i] * r[j] for r in rows) + (ridge if i == j else 0.0) for j in range(n)] for i in range(n)]
    aty = [sum(r[i] * y for r, y in zip(rows, ys)) for i in range(n)]
    # Gaussian elimination
    for c in range(n):
        piv = max(range(c, n), key=lambda r: abs(ata[r][c]))
        ata[c], ata[piv] = ata[piv], ata[c]
        aty[c], aty[piv] = aty[piv], aty[c]
        for r in range(c + 1, n):
            f = ata[r][c] / ata[c][c] if ata[c][c] else 0.0
            for k in range(c, n):
                ata[r][k] -= f * ata[c][k]
            aty[r] -= f * aty[c]
    x = [0.0] * n
    for c in range(n - 1, -1, -1):
        x[c] = (aty[c] - sum(ata[c][k] * x[k] for k in range(c + 1, n))) / ata[c][c] if ata[c][c] else 0.0
    return x


def fit(single: list[dict], rounds: int = 3) -> tuple[dict[str, float], dict]:
    rows_all = []
    for s in single:
        counts = class_counts(s["text"])
        rows_all.append(([counts[k] * s["sz"] * EMU_PER_PT for k in CLASSES], s["inner"], s))
    keep = rows_all
    em = dict(DEFAULT_EM)
    for _ in range(rounds):
        present = [i for i, k in enumerate(CLASSES) if sum(r[0][i] for r in keep) > 0]
        coef = _solve([[r[0][i] for i in present] for r in keep], [r[1] for r in keep])
        for i, c in zip(present, coef):
            em[CLASSES[i]] = round(max(0.1, min(1.3, c)), 4)
        model = WidthModel(em=em)
        resid = [abs(model.width_emu(r[2]["text"], r[2]["sz"]) - r[1]) / max(r[1], 1) for r in rows_all]
        cutoff = max(0.12, statistics.quantiles(resid, n=10)[7]) if len(resid) > 10 else 1.0
        keep = [r for r, e in zip(rows_all, resid) if e <= cutoff]
    model = WidthModel(em=em)
    errors = sorted(abs(model.width_emu(r[2]["text"], r[2]["sz"]) - r[1]) / max(r[1], 1) for r in rows_all)
    stats = {"single_samples": len(rows_all), "kept": len(keep),
             "width_error_median": round(statistics.median(errors), 4) if errors else None,
             "width_error_p90": round(errors[int(len(errors) * 0.9)], 4) if errors else None}
    return em, stats


def evaluate_lines(model: WidthModel, wrapping: list[dict]) -> dict:
    if not wrapping:
        return {"wrapping_samples": 0}
    exact = within1 = 0
    for w in wrapping:
        est = model.estimate_lines(w["text"], w["width"], w["sz"], w["lIns"], w["rIns"])
        exact += est == w["lines"]
        within1 += abs(est - w["lines"]) <= 1
    return {"wrapping_samples": len(wrapping), "lines_exact": round(exact / len(wrapping), 4), "lines_within_1": round(within1 / len(wrapping), 4)}


def calibrate(paths: list[Path], out: Path) -> dict:
    single, wrapping = collect_samples(paths)
    em, stats = fit(single)
    model = WidthModel(em=em)
    stats.update(evaluate_lines(model, wrapping))
    data = {"version": 1, "line_height": LINE_HEIGHT, "em": em, "stats": stats, "sources": len(paths)}
    Path(out).write_text(yaml.safe_dump(data, sort_keys=False, allow_unicode=True), encoding="utf-8")
    return data


import yaml  # noqa: E402  (kept at the end to avoid a circular import during package init)
