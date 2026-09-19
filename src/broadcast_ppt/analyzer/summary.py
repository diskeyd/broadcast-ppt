"""Aggregate per-file analysis JSON into summary.csv and clusters.md.

clusters.md is meant to be pasted into a chat or a note, so it carries counts,
colours, keywords and file names only — never slide text (names, roles).
"""

from __future__ import annotations

import csv
import json
import re
import statistics
from collections import Counter, defaultdict
from pathlib import Path

import yaml

from .core import classify_rgb

UNEDITED_MARK = "미편집"


def load_scope(path: Path | None) -> dict | None:
    if not path or not Path(path).exists():
        return None
    return yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}


def in_scope(doc: dict, scope: dict | None) -> bool:
    if scope is None:
        return True
    name = re.sub(r"\s", "", doc.get("name", ""))
    exc = scope.get("exclude") or {}
    if doc.get("folder", "") in (exc.get("folders") or []):
        return False
    if doc.get("family") in (exc.get("families") or []):
        return False
    if any(name.startswith(x) for x in exc.get("name_startswith") or []):
        return False
    if doc.get("family") in (scope.get("families") or []):
        return True
    return any(x in name for x in scope.get("name_contains") or [])


def load_all(out_dir: Path, scope: dict | None = None) -> list[dict]:
    docs = []
    for p in sorted((out_dir / "files").glob("*.json")):
        doc = json.loads(p.read_text(encoding="utf-8"))
        if in_scope(doc, scope):
            docs.append(doc)
    return docs


def is_chroma_slide(sl: dict) -> bool:
    return any(c.get("area_frac", 0) >= 0.3 for c in sl.get("chroma", []))


def md_table(headers: list[str], rows: list[list]) -> str:
    def esc(v):
        return str(v).replace("|", "\\|").replace("\n", " ")
    lines = ["| " + " | ".join(headers) + " |", "|" + "---|" * len(headers)]
    lines += ["| " + " | ".join(esc(v) for v in r) + " |" for r in rows]
    return "\n".join(lines)


def write_csv(docs: list[dict], out_dir: Path) -> Path:
    path = out_dir / "summary.csv"
    cols = ["path", "folder", "family", "suffix", "date", "mb", "app", "ratio", "slides", "layouts",
            "chroma_slides", "chroma_colors", "bg_types", "pictures", "full_bleed_pics", "tables",
            "groups", "smartart", "video", "text_slides", "error"]
    with path.open("w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        w.writerow(cols)
        for d in docs:
            slides = [s for s in d.get("slides", []) if "error" not in s]
            kinds = Counter()
            for s in slides:
                kinds.update(s.get("kinds", {}))
            chroma_rgb = Counter(c.get("rgb") for s in slides for c in s.get("chroma", []) if c.get("area_frac", 0) >= 0.3)
            w.writerow([
                d["path"], d.get("folder", ""), d["family"], d.get("suffix") or "", d.get("date") or "",
                round(d["bytes"] / 1e6, 1), d.get("Application", ""), d.get("ratio", ""), d.get("slide_count", ""),
                len(d.get("layouts_used", {})), sum(is_chroma_slide(s) for s in slides),
                " ".join(f"{k}:{v}" for k, v in chroma_rgb.most_common(3)),
                " ".join(f"{k}:{v}" for k, v in Counter(s["background"]["type"] for s in slides).most_common()),
                kinds["PICTURE"] + kinds["ph:PICTURE"], sum(s.get("full_bleed_picture", False) for s in slides),
                kinds["TABLE"], kinds["GROUP"], kinds["SMART_ART"], sum((d.get("media") or {}).get(e, 0) for e in (".mp4", ".wmv", ".mov", ".avi")),
                sum(1 for s in slides if s.get("text_len", 0) > 0), d.get("error", ""),
            ])
    return path


def write_markdown(docs: list[dict], out_dir: Path) -> Path:
    ok = [d for d in docs if "error" not in d]
    bad = [d for d in docs if "error" in d]
    all_slides = [(d, s) for d in ok for s in d["slides"] if "error" not in s]
    out: list[str] = ["# broadcast-ppt sample inventory", ""]

    out += ["## Overview", "", md_table(["item", "value"], [
        ["files", len(docs)], ["opened", len(ok)], ["failed", len(bad)], ["slides", len(all_slides)],
        ["slide errors", sum(1 for d in ok for s in d["slides"] if "error" in s)],
        ["folders", ", ".join(f"{k or '(root)'}:{v}" for k, v in Counter(d.get("folder", "") for d in docs).most_common())],
        ["apps", ", ".join(f"{k}:{v}" for k, v in Counter(d.get("Application") for d in docs).most_common())],
    ]), ""]
    if bad:
        out += ["### Failed files", "", md_table(["file", "error"], [[d["path"], d["error"][:120]] for d in bad]), ""]

    fam_rows = []
    by_fam = defaultdict(list)
    for d in ok:
        by_fam[(d.get("folder", ""), d["family"])].append(d)
    for (folder, fam), ds in sorted(by_fam.items(), key=lambda kv: -len(kv[1])):
        dates = sorted(x["date"] for x in ds if x.get("date"))
        counts = [x["slide_count"] for x in ds]
        slides = [s for x in ds for s in x["slides"] if "error" not in s]
        chroma_files = sum(1 for x in ds if any(is_chroma_slide(s) for s in x["slides"] if "error" not in s))
        fam_rows.append([folder or "(root)", fam, len(ds), f"{dates[0]}~{dates[-1]}" if dates else "",
                         int(statistics.median(counts)) if counts else 0,
                         ",".join(sorted({x.get("ratio", "?") for x in ds})), chroma_files,
                         f"{sum(is_chroma_slide(s) for s in slides)}/{len(slides)}"])
    out += ["## Families", "", md_table(["folder", "family", "files", "dates", "median slides", "ratio", "files w/ chroma", "chroma slides"], fam_rows), ""]

    bg = Counter()
    bg_files = defaultdict(set)
    for d, s in all_slides:
        b = s["background"]
        label = f'{b["type"]}@{b["level"]}:' + (b.get("rgb") or (b.get("image") or {}).get("mean_rgb") or "-")
        bg[label] += 1
        bg_files[label].add(d["path"])
    out += ["## Slide backgrounds (top 25)", "", md_table(["background", "key colour", "slides", "files"], [
        [k, _key_of(k), v, len(bg_files[k])] for k, v in bg.most_common(25)]), ""]

    src = Counter()
    rgb = Counter()
    for _, s in all_slides:
        for c in s.get("chroma", []):
            if c.get("area_frac", 0) >= 0.3:
                src[(c["source"], c.get("kind") or c.get("fill"), c["key"])] += 1
                rgb[(c.get("rgb"), c["key"])] += 1
    out += ["## Chroma candidates (area >= 30%)", "", md_table(["source", "shape/fill", "key", "slides"],
            [[a, b, k, v] for (a, b, k), v in src.most_common(20)]), "",
            md_table(["rgb", "key", "slides"], [[a, k, v] for (a, k), v in rgb.most_common(15)]), ""]
    partial = Counter()
    for _, s in all_slides:
        for c in s.get("chroma", []):
            if 0.05 <= c.get("area_frac", 0) < 0.3:
                partial[(c["source"], c.get("kind") or c.get("fill"), c["key"], c.get("rgb"))] += 1
    if partial:
        out += ["### Partial chroma regions (5–30% of slide)", "", md_table(["source", "shape", "key", "rgb", "slides"],
                [[*k, v] for k, v in partial.most_common(15)]), ""]

    kw = Counter()
    kw_files = defaultdict(set)
    for d, s in all_slides:
        for k in s["keywords"]:
            kw[k] += 1
            kw_files[k].add(d["path"])
    sunday = [d for d in ok if d["family"] == "주일오전예배"]
    kw_rows = []
    for k, v in kw.most_common():
        sun_share = sum(1 for d in sunday if d["path"] in kw_files[k])
        kw_rows.append([k, v, len(kw_files[k]), f"{sun_share}/{len(sunday)}"])
    kw_rows.append(["(scripture ref)", sum(s["scripture_ref"] for _, s in all_slides), "", ""])
    kw_rows.append(["(hymn number)", sum(s["hymn_no"] for _, s in all_slides), "", ""])
    kw_rows.append(["(no text)", sum(1 for _, s in all_slides if s["text_len"] == 0), "", ""])
    out += ["## Keywords", "", md_table(["keyword", "slides", "files", "주일오전예배 files"], kw_rows), ""]

    lay = Counter()
    lay_files = defaultdict(set)
    for d, s in all_slides:
        lay[s["layout"]] += 1
        lay_files[s["layout"]].add(d["path"])
    out += ["## Layout names (top 25)", "", md_table(["layout", "slides", "files"],
            [[k, v, len(lay_files[k])] for k, v in lay.most_common(25)]), ""]

    sig = Counter()
    sig_meta: dict[str, dict] = {}
    for d, s in all_slides:
        sig[s["signature"]] += 1
        m = sig_meta.setdefault(s["signature"], {"files": set(), "kw": Counter(), "layout": s["layout"],
                                                 "kinds": s.get("kinds", {}), "chroma": 0, "fams": Counter()})
        m["files"].add(d["path"])
        m["kw"].update(s["keywords"] or ["-"])
        m["chroma"] += is_chroma_slide(s)
        m["fams"][d["family"]] += 1
    out += ["## Shape signatures (top 30)", "", "Signature = layout + top-level shape kinds and positions on a 10% grid.", "",
            md_table(["sig", "slides", "files", "layout", "kinds", "top keywords", "chroma", "families"], [
                [k, v, len(sig_meta[k]["files"]), sig_meta[k]["layout"],
                 " ".join(f"{a}:{b}" for a, b in sorted(sig_meta[k]["kinds"].items())),
                 " ".join(f"{a}:{b}" for a, b in sig_meta[k]["kw"].most_common(3)), sig_meta[k]["chroma"],
                 " ".join(f"{a}:{b}" for a, b in sig_meta[k]["fams"].most_common(2))]
                for k, v in sig.most_common(30)]), ""]

    if sunday:
        stab = []
        for k in kw:
            sigs = Counter(s["signature"] for d in sunday for s in d["slides"] if "error" not in s and k in s["keywords"])
            if sigs:
                top, n = sigs.most_common(1)[0]
                stab.append([k, sum(sigs.values()), len(sigs), top, f"{n / sum(sigs.values()):.0%}"])
        out += ["## 주일오전예배 position stability", "", md_table(["keyword", "slides", "distinct sigs", "top sig", "top share"],
                sorted(stab, key=lambda r: -r[1])), ""]
        seqs = Counter()
        for d in sunday:
            seq = []
            for s in d["slides"]:
                if "error" in s:
                    continue
                k = s["keywords"][0] if s["keywords"] else None
                if k and (not seq or seq[-1] != k):
                    seq.append(k)
            seqs[" > ".join(seq)] += 1
        out += ["### Keyword order per file (top 8)", "", md_table(["files", "order"], [[v, k] for k, v in seqs.most_common(8)]), ""]

    pairs = []
    unedited = [d for d in ok if UNEDITED_MARK in d.get("folder", "")]
    others = [d for d in ok if UNEDITED_MARK not in d.get("folder", "")]
    index = defaultdict(list)
    for d in others:
        index[(d["family"], d.get("date"))].append(d)
    for u in unedited:
        for o in index.get((u["family"], u.get("date")), []):
            ut = {re.sub(r"\s", "", s["text"]) for s in u["slides"] if s.get("text")}
            ot = {re.sub(r"\s", "", s["text"]) for s in o["slides"] if s.get("text")}
            jac = len(ut & ot) / max(len(ut | ot), 1)
            pairs.append([u["family"], u.get("date"), u["slide_count"], o["slide_count"],
                          sum(is_chroma_slide(s) for s in u["slides"] if "error" not in s),
                          sum(is_chroma_slide(s) for s in o["slides"] if "error" not in s), f"{jac:.0%}"])
    out += ["## 미편집 pairs", "", f"미편집 files: {len(unedited)} · paired: {len(pairs)}", ""]
    if pairs:
        out += [md_table(["family", "date", "slides (미편집)", "slides (edited)", "chroma (미편집)", "chroma (edited)", "same slide text"], pairs), ""]

    by_date = defaultdict(lambda: {"u": Counter(), "o": Counter()})
    for d in ok:
        if d.get("date"):
            by_date[d["date"]]["u" if UNEDITED_MARK in d.get("folder", "") else "o"][d["family"]] += 1
    date_rows = [[k, " ".join(f"{f}:{n}" for f, n in v["u"].items()), " ".join(f"{f}:{n}" for f, n in v["o"].items()) or "-"]
                 for k, v in sorted(by_date.items()) if v["u"]]
    if date_rows:
        out += ["### Same-date decks (미편집 dates)", "", md_table(["date", "in 미편집", "elsewhere"], date_rows), ""]

    anomalies = [
        ["full-bleed picture, no text (flattened?)", sum(1 for _, s in all_slides if s["full_bleed_picture"] and s["text_len"] == 0)],
        ["slides with tables", sum(1 for _, s in all_slides if s["kinds"].get("TABLE"))],
        ["slides with groups", sum(1 for _, s in all_slides if s["kinds"].get("GROUP"))],
        ["slides with SmartArt", sum(1 for _, s in all_slides if s["kinds"].get("SMART_ART"))],
        ["slides with media", sum(1 for _, s in all_slides if s["kinds"].get("MEDIA"))],
        ["slides hiding master shapes", sum(1 for _, s in all_slides if not s["show_master_shapes"])],
        ["slides with notes", sum(1 for _, s in all_slides if s["notes"])],
        ["files with embedded fonts", sum(1 for d in ok if d.get("embedded_fonts"))],
    ]
    fonts = Counter()
    for d in ok:
        fonts.update(d.get("fonts", {}))
    media = Counter()
    for d in docs:
        media.update(d.get("media") or {})
    out += ["## Anomalies and inventory", "", md_table(["check", "count"], anomalies), "",
            md_table(["font", "runs"], fonts.most_common(15)), "",
            md_table(["media ext", "parts"], media.most_common()), ""]

    path = out_dir / "clusters.md"
    path.write_text("\n".join(out), encoding="utf-8")
    return path


def _key_of(label: str) -> str:
    m = re.search(r":([0-9A-F]{6})$", label)
    return (classify_rgb(m.group(1)) or "") if m else ""


def summarize(out_dir: Path, scope_path: Path | None = None) -> tuple[Path, Path]:
    docs = load_all(out_dir, load_scope(scope_path))
    return write_csv(docs, out_dir), write_markdown(docs, out_dir)
