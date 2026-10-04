"""broadcast-ppt command line."""

from __future__ import annotations

import argparse
import contextlib
from collections import Counter
import io
import json
import os
import re
import shutil
import sys
import time
import traceback
import unicodedata
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

from .analyzer.core import analyze_file
from .analyzer.summary import summarize
from .ooxml import REPO_ROOT

PPT_SUFFIXES = {".pptx", ".ppt"}
CATALOG = str(REPO_ROOT / "config/templates.yaml")
WIDTHS = str(REPO_ROOT / "config/widths.yaml")


def _json_name(rel: Path) -> str:
    s = unicodedata.normalize("NFC", str(rel)).replace(os.sep, "__")
    return s + ".json"


def _one(path: str, root: str, out: str) -> tuple[str, str | None, int]:
    p, r = Path(path), Path(root)
    target = Path(out) / "files" / _json_name(p.relative_to(r))
    if p.suffix.lower() == ".ppt":
        doc = {"path": str(p.relative_to(r)), "name": p.name, "folder": "", "family": "기타", "bytes": p.stat().st_size,
               "error": "legacy .ppt (binary) — python-pptx cannot read it"}
    else:
        doc = analyze_file(p, r)
    target.write_text(json.dumps(doc, ensure_ascii=False, indent=1, default=list), encoding="utf-8")
    return str(p.relative_to(r)), doc.get("error"), doc.get("slide_count", 0)


def cmd_analyze(args) -> int:
    src = Path(args.input).expanduser().resolve()
    out = Path(args.out).expanduser().resolve()
    (out / "files").mkdir(parents=True, exist_ok=True)
    if src.is_file():
        root, files = src.parent, [src]
    else:
        root = src
        files = sorted(p for p in src.rglob("*") if p.suffix.lower() in PPT_SUFFIXES and not p.name.startswith(("~$", ".")))
    todo = []
    for p in files:
        target = out / "files" / _json_name(p.relative_to(root))
        if target.exists() and not args.force:
            try:
                old = json.loads(target.read_text(encoding="utf-8"))
                if "error" not in old and old.get("bytes") == p.stat().st_size and old.get("mtime") == int(p.stat().st_mtime):
                    continue
            except Exception:
                pass
        todo.append(p)
    print(f"{len(files)} files, {len(todo)} to analyze -> {out}")
    t0 = time.time()
    failed = 0
    with ProcessPoolExecutor(max_workers=args.workers) as ex:
        futs = [ex.submit(_one, str(p), str(root), str(out)) for p in todo]
        for i, f in enumerate(as_completed(futs), 1):
            try:
                rel, err, n = f.result()
            except Exception as exc:
                failed += 1
                print(f"[{i}/{len(todo)}] CRASH {exc}", file=sys.stderr)
                continue
            if err:
                failed += 1
            print(f"[{i}/{len(todo)}] {'FAIL' if err else 'ok  '} {n:>3} slides  {rel}" + (f"  ({err[:80]})" if err else ""))
    csv_path, md_path = summarize(out, args.scope)
    print(f"done in {time.time() - t0:.0f}s, failed {failed}\n{csv_path}\n{md_path}")
    return 1 if failed else 0


def cmd_summarize(args) -> int:
    for p in summarize(Path(args.out).expanduser().resolve(), args.scope):
        print(p)
    return 0


def _stem(path: Path) -> str:
    """Work folder name: Windows forbids <>:"/\\|?* and drops trailing spaces and dots."""
    return re.sub(r'[<>:"/\\|?*]', "_", unicodedata.normalize("NFC", Path(path).stem)).rstrip(" .") or "deck"


def _find_nfc(p: Path) -> Path:
    """NTFS matches names exactly; a copy whose Korean name arrived decomposed (NFD) still counts."""
    if p.exists() or not p.parent.is_dir():
        return p
    want = unicodedata.normalize("NFC", p.name)
    return next((c for c in p.parent.iterdir() if unicodedata.normalize("NFC", c.name) == want), p)


def _default_template() -> Path:
    import yaml
    scope = REPO_ROOT / "config/scope.yaml"
    if scope.exists():
        data = yaml.safe_load(scope.read_text(encoding="utf-8")) or {}
        for t in data.get("templates") or []:
            if str(t).endswith(".pptx"):
                return Path(t) if Path(t).is_absolute() else REPO_ROOT / t
    return REPO_ROOT / "local/templates/24.상반기.pptx"


def _fit_text(fit: dict | None) -> str:
    """Short fit note for the plan table: characters and estimated lines, or the font size a quote was set to."""
    if not fit:
        return ""
    if "est_lines" in fit:
        return f"({fit.get('chars', '?')}자/{fit['est_lines']}줄)"
    if "sz" in fit:
        return f"({fit.get('chars', '?')}자/{fit['sz']:g}pt)"
    return f"({fit.get('chars', '?')}자)"


def _print_plan(plan: dict) -> None:
    by_src = {}
    for op in plan["ops"]:
        by_src.setdefault(op["src"], []).append(op)
    hymn_by_src = {}
    for op in plan.get("hymn_ops", []):
        hymn_by_src.setdefault(op["src"], []).append(op)
    print(f"{'src':>3} {'kind':<20} {'rule':<5} ops")
    for row in plan["slides"]:
        ops = by_src.get(row["src"], [])
        desc = " ".join(("copy" if o["op"] == "copy" else "shapes" if o["op"] == "shapes" else f"T{o['template_slide']}")
                        + _fit_text(o.get("fit")) for o in ops)
        lyrics = hymn_by_src.get(row["src"], [])
        if lyrics:
            desc += f"  +찬송 자막 {len(lyrics)}장"
        flags = sorted({f for o in ops for f in o.get("flags", [])} | {f for o in lyrics for f in o.get("flags", [])})
        ai = f"  ai:{row['ai']['kind']}@{row['ai']['confidence']:.2f}" if row.get("ai") else ""
        print(f"{row['src']:>3} {row['kind']:<20} {row['rule']:<5} {desc or 'skip'}" + (f"  [{', '.join(flags)}]" if flags else "") + ai)
    print("stats:", plan["stats"])
    _print_hymn_note(plan.get("hymn") or {})


def _print_hymn_note(why: dict) -> None:
    """A hymn run nobody could identify would otherwise leave no trace at all in the output."""
    runs, matched = why.get("runs", 0), why.get("matched", 0)
    if not runs or matched == runs:
        return
    head = f"찬송 구간 {runs}개 중 {matched}개만 곡을 찾았습니다."
    if not why.get("template"):
        print(f"{head} 템플릿에 찬송 자막 슬라이드가 없습니다 (catalog를 다시 만드세요).")
    elif not why.get("index"):
        print(f"{head} 곡 색인이 없습니다: local/hymns/index.json을 복사하거나 `broadcast-ppt hymns`를 실행하세요.")
    elif not why.get("songs"):
        print(f"{head} 색인에 곡이 없습니다: local/hymns에 곡 파일을 넣고 `broadcast-ppt hymns`를 실행하세요.")
    elif not why.get("reader"):
        print(f"{head} 이 PC에서는 악보 글자를 읽을 수 없습니다. "
              f"Windows는 설정 > 시간 및 언어 > 언어에 한국어가 있어야 하고, 다시 `uv sync`를 실행하세요.")
    else:
        print(f"{head} 나머지는 악보에서 가사를 읽지 못했습니다 (검수 표시 deferred:hymn 확인).")


def cmd_plan(args) -> int:
    import yaml
    from .mapping.plan import PlanOptions, build_plan, save_plan
    from .rules import ai
    src = Path(args.source)
    if args.overrides and not Path(args.overrides).is_file():
        print(f"overrides file not found: {args.overrides} (create it first, see README)", file=sys.stderr)
        return 2
    try:
        overrides = yaml.safe_load(Path(args.overrides).read_text(encoding="utf-8-sig")) if args.overrides else None
    except UnicodeDecodeError:
        print(f"{args.overrides} 파일을 UTF-8로 저장해 주세요 (메모장: 다른 이름으로 저장 → 인코딩 UTF-8).", file=sys.stderr)
        return 2
    opts = PlanOptions(slide_size=args.slide_size, scripture_both=not args.no_scripture_both, sermon_material=args.sermon_material,
                       ai_model=ai.enabled())
    if opts.ai_model:
        print(f"ai fallback: unclassified slides go to {opts.ai_model} (names masked)", file=sys.stderr)
    try:
        plan = build_plan(src, Path(args.catalog), Path(args.widths), opts, overrides)
    except ValueError as exc:  # a mistake in the overrides file
        print(f"plan refused: {exc}", file=sys.stderr)
        return 2
    out = Path(args.out) if args.out else Path(args.work) / _stem(src) / "plan.json"
    save_plan(plan, out)
    _print_plan(plan)
    print(out)
    return 0


def cmd_build(args) -> int:
    from .builder.build import build_deck
    from .mapping.plan import load_plan
    plan = load_plan(Path(args.plan))
    out = Path(args.out) if args.out else Path(args.plan).with_name("broadcast.pptx")
    template = _find_nfc(Path(args.template) if args.template else _default_template())
    if not template.is_file():
        print(f"템플릿이 없습니다: {template}\n원본 24.상반기.pptx를 PowerPoint로 다시 저장하지 말고 그대로 이 위치에 복사하거나, --template <경로>로 지정하세요.", file=sys.stderr)
        return 2
    try:
        res = build_deck(plan, template, Path(args.catalog), out, force=args.force, widths_path=Path(args.widths))
    except PermissionError as exc:
        print(f"파일을 열 수 없습니다: {exc.filename}\nPowerPoint에서 {out.name}를 닫고 다시 실행하세요.", file=sys.stderr)
        return 1
    except ValueError as exc:
        if str(exc).startswith("template changed"):
            print(f"템플릿이 원본과 다릅니다: {template}\n원본을 PowerPoint로 다시 저장하지 말고 그대로 복사하세요 (SHA256 {plan['template']['sha256'].upper()}).", file=sys.stderr)
            return 1
        print(f"build refused: {exc} (re-run plan, or pass --force)", file=sys.stderr)
        return 1
    flagged = [r for r in res["ops"] if r["flags"]]
    print(f"built {res['slides']} slides -> {res['out']} ({len(flagged)} with flags)")
    hymn_plan = {**plan, "ops": plan.get("hymn_ops") or []}
    hymn_out = out.with_name("hymn.pptx")
    if hymn_plan["ops"]:  # the lyric captions are a deck of their own; the broadcast deck keeps the score pages
        try:
            hres = build_deck(hymn_plan, template, Path(args.catalog), hymn_out, force=args.force, widths_path=Path(args.widths))
        except PermissionError as exc:
            print(f"파일을 열 수 없습니다: {exc.filename}\nPowerPoint에서 {hymn_out.name}를 닫고 다시 실행하세요.", file=sys.stderr)
            return 1
        print(f"찬송 자막 {hres['slides']}장 -> {hres['out']}")
    if args.validate:
        from .validator.check import validate_deck
        rep = validate_deck(out, plan, Path(args.catalog), Path(args.widths))
        if hymn_plan["ops"]:
            hrep = validate_deck(hymn_out, hymn_plan, Path(args.catalog), Path(args.widths))
            hymn_out.with_name("validation-hymn.json").write_text(json.dumps(hrep, ensure_ascii=False, indent=1), encoding="utf-8")
        out.with_name("validation.json").write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")
        _report_validation(rep)
        if hymn_plan["ops"]:
            print("찬송 자막 파일:", end=" ")
            _report_validation(hrep)
            return 0 if rep["ok"] and hrep["ok"] else 1
        return 0 if rep["ok"] else 1
    return 0


def _report_validation(rep: dict) -> None:
    print(f"validation: critical {rep['counts']['critical']} · warning {rep['counts']['warning']} · info {rep['counts']['info']}")
    for f in rep["findings"]:
        if f["level"] != "info":
            print(f"  {f['level']:<8} slide {f['out_slide']}: {f['check']} — {f['message']}")
    reviews = Counter(f["message"] for f in rep["findings"] if f["level"] == "info")
    if reviews:
        print("  review flags:", dict(reviews))


def cmd_validate(args) -> int:
    from .mapping.plan import load_plan
    from .validator.check import validate_deck
    plan = load_plan(Path(args.plan))
    rep = validate_deck(Path(args.deck), plan, Path(args.catalog), Path(args.widths))
    out = Path(args.report) if args.report else Path(args.deck).with_name("validation.json")
    out.write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")
    _report_validation(rep)
    return 0 if rep["ok"] else 1


def cmd_catalog(args) -> int:
    from .mapping.catalog import build_catalog, dump_catalog
    cat = build_catalog(Path(args.template), Path(args.roles))
    print(f"{len(cat['slides'])} template slides catalogued -> {dump_catalog(cat, Path(args.out))}")
    return 0


def cmd_hymns(args) -> int:
    from .mapping.hymns import build_library, load_index, save_index
    lib = Path(args.library)
    if not lib.exists():
        print(f"찬송 라이브러리 폴더가 없습니다: {lib}", file=sys.stderr)
        return 1
    index = load_index(Path(args.out))
    songs, skipped = build_library(lib)
    index["songs"] = songs
    if args.forget:
        index["images"] = {}
    known = sum(1 for v in index["images"].values() if v)
    print(f"{len(songs)}곡 색인 -> {save_index(index, Path(args.out))}"
          f" (악보 이미지 {known}장 확인됨, 미매칭 {len(index['images']) - known}장)")
    for line in skipped:
        print(f"  건너뜀: {line}")
    return 0


def cmd_calibrate(args) -> int:
    from .dataset.calibrate import calibrate
    paths = []
    for p in map(Path, args.paths):
        paths += sorted(x for x in p.rglob("*.pptx") if not x.name.startswith(("~$", "."))) if p.is_dir() else [p]
    data = calibrate(paths, Path(args.out))
    print(f"{len(paths)} decks -> {args.out}: width_scale {data['width_scale']}, lines exact {data['stats'].get('lines_exact')}")
    return 0


def cmd_run(args) -> int:
    build_out, args.out = args.out, None  # --out names the broadcast deck; the plan stays in the work folder
    if cmd_plan(args):
        return 2
    args.plan = str(Path(args.work) / _stem(Path(args.source)) / "plan.json")
    args.out = build_out
    return cmd_build(args)


def _gui_message(reports: list[tuple[str, dict]], outs: list[Path]) -> tuple[bool, str]:
    """The operator reads only this box: where the file went, and which slides to look at.

    info findings are the owner's review flags (dozens a week) and stay in validation.json; the box lists
    only what the operator can act on by opening the deck.
    """
    lines, bad = [], False
    for name, rep in reports:
        crit = sorted({f["out_slide"] for f in rep["findings"] if f["level"] == "critical" and f["out_slide"]})
        rest = sorted({f["out_slide"] for f in rep["findings"] if f["level"] == "warning" and f["out_slide"]})
        if crit:
            bad = True
            lines.append(f"{name}: 방송에 쓰면 안 되는 장 {', '.join(map(str, crit))}")
        if rest:
            lines.append(f"{name}: 확인할 장 {', '.join(map(str, rest))}")
    head = "만들었지만 방송에 쓰면 안 됩니다. 담당자에게 연락하세요." if bad else "변환이 끝났습니다."
    saved = "\n".join(f"  {p.name}" for p in outs)
    tail = "\n\n" + "\n".join(lines) if lines else ""
    return bad, f"{head}\n\n원본과 같은 폴더에 저장했습니다:\n{saved}{tail}"


def _keep_source(src: Path, work: Path) -> Path:
    """Move the original into its work folder so the month's review finds deck, plan and logs together.

    Returns where the original is now; a deck still open in PowerPoint stays where it was.
    """
    kept = work / src.name
    if kept.resolve() == src.resolve():  # re-run picked the copy already kept here
        return src
    try:
        kept.unlink(missing_ok=True)  # Windows rename refuses an existing target
        shutil.move(src, kept)
    except OSError:
        return src
    return kept


def cmd_gui(args) -> int:
    """Double-click entry point: pick a deck, leave the result beside it, keep the JSON in the work folder."""
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()
    picked = filedialog.askopenfilename(title="변환할 예배 PPT를 고르세요", filetypes=[("PowerPoint", "*.pptx")])
    if not picked:
        return 0
    src = Path(picked)
    work = Path(args.work) / _stem(src)
    work.mkdir(parents=True, exist_ok=True)
    for name in ("broadcast.pptx", "hymn.pptx", "validation.json", "validation-hymn.json"):
        (work / name).unlink(missing_ok=True)  # a failed run must not hand out last week's deck or report
    overrides = work / "overrides.yaml"  # no place to type --overrides, so the work folder decides
    run_args = argparse.Namespace(source=str(src), work=args.work, catalog=CATALOG, widths=WIDTHS,
                                  slide_size="template", no_scripture_both=False, sermon_material="lower-third",
                                  overrides=str(overrides) if overrides.is_file() else None,
                                  out=None, template=None, force=False, validate=True)
    log = io.StringIO()
    try:
        # the plan table, the validation line and the hymn note are printed to stdout: keep both streams
        with contextlib.redirect_stdout(log), contextlib.redirect_stderr(log):
            cmd_run(run_args)
    except Exception as exc:
        traceback.print_exc(file=log)
        log.write(f"{exc}\n")
    sys.stderr.write(log.getvalue())  # the console window behind the dialog stays the full log
    (work / "gui.log").write_text(log.getvalue(), encoding="utf-8")  # the console is gone by the time anyone asks
    if not (work / "validation.json").is_file():
        messagebox.showerror("변환하지 못했습니다", f"{log.getvalue().strip()}\n\n담당자에게 연락하세요.")
        return 1
    outs, reports = [], []
    for built, tag, rep in (("broadcast.pptx", "방송용", "validation.json"), ("hymn.pptx", "찬송자막", "validation-hymn.json")):
        if not (work / built).is_file():
            continue
        dest = src.with_name(f"{src.stem}_{tag}.pptx")
        try:
            dest.unlink(missing_ok=True)  # Windows rename refuses an existing target
            shutil.move(work / built, dest)
        except OSError:
            messagebox.showerror("저장하지 못했습니다", f"PowerPoint에서 {dest.name}을(를) 닫고 다시 해 주세요.")
            return 1
        outs.append(dest)
        if (work / rep).is_file():
            reports.append((dest.name, json.loads((work / rep).read_text(encoding="utf-8"))))
    bad, text = _gui_message(reports, outs)
    kept = _keep_source(src, work)
    if kept != src:
        text += f"\n\n원본은 결과 폴더로 옮겼습니다:\n  {kept}"
    (messagebox.showerror if bad else messagebox.showinfo)("확인 필요" if bad else "변환 완료", text)
    return 1 if bad else 0


def _utf8_console() -> None:
    """Redirected output on Korean Windows is cp949, which cannot print every character (same guard as dongdongs)."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass


def main(argv: list[str] | None = None) -> int:
    _utf8_console()
    ap = argparse.ArgumentParser(prog="broadcast-ppt")
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("analyze", help="inventory the structure of .pptx files")
    a.add_argument("input", help="a .pptx file or a folder (searched recursively)")
    a.add_argument("--out", default=str(REPO_ROOT / "local/analysis"))
    a.add_argument("--workers", type=int, default=max((os.cpu_count() or 2) - 2, 1))
    a.add_argument("--force", action="store_true", help="re-analyze files whose JSON is up to date")
    a.add_argument("--scope", default=str(REPO_ROOT / "config/scope.yaml"), help="scope file for the summary (missing file = all decks)")
    a.set_defaults(func=cmd_analyze)
    s = sub.add_parser("summarize", help="rebuild summary.csv and clusters.md from existing JSON")
    s.add_argument("--out", default=str(REPO_ROOT / "local/analysis"))
    s.add_argument("--scope", default=str(REPO_ROOT / "config/scope.yaml"))
    s.set_defaults(func=cmd_summarize)
    def plan_args(p):
        p.add_argument("source")
        p.add_argument("--work", default=str(REPO_ROOT / "local/work"))
        p.add_argument("--catalog", default=CATALOG)
        p.add_argument("--widths", default=WIDTHS)
        p.add_argument("--slide-size", choices=["template", "source"], default="template")
        p.add_argument("--no-scripture-both", action="store_true")
        p.add_argument("--sermon-material", choices=["lower-third", "passthrough"], default="lower-third")
        p.add_argument("--overrides")
        p.add_argument("--out")
    pl = sub.add_parser("plan", help="classify an original deck and write plan.json")
    plan_args(pl)
    pl.set_defaults(func=cmd_plan)
    b = sub.add_parser("build", help="build the broadcast deck from plan.json")
    b.add_argument("plan")
    b.add_argument("--out")
    b.add_argument("--template")
    b.add_argument("--catalog", default=CATALOG)
    b.add_argument("--force", action="store_true")
    b.add_argument("--widths", default=WIDTHS)
    b.add_argument("--no-validate", dest="validate", action="store_false")
    b.set_defaults(func=cmd_build)
    v = sub.add_parser("validate", help="check a built deck against its plan")
    v.add_argument("deck")
    v.add_argument("--plan", required=True)
    v.add_argument("--catalog", default=CATALOG)
    v.add_argument("--widths", default=WIDTHS)
    v.add_argument("--report")
    v.set_defaults(func=cmd_validate)
    r = sub.add_parser("run", help="plan + build")
    plan_args(r)
    r.add_argument("--template")
    r.add_argument("--force", action="store_true")
    r.set_defaults(func=cmd_run, validate=True)
    g = sub.add_parser("gui", help="pick a deck in a file dialog; the result lands next to it (double-click launcher)")
    g.add_argument("--work", default=str(REPO_ROOT / "local/work"))
    g.set_defaults(func=cmd_gui)
    c = sub.add_parser("catalog", help="rebuild config/templates.yaml from the template and config/roles.yaml")
    c.add_argument("template", nargs="?", default=str(REPO_ROOT / "local/templates/24.상반기.pptx"))
    c.add_argument("--roles", default=str(REPO_ROOT / "config/roles.yaml"))
    c.add_argument("--out", default=CATALOG)
    c.set_defaults(func=cmd_catalog)
    hy = sub.add_parser("hymns", help="rebuild the hymn lyric index from local/hymns")
    hy.add_argument("--library", default=str(REPO_ROOT / "local/hymns"))
    hy.add_argument("--out", default=str(REPO_ROOT / "local/hymns/index.json"))
    hy.add_argument("--forget", action="store_true", help="drop the learned score-image map as well")
    hy.set_defaults(func=cmd_hymns)
    cal = sub.add_parser("calibrate", help="refit character widths (config/widths.yaml) from broadcast decks")
    cal.add_argument("paths", nargs="+", help=".pptx files or folders")
    cal.add_argument("--out", default=WIDTHS)
    cal.set_defaults(func=cmd_calibrate)
    args = ap.parse_args(argv)
    try:
        return args.func(args)
    except KeyboardInterrupt:
        return 130
    except Exception as exc:  # the operator reads the last line; the traceback above it is for fixing the code
        traceback.print_exc()
        print(f"\n오류: {exc}\n파일 경로가 맞는지, PowerPoint가 그 파일을 열고 있지 않은지 확인하세요.", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
