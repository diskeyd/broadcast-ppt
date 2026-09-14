"""broadcast-ppt command line."""

from __future__ import annotations

import argparse
from collections import Counter
import json
import os
import sys
import time
import unicodedata
from concurrent.futures import ProcessPoolExecutor, as_completed
from pathlib import Path

from .analyzer.core import analyze_file
from .analyzer.summary import summarize

PPT_SUFFIXES = {".pptx", ".ppt"}


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
    return 0


def cmd_summarize(args) -> int:
    for p in summarize(Path(args.out).expanduser().resolve(), args.scope):
        print(p)
    return 0


def _stem(path: Path) -> str:
    return unicodedata.normalize("NFC", Path(path).stem)


def _default_template() -> Path:
    import yaml
    scope = Path("config/scope.yaml")
    if scope.exists():
        data = yaml.safe_load(scope.read_text(encoding="utf-8")) or {}
        for t in data.get("templates") or []:
            if str(t).endswith(".pptx"):
                return Path(t)
    return Path("local/templates/24.상반기.pptx")


def _print_plan(plan: dict) -> None:
    by_src = {}
    for op in plan["ops"]:
        by_src.setdefault(op["src"], []).append(op)
    print(f"{'src':>3} {'kind':<20} {'rule':<5} ops")
    for row in plan["slides"]:
        ops = by_src.get(row["src"], [])
        desc = " ".join(("copy" if o["op"] == "copy" else f"T{o['template_slide']}") + (f"({o['fit']['chars']}자/{o['fit']['est_lines']}줄)" if o.get("fit") else "") for o in ops)
        flags = sorted({f for o in ops for f in o.get("flags", [])})
        print(f"{row['src']:>3} {row['kind']:<20} {row['rule']:<5} {desc or 'skip'}" + (f"  [{', '.join(flags)}]" if flags else ""))
    print("stats:", plan["stats"])


def cmd_plan(args) -> int:
    import yaml
    from .mapping.plan import PlanOptions, build_plan, save_plan
    src = Path(args.source)
    overrides = yaml.safe_load(Path(args.overrides).read_text(encoding="utf-8")) if args.overrides else None
    opts = PlanOptions(slide_size=args.slide_size, scripture_both=not args.no_scripture_both, sermon_material=args.sermon_material)
    plan = build_plan(src, Path(args.catalog), Path(args.widths), opts, overrides)
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
    res = build_deck(plan, Path(args.template) if args.template else _default_template(), Path(args.catalog), out, force=args.force)
    flagged = [r for r in res["ops"] if r["flags"]]
    print(f"built {res['slides']} slides -> {res['out']} ({len(flagged)} with flags)")
    if getattr(args, "validate", True):
        from .validator.check import validate_deck
        rep = validate_deck(out, plan, Path(args.catalog), Path(getattr(args, "widths", "config/widths.yaml")))
        out.with_name("validation.json").write_text(json.dumps(rep, ensure_ascii=False, indent=1), encoding="utf-8")
        _report_validation(rep)
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
    if args.render:
        print("  render: skipped (no renderer on this machine; Windows PowerPoint check is a later phase)")
    return 0 if rep["ok"] else 1


def cmd_run(args) -> int:
    args.out = None
    cmd_plan(args)
    args.plan = str(Path(args.work) / _stem(Path(args.source)) / "plan.json")
    return cmd_build(args)


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="broadcast-ppt")
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("analyze", help="inventory the structure of .pptx files")
    a.add_argument("input", help="a .pptx file or a folder (searched recursively)")
    a.add_argument("--out", default="local/analysis")
    a.add_argument("--workers", type=int, default=max((os.cpu_count() or 2) - 2, 1))
    a.add_argument("--force", action="store_true", help="re-analyze files whose JSON is up to date")
    a.add_argument("--scope", default="config/scope.yaml", help="scope file for the summary (missing file = all decks)")
    a.set_defaults(func=cmd_analyze)
    s = sub.add_parser("summarize", help="rebuild summary.csv and clusters.md from existing JSON")
    s.add_argument("--out", default="local/analysis")
    s.add_argument("--scope", default="config/scope.yaml")
    s.set_defaults(func=cmd_summarize)
    def plan_args(p):
        p.add_argument("source")
        p.add_argument("--work", default="local/work")
        p.add_argument("--catalog", default="config/templates.yaml")
        p.add_argument("--widths", default="config/widths.yaml")
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
    b.add_argument("--catalog", default="config/templates.yaml")
    b.add_argument("--force", action="store_true")
    b.add_argument("--widths", default="config/widths.yaml")
    b.add_argument("--no-validate", dest="validate", action="store_false")
    b.set_defaults(func=cmd_build)
    v = sub.add_parser("validate", help="check a built deck against its plan")
    v.add_argument("deck")
    v.add_argument("--plan", required=True)
    v.add_argument("--catalog", default="config/templates.yaml")
    v.add_argument("--widths", default="config/widths.yaml")
    v.add_argument("--report")
    v.add_argument("--render", action="store_true")
    v.set_defaults(func=cmd_validate)
    r = sub.add_parser("run", help="plan + build")
    plan_args(r)
    r.add_argument("--template")
    r.add_argument("--force", action="store_true")
    r.set_defaults(func=cmd_run)
    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
