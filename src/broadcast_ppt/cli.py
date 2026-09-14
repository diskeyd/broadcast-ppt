"""broadcast-ppt command line."""

from __future__ import annotations

import argparse
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
                if old.get("bytes") == p.stat().st_size and old.get("mtime") == int(p.stat().st_mtime):
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
    csv_path, md_path = summarize(out)
    print(f"done in {time.time() - t0:.0f}s, failed {failed}\n{csv_path}\n{md_path}")
    return 0


def cmd_summarize(args) -> int:
    for p in summarize(Path(args.out).expanduser().resolve()):
        print(p)
    return 0


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="broadcast-ppt")
    sub = ap.add_subparsers(dest="cmd", required=True)
    a = sub.add_parser("analyze", help="inventory the structure of .pptx files")
    a.add_argument("input", help="a .pptx file or a folder (searched recursively)")
    a.add_argument("--out", default="local/analysis")
    a.add_argument("--workers", type=int, default=max((os.cpu_count() or 2) - 2, 1))
    a.add_argument("--force", action="store_true", help="re-analyze files whose JSON is up to date")
    a.set_defaults(func=cmd_analyze)
    s = sub.add_parser("summarize", help="rebuild summary.csv and clusters.md from existing JSON")
    s.add_argument("--out", default="local/analysis")
    s.set_defaults(func=cmd_summarize)
    args = ap.parse_args(argv)
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
