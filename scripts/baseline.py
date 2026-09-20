"""Classify every analyzed deck and record what the rules decided, to compare across a change.

A rule change touches thousands of slides that are currently correct, and one week's deck cannot
show that. This reads the analysis JSON already in local/analysis/files/ (no .pptx is opened) and
writes one row per slide: rule, kind, how many text boxes the rules saw, and how many characters
reached the fields the mapper turns into captions. Slide text never leaves the process.

Usage:
  uv run python scripts/baseline.py dump before                 # -> local/work/baseline-before.json
  uv run python scripts/baseline.py diff before after
"""

from __future__ import annotations

import argparse
import json
import sys
from collections import Counter
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "src"))

from broadcast_ppt.rules.classify import classify_deck, view_of, DeckContext  # noqa: E402

ANALYSIS = ROOT / "local/analysis/files"
OUT = ROOT / "local/work"


def _body_chars(fields: dict) -> int:
    """Characters the mapper can turn into captions: the body plus every extra passage."""
    n = len(fields.get("body") or "")
    return n + sum(len(e.get("body") or "") for e in (fields.get("extra") or []))


def collect() -> list[dict]:
    rows = []
    for path in sorted(ANALYSIS.glob("*.json")):
        try:
            deck = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as exc:
            print(f"skipped {path.name}: {exc}", file=sys.stderr)
            continue
        try:
            ctx = DeckContext.from_deck(deck)
            decisions = classify_deck(deck)
        except Exception as exc:  # a malformed deck must not stop the sweep
            print(f"skipped {path.name}: {type(exc).__name__}: {exc}", file=sys.stderr)
            continue
        slides = [s for s in deck.get("slides", []) if "error" not in s]
        for slide, d in zip(slides, decisions):
            v = view_of(slide, ctx)
            rows.append({"deck": path.name, "src": slide["index"], "kind": str(d.kind), "rule": d.rule_id,
                         "boxes": len(v.boxes), "box_chars": sum(len(b["text"]) for b in v.boxes),
                         "body_chars": _body_chars(d.fields), "flags": sorted(f for f in d.flags if f)})
    return rows


def cmd_dump(args) -> int:
    rows = collect()
    OUT.mkdir(parents=True, exist_ok=True)
    out = OUT / f"baseline-{args.label}.json"
    out.write_text(json.dumps(rows, ensure_ascii=False, indent=1), encoding="utf-8")
    rules = Counter(r["rule"] for r in rows)
    print(f"{len(rows)} slides from {len({r['deck'] for r in rows})} decks -> {out}")
    print(f"body chars {sum(r['body_chars'] for r in rows)} of {sum(r['box_chars'] for r in rows)} in boxes")
    print("rules: " + ", ".join(f"{k} {n}" for k, n in rules.most_common(8)))
    return 0


def cmd_diff(args) -> int:
    before = {(r["deck"], r["src"]): r for r in json.loads((OUT / f"baseline-{args.before}.json").read_text(encoding="utf-8"))}
    after = {(r["deck"], r["src"]): r for r in json.loads((OUT / f"baseline-{args.after}.json").read_text(encoding="utf-8"))}
    moved, lost, reshaped = Counter(), [], []
    gained = 0
    for key, a in after.items():
        b = before.get(key)
        if b is None or (b["rule"], b["body_chars"]) == (a["rule"], a["body_chars"]):
            continue
        if b["rule"] != a["rule"]:
            moved[f"{b['rule']} -> {a['rule']}"] += 1
        gained += a["body_chars"] - b["body_chars"]
        if a["body_chars"] < b["body_chars"]:
            # a slide that changed kind measures its body differently (a reference moving into the
            # tag box leaves the body shorter and is not a loss), so only same-kind drops are real
            (lost if b["kind"] == a["kind"] else reshaped).append((key, b, a))
    print(f"{len(after)} slides, {sum(moved.values())} changed rule, body chars {gained:+}")
    for k, n in moved.most_common():
        print(f"  {k}: {n}")
    print(f"shorter body after a kind change (check by eye): {len(reshaped)}")
    for key, b, a in reshaped[:10]:
        print(f"  {key[0]} src{key[1]}: {b['kind']}/{b['rule']} {b['body_chars']} -> {a['kind']}/{a['rule']} {a['body_chars']}")
    print(f"slides that lost characters: {len(lost)}" + (" <- REGRESSION" if lost else ""))
    for key, b, a in lost[:20]:
        print(f"  {key[0]} src{key[1]}: {b['rule']} {b['body_chars']} -> {a['rule']} {a['body_chars']}")
    return 1 if lost else 0


def main() -> int:
    ap = argparse.ArgumentParser(prog="baseline")
    sub = ap.add_subparsers(dest="cmd", required=True)
    d = sub.add_parser("dump", help="classify every analyzed deck and write a baseline")
    d.add_argument("label", help="name for this baseline, e.g. before / step1")
    d.set_defaults(func=cmd_dump)
    f = sub.add_parser("diff", help="compare two baselines")
    f.add_argument("before")
    f.add_argument("after")
    f.set_defaults(func=cmd_diff)
    args = ap.parse_args()
    return args.func(args)


if __name__ == "__main__":
    raise SystemExit(main())
