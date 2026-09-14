"""Mirror a Google Drive folder of slide decks into local/input/ via the gws CLI.

Usage:
    uv run python scripts/pull_drive.py <folderId> [--dest local/input] [--list-only] [--workers 3]

- Recurses into subfolders and keeps their names as local folders.
- Downloads .pptx/.ppt (and exports Google Slides to .pptx); other files are only listed.
- Skips files whose md5 already matches (manifest), so re-running after more uploads is incremental.
- gws refuses --output paths outside its working directory, so each download runs with cwd = target folder.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import subprocess
import sys
import unicodedata
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

PPTX = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
WANTED = {PPTX, "application/vnd.ms-powerpoint", "application/haansoftpptx", "application/vnd.google-apps.presentation"}
FOLDER = "application/vnd.google-apps.folder"


def gws_json(args: list[str], cwd: Path | None = None) -> dict:
    r = subprocess.run(["gws", *args], capture_output=True, text=True, cwd=cwd)
    body = r.stdout[r.stdout.find("{"):] if "{" in r.stdout else ""
    data = json.loads(body) if body else {}
    if r.returncode != 0 or "error" in data:
        raise RuntimeError(f"gws {' '.join(args[:3])} failed ({r.returncode}): {(data.get('error') or r.stderr)!s:.300}")
    return data


def list_folder(folder_id: str, rel: Path = Path()) -> list[dict]:
    items, token = [], None
    while True:
        params = {"q": f"'{folder_id}' in parents and trashed=false", "pageSize": 1000,
                  "fields": "nextPageToken,files(id,name,mimeType,size,md5Checksum,modifiedTime)"}
        if token:
            params["pageToken"] = token
        data = gws_json(["drive", "files", "list", "--params", json.dumps(params)])
        for f in data.get("files", []):
            name = unicodedata.normalize("NFC", f["name"]).replace("/", "／")
            if f["mimeType"] == FOLDER:
                items.extend(list_folder(f["id"], rel / name))
            else:
                items.append({**f, "name": name, "rel": str(rel / name)})
        token = data.get("nextPageToken")
        if not token:
            return items


def md5_file(p: Path) -> str:
    h = hashlib.md5()
    with p.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def download(item: dict, dest: Path) -> tuple[str, str]:
    target = dest / item["rel"]
    native = item["mimeType"] == "application/vnd.google-apps.presentation"
    if native:
        target = target.with_suffix(".pptx")
    target.parent.mkdir(parents=True, exist_ok=True)
    part = target.name + ".part"
    if native:
        args = ["drive", "files", "export", "--params", json.dumps({"fileId": item["id"], "mimeType": PPTX}), "--output", part]
    else:
        args = ["drive", "files", "get", "--params", json.dumps({"fileId": item["id"], "alt": "media"}), "--output", part]
    gws_json(args, cwd=target.parent)
    tmp = target.parent / part
    if not native:
        if item.get("size") and tmp.stat().st_size != int(item["size"]):
            tmp.unlink(missing_ok=True)
            raise RuntimeError(f"size mismatch {item['rel']}")
        if item.get("md5Checksum") and md5_file(tmp) != item["md5Checksum"]:
            tmp.unlink(missing_ok=True)
            raise RuntimeError(f"md5 mismatch {item['rel']}")
    tmp.replace(target)
    return item["rel"], str(target)


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("folder_id")
    ap.add_argument("--dest", default="local/input")
    ap.add_argument("--list-only", action="store_true")
    ap.add_argument("--workers", type=int, default=8)
    ap.add_argument("--first", action="append", default=[], help="download paths containing this text first (repeatable)")
    args = ap.parse_args()
    dest = Path(args.dest).expanduser().resolve()
    dest.mkdir(parents=True, exist_ok=True)
    manifest_path = dest / "_manifest.json"
    manifest = json.loads(manifest_path.read_text(encoding="utf-8")) if manifest_path.exists() else {}

    items = list_folder(args.folder_id)
    wanted = [i for i in items if i["mimeType"] in WANTED or i["name"].lower().endswith((".pptx", ".ppt"))]
    skipped = [i for i in items if i not in wanted]
    total = sum(int(i.get("size", 0)) for i in wanted)
    print(f"remote: {len(items)} files, {len(wanted)} decks ({total / 1e9:.2f} GB), {len(skipped)} other files not downloaded")
    for i in skipped:
        print(f"  skip  {i['mimeType']:<40} {i['rel']}")
    if args.list_only:
        return 0

    todo = []
    for i in wanted:
        local = dest / i["rel"]
        known = manifest.get(i["id"])
        if local.exists() and local.stat().st_size == int(i.get("size", -1)):
            if known and known.get("md5") == i.get("md5Checksum"):
                continue
            if i.get("md5Checksum") and md5_file(local) == i["md5Checksum"]:
                manifest[i["id"]] = {"rel": i["rel"], "md5": i["md5Checksum"], "size": i.get("size"), "modified": i.get("modifiedTime")}
                continue
        todo.append(i)
    # priority folders first (e.g. 미편집), then small files first so analysis can start early
    todo.sort(key=lambda i: (not any(k in i["rel"] for k in args.first), int(i.get("size", 0))))
    print(f"to download: {len(todo)} ({sum(int(i.get('size', 0)) for i in todo) / 1e9:.2f} GB)")

    errors = 0
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(download, i, dest): i for i in todo}
        for n, f in enumerate(as_completed(futs), 1):
            i = futs[f]
            try:
                f.result()
                manifest[i["id"]] = {"rel": i["rel"], "md5": i.get("md5Checksum"), "size": i.get("size"), "modified": i.get("modifiedTime")}
                print(f"[{n}/{len(todo)}] ok   {i['rel']}", flush=True)
            except Exception as exc:
                errors += 1
                print(f"[{n}/{len(todo)}] FAIL {i['rel']}: {exc}", file=sys.stderr, flush=True)
            if n % 10 == 0:
                manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=1), encoding="utf-8")
    print(f"done: {len(todo) - errors} downloaded, {errors} failed")
    return 1 if errors else 0


if __name__ == "__main__":
    raise SystemExit(main())
