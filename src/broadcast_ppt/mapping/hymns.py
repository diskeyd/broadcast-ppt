"""Hymn runs: find the song in the chroma-key library and read its lyric lines.

A hymn is a run of score images with no text on the slide, so the song is identified from the image:
a cached sha1 -> song map, filled the first time a score is seen by reading the title and the lyric
lines off the picture with the system text recogniser (Vision on the Mac, Windows.Media.Ocr on the
broadcast PC). The library pptx files are never copied — only their lyric text is used; the broadcast
slides come from the template like every other caption. A song the library does not have is captioned
from the lyric lines read off its score pages.
"""

from __future__ import annotations

import asyncio
import difflib
import functools
import hashlib
import json
import re
import subprocess
import sys
import unicodedata
from dataclasses import asdict, dataclass, field
from pathlib import Path

from pptx import Presentation

from ..ooxml import REPO_ROOT

LIBRARY = REPO_ROOT / "local/hymns"
INDEX = LIBRARY / "index.json"
OCR_SCRIPT = REPO_ROOT / "scripts/ocr_vision.js"
LYRIC_MIN_PT = 40.0   # library lyrics are 44pt; smaller boxes are a creed, a notice or a section card
LABEL_MAX_PT = 24.0   # the song label sits under the lyric at 20-24pt
MIN_LINES = 3         # anything shorter is a stray pair of boxes, not a song
MIN_KEY = 6           # a shorter lyric line matches too many songs to identify one
TITLE_RATIO = 0.8
NUMBER_RE = re.compile(r"^\s*(\d{1,3})\s*[_장.]")
# a lyric line as the recogniser reads it off a score: Korean, a few words long, dashes for held notes.
# Two of them are what tells a score page from a photo of a banner or a sign.
# a weekly deck runs the songs together with the rest of the service, so a song can be followed by the
# slide naming who prays. That is a person, never a lyric: it ends the song and never enters the index.
PERSON_ROLE = re.compile(r"[가-힣]{2,4}\s*(?:권사|집사|장로|목사|전도사|사모|성도)\s*$")
LYRIC_LINE = re.compile(r"^[가-힣][가-힣0-9 .,!?~·\-–—()'\"]{3,29}$")


def key(text: str) -> str:
    """Comparison form: OCR drops spaces and punctuation differently from the library text."""
    return re.sub(r"[^가-힣0-9a-z]", "", unicodedata.normalize("NFC", text).lower())


@dataclass
class Song:
    id: str
    file: str
    title: str
    number: int | None
    label: str = ""      # what the first broadcast slide shows: "108장 그 어린 주 예수"
    lines: list[str] = field(default_factory=list)


def _boxes(slide) -> list[tuple[float, str]]:
    out = []
    for sh in slide.shapes:
        if not sh.has_text_frame or not sh.text.strip():
            continue
        sizes = [r.font.size.pt for p in sh.text_frame.paragraphs for r in p.runs if r.font.size]
        out.append((max(sizes) if sizes else 0.0, sh.text.strip()))
    return out


def read_deck(path: Path, rel: str) -> list[Song]:
    """Songs in one library deck. A weekly deck holds several: each starts at the slide with a label box."""
    songs: list[Song] = []
    cur: Song | None = None
    for slide in Presentation(path).slides:
        boxes = _boxes(slide)
        if not boxes:
            continue
        top = max(pt for pt, _ in boxes)
        if top < LYRIC_MIN_PT:
            continue  # creed, notice or section card inside a weekly deck
        lyric = [t for pt, t in boxes if pt == top]
        label = [t for pt, t in boxes if pt < top and pt <= LABEL_MAX_PT]
        if label:
            m = NUMBER_RE.match(label[0]) or NUMBER_RE.match(Path(rel).name)
            cur = Song(id=f"{rel}#{len(songs) + 1}", file=rel, title=label[0].strip(),
                       number=int(m.group(1)) if m else None)
            songs.append(cur)
        if cur is None:
            continue  # an intro slide before the first song (weekly decks open with one)
        lines = [ln.strip() for ln in lyric[0].splitlines() if ln.strip()]
        if any(PERSON_ROLE.match(ln) for ln in lines):
            cur = None  # a name after the song: the song ended here and the name is not kept
            continue
        cur.lines += lines
    songs = [s for s in songs if len(s.lines) >= MIN_LINES]
    if not songs:  # a deck whose first slide carries no label: name it after the file
        return _from_filename(path, rel)
    return songs


def _from_filename(path: Path, rel: str) -> list[Song]:
    stem = Path(rel).stem
    m = NUMBER_RE.match(stem)
    song = Song(id=f"{rel}#1", file=rel, title=stem, number=int(m.group(1)) if m else None)
    for slide in Presentation(path).slides:
        boxes = _boxes(slide)
        if boxes:
            top = max(pt for pt, _ in boxes)
            for line in next(t for pt, t in boxes if pt == top).splitlines():
                if line.strip():
                    song.lines.append(line.strip())
    return [song] if len(song.lines) >= MIN_LINES else []


def build_library(lib_dir: Path = LIBRARY) -> tuple[list[dict], list[str]]:
    """Every song in the library, plus the files that could not be read."""
    songs, skipped = [], []
    for p in sorted(Path(lib_dir).rglob("*")):
        if not p.is_file() or p.name.startswith(("~$", ".")):
            continue
        rel = unicodedata.normalize("NFC", p.relative_to(lib_dir).as_posix())
        if p.suffix.lower() == ".ppt":
            skipped.append(f"{rel}: 구형식 .ppt (PowerPoint에서 .pptx로 저장하면 색인됨)")
            continue
        if p.suffix.lower() != ".pptx":
            continue
        try:
            found = read_deck(p, rel)
        except Exception as exc:
            skipped.append(f"{rel}: {exc}")
            continue
        # a single-song file names the song better than its own label box does (108_… is labelled "524장")
        if len(found) == 1:
            stem = unicodedata.normalize("NFC", p.stem)
            m = NUMBER_RE.match(stem)
            found[0].number = int(m.group(1)) if m else None
            found[0].title = stem[m.end():].strip() if m else stem
        for song in found:
            song.label = f"{song.number}장 {song.title}" if song.number else song.title
        songs += [asdict(s) for s in found]
    return songs, skipped


def load_index(path: Path = INDEX) -> dict:
    path = Path(path)
    if not path.exists():
        return {"songs": [], "images": {}}
    data = json.loads(path.read_text(encoding="utf-8"))
    data.setdefault("songs", [])
    data.setdefault("images", {})
    return data


def save_index(data: dict, path: Path = INDEX) -> Path:
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding="utf-8")
    return path


# --- identifying a score image -------------------------------------------------------------------

def ocr_available() -> bool:
    """Vision on the Mac, the Windows OCR engine (Korean, else the user's languages) on the broadcast PC."""
    if sys.platform == "darwin":
        return OCR_SCRIPT.exists()
    return sys.platform == "win32" and _win_engine() is not None


@functools.cache
def _win_engine():
    try:
        from winrt.windows.globalization import Language
        from winrt.windows.media.ocr import OcrEngine
    except ImportError:
        return None
    return OcrEngine.try_create_from_language(Language("ko")) or OcrEngine.try_create_from_user_profile_languages()


def reader_language() -> str | None:
    """The language the score reader reads in, for the plan's hymn note ("ko" is what the lyrics need)."""
    if sys.platform == "darwin":
        return "ko-KR (Vision)" if OCR_SCRIPT.exists() else None
    engine = _win_engine() if sys.platform == "win32" else None
    return engine.recognizer_language.language_tag if engine is not None else None


def _enlarged(image: Path) -> Path:
    """A score page is small (960x720 on a 4:3 deck) and the Windows engine misses lyric lines that small;
    twice the size reads them. The engine's own size ceiling is kept."""
    from PIL import Image
    with Image.open(image) as im:
        scale = min(2.0, 2600 / max(im.size))
        if scale <= 1.0:
            return image
        out = image.with_name(image.stem + "@2x.png")
        im.convert("RGB").resize((int(im.width * scale), int(im.height * scale)), Image.LANCZOS).save(out)
    return out


async def _win_read(engine, image: Path) -> list[str]:
    from winrt.windows.graphics.imaging import BitmapDecoder
    from winrt.windows.storage import FileAccessMode, StorageFile
    file = await StorageFile.get_file_from_path_async(str(Path(image).resolve()))
    stream = await file.open_async(FileAccessMode.READ)
    bitmap = await (await BitmapDecoder.create_async(stream)).get_software_bitmap_async()
    return [line.text for line in (await engine.recognize_async(bitmap)).lines]


def ocr(image: Path) -> list[str]:
    """Text lines on the picture, top to bottom. Empty where no recogniser is available."""
    if not ocr_available():
        return []
    if sys.platform == "win32":
        lines = asyncio.run(_win_read(_win_engine(), _enlarged(Path(image))))
    else:
        r = subprocess.run(["osascript", "-l", "JavaScript", str(OCR_SCRIPT), str(image)],
                           capture_output=True, encoding="utf-8", errors="replace")
        if r.returncode != 0:
            return []
        lines = r.stdout.splitlines()
    return [ln.strip() for ln in lines if ln.strip()]


def clean_lyric(line: str) -> str:
    """A lyric line as it is sung, from the line the recogniser read under the staff: repeat signs read
    as ': ' go, a held syllable (허-락) closes up, held notes between words become one space."""
    t = re.sub(r"^[\s:|.]+", "", line)
    t = re.sub(r"(?<=[가-힣])[-–—~]+(?=[가-힣])", "", t)
    t = re.sub(r"\s*[-–—~]+\s*", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _line_map(songs: list[dict]) -> dict[str, list[tuple[str, int]]]:
    out: dict[str, list[tuple[str, int]]] = {}
    for s in songs:
        for i, line in enumerate(s["lines"]):
            k = key(line)
            if len(k) >= MIN_KEY:
                out.setdefault(k, []).append((s["id"], i))
    return out


def match(songs: list[dict], text: list[str]) -> tuple[str, int, int] | None:
    """(song id, first, last matching lyric line) for the OCR text of one score page, or None.

    Two lyric lines of the same song, or one line plus a title that reads like the song's, is the gate:
    a single shared line ("주의 은혜라") shows up in a dozen songs."""
    lines = _line_map(songs)
    hits: dict[str, set[int]] = {}
    for raw in text:
        k = key(re.sub(r"^\s*\d+[.)]", "", raw))
        if len(k) < MIN_KEY:
            continue
        for lk, places in lines.items():
            if k in lk or lk in k:
                for song_id, i in places:
                    hits.setdefault(song_id, set()).add(i)
    if not hits:
        return None
    titles = {s["id"]: key(s["title"]) for s in songs}
    def score(song_id: str) -> tuple[int, float]:
        best = max((difflib.SequenceMatcher(None, titles[song_id], key(t)).ratio() for t in text), default=0.0)
        return len(hits[song_id]), best
    song_id = max(hits, key=score)
    n, ratio = score(song_id)
    if n >= 2 or (n == 1 and ratio >= TITLE_RATIO):
        return song_id, min(hits[song_id]), max(hits[song_id])
    return None


# --- a run of score slides in the source deck ------------------------------------------------------

@dataclass
class HymnMatch:
    song: str
    title: str
    label: str
    lines: list[str]
    start: int
    end: int
    guessed_end: bool
    ocr: bool = False  # captioned from the lyric lines read off the score, not from the library


def _images(src: Path, indices: list[int]) -> dict[int, bytes]:
    prs = Presentation(src)
    slides = list(prs.slides)
    out = {}
    for i in indices:
        for sh in slides[i - 1].shapes:
            if sh.shape_type == 13:  # PICTURE
                try:
                    out[i] = sh.image.blob
                except Exception:
                    pass
                break
    return out


def _page(index: dict, sha: str, blob: bytes, work: Path) -> tuple[str, int, int] | None:
    """Which song and which lyric lines this score page shows, from the cache or from OCR."""
    if sha in index["images"]:
        hit = index["images"][sha] or {}
        return (hit["song"], hit["line"], hit["last"]) if hit.get("song") else None
    work.mkdir(parents=True, exist_ok=True)
    path = work / f"{sha}.png"
    path.write_bytes(blob)
    text = ocr(path)
    found = match(index["songs"], text)
    # an unmatched page keeps the first lines read, so a run that went wrong shows what the reader saw
    index["images"][sha] = ({"song": found[0], "line": found[1], "last": found[2]} if found
                            else {"song": None, "ocr": sum(1 for t in text if LYRIC_LINE.match(t.strip())), "read": text[:4]})
    return found


@dataclass
class RunResult:
    match: HymnMatch | None
    readable: bool = False  # the score was read but no song in the library fits it


def match_run(src: Path, indices: list[int], index: dict, work: Path) -> RunResult:
    """Identify a run of score slides and the stretch of the song it covers."""
    if not index["songs"]:
        return RunResult(None)
    blobs = _images(Path(src), [indices[0], indices[-1]])
    if indices[0] not in blobs:
        return RunResult(None)
    sha = hashlib.sha1(blobs[indices[0]]).hexdigest()
    first = _page(index, sha, blobs[indices[0]], work)
    readable = (index["images"].get(sha) or {}).get("ocr", 0) >= 2  # lyric-like lines: a score, not a photo
    if first is None:
        return RunResult(None, readable)
    song = next((s for s in index["songs"] if s["id"] == first[0]), None)
    if song is None:  # the library was rebuilt under a cached id
        return RunResult(None)
    last = None
    if len(indices) > 1 and indices[-1] in blobs:
        found = _page(index, hashlib.sha1(blobs[indices[-1]]).hexdigest(), blobs[indices[-1]], work)
        last = found[2] if found and found[0] == first[0] else None
    end = len(song["lines"]) - 1 if last is None else last
    return RunResult(HymnMatch(song=song["id"], title=song["title"], label=song.get("label") or song["title"],
                               lines=song["lines"], start=first[1], end=max(end, first[1]), guessed_end=last is None))


def ocr_run(src: Path, indices: list[int], index: dict, work: Path) -> HymnMatch | None:
    """A score run the library has no song for, captioned from its own pages: the lyric lines read off
    every page, in order. There is no title, so the first caption carries none."""
    blobs = _images(Path(src), indices)
    lines: list[str] = []
    for i in indices:
        if i not in blobs:
            continue
        sha = hashlib.sha1(blobs[i]).hexdigest()
        work.mkdir(parents=True, exist_ok=True)
        path = work / f"{sha}.png"
        path.write_bytes(blobs[i])
        lines += [t for t in map(clean_lyric, ocr(path)) if LYRIC_LINE.match(t)]
    if not lines:
        return None
    return HymnMatch(song="ocr", title=lines[0], label="", lines=lines, start=0, end=len(lines) - 1,
                     guessed_end=False, ocr=True)
