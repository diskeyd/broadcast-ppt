"""Field extractors. Pure functions; return None / [] instead of raising."""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

from .books import ABBREV, BOOK_NAMES_BY_LENGTH, BOOKS

ROLES = ["담임목사", "부목사", "안수집사", "목사", "전도사", "강도사", "장로", "권사", "집사", "성도", "사모", "교사", "선교사"]
_ROLE_RE = re.compile(r"([가-힣]{2,4})\s*(" + "|".join(ROLES) + r")(?![가-힣])")
_BOOK_ALT = "|".join(re.escape(b) for b in BOOK_NAMES_BY_LENGTH)
_REF_RE = re.compile(
    r"(?P<book>" + _BOOK_ALT + r")\s*"
    r"(?P<ch>\d{1,3})\s*(?:[:：]|장\s*)\s*"
    r"(?P<v1>\d{1,3})\s*절?"
    r"(?:\s*(?:[-~–—]|부터)\s*(?:(?P<ch2>\d{1,3})\s*(?:[:：]|장\s*))?\s*(?P<v2>\d{1,3})\s*절?)?"
)
_PAGE_RE = re.compile(
    r"(?:(?P<test>구약|신약|구|신)\s*(?:성경)?\s*)?"
    r"(?:[pP]\s*\.?\s*(?P<page>\d{1,4})|(?P<page2>\d{1,4})\s*(?:쪽|페이지))"
)
_QUOTE_RE = re.compile(r"[“\"「『](?P<t>[^”\"」』]{1,60})[”\"」』]")
_ITEM_START_RE = re.compile(r"(?:^|(?<=\s))(?P<no>\d{1,2})\s*[.)]\s*(?=\S)")
_BRACKET_TITLE_RE = re.compile(r"^\s*[\[［](?P<t>[^\]］]{1,30})[\]］]\s*[:：]?\s*")
_VERSE_NO_RE = re.compile(r"(?:^|(?<=\s))(?P<no>\d{1,3})\s(?=\S)")


def norm(s: str) -> str:
    return unicodedata.normalize("NFC", s or "")


def compact(s: str) -> str:
    return re.sub(r"\s+", "", norm(s))


@dataclass(frozen=True)
class ScriptureRef:
    book: str          # full name
    abbrev: str
    testament: str
    chapter: int
    verse_start: int
    verse_end: int
    raw: str

    @property
    def short(self) -> str:
        return f"{self.abbrev}{self.chapter}:{self.verse_start}" + (f"-{self.verse_end}" if self.verse_end != self.verse_start else "")

    @property
    def long(self) -> str:
        return f"{self.book} {self.chapter}:{self.verse_start}" + (f"-{self.verse_end}" if self.verse_end != self.verse_start else "")


def scripture_ref(text: str) -> ScriptureRef | None:
    """First scripture reference in `text` (full names, abbreviations, 장/절, ranges)."""
    for m in _REF_RE.finditer(norm(text)):
        book = m.group("book")
        # single-syllable abbreviations are ambiguous inside words ("시간 5:3"); require a boundary before them
        start = m.start("book")
        prev = m.string[start - 1] if start > 0 else " "
        if len(book) == 1 and prev.strip() and re.match(r"[가-힣]", prev):
            continue
        full, test, _ = BOOKS[book]
        v1 = int(m.group("v1"))
        v2 = int(m.group("v2")) if m.group("v2") else v1
        return ScriptureRef(full, ABBREV[full], test, int(m.group("ch")), v1, max(v1, v2), m.group(0).strip())
    return None


def page_ref(text: str) -> dict | None:
    m = _PAGE_RE.search(norm(text))
    if not m:
        return None
    test = {"구약": "old", "구": "old", "신약": "new", "신": "new"}.get(m.group("test") or "")
    return {"testament": test, "page": int(m.group("page") or m.group("page2")), "raw": m.group(0).strip()}


def name_role(text: str) -> list[dict]:
    out = []
    for m in _ROLE_RE.finditer(norm(text)):
        name = m.group(1)
        if name in ("담임", "부", "안수"):
            continue
        out.append({"name": name, "role": m.group(2)})
    return out


def quoted_title(text: str) -> str | None:
    m = _QUOTE_RE.search(norm(text))
    return m.group("t").strip() if m else None


DETAIL_TITLES = {"강사", "일시", "장소", "시간", "대상", "문의", "주제", "본문", "기간", "신청", "회비", "준비물"}
_ITEM_LINE_RE = re.compile(
    r"^\s*(?:(?P<no>\d{1,2})\s*[.)]\s*)?"
    r"(?:[\[［](?P<bt>[^\]］]{1,40})[\]］]|(?P<wt>[가-힣A-Za-z0-9·()/]{1,10}(?: [가-힣A-Za-z0-9·()/]{1,10})?))"
    r"\s*[:：]\s*(?P<body>.*)$"
)


def announcement_items(text: str) -> list[dict]:
    """Announcement block -> items. One item per line that starts with "[제목] :" or a short "제목 :"
    (optionally numbered); lines without a title continue the previous item. Lines before the first
    item (slogans, headers) are ignored. Originals are not numbered; numbers are assigned in order."""
    items: list[dict] = []
    for raw in norm(text).splitlines():
        line = raw.strip()
        if not line:
            continue
        m = _ITEM_LINE_RE.match(line)
        if m and not m.group("bt") and items and (m.group("wt") or "").strip() in DETAIL_TITLES:
            m = None  # "강사 : …", "일시 : …" lines describe the previous item
        if m:
            title = (m.group("bt") or m.group("wt") or "").strip()
            items.append({"no": len(items) + 1, "title": title, "body": m.group("body").strip(),
                          "bracketed": bool(m.group("bt")), "flag": None})
        elif items:
            items[-1]["body"] = (items[-1]["body"] + " " + line).strip()
    for it in items:
        if not it["body"]:
            it["flag"] = "review:empty_item_body"
    return items


def verse_chunks(text: str, ref: ScriptureRef | None) -> list[tuple[int | None, str]]:
    """Split scripture text on leading verse numbers that fall inside the reference range."""
    t = norm(text)
    lo, hi = (ref.verse_start, ref.verse_end) if ref else (1, 176)
    cuts = [m for m in _VERSE_NO_RE.finditer(t) if lo <= int(m.group("no")) <= hi]
    if not cuts:
        return [(None, t.strip())]
    out: list[tuple[int | None, str]] = []
    if cuts[0].start() > 0 and t[: cuts[0].start()].strip():
        out.append((None, t[: cuts[0].start()].strip()))
    for i, m in enumerate(cuts):
        end = cuts[i + 1].start() if i + 1 < len(cuts) else len(t)
        out.append((int(m.group("no")), t[m.start():end].strip()))
    return out


def english_line(text: str) -> str | None:
    """A line made only of Latin letters/punctuation (the English subtitle on cards)."""
    for line in norm(text).splitlines():
        s = line.strip()
        if s and re.fullmatch(r"[A-Za-z][A-Za-z '’&,.\-]{2,60}", s):
            return s
    return None


def korean_title_line(text: str) -> str | None:
    """First line containing Hangul, with inner spacing normalised ('기 도' -> '기도' only for 2-3 syllable words)."""
    for line in norm(text).splitlines():
        s = line.strip()
        if s and re.search(r"[가-힣]", s):
            if re.fullmatch(r"(?:[가-힣]\s+){1,3}[가-힣]", s):
                return re.sub(r"\s+", "", s)
            return re.sub(r"\s{2,}", " ", s)
    return None
