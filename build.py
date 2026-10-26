#!/usr/bin/env python3
"""Build the personal academic website from the CV (.docx) into dist/.

Usage:
    python build.py              # reads cv/cv.docx
    python build.py other.docx   # reads another copy of the CV

Standard library only, so it runs on any machine with Python 3.8+.
"""
from __future__ import annotations

import hashlib
import html
import json
import re
import struct
import shutil
import sys
import zipfile
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path
from urllib.parse import quote

ROOT = Path(__file__).resolve().parent
CV_PATH = ROOT / "cv" / "cv.docx"
ASSETS = ROOT / "assets"
PHOTOS_DIR = ASSETS / "photos"
PHOTOS_CONFIG = ROOT / "photos.json"
DIST = ROOT / "dist"

# Section key -> headings (compared case-insensitively, ignoring a trailing
# "." or ":") that start that section in the CV. Add a heading here if you
# rename one in Word.
SECTION_HEADINGS = {
    "summary": ["professional summary", "summary", "profile"],
    "education": ["education"],
    "interests": ["current research interest", "current research interests", "research interests"],
    "experience": ["professional experience"],  # umbrella heading; its own content is ignored
    "research": ["research experience"],
    "teaching": ["teaching experience"],
    "service": ["leadership and community service", "leadership and service", "service"],
    "publications": ["publications"],
    "grants": ["grants and contracts", "grants"],
    "funding": ["research and conference funding", "conference funding"],
    "review": ["scholarly peer review and academic activities", "peer review", "professional service"],
    "awards": ["awards, scholarships and fellowships", "awards and honors", "honors and awards", "awards"],
    "memberships": ["academic professional membership and position", "professional memberships", "memberships"],
    "presentations": ["academic presentation and conferences", "conference presentations", "presentations"],
    "skills": ["technical and professional skills", "skills"],
}
REQUIRED_SECTIONS = ["summary", "education", "research", "teaching", "publications"]

# Heading for publications listed before the first publications subheading.
FIRST_PUBLICATION_GROUP = "Journal Articles and Research Papers"

PAGES = [
    ("index", "Home", "./"),
    ("about", "About", "about"),
    ("research", "Research", "research"),
    ("teaching", "Teaching", "teaching"),
    ("cv", "CV", "cv"),
    ("contact", "Contact", "contact"),
]

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"

# Phone numbers never reach the pages. The lookarounds keep DOIs, page
# ranges and URLs from matching.
PHONE_RE = re.compile(
    r"(?<![\w/.\-])(?:\+\d{1,3}[\s.\-]?)?(?:\(\d{2,4}\)|\d{2,4})[\s.\-]\d{3,4}[\s.\-]\d{3,4}(?![\w/\-])"
)
PHONE_LABEL_RE = re.compile(r"(?i)\b(?:phone|tel(?:ephone)?|mobile|cell)\s*(?:no\.?|number)?\s*:\s*[,;|]?\s*")
ADDRESS_LINE_RE = re.compile(r"(?im)^\s*(?:home\s+|mailing\s+)?address\s*:.*$")

MONTH = (r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|June?|July?|Aug(?:ust)?|"
         r"Sept?(?:ember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?|Spring|Summer|Fall|Autumn|Winter)\.?")
YEAR = r"(?:19|20)\d{2}(?:\s*/\s*(?:19|20)\d{2})?"
POINT = rf"(?:{MONTH}\s+(?:\d{{1,2}}(?:\s*[-–]\s*\d{{1,2}})?,?\s+)?)?{YEAR}"
SEP = r"\s*(?:-|–|—|to)\s*"
END = rf"(?:{POINT}|present|date|current|now)"
RANGE = rf"(?:{POINT}(?:{SEP}{END})?|{MONTH}{SEP}{POINT})"
DATE_FULL_RE = re.compile(rf"(?i)^\(?\s*{RANGE}\s*\)?\.?$")
TRAILING_DATE_RE = re.compile(rf"(?i)(?:^|[\s,(\-–—|])({RANGE})\s*[.)]?\s*$")
AMOUNT_RE = re.compile(r"^(?:\$|USD|RMB|NGN|£|€|₦)\s?[\d,]+")
YEAR_RE = re.compile(r"\b(?:19|20)\d{2}\b")

BULLET_RE = re.compile(r"^\s*[·•▪●◦‣■➢➤]\s*")
NUM_PREFIX_RE = re.compile(r"^\s*\d{1,3}[.)]\s+")
CONTINUATION_RE = re.compile(r"(?i)^\s*(?:research topic|presented|presentation topic|field alignment)\b")
PRESENTATION_LABEL_RE = re.compile(
    r"(?i)\b(accepted [a-z\- ]*?presentation|presented topic|presentation topic|presented|research topic)\s*:"
)
EMAIL_RE = re.compile(r"[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
URL_RE = re.compile(r"(?:https?://|www\.)[^\s,;<>\"]+")


class BuildError(Exception):
    pass


# --------------------------------------------------------------------------
# Reading the .docx
# --------------------------------------------------------------------------

class Para:
    __slots__ = ("text", "bold", "numbered")

    def __init__(self, text: str, bold: bool, numbered: bool):
        self.text = text
        self.bold = bold
        self.numbered = numbered


def scrub(text: str) -> str:
    """Remove phone numbers and address lines from a paragraph."""
    text = ADDRESS_LINE_RE.sub("", text)
    if PHONE_RE.search(text):
        text = PHONE_LABEL_RE.sub("", PHONE_RE.sub("", text))
    return text


def _run_is_bold(run) -> bool:
    b = run.find(f"{W}rPr/{W}b")
    return b is not None and b.get(f"{W}val", "true") not in ("0", "false", "none")


def read_paragraphs(path: Path) -> list[Para]:
    """Return the CV's paragraphs in reading order, text boxes included once."""
    try:
        with zipfile.ZipFile(path) as z:
            data = z.read("word/document.xml")
    except FileNotFoundError:
        raise BuildError(f"CV not found at {path}. Put your CV there (see README.md).")
    except (zipfile.BadZipFile, KeyError):
        raise BuildError(f"{path} is not a Word .docx file.")

    paras: list[Para] = []

    def read_p(p):
        pieces: list[str] = []
        nested = []
        counts = [0, 0]  # bold chars, total chars

        def walk(el, bold):
            for c in el:
                tag = c.tag
                if tag in (MC_FALLBACK, f"{W}pPr"):
                    continue
                if tag == f"{W}p":
                    nested.append(c)
                elif tag == f"{W}r":
                    walk(c, _run_is_bold(c))
                elif tag == f"{W}t":
                    t = c.text or ""
                    pieces.append(t)
                    n = len(t.strip())
                    counts[1] += n
                    if bold:
                        counts[0] += n
                elif tag in (f"{W}tab", f"{W}ptab"):
                    pieces.append("\t")
                elif tag in (f"{W}br", f"{W}cr"):
                    pieces.append("\n")
                elif tag == f"{W}noBreakHyphen":
                    pieces.append("-")
                else:
                    walk(c, bold)

        walk(p, False)
        num = p.find(f"{W}pPr/{W}numPr/{W}numId")
        numbered = num is not None and num.get(f"{W}val") not in (None, "0")
        text = scrub("".join(pieces).replace(" ", " "))
        paras.append(Para(text, counts[1] > 0 and counts[0] * 2 >= counts[1], numbered))
        for n in nested:
            read_p(n)

    def visit(el):
        for child in el:
            if child.tag == MC_FALLBACK:
                continue
            if child.tag == f"{W}p":
                read_p(child)
            else:
                visit(child)

    visit(ET.fromstring(data).find(f"{W}body"))
    return paras


# --------------------------------------------------------------------------
# Text helpers
# --------------------------------------------------------------------------

def clean(s: str) -> str:
    return re.sub(r"\s+", " ", s.replace("\t", " ")).strip().strip("|").strip()


def tidy(s: str) -> str:
    return s.strip().rstrip(" ,;:-–—|(").strip()


def heading_case(s: str) -> str:
    s = clean(s).rstrip(".:").strip()
    if s.isupper():
        small = {"and", "or", "of", "the", "in", "for", "to", "a", "an", "on"}
        words = s.lower().split()
        s = " ".join(w if (i and w in small) else w.capitalize() for i, w in enumerate(words))
    return s[:1].upper() + s[1:]


def normalize_heading(s: str) -> str:
    return re.sub(r"\s+", " ", s.strip().lower().replace("&", "and")).rstrip(".:").strip()


def is_meta(seg: str) -> bool:
    return bool(DATE_FULL_RE.match(seg) or AMOUNT_RE.match(seg)
                or (seg.startswith("(") and YEAR_RE.search(seg) and len(seg) <= 90))


def meta_text(seg: str) -> str:
    seg = clean(seg).rstrip(".")
    if seg.startswith("(") and seg.endswith(")"):
        seg = seg[1:-1].strip()
    return seg


def split_meta(line: str, trailing: bool = True) -> tuple[list[str], str]:
    """Split a line into its text parts and a date (or amount) set off at the end."""
    segs = [s for s in (clean(x) for x in line.split("\t")) if s]
    if not segs:
        return [], ""
    if len(segs) == 1 and DATE_FULL_RE.match(segs[0]):
        return [], meta_text(segs[0])
    if len(segs) > 1 and is_meta(segs[-1]):
        return [p for p in (tidy(s) for s in segs[:-1]) if p], meta_text(segs[-1])
    joined = clean(" ".join(segs))
    if trailing:
        m = TRAILING_DATE_RE.search(joined)
        if m and m.start(1) > 0:
            head = tidy(joined[:m.start(1)])
            if head:
                return [head], m.group(1)
    return [joined], ""


def extract_lines(raw_lines: list[str]) -> tuple[list[str], str]:
    """Clean a block of lines and pull out one date or amount.

    A value set off by a tab anywhere in the block wins; otherwise a date at
    the end of a line is used.
    """
    raw_lines = [BULLET_RE.sub("", r) for r in raw_lines if clean(r)]
    for trailing in (False, True):
        lines, meta = [], ""
        for raw in raw_lines:
            if meta:
                parts = [clean(raw)]
            else:
                parts, meta = split_meta(raw, trailing=trailing)
            lines.extend(p for p in parts if p)
        if meta:
            break
    return lines, meta


def finalize_header(raw_lines: list[str]) -> dict:
    lines, meta = extract_lines(raw_lines)
    sub = [s.rstrip(":").strip() for s in lines[1:]]
    return {"title": lines[0] if lines else "", "sub": sub, "meta": meta, "blocks": []}


# --------------------------------------------------------------------------
# Section parsers
# --------------------------------------------------------------------------

def parse_paragraph_text(paras: list[Para]) -> list[str]:
    return [t for t in (clean(p.text) for p in paras) if t]


def parse_interests(paras: list[Para]) -> list[str]:
    items: list[str] = []
    for p in paras:
        t = clean(BULLET_RE.sub("", p.text.replace("\n", " "))).rstrip(".")
        if t and t not in items:
            items.append(t)
    return items


def parse_education(paras: list[Para]) -> list[dict]:
    groups: list[list[str]] = []
    for p in paras:
        if not clean(p.text):
            continue
        if p.numbered or not groups:
            groups.append([])
        groups[-1].extend(p.text.split("\n"))
    return [finalize_header(g) for g in groups]


def _bullet_blocks(p: Para) -> list[tuple[str, str, str]]:
    out = []
    for i, raw in enumerate(p.text.split("\n")):
        raw = BULLET_RE.sub("", raw)
        if not clean(raw):
            continue
        parts, meta = split_meta(raw, trailing=False)
        text = " ".join(parts)
        if i == 0:
            out.append(("bullet", text, meta))
        elif text.endswith(":"):
            out.append(("sub", text.rstrip(":"), ""))
        else:
            out.append(("note", text, meta))
    return out


def parse_experience(paras: list[Para]) -> list[dict]:
    """Entries are a header (bold, or carrying a date) followed by bullets."""
    entries: list[dict] = []
    cur = None
    in_header = False
    after_gap = True   # previous paragraph was blank (or this is the section start)
    pending = None     # plain paragraph after a gap; joins the next header if one follows

    def flush_pending():
        nonlocal pending, cur
        if pending is None:
            return
        if cur is None:
            cur = {"raw": pending.text.split("\n"), "blocks": []}
            entries.append(cur)
        else:
            cur["blocks"].append(("note", clean(pending.text), ""))
        pending = None

    for p in paras:
        if not clean(p.text):
            flush_pending()
            in_header = False
            after_gap = True
            continue
        gap, after_gap = after_gap, False
        first = clean(p.text.strip().split("\n")[0])
        single = "\n" not in p.text.strip()
        manual_bullet = bool(BULLET_RE.match(p.text))
        if p.numbered or manual_bullet:
            flush_pending()
            if cur is None:
                cur = {"raw": [], "blocks": []}
                entries.append(cur)
            if p.bold and not manual_bullet and single and len(first) < 80 and not first.endswith("."):
                cur["blocks"].append(("sub", first, ""))
            else:
                cur["blocks"].extend(_bullet_blocks(p))
            in_header = False
            continue
        has_date = any(split_meta(line)[1] for line in p.text.split("\n"))
        if p.bold or has_date:
            if (pending is None and cur is not None and not in_header and p.bold and single
                    and len(first) < 45 and first.endswith(":")):
                cur["blocks"].append(("sub", first.rstrip(":"), ""))
                continue
            if in_header and cur is not None:
                cur["raw"].extend(p.text.split("\n"))
            else:
                lead = pending.text.split("\n") if pending is not None else []
                pending = None
                cur = {"raw": lead + p.text.split("\n"), "blocks": []}
                entries.append(cur)
                in_header = True
            continue
        flush_pending()
        if gap:
            pending = p
        elif cur is not None:
            cur["blocks"].append(("note", clean(p.text), ""))
        in_header = False
    flush_pending()

    out = []
    for e in entries:
        entry = finalize_header(e["raw"])
        entry["blocks"] = e["blocks"]
        if entry["title"] or entry["blocks"]:
            out.append(entry)
    return out


def parse_publications(paras: list[Para]) -> list[dict]:
    groups: list[dict] = []

    def items() -> list[str]:
        if not groups:
            groups.append({"title": FIRST_PUBLICATION_GROUP, "items": []})
        return groups[-1]["items"]

    for p in paras:
        for i, raw in enumerate(p.text.split("\n")):
            text = clean(raw)
            if not text:
                continue
            if NUM_PREFIX_RE.match(text):
                items().append(NUM_PREFIX_RE.sub("", text))
            elif i == 0 and p.numbered:
                items().append(text)
            elif i == 0 and len(text) < 90 and not YEAR_RE.search(text):
                groups.append({"title": heading_case(text), "items": []})
            elif groups and groups[-1]["items"]:
                groups[-1]["items"][-1] += " " + text
            else:
                items().append(text)
    return [g for g in groups if g["items"]]


def parse_grants(paras: list[Para]) -> list[dict]:
    texts: list[str] = []
    for p in paras:
        t = clean(p.text)
        if not t:
            continue
        if p.numbered or NUM_PREFIX_RE.match(t) or not texts:
            texts.append(NUM_PREFIX_RE.sub("", t))
        else:
            texts[-1] += " " + t
    out = []
    for t in texts:
        m = re.match(r"^(.{5,160}?),\s*((?:19|20)\d{2}(?:\s*[-–]\s*(?:19|20)\d{2})?)\s*,\s*(.*)$", t)
        title, years, body = m.groups() if m else (t, "", "")
        amount = re.search(r"\$\s?[\d,]+|RMB\s?[\d,]+(?:\s*\(approximately USD \$[\d,]+\))?", t)
        out.append({
            "title": title.strip(),
            "meta": years,
            "amount": amount.group(0) if amount else "",
            "body": body[:1].upper() + body[1:],
        })
    return out


def parse_list(paras: list[Para], with_meta: bool = True) -> list[dict]:
    """One entry per numbered paragraph; unnumbered lines continue the entry."""
    groups: list[list[str]] = []
    for p in paras:
        if not clean(p.text):
            continue
        if groups and CONTINUATION_RE.match(p.text):
            groups[-1].append(p.text)
        elif p.numbered or NUM_PREFIX_RE.match(p.text) or not groups:
            groups.append([p.text])
        else:
            groups[-1].append(p.text)
    out = []
    for g in groups:
        raw_lines = "\n".join(g).split("\n")
        raw_lines[0] = NUM_PREFIX_RE.sub("", raw_lines[0])
        if with_meta:
            lines, meta = extract_lines(raw_lines)
        else:
            lines, meta = [clean(BULLET_RE.sub("", r)) for r in raw_lines if clean(r)], ""
        if lines:
            out.append({"title": lines[0], "sub": lines[1:], "meta": meta, "blocks": []})
    return out


def parse_presentations(paras: list[Para]) -> list[dict]:
    out = []
    for e in parse_list(paras, with_meta=False):
        text = " ".join([e["title"]] + e["sub"])
        year = YEAR_RE.search(text)
        item = {"venue": text, "label": "", "topic": "", "authors": "", "year": year.group(0) if year else ""}
        m = PRESENTATION_LABEL_RE.search(text)
        if m:
            rest = text[m.end():].strip().lstrip("“\"")
            close = re.search(r"[”\"]", rest)
            if close:
                topic, tail = rest[:close.start()], rest[close.end():]
            else:
                bits = re.split(r"\s+[Bb]y\s+", rest, maxsplit=1)
                topic, tail = bits[0], (bits[1] if len(bits) > 1 else "")
            tail = re.sub(r"^[\s.,]*(?:\[[^\]]*\][.\s]*)?(?:Research Presentation\s+)?(?:[Bb]y\s+)?", "", tail)
            item.update(
                venue=text[:m.start()].strip(" ,.;"),
                label=m.group(1).strip()[:1].upper() + m.group(1).strip()[1:],
                topic=topic.strip().rstrip(".").strip(),
                authors=tail.strip().rstrip("."),
            )
            year = YEAR_RE.search(item["venue"])
            if year:
                item["year"] = year.group(0)
        out.append(item)
    return out


def parse_skills(paras: list[Para]) -> list[dict]:
    groups: list[dict] = []
    for p in paras:
        t = clean(BULLET_RE.sub("", p.text))
        if not t:
            continue
        if p.numbered or BULLET_RE.match(p.text):
            if not groups:
                groups.append({"title": "", "items": []})
            groups[-1]["items"].append(t.rstrip("."))
        else:
            g = groups[-1] if groups else None
            if g and g["items"] and re.search(r"(?:\b(?:for|and|of|to|in|the|with|on|or)|,)$", g["items"][-1]):
                g["items"][-1] += " " + t.rstrip(".")
            else:
                groups.append({"title": t.rstrip(":").strip(), "items": []})
    return [g for g in groups if g["items"]]


# --------------------------------------------------------------------------
# Whole-CV parsing
# --------------------------------------------------------------------------

def split_sections(paras: list[Para]) -> tuple[list[Para], dict]:
    lookup = {alias: key for key, aliases in SECTION_HEADINGS.items() for alias in aliases}
    preamble: list[Para] = []
    sections: dict = {}
    current = None
    for p in paras:
        stripped = p.text.strip()
        norm = normalize_heading(stripped)
        if norm.startswith("references"):
            current = "_end"
            continue
        if "\t" not in stripped and "\n" not in stripped and len(stripped) < 80 and norm in lookup:
            current = lookup[norm]
            sections.setdefault(current, {"title": heading_case(stripped), "paras": []})
            continue
        if current is None:
            preamble.append(p)
        elif current != "_end":
            sections[current]["paras"].append(p)
    return preamble, sections


def section_outline(path: Path) -> list[str]:
    """Plain-text outline of the CV, used by publish.py to show what changed."""
    preamble, sections = split_sections(read_paragraphs(path))
    lines = [clean(p.text) for p in preamble if clean(p.text)]
    for key, sec in sections.items():
        lines.extend(f"[{sec['title']}] {clean(p.text)}" for p in sec["paras"] if clean(p.text))
    return lines


def parse_citations(paras: list[Para]) -> str:
    """Citation count as the CV states it, e.g. "over 600 citations" -> "600+"."""
    for p in paras:
        m = re.search(r"(over|more than)?\s*([\d,]+)\s*(\+)?\s+citations", p.text, re.I)
        if m:
            return m.group(2) + ("+" if m.group(1) or m.group(3) else "")
    return ""


def parse_contact(preamble: list[Para]) -> dict:
    lines = [clean(l) for p in preamble for l in p.text.split("\n") if clean(l)]
    if not lines:
        raise BuildError("The CV should start with your name on the first line.")
    rest = " ".join(lines[1:])
    links = []
    for m in URL_RE.finditer(rest):
        url = m.group(0).rstrip(".,;)")
        written = re.search(r"([A-Za-z][A-Za-z ]{1,40}?)\s*:\s*$", rest[:m.start()])
        if not url.startswith("http"):
            url = "https://" + url
        low = url.lower()
        if "linkedin.com" in low:
            label = "LinkedIn"
        elif "scholar.google" in low:
            label = "Google Scholar"
        elif "orcid.org" in low:
            label = "ORCID"
        elif "researchgate" in low:
            label = "ResearchGate"
        elif "github.com" in low:
            label = "GitHub"
        else:
            # Use the CV's own label, e.g. "Academic Website: https://..."
            label = heading_case(written.group(1)) if written else "Website"
        links.append({"label": label, "url": url})
    email = EMAIL_RE.search(rest)
    return {"name": lines[0], "email": email.group(0) if email else "", "links": links}


def parse_cv(path: Path) -> dict:
    paras = read_paragraphs(path)
    preamble, sections = split_sections(paras)
    missing = [k for k in REQUIRED_SECTIONS if k not in sections]
    if missing:
        detail = "\n".join(f"  - {k}: one of {', '.join(repr(h.title()) for h in SECTION_HEADINGS[k])}"
                           for k in missing)
        raise BuildError(
            "The CV is missing required section(s). Add a heading paragraph (on its own line) for:\n"
            + detail + "\nor add your heading's wording to SECTION_HEADINGS in build.py."
        )

    def paras_of(key):
        return sections.get(key, {}).get("paras", [])

    updated = ""
    for p in paras:
        m = re.search(r"Last updated:\s*([A-Z][a-z]+ \d{1,2}, \d{4})", p.text)
        if m:
            updated = m.group(1)

    cv = {
        "contact": parse_contact(preamble),
        "titles": {k: v["title"] for k, v in sections.items()},
        "summary": parse_paragraph_text(paras_of("summary")),
        "education": parse_education(paras_of("education")),
        "interests": parse_interests(paras_of("interests")),
        "research": parse_experience(paras_of("research")),
        "teaching": parse_experience(paras_of("teaching")),
        "service": parse_experience(paras_of("service")),
        "publications": parse_publications(paras_of("publications")),
        "grants": parse_grants(paras_of("grants")),
        "funding": parse_list(paras_of("funding")),
        "review": parse_list(paras_of("review")),
        "awards": parse_list(paras_of("awards")),
        "memberships": parse_list(paras_of("memberships")),
        "presentations": parse_presentations(paras_of("presentations")),
        "skills": parse_skills(paras_of("skills")),
        "updated": updated,
        "citations": parse_citations(paras),
    }
    for key in REQUIRED_SECTIONS:
        if not cv[key]:
            raise BuildError(f"The '{cv['titles'][key]}' section in the CV has no entries.")
    return cv


# --------------------------------------------------------------------------
# HTML rendering
# --------------------------------------------------------------------------

def esc(s: str) -> str:
    return html.escape(s, quote=True)


def name_regex(full_name: str):
    parts = full_name.split()
    if len(parts) < 2:
        return None
    first, last, middle = parts[0], parts[-1], parts[1:-1]
    fi = re.escape(first[0])
    mi = re.escape(middle[0][0]) if middle else None
    lead = rf"(?:{re.escape(first)}\s+(?:{re.escape(middle[0])}\s+|{mi}\.\s+)?)?" if middle else rf"(?:{re.escape(first)}\s+)?"
    initials = rf"{fi}\b\.?,?\s?(?:{mi}\b\.?)?" if mi else rf"{fi}\b\.?"
    tail = rf"(?:,?\s*(?:{re.escape(first)}\s+{mi}\b\.?|{initials}))?" if mi else rf"(?:,?\s*{initials})?"
    return re.compile(lead + re.escape(last) + tail)


NAME_RE = None


def rich(text: str, highlight: bool = False) -> str:
    """Escape text, turn URLs into links, and optionally bold the owner's name."""
    out = []
    pos = 0
    for m in URL_RE.finditer(text):
        url = m.group(0).rstrip(".,;)")
        out.append(_plain(text[pos:m.start()], highlight))
        href = url if url.startswith("http") else "https://" + url
        out.append(f'<a href="{esc(href)}" rel="noopener">{esc(url)}</a>')
        pos = m.start() + len(url)
    out.append(_plain(text[pos:], highlight))
    return "".join(out)


def _plain(text: str, highlight: bool) -> str:
    if not (highlight and NAME_RE):
        return esc(text)
    out, pos = [], 0
    for m in NAME_RE.finditer(text):
        out.append(esc(text[pos:m.start()]))
        out.append(f'<strong class="me">{esc(m.group(0).rstrip(", "))}</strong>')
        out.append(esc(m.group(0)[len(m.group(0).rstrip(", ")):]))
        pos = m.end()
    out.append(esc(text[pos:]))
    return "".join(out)


def entry_html(e: dict, level: int = 3) -> str:
    h = f"h{level}"
    parts = ['<article class="entry">', '<div class="entry-head">']
    parts.append(f'<{h} class="entry-title">{rich(e["title"])}</{h}>')
    if e.get("meta"):
        parts.append(f'<p class="entry-meta">{esc(e["meta"])}</p>')
    parts.append("</div>")
    if e.get("sub"):
        parts.append('<p class="entry-sub">' + "<br>".join(rich(s) for s in e["sub"]) + "</p>")
    bullets: list[str] = []

    def flush():
        if bullets:
            parts.append("<ul>" + "".join(bullets) + "</ul>")
            bullets.clear()

    for kind, text, meta in e.get("blocks", []):
        if kind == "bullet":
            m = f' <span class="inline-meta">{esc(meta)}</span>' if meta else ""
            bullets.append(f"<li>{rich(text)}{m}</li>")
            continue
        flush()
        if kind == "sub":
            parts.append(f'<h{level + 1} class="entry-role">{rich(text)}</h{level + 1}>')
        else:
            parts.append(f'<p class="entry-note">{rich(text)}</p>')
    flush()
    parts.append("</article>")
    return "".join(parts)


def section(title: str, inner: str, sid: str = "", lead: str = "") -> str:
    ida = f' id="{sid}"' if sid else ""
    lead_html = f'<p class="section-lead">{lead}</p>' if lead else ""
    return f'<section class="block"{ida}><h2>{esc(title)}</h2>{lead_html}{inner}</section>'


def entries_html(entries: list[dict]) -> str:
    return '<div class="entries">' + "".join(entry_html(e) for e in entries) + "</div>"


def publications_html(groups: list[dict], filterable: bool, limit: int = 0) -> str:
    parts = []
    if filterable:
        total = sum(len(g["items"]) for g in groups)
        parts.append(
            '<div class="filter"><label for="pub-filter">Filter publications</label>'
            '<input type="search" id="pub-filter" autocomplete="off" '
            'placeholder="Keyword, co-author, journal, or year">'
            f'<p id="pub-count" class="filter-count" aria-live="polite" data-total="{total}">{total} publications</p></div>'
        )
    for g in groups:
        items = g["items"][:limit] if limit else g["items"]
        lis = "".join(f'<li value="{i}">{rich(t, highlight=True)}</li>' for i, t in enumerate(items, 1))
        parts.append(f'<div class="pub-group"><h3>{esc(g["title"])}</h3><ol class="pubs">{lis}</ol></div>')
    return "".join(parts)


def presentations_html(items: list[dict], limit: int = 0) -> str:
    rows = []
    for p in items[:limit] if limit else items:
        if p["topic"]:
            body = f'<h3 class="entry-title">{rich(p["topic"])}</h3><p class="entry-sub">{rich(p["venue"])}</p>'
            if p["authors"]:
                body += f'<p class="entry-note">{rich(p["authors"], highlight=True)}</p>'
        else:
            body = f'<p class="entry-title">{rich(p["venue"], highlight=True)}</p>'
        year = f'<p class="entry-meta">{esc(p["year"])}</p>' if p["year"] else ""
        rows.append(f'<article class="entry talk"><div class="entry-head"><div>{body}</div>{year}</div></article>')
    return '<div class="entries">' + "".join(rows) + "</div>"


def grants_html(grants: list[dict]) -> str:
    rows = []
    for g in grants:
        meta = " · ".join(x for x in (g["meta"], g["amount"]) if x)
        rows.append(
            '<article class="entry"><div class="entry-head">'
            f'<h3 class="entry-title">{rich(g["title"])}</h3>'
            + (f'<p class="entry-meta">{esc(meta)}</p>' if meta else "")
            + "</div>"
            + (f'<p class="entry-note">{rich(g["body"], highlight=True)}</p>' if g["body"] else "")
            + "</article>"
        )
    return '<div class="entries">' + "".join(rows) + "</div>"


def chips(items: list[str]) -> str:
    return '<ul class="chips">' + "".join(f"<li>{esc(i)}</li>" for i in items) + "</ul>"


def skills_html(groups: list[dict]) -> str:
    cols = []
    for g in groups:
        head = f"<h3>{esc(g['title'])}</h3>" if g["title"] else ""
        cols.append(f'<div class="skill-group">{head}{chips(g["items"])}</div>')
    return '<div class="skills">' + "".join(cols) + "</div>"


def contact_links_html(contact: dict) -> str:
    items = []
    if contact["email"]:
        e = esc(contact["email"])
        items.append(f'<li><span class="label">Email</span><a href="mailto:{e}">{e}</a></li>')
    for link in contact["links"]:
        shown = re.sub(r"^https?://(www\.)?", "", link["url"]).split("?")[0].rstrip("/")
        items.append(f'<li><span class="label">{esc(link["label"])}</span>'
                     f'<a href="{esc(link["url"])}" rel="noopener">{esc(shown)}</a></li>')
    return '<ul class="contact-list">' + "".join(items) + "</ul>"


# --------------------------------------------------------------------------
# Photos (optional): photos.json says which picture goes where; the files are
# the web-sized copies made by optimize-photos.ps1 in assets/photos/.
# --------------------------------------------------------------------------

def photo_slug(src: str) -> str:
    """Must match Get-Slug in optimize-photos.ps1."""
    stem = re.sub(r"\.[^.]+$", "", src)
    return re.sub(r"[^a-z0-9]+", "-", stem.lower()).strip("-") + ".jpg"


def jpeg_size(path: Path) -> tuple[int, int]:
    """(width, height) from a JPEG's frame header, or (0, 0) if unreadable."""
    data = path.read_bytes()
    i = 2
    while i + 9 < len(data):
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in (0xC0, 0xC1, 0xC2):
            h, w = struct.unpack(">HH", data[i + 5:i + 9])
            return w, h
        i += 2 + struct.unpack(">H", data[i + 2:i + 4])[0]
    return 0, 0


def load_photos() -> dict:
    """Read photos.json, keeping only pictures whose web copy exists."""
    if not PHOTOS_CONFIG.exists():
        return {"portrait": None, "galleries": {}}
    try:
        config = json.loads(PHOTOS_CONFIG.read_text(encoding="utf-8"))
    except json.JSONDecodeError as e:
        raise BuildError(f"photos.json is not valid JSON: {e}")

    def resolve(item):
        if not item or not item.get("src"):
            return None
        path = PHOTOS_DIR / photo_slug(item["src"])
        if not path.exists():
            print(f"Warning: no web copy of '{item['src']}' yet; run optimize-photos.ps1. Skipping it.")
            return None
        w, h = jpeg_size(path)
        return {**item, "file": f"assets/photos/{path.name}", "w": w, "h": h}

    galleries = {}
    for name, items in (config.get("galleries") or {}).items():
        resolved = [p for p in (resolve(i) for i in items) if p]
        if resolved:
            galleries[name] = resolved
    return {"portrait": resolve(config.get("portrait")), "galleries": galleries}


PHOTOS: dict = {"portrait": None, "galleries": {}}


def asset_version(name: str) -> str:
    """Short content hash, so browsers fetch a changed stylesheet or script right away."""
    return hashlib.sha1((ASSETS / name).read_bytes()).hexdigest()[:10]


def img_tag(p: dict, cls: str = "", eager: bool = False) -> str:
    size = f' width="{p["w"]}" height="{p["h"]}"' if p["w"] else ""
    loading = "" if eager else ' loading="lazy"'
    c = f' class="{cls}"' if cls else ""
    return f'<img src="{esc(p["file"])}" alt="{esc(p.get("alt", ""))}"{size}{loading} decoding="async"{c}>'


def gallery(name: str) -> str:
    items = PHOTOS["galleries"].get(name)
    if not items:
        return ""
    figs = "".join(
        f'<figure class="photo">{img_tag(p)}'
        + (f"<figcaption>{esc(p['caption'])}</figcaption>" if p.get("caption") else "")
        + "</figure>"
        for p in items
    )
    return f'<div class="gallery gallery-{len(items)}">{figs}</div>'


def portrait(cls: str, eager: bool = False) -> str:
    p = PHOTOS["portrait"]
    return f'<div class="{cls}">{img_tag(p, eager=eager)}</div>' if p else ""


def contact_line_html(contact: dict) -> str:
    """Compact, clickable contact line, like the header of the CV."""
    items = [f'<a href="{esc(l["url"])}" rel="noopener">{esc(l["label"])}</a>' for l in contact["links"]]
    if contact["email"]:
        e = esc(contact["email"])
        items.insert(0, f'<a href="mailto:{e}">{e}</a>')
    return '<ul class="contact-line">' + "".join(f"<li>{i}</li>" for i in items) + "</ul>" if items else ""


ICON_SUN = ('<svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true" focusable="false">'
            '<circle cx="12" cy="12" r="4.5" fill="none" stroke="currentColor" stroke-width="1.8"/>'
            '<path d="M12 2.5v2.2M12 19.3v2.2M4.6 4.6l1.6 1.6M17.8 17.8l1.6 1.6M2.5 12h2.2M19.3 12h2.2'
            'M4.6 19.4l1.6-1.6M17.8 6.2l1.6-1.6" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>')
ICON_MOON = ('<svg viewBox="0 0 24 24" width="20" height="20" aria-hidden="true" focusable="false">'
             '<path d="M20 14.5A8 8 0 0 1 9.5 4a8 8 0 1 0 10.5 10.5z" fill="none" stroke="currentColor" '
             'stroke-width="1.8" stroke-linejoin="round"/></svg>')


def favicon(name: str) -> str:
    initials = "".join(w[0] for w in (name.split()[0], name.split()[-1])).upper()
    svg = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64"><rect width="64" height="64" rx="14" '
           'fill="#1f4e5f"/><text x="32" y="42" font-family="Georgia,serif" font-size="28" text-anchor="middle" '
           f'fill="#faf7f2">{initials}</text></svg>')
    return "data:image/svg+xml," + quote(svg)


def short_name(name: str) -> str:
    parts = name.split()
    if len(parts) > 2:
        return f"{parts[0]} {parts[1][0]}. {parts[-1]}"
    return name


def layout(slug: str, title: str, body: str, cv: dict, description: str) -> str:
    name = cv["contact"]["name"]
    nav = "".join(
        f'<li><a href="{href}"' + (' aria-current="page"' if s == slug else "") + f">{label}</a></li>"
        for s, label, href in PAGES
    )
    page_title = name if slug == "index" else f"{title} | {name}"
    updated = f' · CV updated {esc(cv["updated"])}' if cv["updated"] else ""
    year = date.today().year
    p = PHOTOS["portrait"]
    og_image = f'\n<meta property="og:image" content="{esc(p["file"])}">' if p else ""
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(page_title)}</title>
<meta name="description" content="{esc(description)}">
<meta name="color-scheme" content="light dark">
<meta property="og:title" content="{esc(page_title)}">
<meta property="og:description" content="{esc(description)}">
<meta property="og:type" content="website">{og_image}
<link rel="icon" href="{favicon(name)}">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,560;9..144,650&family=Source+Sans+3:ital,wght@0,400;0,600;0,700;1,400&display=swap">
<link rel="stylesheet" href="assets/style.css?v={asset_version('style.css')}">
<script src="assets/site.js?v={asset_version('site.js')}"></script>
</head>
<body>
<a class="skip-link" href="#main">Skip to content</a>
<header class="site-header">
  <div class="wrap header-inner">
    <a class="brand" href="./">{esc(short_name(name))}</a>
    <div class="header-actions">
      <button class="theme-toggle" type="button" aria-label="Switch between light and dark mode">
        <span class="icon-sun">{ICON_SUN}</span><span class="icon-moon">{ICON_MOON}</span>
      </button>
      <button class="nav-toggle" type="button" aria-expanded="false" aria-controls="site-nav">Menu</button>
    </div>
    <nav id="site-nav" class="site-nav" aria-label="Main"><ul>{nav}</ul></nav>
  </div>
</header>
<main id="main" class="wrap">
{body}
</main>
<footer class="site-footer">
  <div class="wrap footer-inner">
    <p>&copy; {year} {esc(name)}{updated}</p>
    <p><a href="contact">Contact</a></p>
  </div>
</footer>
</body>
</html>
"""


def tagline(cv: dict) -> str:
    first = cv["summary"][0]
    m = re.match(r"^(.{20,200}?)\s+with\s", first)
    if m:
        return m.group(1).strip().rstrip(",")
    return re.split(r"(?<=[.!?])\s", first, maxsplit=1)[0].rstrip(".")


def current_position(cv: dict) -> tuple[str, list[str]]:
    """Organization and roles from the first Research Experience entry."""
    e = cv["research"][0]
    org, _, role = e["title"].partition(": ")
    roles = [role.strip().rstrip(".")] if role.strip() else []
    roles += [t for k, t, _ in e["blocks"] if k == "sub"]
    return org.strip().rstrip("."), roles


def intro_paragraph(cv: dict) -> str:
    """First first-person summary paragraph (the opening one repeats the tagline)."""
    for p in cv["summary"]:
        if re.match(r"^(I|My)\b", p):
            return p
    return cv["summary"][0]


def render_home(cv: dict, cv_file: str) -> str:
    name = cv["contact"]["name"]
    affiliation, roles = current_position(cv)
    pubs = cv["publications"]
    n_pubs = sum(len(g["items"]) for g in pubs)
    stats = [(n_pubs, "Publications"), (len(cv["presentations"]), "Conference presentations"),
             (len(cv["grants"]), "Research grants")]
    stats_html = "".join(f'<div class="stat"><span class="stat-num">{n}</span><span class="stat-label">{l}</span></div>'
                         for n, l in stats if n)
    if cv["citations"]:
        scholar = next((l["url"] for l in cv["contact"]["links"] if l["label"] == "Google Scholar"), "")
        num = f'<span class="stat-num">{esc(cv["citations"])}</span><span class="stat-label">Google Scholar citations</span>'
        stats_html += (f'<a class="stat stat-link" href="{esc(scholar)}" rel="noopener">{num}</a>' if scholar
                       else f'<div class="stat">{num}</div>')
    roles_html = "".join(f"<li>{esc(r)}</li>" for r in roles)
    body = f"""
<section class="hero{' has-portrait' if PHOTOS['portrait'] else ''}">
 <div class="hero-text">
  <p class="eyebrow">{esc(affiliation)}</p>
  <h1>{esc(name)}</h1>
  <p class="hero-tagline">{esc(tagline(cv))}</p>
  {f'<ul class="roles">{roles_html}</ul>' if roles else ''}
  <div class="actions">
    <a class="button" href="research">Research and publications</a>
    <a class="button button-quiet" href="{esc(cv_file)}" download>Download CV</a>
  </div>
 </div>
 {portrait("hero-portrait", eager=True)}
</section>
<section class="intro">
  <div class="prose"><p>{rich(intro_paragraph(cv))}</p></div>
  {f'<div class="stats">{stats_html}</div>' if stats_html else ''}
</section>
"""
    if cv["interests"]:
        body += section(cv["titles"].get("interests", "Research Interests"), chips(cv["interests"]))
    if gallery("home"):
        body += section("Highlights", gallery("home"))
    body += section("Recent publications",
                    publications_html(pubs[:1], filterable=False, limit=4)
                    + '<p class="more"><a href="research#publications">All publications</a></p>')
    if cv["presentations"]:
        body += section("Recent presentations", presentations_html(cv["presentations"], limit=3)
                        + '<p class="more"><a href="research#presentations">All presentations</a></p>')
    return body


def render_about(cv: dict) -> str:
    t = cv["titles"]
    body = ('<header class="page-head has-portrait"><div><h1>About</h1>'
            f'<p class="page-lead">{esc(tagline(cv))}.</p></div>{portrait("page-portrait")}</header>')
    body += section(t["summary"], '<div class="prose">' + "".join(f"<p>{rich(p)}</p>" for p in cv["summary"]) + "</div>")
    body += section(t["education"], gallery("about-education") + entries_html(cv["education"]), "education")
    if cv["awards"]:
        body += section(t["awards"], entries_html(cv["awards"]), "awards")
    if cv["service"]:
        body += section(t["service"], gallery("about-service") + entries_html(cv["service"]), "service")
    if cv["memberships"]:
        body += section(t["memberships"], entries_html(cv["memberships"]), "memberships")
    return body


def render_research(cv: dict) -> str:
    t = cv["titles"]
    lead = rich(cv["summary"][1]) if len(cv["summary"]) > 1 else ""
    body = '<header class="page-head"><h1>Research</h1>' + (f'<p class="page-lead">{lead}</p>' if lead else "") + "</header>"
    if cv["interests"]:
        body += section(t["interests"], chips(cv["interests"]), "interests")
    body += section(t["publications"], publications_html(cv["publications"], filterable=True), "publications")
    if cv["grants"]:
        body += section(t["grants"], grants_html(cv["grants"]), "grants")
    if cv["presentations"]:
        body += section(t["presentations"], gallery("research-presentations")
                        + presentations_html(cv["presentations"]), "presentations")
    body += section(t["research"], gallery("research-epc") + entries_html(cv["research"]), "experience")
    if cv["funding"]:
        body += section(t["funding"], entries_html(cv["funding"]), "funding")
    if cv["review"]:
        body += section(t["review"], entries_html(cv["review"]), "review")
    return body


def render_teaching(cv: dict) -> str:
    body = '<header class="page-head"><h1>Teaching</h1></header>'
    body += section(cv["titles"]["teaching"], entries_html(cv["teaching"]), "teaching")
    return body


def render_cv(cv: dict, cv_file: str) -> str:
    t = cv["titles"]
    order = [
        ("education", lambda: entries_html(cv["education"])),
        ("interests", lambda: chips(cv["interests"])),
        ("research", lambda: entries_html(cv["research"])),
        ("teaching", lambda: entries_html(cv["teaching"])),
        ("service", lambda: entries_html(cv["service"])),
        ("publications", lambda: publications_html(cv["publications"], filterable=False)),
        ("grants", lambda: grants_html(cv["grants"])),
        ("funding", lambda: entries_html(cv["funding"])),
        ("review", lambda: entries_html(cv["review"])),
        ("awards", lambda: entries_html(cv["awards"])),
        ("memberships", lambda: entries_html(cv["memberships"])),
        ("presentations", lambda: presentations_html(cv["presentations"])),
        ("skills", lambda: skills_html(cv["skills"])),
    ]
    present = [(k, f) for k, f in order if cv.get(k)]
    toc = "".join(f'<li><a href="#cv-{k}">{esc(t[k])}</a></li>' for k, _ in present)
    body = (
        '<header class="page-head"><h1>Curriculum Vitae</h1>'
        f'<p class="page-lead">{esc(cv["contact"]["name"])}</p>'
        f'{contact_line_html(cv["contact"])}'
        f'<div class="actions"><a class="button" href="{esc(cv_file)}" download>Download CV (.docx)</a></div>'
        "</header>"
        f'<nav class="toc" aria-label="CV sections"><ul>{toc}</ul></nav>'
    )
    body += section(t["summary"], '<div class="prose">' + "".join(f"<p>{rich(p)}</p>" for p in cv["summary"]) + "</div>", "cv-summary")
    for k, f in present:
        body += section(t[k], f(), f"cv-{k}")
    return body


def render_contact(cv: dict) -> str:
    c = cv["contact"]
    affiliation, roles = current_position(cv)
    body = '<header class="page-head"><h1>Contact</h1></header>'
    body += '<div class="contact-grid">'
    body += section("Reach me", contact_links_html(c))
    aff = f"<p>{esc(affiliation)}</p>" + "".join(f'<p class="entry-sub">{esc(r)}</p>' for r in roles)
    body += section("Affiliation", portrait("contact-portrait") + f'<div class="prose">{aff}</div>')
    body += "</div>"
    if cv["interests"]:
        body += section(cv["titles"].get("interests", "Research Interests"), chips(cv["interests"]))
    return body


# --------------------------------------------------------------------------
# Clickable links in the published CV copy
# --------------------------------------------------------------------------

HYPERLINK_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
# A run holding only formatting and one piece of text; other runs are left alone.
SIMPLE_RUN_RE = re.compile(
    r"(<w:r(?:\s[^>]*)?>)((?:<w:rPr>(?:(?!</w:rPr>).)*</w:rPr>)?)<w:t(?:\s[^>]*)?>([^<]*)</w:t></w:r>", re.S
)
LINK_TARGET_RE = re.compile(r"(?:https?://|www\.)[^\s<>\"]+|[\w.+-]+@[\w-]+(?:\.[\w-]+)+")
# Word requires run properties in schema order, so color and underline go
# before the first property that must follow them.
AFTER_COLOR = r"spacing|w|kern|position|sz|szCs|highlight|u|effect|bdr|shd|fitText|vertAlign|rtl|cs|em|lang|eastAsianLayout|specVanish|oMath"
AFTER_UNDERLINE = r"effect|bdr|shd|fitText|vertAlign|rtl|cs|em|lang|eastAsianLayout|specVanish|oMath"


def _insert_before(rpr: str, tag: str, followers: str) -> str:
    m = re.search(rf"<w:(?:{followers})[\s/>]", rpr)
    pos = m.start() if m else rpr.rindex("</w:rPr>")
    return rpr[:pos] + tag + rpr[pos:]


def _link_rpr(rpr: str) -> str:
    if not rpr:
        return '<w:rPr><w:color w:val="0563C1"/><w:u w:val="single"/></w:rPr>'
    rpr = re.sub(r"<w:(?:color|u)\s[^>]*/>", "", rpr)
    rpr = _insert_before(rpr, '<w:color w:val="0563C1"/>', AFTER_COLOR)
    return _insert_before(rpr, '<w:u w:val="single"/>', AFTER_UNDERLINE)


def linkify_docx(src: Path, dst: Path) -> int:
    """Copy the CV, turning web addresses and emails into clickable links.

    Returns how many links were added. Falls back to a plain copy if the
    result would not be valid XML.
    """
    with zipfile.ZipFile(src) as z:
        xml = z.read("word/document.xml").decode("utf-8")
        rels = z.read("word/_rels/document.xml.rels").decode("utf-8")

    # Leave text that is already inside a link alone.
    linked = [(m.start(), m.end()) for m in re.finditer(r"<w:hyperlink\b.*?</w:hyperlink>", xml, re.S)]
    new_rels: list[str] = []

    def rewrite(m: re.Match) -> str:
        if any(a <= m.start() < b for a, b in linked):
            return m.group(0)
        open_tag, rpr, text = m.groups()
        if not LINK_TARGET_RE.search(text):
            return m.group(0)
        out, pos = [], 0

        def plain(t: str) -> str:
            return f'{open_tag}{rpr}<w:t xml:space="preserve">{t}</w:t></w:r>' if t else ""

        # An address running into the next run was split by Word; linking
        # only the first part would point somewhere wrong, so skip the run.
        if re.search(r"[^\s.,;)]$", text) and LINK_TARGET_RE.search(text[-80:].split()[-1] if text.split() else ""):
            nxt = re.match(r"(?:(?!</w:p>|<w:t[\s>]|<w:br\b|<w:tab\b).)*<w:t(?:\s[^>]*)?>([^<]{0,2})",
                           xml[m.end():m.end() + 2000], re.S)
            if nxt and re.match(r"[\w/%?=&#~+\-]|\.[^\s<]", nxt.group(1)):
                return m.group(0)
        for lm in LINK_TARGET_RE.finditer(text):
            shown = lm.group(0).rstrip(".,;)")
            if "@" in shown and "/" not in shown:
                target = "mailto:" + shown
            else:
                target = shown if shown.startswith("http") else "https://" + shown
            rid = f"rIdWebLink{len(new_rels) + 1}"
            new_rels.append(f'<Relationship Id="{rid}" Type="{HYPERLINK_TYPE}" Target="{target}" TargetMode="External"/>')
            out.append(plain(text[pos:lm.start()]))
            out.append(f'<w:hyperlink r:id="{rid}" w:history="1">{open_tag}{_link_rpr(rpr)}'
                       f'<w:t xml:space="preserve">{shown}</w:t></w:r></w:hyperlink>')
            pos = lm.start() + len(shown)
        out.append(plain(text[pos:]))
        return "".join(out)

    new_xml = SIMPLE_RUN_RE.sub(rewrite, xml)
    if not new_rels:
        shutil.copy2(src, dst)
        return 0
    root_tag = re.search(r"<w:document\b[^>]*>", new_xml).group(0)
    if "xmlns:r=" not in root_tag:
        new_xml = new_xml.replace(root_tag, root_tag[:-1] + f' xmlns:r="{R_NS}">', 1)
    new_rels_xml = rels.replace("</Relationships>", "".join(new_rels) + "</Relationships>")
    try:
        ET.fromstring(new_xml.encode("utf-8"))
        ET.fromstring(new_rels_xml.encode("utf-8"))
    except ET.ParseError as e:
        print(f"Warning: could not add links to the CV copy ({e}); publishing it unchanged.")
        shutil.copy2(src, dst)
        return 0

    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            if item.filename == "word/document.xml":
                data = new_xml.encode("utf-8")
            elif item.filename == "word/_rels/document.xml.rels":
                data = new_rels_xml.encode("utf-8")
            else:
                data = zin.read(item.filename)
            zout.writestr(item, data)
    return len(new_rels)


def visible_text(page: str) -> str:
    page = re.sub(r"(?is)<(script|style)\b.*?</\1>", " ", page)
    page = re.sub(r'(?i)\s(href|src)="[^"]*"', " ", page)
    return html.unescape(re.sub(r"<[^>]+>", " ", page))


def build(cv_path: Path = CV_PATH, out_dir: Path = DIST) -> Path:
    global NAME_RE, PHOTOS
    cv = parse_cv(cv_path)
    PHOTOS = load_photos()
    name = cv["contact"]["name"]
    NAME_RE = name_regex(name)
    cv_file = re.sub(r"[^A-Za-z0-9]+", "-", name).strip("-") + "-CV.docx"
    desc = tagline(cv)

    pages = {
        "index": ("Home", render_home(cv, cv_file)),
        "about": ("About", render_about(cv)),
        "research": ("Research", render_research(cv)),
        "teaching": ("Teaching", render_teaching(cv)),
        "cv": ("CV", render_cv(cv, cv_file)),
        "contact": ("Contact", render_contact(cv)),
    }

    if out_dir.exists():
        shutil.rmtree(out_dir)
    (out_dir / "assets").mkdir(parents=True)
    for slug, (title, body) in pages.items():
        page = layout(slug, title, body, cv, desc)
        found = PHONE_RE.search(visible_text(page))
        if found:
            raise BuildError(f"Refusing to publish: '{found.group(0)}' on the {title} page looks like a phone number.")
        (out_dir / f"{slug}.html").write_text(page, encoding="utf-8")
    for asset in ("style.css", "site.js"):
        shutil.copy2(ASSETS / asset, out_dir / "assets" / asset)
    used = {PHOTOS["portrait"]["file"]} if PHOTOS["portrait"] else set()
    used |= {p["file"] for items in PHOTOS["galleries"].values() for p in items}
    for rel in used:
        (out_dir / rel).parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / rel, out_dir / rel)
    n_links = linkify_docx(cv_path, out_dir / cv_file)
    (out_dir / ".nojekyll").write_text("", encoding="utf-8")

    n_pubs = sum(len(g["items"]) for g in cv["publications"])
    print(f"Built {len(pages)} pages into {out_dir.name}/ "
          f"({n_pubs} publications, {len(cv['presentations'])} presentations, "
          f"{len(cv['research']) + len(cv['teaching'])} positions; "
          f"{n_links} clickable links in the CV download).")
    return out_dir


def main(argv: list[str]) -> int:
    cv_path = Path(argv[1]) if len(argv) > 1 else CV_PATH
    try:
        build(cv_path)
    except BuildError as e:
        print(f"Build failed: {e}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
