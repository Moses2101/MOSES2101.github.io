#!/usr/bin/env python3
"""Publish the newest CV in one command.

Copies the newest .docx from your CV folder into cv/cv.docx, shows what
changed, rebuilds the site, then commits and pushes so GitHub Pages
redeploys.

Usage:
    python publish.py                      # newest CV from the default folder
    python publish.py "D:\\Other folder"    # newest CV from another folder
    python publish.py --remove-phone       # strip phone numbers from the published CV copy
    python publish.py --no-push            # commit locally, push later
    python publish.py --yes                # skip the confirmation prompt
"""
from __future__ import annotations

import argparse
import difflib
import re
import shutil
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path

import build

DEFAULT_SOURCE = Path(r"C:\Users\ogunn\OneDrive - The University of Alabama\Desktop\JESUS jobs application\Moses CV")

T_RE = re.compile(r"(<w:t(?:\s[^>]*)?>)([^<]*)(</w:t>)")
# No leading \b: text runs from adjacent paragraphs are joined without a separator.
LABEL_BEFORE_RE = re.compile(
    r"(?i)(?:phone|\btel(?:ephone)?|mobile|cell(?:phone)?|whatsapp)\s*(?:no\.?|number|#)?\s*[:.]?\s*$"
)
REDACT_PARTS = re.compile(r"word/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml")


def newest_cv(folder: Path) -> Path:
    if not folder.is_dir():
        sys.exit(f"CV folder not found: {folder}")
    docs = [p for p in folder.glob("*.docx") if not p.name.startswith("~$")]
    if not docs:
        sys.exit(f"No .docx files in {folder}")
    return max(docs, key=lambda p: p.stat().st_mtime)


def redact_xml(xml: str) -> tuple[str, int]:
    """Delete phone numbers (and a 'Phone:' label before them) from Word XML text runs."""
    matches = list(T_RE.finditer(xml))
    texts = [m.group(2) for m in matches]
    full = "".join(texts)
    keep = [True] * len(full)
    found = 0
    for m in build.PHONE_RE.finditer(full):
        found += 1
        start = m.start()
        window = max(0, start - 30)
        label = LABEL_BEFORE_RE.search(full[window:start])
        if label:
            start = window + label.start()
        for i in range(start, m.end()):
            keep[i] = False
    if not found:
        return xml, 0
    out, pos, offset = [], 0, 0
    for m, t in zip(matches, texts):
        out.append(xml[pos:m.start(2)])
        out.append("".join(ch for i, ch in enumerate(t) if keep[offset + i]))
        offset += len(t)
        pos = m.end(2)
    out.append(xml[pos:])
    return "".join(out), found


def remove_phone_numbers(src: Path, dst: Path) -> int:
    total = 0
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if REDACT_PARTS.fullmatch(item.filename):
                xml, n = redact_xml(data.decode("utf-8"))
                data, total = xml.encode("utf-8"), total + n
            zout.writestr(item, data)
    return total


def show_changes(old: Path, new: Path) -> int:
    before = build.section_outline(old) if old.exists() else []
    after = build.section_outline(new)
    diff = [l for l in difflib.unified_diff(before, after, "published CV", "new CV", n=0, lineterm="")
            if not l.startswith("@@")]
    if not diff:
        print("The text of the CV is unchanged (formatting may differ).")
        return 0
    added = sum(1 for l in diff if l.startswith("+") and not l.startswith("+++"))
    removed = sum(1 for l in diff if l.startswith("-") and not l.startswith("---"))
    print(f"\nChanges: {added} paragraph(s) added or edited, {removed} removed or replaced.\n")
    for line in diff[2:202]:
        print(("  " if line[0] not in "+-" else "") + line[:240])
    if len(diff) > 202:
        print(f"  ... and {len(diff) - 202} more lines")
    print()
    return added + removed


def git(*args: str, check: bool = True) -> subprocess.CompletedProcess:
    return subprocess.run(["git", *args], cwd=build.ROOT, check=check, text=True)


def main() -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(errors="replace")
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("folder", nargs="?", type=Path, default=DEFAULT_SOURCE, help="folder holding your CV .docx")
    parser.add_argument("--remove-phone", action="store_true", help="remove phone numbers from the published CV copy")
    parser.add_argument("-m", "--message", help="commit message")
    parser.add_argument("--no-push", action="store_true", help="commit but do not push")
    parser.add_argument("-y", "--yes", action="store_true", help="do not ask for confirmation")
    args = parser.parse_args()

    if shutil.which("git") is None:
        sys.exit("Git is not installed or not on PATH. See README.md, 'One-time setup'.")
    if not (build.ROOT / ".git").exists():
        sys.exit("This folder is not a Git repository yet. See README.md, 'One-time setup'.")

    source = newest_cv(args.folder)
    print(f"Newest CV: {source.name}")

    with tempfile.TemporaryDirectory() as tmp:
        staged = Path(tmp) / "cv.docx"
        shutil.copyfile(source, staged)
        if args.remove_phone:
            redacted = Path(tmp) / "cv-redacted.docx"
            n = remove_phone_numbers(staged, redacted)
            staged = redacted
            print(f"Removed {n} phone number(s) from the published copy." if n else "No phone numbers found in the CV.")

        target = build.CV_PATH
        if target.exists() and target.read_bytes() == staged.read_bytes():
            print("This CV is identical to the published one. Nothing to do.")
            return 0
        show_changes(target, staged)

        # Build from the staged copy first so a broken CV never replaces the good one.
        try:
            build.build(staged)
        except build.BuildError as e:
            print(f"Build failed, nothing was published: {e}", file=sys.stderr)
            return 1

        if not args.yes:
            answer = input("Publish this version? [y/N] ").strip().lower()
            if answer not in ("y", "yes"):
                print("Cancelled. Nothing was committed.")
                return 0
        target.parent.mkdir(exist_ok=True)
        shutil.copyfile(staged, target)

    build.build()
    git("add", "-A")
    if git("diff", "--cached", "--quiet", "--no-ext-diff", "--no-textconv", check=False).returncode == 0:
        print("Nothing to commit.")
        return 0
    git("commit", "-m", args.message or f"Update CV from {source.name}")
    if args.no_push:
        print("Committed. Run 'git push' when you are ready to publish.")
        return 0
    git("push")
    print("Pushed. GitHub Pages will redeploy in a minute or two (see the Actions tab).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
