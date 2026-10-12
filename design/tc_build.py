#!/usr/bin/env python3
"""
tc_build.py: assemble Technology Consigliere install pages.

Every published page is one self-contained HTML file. This script fills two
kinds of marked regions in each page, so nothing is ever hand-pasted:

  <!-- TC:STYLE -->  ...  <!-- /TC:STYLE -->
      Replaced with design/tc.css, inlined in a <style> element.

  <!-- TC:BOOKMARKLET src="dist/crosslist.bookmarklet.txt" label="Cross-list sections" -->
  ...
  <!-- /TC:BOOKMARKLET -->
      Replaced with the drag button. src is relative to the repository root
      and holds the encoded javascript: URL. label becomes the bookmark name,
      so keep it short and verb-first.

Usage, from the repository root:

    python3 design/tc_build.py            rebuild every page listed in PAGES
    python3 design/tc_build.py --check    exit 1 if any page is out of date
    python3 design/tc_build.py some.html  build just the pages you name

Run --check before you push. A page that fails it is either stale or was
edited by hand inside a marked region.

Same file in every repo. Only the PAGES list changes.
"""

from pathlib import Path
import html
import re
import sys

ROOT = Path(__file__).resolve().parent.parent
CSS = ROOT / "design" / "tc.css"

# Edit per repository. Paths are relative to the repository root.
PAGES = [
    "index.html",
    "teacher/index.html",
    "teacher/install-exporter.html",
    "design/template.html",
]

STYLE_RE = re.compile(r"<!-- TC:STYLE -->.*?<!-- /TC:STYLE -->", re.S)
BOOK_RE = re.compile(
    r'<!-- TC:BOOKMARKLET src="(?P<src>[^"]+)" label="(?P<label>[^"]+)" -->'
    r".*?<!-- /TC:BOOKMARKLET -->",
    re.S,
)
# A bare double quote ends the href attribute and a bare < can open a tag.
# Single quotes are safe inside a double-quoted attribute, and the exporter
# build contains them, so they are allowed.
UNSAFE = ('"', "<", ">", " ", "\n")


def style_block(css: str) -> str:
    return "<!-- TC:STYLE -->\n<style>\n" + css.strip() + "\n</style>\n<!-- /TC:STYLE -->"


def bookmarklet_block(m: re.Match) -> str:
    src, label = m.group("src"), m.group("label")
    path = ROOT / src
    if not path.exists():
        raise SystemExit(f"Missing bookmarklet build: {src}")
    url = path.read_text(encoding="utf-8").strip()
    if not url.startswith("javascript:"):
        raise SystemExit(f"{src} does not start with javascript:")
    for bad in UNSAFE:
        if bad in url:
            raise SystemExit(f"{src} contains {bad!r}, which would break the href.")
    anchor = f'<a class="tc-grab" href="{url}">{html.escape(label, quote=False)}</a>'
    return (
        f'<!-- TC:BOOKMARKLET src="{src}" label="{label}" -->\n'
        f"{anchor}\n<!-- /TC:BOOKMARKLET -->"
    )


def build(page: str, css: str) -> str:
    if not STYLE_RE.search(page):
        raise SystemExit("Page has no TC:STYLE markers.")
    out = STYLE_RE.sub(lambda _: style_block(css), page, count=1)
    out = BOOK_RE.sub(bookmarklet_block, out)
    return out


def main(argv: list[str]) -> int:
    check = "--check" in argv
    pages = [a for a in argv if not a.startswith("--")] or PAGES
    if not pages:
        print("PAGES is empty. List this repo's pages at the top of the script.", file=sys.stderr)
        return 1
    css = CSS.read_text(encoding="utf-8")
    stale = []
    for rel in pages:
        path = ROOT / rel
        before = path.read_text(encoding="utf-8")
        after = build(before, css)
        if before == after:
            print(f"ok       {rel}")
        elif check:
            stale.append(rel)
            print(f"STALE    {rel}")
        else:
            path.write_text(after, encoding="utf-8")
            print(f"rebuilt  {rel} ({len(after.encode()) / 1024:.0f} KB)")
    if stale:
        print("\nRun python3 design/tc_build.py and commit the result.", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
