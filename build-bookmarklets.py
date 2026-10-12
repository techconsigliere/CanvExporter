#!/usr/bin/env python3
"""
build-bookmarklets.py: encode every bookmarklet from its readable source.

Reads each source file in TOOLS, URL-encodes it, prefixes javascript:, and
writes the result next to the other builds. Then run design/tc_build.py to
put the new builds into the install pages.

Run from the repository root:

    python3 build-bookmarklets.py
    python3 design/tc_build.py

Comments are kept on purpose. Stripping JavaScript comments correctly means
parsing the language, because // and /* show up inside strings and regular
expressions, and a naive regex quietly breaks the tool. So every build
decodes back to its source byte for byte, which is also the easiest thing
for a cautious reader to verify.
"""

from pathlib import Path
from urllib.parse import quote
import sys

ROOT = Path(__file__).parent

# source file -> encoded build. Paths are relative to the repository root.
TOOLS = {
    # Not dist/: this repo's .gitignore excludes every dist/ folder.
    "teacher/exporter_src.js": "teacher/exporter.bookmarklet.txt",
}

# Characters left unencoded. Everything else is percent-encoded, including
# quotes and angle brackets (the URL lands inside href="...") and # (some
# browsers treat a bare # as the start of a fragment and drop the rest).
SAFE = "!$&*+,-./:;=?@_~"


def main() -> int:
    for src, out in TOOLS.items():
        source = ROOT / src
        if not source.exists():
            print(f"Missing {src}. Run this from the repository root.", file=sys.stderr)
            return 1
        url = "javascript:" + quote(source.read_text(encoding="utf-8"), safe=SAFE)
        for bad in ('"', "<", ">", " ", "#", "\n"):
            if bad in url:
                print(f"{out} would contain {bad!r}, which breaks the bookmark.", file=sys.stderr)
                return 1
        target = ROOT / out
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(url + "\n", encoding="utf-8")
        print(f"wrote    {out} ({len(url) / 1024:.0f} KB)")
    print("\nNow run: python3 design/tc_build.py")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
