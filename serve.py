#!/usr/bin/env python3
"""Preview the site locally, with clean addresses like GitHub Pages.

Usage:
    python serve.py              # rebuild, then serve dist/ at http://localhost:8000
    python serve.py --port 9000
    python serve.py --no-build   # serve dist/ as it is

/about serves about.html, / serves index.html, exactly as on GitHub Pages.
"""
from __future__ import annotations

import argparse
import functools
import http.server
import os
import sys
import webbrowser

import build


class CleanURLHandler(http.server.SimpleHTTPRequestHandler):
    def translate_path(self, path):
        fs_path = super().translate_path(path)
        if not os.path.exists(fs_path) and os.path.isfile(fs_path + ".html"):
            return fs_path + ".html"
        return fs_path

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        super().end_headers()


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--port", type=int, default=8000)
    parser.add_argument("--no-build", action="store_true", help="serve dist/ without rebuilding")
    parser.add_argument("--no-open", action="store_true", help="do not open a browser window")
    args = parser.parse_args()

    if not args.no_build:
        try:
            build.build()
        except build.BuildError as e:
            print(f"Build failed: {e}", file=sys.stderr)
            return 1
    if not build.DIST.is_dir():
        print("dist/ does not exist yet. Run without --no-build.", file=sys.stderr)
        return 1

    handler = functools.partial(CleanURLHandler, directory=str(build.DIST))
    server = http.server.ThreadingHTTPServer(("127.0.0.1", args.port), handler)
    url = f"http://localhost:{args.port}/"
    print(f"Previewing at {url}  (press Ctrl+C to stop)")
    if not args.no_open:
        webbrowser.open(url)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nStopped.")
    finally:
        server.server_close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
