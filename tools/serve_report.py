#!/usr/bin/env python3
"""Serve reports/ on 0.0.0.0:8000; '/' redirects to the dashboard."""
from __future__ import annotations

import http.server
import os

ROOT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "reports")


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=ROOT, **kwargs)

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self.send_response(302)
            self.send_header("Location", "/dashboard.html")
            self.end_headers()
            return
        super().do_GET()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "8000"))
    print(f"http://0.0.0.0:{port}/dashboard.html -> {os.path.abspath(ROOT)}", flush=True)
    http.server.ThreadingHTTPServer(("0.0.0.0", port), Handler).serve_forever()
