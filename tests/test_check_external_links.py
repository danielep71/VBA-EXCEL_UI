"""Tests for the external link check (tools/check_external_links.py)."""

from __future__ import annotations

import http.server
import sys
import threading
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import check_external_links as links  # noqa: E402


class Handler(http.server.BaseHTTPRequestHandler):
    def do_GET(self) -> None:  # noqa: N802 - name required by BaseHTTPRequestHandler
        if self.path == "/ok":
            self.send_response(200)
        elif self.path == "/moved":
            self.send_response(301)
            self.send_header("Location", "/ok")
        elif self.path == "/loop":
            self.send_response(302)
            self.send_header("Location", "/loop")
        elif self.path == "/nowhere":
            self.send_response(302)
        elif self.path == "/gone":
            self.send_response(410)
        elif self.path == "/forbidden":
            self.send_response(403)
        else:
            self.send_response(404)
        self.send_header("Content-Length", "0")
        self.end_headers()

    def log_message(self, format: str, *args: object) -> None:  # noqa: A002
        pass


class CheckUrlTests(unittest.TestCase):
    server: http.server.ThreadingHTTPServer

    @classmethod
    def setUpClass(cls) -> None:
        cls.server = http.server.ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        threading.Thread(target=cls.server.serve_forever, daemon=True).start()

    @classmethod
    def tearDownClass(cls) -> None:
        cls.server.shutdown()
        cls.server.server_close()

    def outcome(self, path: str) -> str:
        url = f"http://127.0.0.1:{self.server.server_address[1]}{path}"
        return links.check_url(url, attempts=2, timeout=5, delay=0)[0]

    def test_success_and_resolved_redirect_pass(self) -> None:
        self.assertEqual(self.outcome("/ok"), "ok")
        self.assertEqual(self.outcome("/moved"), "ok")

    def test_redirect_that_never_resolves_is_broken(self) -> None:
        self.assertEqual(self.outcome("/loop"), "broken")
        self.assertEqual(self.outcome("/nowhere"), "broken")

    def test_missing_pages_are_broken(self) -> None:
        self.assertEqual(self.outcome("/missing"), "broken")
        self.assertEqual(self.outcome("/gone"), "broken")

    def test_access_controlled_pages_are_unverified(self) -> None:
        self.assertEqual(self.outcome("/forbidden"), "unverified")


class ExtractLinksTests(unittest.TestCase):
    def test_markdown_forms_are_extracted(self) -> None:
        text = ('[a](https://a.example/x#frag "title") <https://b.example>\n\n'
                "[ref]: https://c.example/y\n")
        self.assertEqual(links.extract_links(text),
                         ["https://a.example/x#frag", "https://b.example", "https://c.example/y"])

    def test_html_sources_spanning_lines_are_extracted(self) -> None:
        text = ('<p align="center">\n  <img width="100%"\n       alt="Demo"\n'
                '       src="https://img.example/demo.png">\n</p>\n'
                "<a href='https://d.example/page'>x</a>\n")
        self.assertEqual(links.extract_links(text),
                         ["https://img.example/demo.png", "https://d.example/page"])

    def test_code_is_ignored(self) -> None:
        text = ("`https://code.example` and `<img src=\"https://code.example/i.png\">`\n"
                "```\n[x](https://fenced.example)\n<img src=\"https://fenced.example/i.png\">\n```\n")
        self.assertEqual(links.extract_links(text), [])

    def test_readme_images_are_collected(self) -> None:
        urls = links.collect(ROOT, ["README.md"])
        self.assertTrue(any("user-attachments/assets" in url for url in urls))


if __name__ == "__main__":
    unittest.main()
