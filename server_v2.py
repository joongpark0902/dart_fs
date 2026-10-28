"""Versioned calendar-style UI; reuse the existing API without modifying it."""
from pathlib import Path
from urllib.parse import urlparse

import server


class CalendarHandler(server.Handler):
    def do_GET(self):
        if urlparse(self.path).path == "/":
            page = Path(__file__).parent / "web" / "index_v2.html"
            return self._send(200, page.read_bytes(), "text/html; charset=utf-8")
        return super().do_GET()


if __name__ == "__main__":
    server.Handler = CalendarHandler
    server.main()
