"""Local static server that mimics GitHub Pages' extensionless URLs
(/merge-into-tabs -> merge-into-tabs.html). Run from the repo root:
    python3 test/serve.py [port]
"""
import http.server
import os
import sys


class Handler(http.server.SimpleHTTPRequestHandler):
    def translate_path(self, path):
        local = super().translate_path(path)
        if not os.path.exists(local) and os.path.exists(local + ".html"):
            return local + ".html"
        return local

    def log_message(self, *args):
        pass


port = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
http.server.ThreadingHTTPServer(("127.0.0.1", port), Handler).serve_forever()
