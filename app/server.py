"""HTTP routing and server startup."""

import os
from http.server import ThreadingHTTPServer

from .legacy import AppHandler, ensure_dirs


HOST = "0.0.0.0"
PORT = int(os.environ.get("PORT", "8000"))


def run(host=HOST, port=PORT):
    ensure_dirs()
    server = ThreadingHTTPServer((host, port), AppHandler)
    print(f"Tata Inventory Criteria Checker running at http://{host}:{port}")
    server.serve_forever()
