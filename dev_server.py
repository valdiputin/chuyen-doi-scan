"""Local development server phục vụ thư mục public/ và các API routes.

Chạy:
    python dev_server.py
Mở trình duyệt:
    http://localhost:8000
"""
from http.server import HTTPServer, SimpleHTTPRequestHandler
import argparse
import os
import sys
import urllib.parse

# Đảm bảo root dir trong sys.path
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

PUBLIC_DIR = os.path.join(BASE_DIR, "public")

# Tự đọc file .env nếu có
env_path = os.path.join(BASE_DIR, ".env")
if os.path.exists(env_path):
    with open(env_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                k, v = line.split("=", 1)
                os.environ.setdefault(k.strip(), v.strip())

import api.health as api_health
import api.office_convert as api_office_convert
import api.pdf_to_excel as api_pdf_to_excel

# Đảm bảo Content-Type có charset=utf-8 cho static files
SimpleHTTPRequestHandler.extensions_map.update({
    ".html": "text/html; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".json": "application/json; charset=utf-8",
})


class DevHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=PUBLIC_DIR, **kwargs)

    def _get_api_handler(self):
        parsed = urllib.parse.urlparse(self.path)
        p = parsed.path.rstrip("/")
        if p in ("/api/pdf-to-excel", "/api/pdf_to_excel"):
            return api_pdf_to_excel.handler
        elif p in ("/api/office-convert", "/api/office_convert"):
            return api_office_convert.handler
        elif p in ("/api/health",):
            return api_health.handler
        return None

    def _dispatch(self, handler_cls, method_name):
        inst = object.__new__(handler_cls)
        inst.request = self.request
        inst.client_address = self.client_address
        inst.server = self.server
        inst.rfile = self.rfile
        inst.wfile = self.wfile
        inst.headers = self.headers
        inst.path = self.path
        inst.command = self.command
        inst.requestline = getattr(self, "requestline", f"{self.command} {self.path} {self.request_version}")
        inst.request_version = self.request_version
        inst.close_connection = self.close_connection
        method = getattr(inst, method_name, None)
        if method:
            method()
        else:
            self.send_response(405)
            self.end_headers()

    def do_OPTIONS(self):
        h = self._get_api_handler()
        if h:
            self._dispatch(h, "do_OPTIONS")
        else:
            self.send_response(204)
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
            self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Filename, Authorization")
            self.end_headers()

    def do_GET(self):
        h = self._get_api_handler()
        if h:
            self._dispatch(h, "do_GET")
        else:
            super().do_GET()

    def do_POST(self):
        h = self._get_api_handler()
        if h:
            self._dispatch(h, "do_POST")
        else:
            self.send_response(404)
            self.send_header("Content-Type", "application/json")
            self.end_headers()
            self.wfile.write(b'{"error": "API route not found"}')


def run(port: int = 8000):
    os.makedirs(PUBLIC_DIR, exist_ok=True)
    server_address = ("", port)
    httpd = HTTPServer(server_address, DevHandler)
    print(f"=====================================================")
    print(f" [PDF-TO-EXCEL] Dev server running at http://localhost:{port}")
    print(f" Serving static files from: {PUBLIC_DIR}")
    print(f" API Routes: /api/pdf-to-excel, /api/office-convert, /api/health")
    print(f" GOTENBERG_URL: {os.environ.get('GOTENBERG_URL', '(chưa cấu hình - chế độ dự phòng trình duyệt)')}")
    print(f" Nhấn Ctrl+C để dừng server.")
    print(f"=====================================================")
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        print("\nĐang dừng dev server...")
        httpd.server_close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Chạy local dev server")
    parser.add_argument("--port", type=int, default=int(os.environ.get("PORT", 8000)), help="Cổng server (mặc định 8000)")
    args = parser.parse_args()
    run(args.port)
