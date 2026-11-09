"""Vercel Serverless Function: Proxy chuyển đổi Office/HTML/URL sang PDF qua Gotenberg 8."""
from http.server import BaseHTTPRequestHandler
import json
import os
import time
import urllib.parse
import requests

GOTENBERG_URL = os.environ.get("GOTENBERG_URL", "").rstrip("/")
GOTENBERG_TOKEN = os.environ.get("GOTENBERG_TOKEN", "")

# Timeout tối thiểu 90s cho Hugging Face Space khi thức dậy
TIMEOUT_SECONDS = 100


class handler(BaseHTTPRequestHandler):
    def _send_json_error(self, status_code: int, error_code: str, message: str):
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        payload = json.dumps({"error": error_code, "message": message}, ensure_ascii=False).encode("utf-8")
        self.wfile.write(payload)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Filename, Authorization")
        self.end_headers()

    def do_GET(self):
        # Kiểm tra trạng thái cấu hình
        if not GOTENBERG_URL:
            self._send_json_error(503, "NOT_CONFIGURED", "GOTENBERG_URL chưa được cấu hình.")
            return

        headers = {}
        if GOTENBERG_TOKEN:
            headers["Authorization"] = f"Bearer {GOTENBERG_TOKEN}"

        try:
            r = requests.get(f"{GOTENBERG_URL}/health", headers=headers, timeout=10)
            self.send_response(r.status_code)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            self.wfile.write(r.content)
        except Exception as e:
            self._send_json_error(502, "UNREACHABLE", f"Không thể kết nối đến Gotenberg: {e}")

    def do_POST(self):
        if not GOTENBERG_URL:
            self._send_json_error(503, "NOT_CONFIGURED", "GOTENBERG_URL chưa được cấu hình. Sử dụng phương án xử lý trên trình duyệt.")
            return

        parsed_url = urllib.parse.urlparse(self.path)
        qs = urllib.parse.parse_qs(parsed_url.query)

        conv_type = qs.get("type", ["libreoffice"])[0].lower()
        landscape = qs.get("landscape", ["false"])[0].lower() in ("true", "1", "yes")

        try:
            content_length = int(self.headers.get("Content-Length", 0))
        except (ValueError, TypeError):
            content_length = 0

        # Max body 4.5MB
        if content_length > 4600000:
            self._send_json_error(413, "FILE_TOO_LARGE", "Kích thước request vượt quá 4.5MB.")
            return

        body = self.rfile.read(content_length)

        # Chuẩn bị request tới Gotenberg
        req_headers = {}
        if GOTENBERG_TOKEN:
            req_headers["Authorization"] = f"Bearer {GOTENBERG_TOKEN}"

        target_endpoint = None
        data_fields = {}
        files_data = None

        if conv_type == "libreoffice":
            target_endpoint = f"{GOTENBERG_URL}/forms/libreoffice/convert"
            filename = self.headers.get("X-Filename") or qs.get("filename", ["document.docx"])[0]
            # Giữ tên file an toàn
            filename = os.path.basename(urllib.parse.unquote(filename))
            files_data = {"files": (filename, body)}
            if landscape:
                data_fields["landscape"] = "true"

        elif conv_type == "url":
            target_endpoint = f"{GOTENBERG_URL}/forms/chromium/convert/url"
            url_to_convert = None
            if body:
                try:
                    j = json.loads(body.decode("utf-8"))
                    url_to_convert = j.get("url")
                except Exception:
                    url_to_convert = body.decode("utf-8").strip()
            if not url_to_convert:
                url_to_convert = qs.get("url", [None])[0]

            if not url_to_convert:
                self._send_json_error(400, "MISSING_URL", "Thiếu tham số url để chuyển đổi.")
                return

            data_fields["url"] = url_to_convert
            data_fields["printBackground"] = "true"
            if landscape:
                data_fields["landscape"] = "true"

        elif conv_type == "html":
            target_endpoint = f"{GOTENBERG_URL}/forms/chromium/convert/html"
            files_data = {"files": ("index.html", body)}
            data_fields["printBackground"] = "true"
            if landscape:
                data_fields["landscape"] = "true"

        else:
            self._send_json_error(400, "INVALID_TYPE", f"Loại chuyển đổi không hỗ trợ: {conv_type}")
            return

        # Thực hiện gọi Gotenberg với retry 1 lần nếu gặp 502/503 (Space đang khởi động)
        max_attempts = 2
        last_error = None
        for attempt in range(1, max_attempts + 1):
            try:
                if files_data:
                    res = requests.post(
                        target_endpoint,
                        headers=req_headers,
                        data=data_fields,
                        files=files_data,
                        timeout=TIMEOUT_SECONDS,
                    )
                else:
                    res = requests.post(
                        target_endpoint,
                        headers=req_headers,
                        data=data_fields,
                        timeout=TIMEOUT_SECONDS,
                    )

                # Nếu gặp 502/503 (Space đang thức dậy), thử lại 1 lần sau khi chờ
                if res.status_code in (502, 503) and attempt < max_attempts:
                    time.sleep(3)
                    continue

                if res.status_code == 200:
                    self.send_response(200)
                    self.send_header("Content-Type", "application/pdf")
                    self.send_header("Content-Disposition", 'attachment; filename="converted.pdf"')
                    self.send_header("Access-Control-Allow-Origin", "*")
                    self.end_headers()
                    self.wfile.write(res.content)
                    return
                else:
                    msg = res.text or f"Gotenberg error HTTP {res.status_code}"
                    self._send_json_error(res.status_code, "GOTENBERG_ERROR", msg)
                    return

            except (requests.Timeout, requests.ConnectionError) as e:
                last_error = e
                if attempt < max_attempts:
                    time.sleep(3)
                    continue
            except Exception as e:
                self._send_json_error(500, "PROXY_EXCEPTION", str(e))
                return

        self._send_json_error(504, "TIMEOUT", f"Server Gotenberg không phản hồi sau {TIMEOUT_SECONDS}s: {last_error}")

