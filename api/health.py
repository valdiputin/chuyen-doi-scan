"""Vercel Serverless Function: Kiểm tra sức khỏe hệ thống và Gotenberg."""
from http.server import BaseHTTPRequestHandler
import json
import os
import requests

GOTENBERG_URL = os.environ.get("GOTENBERG_URL", "").rstrip("/")
GOTENBERG_TOKEN = os.environ.get("GOTENBERG_TOKEN", "")


class handler(BaseHTTPRequestHandler):
    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.end_headers()

    def do_GET(self):
        result = {
            "status": "ok",
            "app": "Document Converter Web App",
            "gotenberg": {
                "configured": bool(GOTENBERG_URL),
                "url": GOTENBERG_URL if GOTENBERG_URL else None,
                "status": "not_configured",
                "message": "GOTENBERG_URL chưa được cấu hình. Các tác vụ Office sẽ dùng phương án dự phòng trên trình duyệt.",
            },
        }

        if GOTENBERG_URL:
            headers = {}
            if GOTENBERG_TOKEN:
                headers["Authorization"] = f"Bearer {GOTENBERG_TOKEN}"

            try:
                # Timeout ngắn để không làm chậm health check
                r = requests.get(f"{GOTENBERG_URL}/health", headers=headers, timeout=6)
                if r.status_code == 200:
                    result["gotenberg"]["status"] = "ok"
                    result["gotenberg"]["message"] = "Gotenberg đang hoạt động bình thường."
                    try:
                        result["gotenberg"]["details"] = r.json()
                    except Exception:
                        pass
                elif r.status_code in (502, 503):
                    result["gotenberg"]["status"] = "sleeping"
                    result["gotenberg"]["message"] = "Hugging Face Space đang ngủ hoặc đang khởi động."
                else:
                    result["gotenberg"]["status"] = "error"
                    result["gotenberg"]["message"] = f"Gotenberg trả về mã lỗi HTTP {r.status_code}."
            except requests.Timeout:
                result["gotenberg"]["status"] = "sleeping"
                result["gotenberg"]["message"] = "Hugging Face Space đang ngủ (timeout khi kết nối)."
            except Exception as e:
                result["gotenberg"]["status"] = "unreachable"
                result["gotenberg"]["message"] = f"Không thể kết nối đến Gotenberg: {e}"

        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(json.dumps(result, ensure_ascii=False, indent=2).encode("utf-8"))

