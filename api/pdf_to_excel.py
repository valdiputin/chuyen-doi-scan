"""Vercel Serverless Function: Chuyển đổi PDF sang Excel dùng converter.py."""
from http.server import BaseHTTPRequestHandler
import json
import os
import sys
import tempfile
import urllib.parse

# Đảm bảo import được converter.py từ thư mục gốc
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))
for p in [ROOT_DIR, CURRENT_DIR, os.getcwd()]:
    if p not in sys.path:
        sys.path.insert(0, p)

try:
    from converter import Options, convert
except ImportError as e:
    convert = None
    import_err = str(e)
else:
    import_err = None


class handler(BaseHTTPRequestHandler):
    def _send_json_error(self, status_code: int, message: str):
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        payload = json.dumps({"error": message}, ensure_ascii=False).encode("utf-8")
        self.wfile.write(payload)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS, GET")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Filename, Authorization")
        self.end_headers()

    def do_GET(self):
        # Thông tin endpoint hoặc hướng dẫn
        self.send_response(200)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        res = {
            "status": "ready" if convert is not None else "error",
            "message": "Gửi POST với raw PDF body để chuyển sang Excel (.xlsx).",
            "import_error": import_err,
        }
        self.wfile.write(json.dumps(res, ensure_ascii=False).encode("utf-8"))

    def do_POST(self):
        if convert is None:
            self._send_json_error(500, f"Không thể nạp module converter: {import_err}")
            return

        try:
            content_length = int(self.headers.get("Content-Length", 0))
        except (ValueError, TypeError):
            content_length = 0

        # Giới hạn body 4.5MB của Vercel hobby
        MAX_BODY = 4600000
        if content_length > MAX_BODY:
            self._send_json_error(413, "Kích thước file vượt quá giới hạn 4.5MB của Vercel.")
            return

        # Đọc raw body (không dùng module cgi)
        pdf_bytes = self.rfile.read(content_length)
        if not pdf_bytes:
            self._send_json_error(400, "Dữ liệu PDF rỗng. Vui lòng gửi raw body PDF.")
            return

        # Kiểm tra magic bytes %PDF
        if not pdf_bytes.startswith(b"%PDF"):
            self._send_json_error(400, "Định dạng không hợp lệ, dữ liệu không phải file PDF.")
            return

        # Parse query string
        parsed_url = urllib.parse.urlparse(self.path)
        qs = urllib.parse.parse_qs(parsed_url.query)

        mode = qs.get("mode", ["auto"])[0]
        if mode not in ("auto", "lines", "text"):
            mode = "auto"

        one_sheet_val = qs.get("one_sheet", ["false"])[0].lower()
        one_sheet = one_sheet_val in ("true", "1", "yes")

        include_text_val = qs.get("include_text", ["false"])[0].lower()
        include_text = include_text_val in ("true", "1", "yes")

        pages = qs.get("pages", [None])[0]
        if pages:
            pages = pages.strip() or None

        opts = Options(
            mode=mode,
            one_sheet=one_sheet,
            include_text=include_text,
            pages=pages,
        )

        tmp_pdf = None
        tmp_xlsx = None
        try:
            # Ghi PDF vào file tạm
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as f_pdf:
                f_pdf.write(pdf_bytes)
                tmp_pdf = f_pdf.name

            # Tạo đường dẫn file tạm xlsx
            tmp_xlsx = tmp_pdf + ".xlsx"

            # Chạy converter.convert
            table_count = convert(tmp_pdf, tmp_xlsx, opts)

            # Đọc lại xlsx
            with open(tmp_xlsx, "rb") as f_out:
                xlsx_bytes = f_out.read()

            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            self.send_header("Content-Disposition", 'attachment; filename="converted.xlsx"')
            self.send_header("X-Table-Count", str(table_count))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.send_header("Access-Control-Expose-Headers", "X-Table-Count, Content-Disposition")
            self.end_headers()
            self.wfile.write(xlsx_bytes)

        except Exception as e:
            self._send_json_error(500, f"Lỗi trong quá trình chuyển đổi: {e}")
        finally:
            if tmp_pdf and os.path.exists(tmp_pdf):
                try:
                    os.unlink(tmp_pdf)
                except OSError:
                    pass
            if tmp_xlsx and os.path.exists(tmp_xlsx):
                try:
                    os.unlink(tmp_xlsx)
                except OSError:
                    pass

