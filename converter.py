"""Chuyển bảng trong PDF sang Excel (pdfplumber + openpyxl)."""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Optional

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

LINES = {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
TEXT = {"vertical_strategy": "text", "horizontal_strategy": "text"}


@dataclass
class Options:
    mode: str = "auto"          # auto | lines | text
    one_sheet: bool = False     # gộp tất cả bảng vào 1 sheet
    include_text: bool = False  # trang không có bảng -> xuất text từng dòng
    pages: Optional[str] = None # "1-3,5"


def parse_pages(spec: Optional[str], total: int) -> list[int]:
    if not spec or not spec.strip():
        return list(range(total))
    result: list[int] = []
    for part in spec.replace(" ", "").split(","):
        if "-" in part:
            a, b = part.split("-", 1)
            result.extend(range(int(a) - 1, min(int(b), total)))
        elif part:
            result.append(int(part) - 1)
    return [p for p in dict.fromkeys(result) if 0 <= p < total]


def _clean(cell) -> str:
    return "" if cell is None else str(cell).replace("\r", "").strip()


def _unwrap(text: str) -> str:
    """Nối dòng bị ngắt do cột hẹp: '07:\\n59:17' -> '07:59:17'."""
    lines = [l.strip() for l in text.split("\n") if l.strip()]
    if len(lines) < 2:
        return text
    out = lines[0]
    for l in lines[1:]:
        num_split = (out[-1] == "," and len(out) > 1 and out[-2].isdigit() and l[0].isdigit()) or \
                    (l[0] in ",." and out[-1].isdigit() and l[1:2].isdigit())  # '6,098,429\n,598'
        if out.endswith((":", "/", "-", ".")) or (out[-1].isdigit() and l[0].isdigit()) or num_split:
            out += l
        else:
            out += " " + l
    return out


def _to_value(text: str):
    """Chuyển số kiểu '1.234,5' hoặc '1,234.5' thành số nếu được."""
    t = text.replace(" ", "")
    if not t or any(c.isalpha() for c in t):
        return text
    # Mã/số tài khoản: bắt đầu bằng 0 hoặc dãy số dài -> giữ dạng text
    # Có thể kèm dấu +/- ở đầu, vd '-0201245001' (không được thành số âm)
    digits = t.lstrip("+-")
    if digits.isdigit() and (len(digits) >= 10 or (digits.startswith("0") and len(digits) > 1)):
        return text
    candidates = [t, t.replace(",", ""), t.replace(".", "").replace(",", ".")]
    for c in candidates:
        try:
            v = float(c)
            return int(v) if v.is_integer() and "." not in c else v
        except ValueError:
            continue
    return text


def _extract(page, mode: str) -> list[list[list[str]]]:
    if mode == "lines":
        tables = page.extract_tables(LINES)
    elif mode == "text":
        tables = page.extract_tables(TEXT)
    else:
        tables = page.extract_tables(LINES) or page.extract_tables(TEXT)
    out = []
    for t in tables:
        rows = [[_clean(c) for c in row] for row in t]
        rows = [rows[0]] + [[_unwrap(c) for c in r] for r in rows[1:]]
        rows = [r for r in rows if any(r)]
        if rows:
            out.append(rows)
    return out


HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
THIN = Side(style="thin", color="999999")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def _account_cols(header) -> set[int]:
    """Vị trí các cột 'Tài khoản' theo header."""
    return {i for i, h in enumerate(header) if "tài khoản" in h.lower()}


def _account_value(text: str) -> str:
    """Cột tài khoản: luôn là text; bỏ dấu +/- ở đầu, vd '- 0123' -> '0123'."""
    t = text.replace(" ", "")
    digits = t.lstrip("+-")
    return digits if digits.isdigit() else text


def _cell_value(text: str, c_i: int, acc_cols: set[int]):
    return _account_value(text) if c_i in acc_cols else _to_value(text)


def _write_table(ws, rows, start_row: int, skip_header: bool = False) -> int:
    """Ghi bảng từ start_row; trả về dòng cuối cùng đã ghi."""
    acc_cols = _account_cols(rows[0])
    if skip_header:
        rows = [[""] * len(rows[0])] + rows[1:]  # giữ chỉ số r_i, không ghi header
        for r_i, row in enumerate(rows[1:], 1):
            for c_i, text in enumerate(row):
                cell = ws.cell(row=start_row + r_i, column=c_i + 1,
                               value=_cell_value(text, c_i, acc_cols))
                cell.border = BORDER
                cell.alignment = Alignment(wrap_text=True, vertical="top")
        return start_row + len(rows) - 1
    for r_i, row in enumerate(rows):
        for c_i, text in enumerate(row):
            cell = ws.cell(row=start_row + r_i, column=c_i + 1,
                           value=text if r_i == 0 else _cell_value(text, c_i, acc_cols))
            cell.border = BORDER
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            if r_i == 0:
                cell.font = Font(bold=True)
                cell.fill = HEADER_FILL
    return start_row + len(rows) - 1


def _autosize(ws):
    widths: dict[int, int] = {}
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None:
                longest = max(len(s) for s in str(cell.value).split("\n"))
                widths[cell.column] = max(widths.get(cell.column, 0), longest)
    for col, w in widths.items():
        ws.column_dimensions[get_column_letter(col)].width = min(max(w + 2, 8), 60)


def convert(pdf_path: str | Path, xlsx_path: str | Path, opts: Options = Options(),
            progress: Optional[Callable[[int, int], None]] = None) -> int:
    """Trả về số bảng tìm được."""
    wb = Workbook()
    wb.remove(wb.active)
    count = 0
    merged = wb.create_sheet("TatCa") if opts.one_sheet else None
    if merged is not None:
        merged.freeze_panes = None
    merged_row = 1
    last_header = None

    with pdfplumber.open(str(pdf_path)) as pdf:
        pages = parse_pages(opts.pages, len(pdf.pages))
        for i, p_idx in enumerate(pages):
            page = pdf.pages[p_idx]
            tables = _extract(page, opts.mode)
            for t_idx, rows in enumerate(tables, 1):
                count += 1
                if merged is not None:
                    if rows[0] == last_header:
                        # Bảng nối tiếp từ trang trước: bỏ header lặp, ghi liền
                        merged_row = _write_table(merged, rows, merged_row - 1, skip_header=True)
                    else:
                        if merged_row > 1:
                            merged_row += 1
                        merged_row = _write_table(merged, rows, merged_row)
                        if merged.freeze_panes is None:
                            merged.freeze_panes = merged.cell(row=merged_row - len(rows) + 2, column=1)
                    merged_row += 1
                    last_header = rows[0]
                else:
                    ws = wb.create_sheet(f"Tr{p_idx + 1}_Bang{t_idx}"[:31])
                    _write_table(ws, rows, 1)
                    ws.freeze_panes = "A2"
                    _autosize(ws)
            if not tables and opts.include_text:
                text = page.extract_text() or ""
                if text.strip():
                    ws = wb.create_sheet(f"Tr{p_idx + 1}_Text")
                    for r, line in enumerate(text.splitlines(), 1):
                        ws.cell(row=r, column=1, value=line)
                    ws.column_dimensions["A"].width = 100
            if progress:
                progress(i + 1, len(pages))

    if merged is not None:
        _autosize(merged)
    if not wb.sheetnames:
        wb.create_sheet("TrongRong").cell(row=1, column=1,
            value="Không tìm thấy bảng. PDF có thể là bản scan (ảnh) — cần OCR.")
    wb.save(str(xlsx_path))
    return count


# ---------------------------------------------------------------------------
# PDF -> Word (giữ bảng)
# ---------------------------------------------------------------------------

def _page_blocks(page, mode: str) -> list[tuple[float, str, object]]:
    """Trả về các khối (top, 'table'|'text', nội dung) theo thứ tự trên → dưới."""
    if mode == "text":
        found = page.find_tables(TEXT)
    else:
        found = page.find_tables(LINES)
        if not found and mode == "auto":
            found = []  # auto: chỉ dùng bảng có kẻ viền, tránh biến đoạn văn thành bảng
    blocks: list[tuple[float, str, object]] = []
    rest = page
    for t in found:
        raw = t.extract()
        rows = [[_clean(c) for c in row] for row in raw]
        if rows:
            rows = [rows[0]] + [[_unwrap(c) for c in r] for r in rows[1:]]
            rows = [r for r in rows if any(r)]
        if rows:
            blocks.append((t.bbox[1], "table", rows))
            rest = rest.outside_bbox(t.bbox)
    for line in rest.extract_text_lines():
        text = line["text"].strip()
        if text:
            blocks.append((line["top"], "text", text))
    blocks.sort(key=lambda b: (b[0], b[1] == "text"))
    return blocks


def _repeat_header(row) -> None:
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    tr_pr = row._tr.get_or_add_trPr()
    el = OxmlElement("w:tblHeader")
    el.set(qn("w:val"), "true")
    tr_pr.append(el)


def convert_to_docx(pdf_path: str | Path, docx_path: str | Path, mode: str = "auto",
                    pages: Optional[str] = None) -> int:
    """PDF -> Word: bảng thành bảng Word thật, chữ ngoài bảng thành đoạn văn.
    Bảng nối tiếp qua nhiều trang (cùng header) được ghép thành một bảng.
    Trả về số bảng Word đã tạo."""
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.shared import Cm, Pt

    doc = Document()
    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(9)
    normal.paragraph_format.space_after = Pt(0)

    table_count = 0
    cur_table = None   # bảng Word đang ghi dở
    cur_header = None  # header của bảng đó

    with pdfplumber.open(str(pdf_path)) as pdf:
        idx = parse_pages(pages, len(pdf.pages))
        if idx:
            first = pdf.pages[idx[0]]
            sec = doc.sections[0]
            if first.width > first.height:
                sec.orientation = WD_ORIENT.LANDSCAPE
                sec.page_width, sec.page_height = sec.page_height, sec.page_width
            for side in ("left_margin", "right_margin", "top_margin", "bottom_margin"):
                setattr(sec, side, Cm(1.2))

        all_blocks = [_page_blocks(pdf.pages[p], mode) for p in idx]

    # Dòng chữ lặp ở hầu hết các trang (header/footer): không chèn giữa bảng, ghi 1 lần ở cuối
    freq: dict[str, int] = {}
    for blocks in all_blocks:
        for t in {c for _, k, c in blocks if k == "text"}:
            freq[t] = freq.get(t, 0) + 1
    repeated = {t for t, c in freq.items() if len(all_blocks) >= 3 and c >= len(all_blocks) * 0.5}
    footer: list[str] = []

    for n, blocks in enumerate(all_blocks):
        first_block = True
        for _, kind, content in blocks:
            if kind == "text" and content in repeated:
                if content not in footer:
                    footer.append(content)
                continue
            if kind == "table":
                rows = content
                if cur_table is not None and rows[0] == cur_header:
                    rows = rows[1:]  # bảng nối tiếp trang trước: bỏ header lặp
                elif cur_table is not None and first_block and len(rows[0]) == len(cur_table.columns):
                    pass  # bảng nối tiếp không lặp header: ghi tiếp
                else:
                    if n > 0 and first_block:
                        doc.add_page_break()
                    cur_table = doc.add_table(rows=0, cols=max(len(r) for r in rows))
                    cur_table.style = "Table Grid"
                    cur_header = rows[0]
                    table_count += 1
                    hdr = cur_table.add_row()
                    _repeat_header(hdr)
                    for c, text in zip(hdr.cells, rows[0]):
                        c.text = text
                        for r in c.paragraphs[0].runs:
                            r.font.bold = True
                    rows = rows[1:]
                width = len(cur_table.columns)
                for row in rows:
                    for c, text in zip(cur_table.add_row().cells, row[:width]):
                        c.text = text
            else:
                if n > 0 and first_block:
                    doc.add_page_break()
                cur_table = None
                cur_header = None
                doc.add_paragraph(content)
            first_block = False

    for t in footer:
        doc.add_paragraph(t)

    doc.save(str(docx_path))
    return table_count
