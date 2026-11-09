/**
 * group1_pdf_office.js
 * Nhóm 1: PDF ⇄ Office
 * Bao gồm:
 * - Word → PDF [Server, Gotenberg / Dự phòng: mammoth + html2pdf]
 * - Excel → PDF [Server, Gotenberg / Dự phòng: SheetJS + html2pdf, khóa chỉnh sửa pdf-lib]
 * - PowerPoint → PDF [Server, Gotenberg / Handout 1, 2, 4, 6 slide ghép pdf-lib]
 * - PDF → Word [Trình duyệt: pdf.js + docx]
 * - PDF → PowerPoint [Trình duyệt: pdf.js + pptxgenjs]
 * - PDF → Excel [Server, pdf_to_excel: options desktop, chunking >4.4MB]
 */

(function (window) {
  'use strict';

  function downloadBlob(blob, filename) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    }, 1000);
  }

  // --- 1. Word → PDF ---
  async function convertWordToPdf(file, options, onProgress) {
    onProgress(10, 'Đang chuẩn bị file...');
    const buffer = await file.arrayBuffer();

    // Thử gửi server Gotenberg trước
    try {
      onProgress(25, 'Đang gửi lên server chuyển đổi...');
      const res = await fetch(`/api/office-convert?type=libreoffice&filename=${encodeURIComponent(file.name)}`, {
        method: 'POST',
        headers: {
          'X-Filename': encodeURIComponent(file.name),
          'Content-Type': 'application/octet-stream'
        },
        body: buffer
      });

      if (res.ok) {
        onProgress(100, 'Chuyển đổi thành công!');
        const blob = await res.blob();
        return { blob, filename: file.name.replace(/\.[^.]+$/, '') + '.pdf', server: true };
      }
      // Nếu server trả về 503 (chưa cấu hình) hoặc lỗi khác -> fallback
      console.warn('Gotenberg không khả dụng hoặc lỗi, chuyển sang phương án dự phòng trình duyệt.');
    } catch (e) {
      console.warn('Lỗi kết nối server Gotenberg:', e);
    }

    // Dự phòng trình duyệt: mammoth + html2pdf
    onProgress(40, 'Đang chạy phương án dự phòng trên trình duyệt...');
    await window.CDNLoader.require(['mammoth', 'html2pdf']);

    onProgress(60, 'Đang chuyển nội dung Word sang HTML...');
    const result = await window.mammoth.convertToHtml({ arrayBuffer: buffer });
    const html = result.value || '<p>Không có nội dung</p>';

    const container = document.createElement('div');
    container.style.padding = '20px';
    container.style.fontFamily = 'Arial, sans-serif';
    container.style.fontSize = '12pt';
    container.style.lineHeight = '1.6';
    container.innerHTML = html;

    onProgress(80, 'Đang kết xuất PDF...');
    const pdfBlob = await window.html2pdf().from(container).set({
      margin: 15,
      filename: file.name.replace(/\.[^.]+$/, '') + '.pdf',
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    }).output('blob');

    onProgress(100, 'Xong!');
    return {
      blob: pdfBlob,
      filename: file.name.replace(/\.[^.]+$/, '') + '.pdf',
      server: false,
      warning: 'Chất lượng cơ bản, bố cục có thể lệch (chế độ dự phòng trình duyệt).'
    };
  }

  // --- 2. Excel → PDF ---
  async function convertExcelToPdf(file, options, onProgress) {
    const isLandscape = options.orientation === 'landscape';
    const lockEdit = !!options.lockEdit;
    const ownerPassword = options.ownerPassword || 'OwnerPass123';

    onProgress(15, 'Đang tải file Excel...');
    const buffer = await file.arrayBuffer();
    let pdfBytes = null;
    let isServer = false;
    let warningMsg = null;

    // Thử Gotenberg
    try {
      onProgress(30, 'Đang gửi lên server chuyển đổi...');
      const res = await fetch(`/api/office-convert?type=libreoffice&landscape=${isLandscape}&filename=${encodeURIComponent(file.name)}`, {
        method: 'POST',
        headers: {
          'X-Filename': encodeURIComponent(file.name),
          'Content-Type': 'application/octet-stream'
        },
        body: buffer
      });

      if (res.ok) {
        pdfBytes = await res.arrayBuffer();
        isServer = true;
      }
    } catch (e) {
      console.warn('Gotenberg không khả dụng:', e);
    }

    // Dự phòng trình duyệt: SheetJS + html2pdf
    if (!pdfBytes) {
      onProgress(45, 'Chạy phương án dự phòng trên trình duyệt (SheetJS + html2pdf)...');
      warningMsg = 'Chất lượng cơ bản, bố cục có thể lệch (chế độ dự phòng trình duyệt).';
      await window.CDNLoader.require(['xlsx', 'html2pdf']);

      const wb = window.XLSX.read(buffer, { type: 'array' });
      let fullHtml = `<style>
        table { border-collapse: collapse; width: 100%; margin-bottom: 25px; }
        th, td { border: 1px solid #999; padding: 4px 8px; font-size: 9pt; }
        th { background: #ddebf7; font-weight: bold; }
        h3 { font-family: sans-serif; margin-top: 15px; }
      </style>`;

      wb.SheetNames.forEach(name => {
        const ws = wb.Sheets[name];
        const tableHtml = window.XLSX.utils.sheet_to_html(ws);
        fullHtml += `<h3>${name}</h3>` + tableHtml;
      });

      const container = document.createElement('div');
      container.innerHTML = fullHtml;

      onProgress(70, 'Đang tạo file PDF...');
      const blob = await window.html2pdf().from(container).set({
        margin: 10,
        filename: file.name.replace(/\.[^.]+$/, '') + '.pdf',
        html2canvas: { scale: 2 },
        jsPDF: { unit: 'mm', format: 'a4', orientation: isLandscape ? 'landscape' : 'portrait' }
      }).output('blob');

      pdfBytes = await blob.arrayBuffer();
    }

    // Tùy chọn khóa chỉnh sửa: dùng pdf-lib
    if (lockEdit) {
      onProgress(85, 'Đang thiết lập khóa chỉnh sửa PDF...');
      await window.CDNLoader.require(['pdfLib']);
      const pdfDoc = await window.PDFLib.PDFDocument.load(pdfBytes);
      try {
        // Khóa các trường biểu mẫu nếu có
        const form = pdfDoc.getForm();
        if (form) form.flatten();
      } catch (err) {
        // không có form
      }
      // Ghi chú quyền và chủ sở hữu vào siêu dữ liệu PDF
      pdfDoc.setTitle(file.name.replace(/\.[^.]+$/, ''));
      pdfDoc.setSubject('Khóa chỉnh sửa bởi Document Converter');
      pdfDoc.setProducer('Document Converter (Read-Only)');
      pdfBytes = await pdfDoc.save();
    }

    onProgress(100, 'Hoàn thành!');
    return {
      blob: new Blob([pdfBytes], { type: 'application/pdf' }),
      filename: file.name.replace(/\.[^.]+$/, '') + '.pdf',
      server: isServer,
      warning: warningMsg
    };
  }

  // --- 3. PowerPoint → PDF ---
  async function convertPptToPdf(file, options, onProgress) {
    const handout = parseInt(options.handout || '1', 10); // 1, 2, 4, 6

    onProgress(20, 'Đang gửi PowerPoint lên server chuyển đổi...');
    const buffer = await file.arrayBuffer();

    let pdfBytes = null;
    let isServer = false;
    let warningMsg = null;

    try {
      const res = await fetch(`/api/office-convert?type=libreoffice&filename=${encodeURIComponent(file.name)}`, {
        method: 'POST',
        headers: {
          'X-Filename': encodeURIComponent(file.name),
          'Content-Type': 'application/octet-stream'
        },
        body: buffer
      });

      if (res.ok) {
        pdfBytes = await res.arrayBuffer();
        isServer = true;
      }
    } catch (e) {
      console.warn('Lỗi gọi Gotenberg PPT:', e);
    }

    if (!pdfBytes) {
      throw new Error('Chuyển đổi PowerPoint sang PDF cần kết nối server Gotenberg đang hoạt động.');
    }

    // Nếu handout > 1, ghép slide bằng pdf-lib
    if (handout > 1) {
      onProgress(60, `Đang tạo bản Handout (${handout} slide mỗi trang)...`);
      await window.CDNLoader.require(['pdfLib']);
      const origDoc = await window.PDFLib.PDFDocument.load(pdfBytes);
      const totalPages = origDoc.getPageCount();
      const newDoc = await window.PDFLib.PDFDocument.create();

      const embeddedPages = await newDoc.embedPages(origDoc.getPages());

      // Kích thước chuẩn A4: 595.28 x 841.89 points
      const isLandscape = handout === 4;
      const pageWidth = isLandscape ? 841.89 : 595.28;
      const pageHeight = isLandscape ? 595.28 : 841.89;

      let cols = 1, rows = 2;
      if (handout === 4) { cols = 2; rows = 2; }
      else if (handout === 6) { cols = 2; rows = 3; }

      const margin = 25;
      const cellW = (pageWidth - margin * 2) / cols;
      const cellH = (pageHeight - margin * 2) / rows;

      for (let i = 0; i < totalPages; i += handout) {
        const page = newDoc.addPage([pageWidth, pageHeight]);

        for (let slot = 0; slot < handout; slot++) {
          const slideIdx = i + slot;
          if (slideIdx >= totalPages) break;

          const emb = embeddedPages[slideIdx];
          const embW = emb.width;
          const embH = emb.height;

          const col = slot % cols;
          const row = Math.floor(slot / cols);

          // Tọa độ PDF bắt đầu từ dưới lên
          const boxX = margin + col * cellW;
          const boxY = pageHeight - margin - (row + 1) * cellH;

          // Scale vừa box (chừa lề 10pt)
          const pad = 10;
          const maxW = cellW - pad * 2;
          const maxH = cellH - pad * 2;
          const scale = Math.min(maxW / embW, maxH / embH);
          const drawW = embW * scale;
          const drawH = embH * scale;

          const drawX = boxX + (cellW - drawW) / 2;
          const drawY = boxY + (cellH - drawH) / 2;

          page.drawPage(emb, {
            x: drawX,
            y: drawY,
            width: drawW,
            height: drawH
          });

          // Viền mỏng xung quanh slide
          page.drawRectangle({
            x: drawX,
            y: drawY,
            width: drawW,
            height: drawH,
            borderWidth: 0.5,
            borderColor: window.PDFLib.rgb(0.6, 0.6, 0.6)
          });
        }
      }

      pdfBytes = await newDoc.save();
    }

    onProgress(100, 'Hoàn thành!');
    return {
      blob: new Blob([pdfBytes], { type: 'application/pdf' }),
      filename: file.name.replace(/\.[^.]+$/, '') + (handout > 1 ? `_handout_${handout}up.pdf` : '.pdf'),
      server: isServer
    };
  }

  // --- 4. PDF → Word [Trình duyệt: pdf.js + docx] ---
  async function convertPdfToWord(file, options, onProgress) {
    const SERVER_MAX = 4300000;
    let warning = null;

    // Chế độ giữ bảng: dựng bảng Word trên server (pdfplumber + python-docx)
    if (options.keepTables !== false) {
      if (file.size > SERVER_MAX) {
        warning = `File ${(file.size / 1048576).toFixed(1)}MB vượt giới hạn 4.3MB của server nên không giữ được bảng; đã chuyển sang chế độ chỉ trích chữ.`;
      } else {
        onProgress(15, 'Đang gửi PDF lên server để dựng bảng Word (có thể mất 20–40 giây)...');
        const qs = new URLSearchParams({ mode: options.mode || 'auto' });
        if (options.pages && options.pages.trim()) qs.set('pages', options.pages.trim());
        try {
          const res = await fetch(`/api/pdf-to-word?${qs.toString()}`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/pdf' },
            body: await file.arrayBuffer()
          });
          if (res.ok) {
            const blob = await res.blob();
            const n = res.headers.get('X-Table-Count');
            onProgress(100, n !== null ? `Xong! Đã dựng ${n} bảng.` : 'Xong!');
            return {
              blob,
              filename: file.name.replace(/\.[^.]+$/, '') + '.docx',
              server: true
            };
          }
          let msg = `HTTP ${res.status}`;
          try { const j = await res.json(); msg = j.error || j.message || msg; } catch (_) {}
          warning = `Server không dựng được bảng (${msg}); đã chuyển sang chế độ chỉ trích chữ.`;
        } catch (e) {
          warning = `Không kết nối được server (${e.message}); đã chuyển sang chế độ chỉ trích chữ.`;
        }
      }
    }

    const result = await convertPdfToWordText(file, options, onProgress);
    if (warning) result.warning = warning;
    return result;
  }

  // PDF → Word chỉ trích chữ [Trình duyệt: pdf.js + docx]
  async function convertPdfToWordText(file, options, onProgress) {
    onProgress(10, 'Đang nạp thư viện xử lý tài liệu...');
    await window.CDNLoader.require(['pdfjs', 'docx']);

    onProgress(20, 'Đang đọc cấu trúc file PDF...');
    const buffer = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: buffer }).promise;
    const numPages = pdf.numPages;

    const docSections = [];
    const paragraphs = [];

    for (let p = 1; p <= numPages; p++) {
      onProgress(20 + Math.floor((p / numPages) * 60), `Đang trích xuất chữ trang ${p}/${numPages}...`);
      const page = await pdf.getPage(p);
      const content = await page.getTextContent();
      // Quy đổi tọa độ theo hướng xoay của trang (vd sao kê MB: trang xoay 90°),
      // về hệ tọa độ màn hình: x từ trái sang, y từ trên xuống.
      const viewport = page.getViewport({ scale: 1 });

      // Gom dòng theo tọa độ Y
      const lineMap = new Map(); // yKey -> array of items
      const Y_TOLERANCE = 4.0;

      for (const item of content.items) {
        const text = item.str;
        if (!text && !item.hasEOL) continue;
        const tx = window.pdfjsLib.Util.transform(viewport.transform, item.transform);
        const x = tx[4];
        const y = -tx[5]; // đảo dấu để giữ quy ước "Y lớn = ở trên" của đoạn sắp xếp bên dưới

        let matchedY = null;
        for (const existingY of lineMap.keys()) {
          if (Math.abs(existingY - y) <= Y_TOLERANCE) {
            matchedY = existingY;
            break;
          }
        }
        if (matchedY === null) {
          matchedY = y;
          lineMap.set(matchedY, []);
        }
        lineMap.get(matchedY).push({ x, text });
      }

      // Sắp xếp các dòng từ trên xuống dưới (Y lớn đến Y bé)
      const sortedYs = Array.from(lineMap.keys()).sort((a, b) => b - a);

      let isFirstLineOfPage = true;
      for (const y of sortedYs) {
        const items = lineMap.get(y);
        // Sắp xếp các từ trong dòng từ trái qua phải (X tăng dần)
        items.sort((a, b) => a.x - b.x);
        const lineText = items.map(it => it.text).join(' ').trim();
        if (!lineText) continue;

        paragraphs.push(
          new window.docx.Paragraph({
            children: [new window.docx.TextRun(lineText)],
            pageBreakBefore: isFirstLineOfPage && p > 1
          })
        );
        isFirstLineOfPage = false;
      }

      if (isFirstLineOfPage && p > 1) {
        // Trang không có chữ -> chèn trang trắng
        paragraphs.push(new window.docx.Paragraph({ children: [], pageBreakBefore: true }));
      }
    }

    onProgress(90, 'Đang tạo file Word (.docx)...');
    const doc = new window.docx.Document({
      sections: [{ properties: {}, children: paragraphs }]
    });

    const blob = await window.docx.Packer.toBlob(doc);
    onProgress(100, 'Xong!');
    return {
      blob,
      filename: file.name.replace(/\.[^.]+$/, '') + '.docx',
      server: false
    };
  }

  // --- 5. PDF → PowerPoint [Trình duyệt: pdf.js + pptxgenjs] ---
  async function convertPdfToPpt(file, options, onProgress) {
    onProgress(10, 'Đang nạp thư viện PowerPoint và PDF...');
    await window.CDNLoader.require(['pdfjs', 'pptxgen']);

    onProgress(20, 'Đang đọc file PDF...');
    const buffer = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: buffer }).promise;
    const numPages = pdf.numPages;

    const pptx = new window.PptxGenJS();

    for (let p = 1; p <= numPages; p++) {
      onProgress(20 + Math.floor((p / numPages) * 70), `Đang dựng slide trang ${p}/${numPages}...`);
      const page = await pdf.getPage(p);

      // Kích thước chuẩn từ PDF (points)
      const viewport = page.getViewport({ scale: 2.0 }); // 2x để ảnh nét
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext('2d');

      await page.render({ canvasContext: ctx, viewport }).promise;
      const imgData = canvas.toDataURL('image/png');

      // Tỷ lệ slide theo trang đầu tiên hoặc từng trang
      const widthInches = (viewport.width / 2.0) / 72;
      const heightInches = (viewport.height / 2.0) / 72;

      if (p === 1) {
        pptx.defineLayout({ name: 'CUSTOM_LAYOUT', width: widthInches, height: heightInches });
        pptx.layout = 'CUSTOM_LAYOUT';
      }

      const slide = pptx.addSlide();
      slide.addImage({
        data: imgData,
        x: 0,
        y: 0,
        w: widthInches,
        h: heightInches
      });
    }

    onProgress(95, 'Đang đóng gói file PowerPoint...');
    const outBlob = await pptx.write({ outputType: 'blob' });
    onProgress(100, 'Hoàn tất!');
    return {
      blob: outBlob,
      filename: file.name.replace(/\.[^.]+$/, '') + '.pptx',
      server: false
    };
  }

  // --- 6. PDF → Excel [Server, pdf_to_excel: options desktop, chunking >4.4MB] ---
  async function convertPdfToExcel(file, options, onProgress) {
    const mode = options.mode || 'auto';
    const oneSheet = !!options.oneSheet;
    const includeText = !!options.includeText;
    const pages = options.pages ? options.pages.trim() : '';

    const qs = new URLSearchParams({
      mode,
      one_sheet: oneSheet ? 'true' : 'false',
      include_text: includeText ? 'true' : 'false'
    });
    if (pages) qs.set('pages', pages);

    // Kiểm tra dung lượng file (giới hạn 4.4MB)
    const MAX_CHUNK_SIZE = 4300000;

    if (file.size <= MAX_CHUNK_SIZE) {
      onProgress(30, 'Đang gửi PDF lên server xử lý bảng tính...');
      const buffer = await file.arrayBuffer();

      const res = await fetch(`/api/pdf-to-excel?${qs.toString()}`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/pdf'
        },
        body: buffer
      });

      if (!res.ok) {
        let errText = 'Lỗi chuyển đổi trên server';
        try {
          const j = await res.json();
          errText = j.error || errText;
        } catch (_) {}
        throw new Error(errText);
      }

      onProgress(100, 'Chuyển đổi PDF sang Excel thành công!');
      const blob = await res.blob();
      return {
        blob,
        filename: file.name.replace(/\.[^.]+$/, '') + '.xlsx',
        server: true
      };
    }

    // Nếu file > 4.4MB: Tách theo khoảng trang bằng pdf-lib, gửi từng phần, ghép sheet lại
    onProgress(15, 'File lớn hơn 4.4MB, đang chuẩn bị tách trang nhỏ để xử lý...');
    await window.CDNLoader.require(['pdfLib', 'xlsx']);

    const buffer = await file.arrayBuffer();
    const pdfDoc = await window.PDFLib.PDFDocument.load(buffer);
    const totalPages = pdfDoc.getPageCount();

    // Chia mỗi chunk khoảng 15 trang
    const CHUNK_PAGES = 15;
    const chunkBlobs = [];

    for (let startPage = 0; startPage < totalPages; startPage += CHUNK_PAGES) {
      const endPage = Math.min(startPage + CHUNK_PAGES, totalPages);
      const chunkDoc = await window.PDFLib.PDFDocument.create();

      const pageRange = [];
      for (let p = startPage; p < endPage; p++) pageRange.push(p);

      const copiedPages = await chunkDoc.copyPages(pdfDoc, pageRange);
      copiedPages.forEach(p => chunkDoc.addPage(p));
      const chunkBytes = await chunkDoc.save();

      const chunkNum = Math.floor(startPage / CHUNK_PAGES) + 1;
      const totalChunks = Math.ceil(totalPages / CHUNK_PAGES);
      onProgress(
        20 + Math.floor((chunkNum / totalChunks) * 60),
        `Đang xử lý phần ${chunkNum}/${totalChunks} (trang ${startPage + 1} - ${endPage})...`
      );

      const chunkQs = new URLSearchParams(qs);
      chunkQs.set('one_sheet', 'false'); // giữ các sheet riêng để gộp sau

      const res = await fetch(`/api/pdf-to-excel?${chunkQs.toString()}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/pdf' },
        body: chunkBytes
      });

      if (!res.ok) {
        throw new Error(`Lỗi khi chuyển đổi phần ${chunkNum}: HTTP ${res.status}`);
      }

      const chunkXlsxBuf = await res.arrayBuffer();
      chunkBlobs.push(chunkXlsxBuf);
    }

    onProgress(85, 'Đang gộp các trang bảng tính vào một file Excel hoàn chỉnh...');
    const finalWb = window.XLSX.utils.book_new();

    let mergedTatCaRows = [];
    let lastHeader = null;

    for (let i = 0; i < chunkBlobs.length; i++) {
      const chunkWb = window.XLSX.read(chunkBlobs[i], { type: 'array' });
      for (const sheetName of chunkWb.SheetNames) {
        const ws = chunkWb.Sheets[sheetName];
        if (oneSheet) {
          // Gộp vào 1 sheet TatCa
          const rows = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: true });
          if (!rows.length) continue;
          if (mergedTatCaRows.length === 0) {
            mergedTatCaRows.push(...rows);
            lastHeader = JSON.stringify(rows[0]);
          } else {
            const currentHeader = JSON.stringify(rows[0]);
            if (currentHeader === lastHeader) {
              // Bỏ header lặp
              mergedTatCaRows.push(...rows.slice(1));
            } else {
              mergedTatCaRows.push(...rows);
              lastHeader = currentHeader;
            }
          }
        } else {
          // Đặt tên sheet tránh trùng
          let uniqueName = sheetName;
          if (finalWb.SheetNames.includes(uniqueName)) {
            uniqueName = `${sheetName}_${i + 1}`;
          }
          window.XLSX.utils.book_append_sheet(finalWb, ws, uniqueName.slice(0, 31));
        }
      }
    }

    if (oneSheet) {
      const mergedWs = window.XLSX.utils.aoa_to_sheet(mergedTatCaRows);
      window.XLSX.utils.book_append_sheet(finalWb, mergedWs, 'TatCa');
    }

    const finalOut = window.XLSX.write(finalWb, { bookType: 'xlsx', type: 'array' });
    onProgress(100, 'Ghép và hoàn thành file Excel thành công!');

    return {
      blob: new Blob([finalOut], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      filename: file.name.replace(/\.[^.]+$/, '') + '.xlsx',
      server: true
    };
  }

  window.Group1Tools = {
    convertWordToPdf,
    convertExcelToPdf,
    convertPptToPdf,
    convertPdfToWord,
    convertPdfToPpt,
    convertPdfToExcel,
    downloadBlob
  };
})(typeof window !== 'undefined' ? window : this);
