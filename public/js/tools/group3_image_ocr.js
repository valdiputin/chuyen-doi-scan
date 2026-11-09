/**
 * group3_image_ocr.js
 * Nhóm 3: Ảnh & OCR [Trình duyệt]
 * - Ảnh/Scan → Word/Excel/PDF/Text có OCR (tesseract.js vie+eng, tiền xử lý tương phản & xoay)
 * - PDF → Ảnh (pdf.js: PNG/JPG, DPI 96/150/300, nén ZIP)
 * - Ảnh → PDF (pdf-lib: nhiều ảnh, sắp xếp, xoay, khổ A4 / gốc, nén chất lượng)
 */

(function (window) {
  'use strict';

  // --- Tiền xử lý ảnh (tăng tương phản và grayscale) trên canvas ---
  function preprocessImageCanvas(canvas, autoRotateDeg = 0) {
    const width = canvas.width;
    const height = canvas.height;

    let targetCanvas = canvas;
    if (autoRotateDeg !== 0) {
      targetCanvas = document.createElement('canvas');
      if (autoRotateDeg === 90 || autoRotateDeg === 270) {
        targetCanvas.width = height;
        targetCanvas.height = width;
      } else {
        targetCanvas.width = width;
        targetCanvas.height = height;
      }
      const tCtx = targetCanvas.getContext('2d');
      tCtx.translate(targetCanvas.width / 2, targetCanvas.height / 2);
      tCtx.rotate((autoRotateDeg * Math.PI) / 180);
      tCtx.drawImage(canvas, -width / 2, -height / 2);
    }

    const ctx = targetCanvas.getContext('2d');
    const imgData = ctx.getImageData(0, 0, targetCanvas.width, targetCanvas.height);
    const d = imgData.data;

    // Tăng tương phản và đưa về Grayscale
    const contrastFactor = 1.35; // tăng tương phản 35%
    for (let i = 0; i < d.length; i += 4) {
      // Grayscale theo trọng số cảm nhận mắt người
      const gray = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
      // Tăng tương phản: f(x) = factor * (x - 128) + 128
      let c = contrastFactor * (gray - 128) + 128;
      if (c < 0) c = 0;
      if (c > 255) c = 255;
      d[i] = c;
      d[i + 1] = c;
      d[i + 2] = c;
    }
    ctx.putImageData(imgData, 0, 0);
    return targetCanvas;
  }

  // Chuyển file ảnh hoặc trang PDF scan thành canvas
  async function loadImagesAsCanvases(file, onProgress) {
    const canvases = [];
    const isPdf = file.name.toLowerCase().endsWith('.pdf');

    if (isPdf) {
      await window.CDNLoader.require(['pdfjs']);
      const buffer = await file.arrayBuffer();
      const pdf = await window.pdfjsLib.getDocument({ data: buffer }).promise;
      const total = pdf.numPages;

      for (let p = 1; p <= total; p++) {
        onProgress(Math.floor((p / total) * 30), `Đang dựng trang ${p}/${total} từ PDF...`);
        const page = await pdf.getPage(p);
        const viewport = page.getViewport({ scale: 2.0 }); // 2x để OCR nét hơn
        const canvas = document.createElement('canvas');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        const ctx = canvas.getContext('2d');
        await page.render({ canvasContext: ctx, viewport }).promise;
        canvases.push({ canvas, originalWidth: viewport.width, originalHeight: viewport.height, pageNum: p });
      }
    } else {
      // File ảnh thông thường: JPG/PNG/WEBP
      onProgress(15, 'Đang đọc file ảnh...');
      const img = new Image();
      const url = URL.createObjectURL(file);
      await new Promise((resolve, reject) => {
        img.onload = resolve;
        img.onerror = reject;
        img.src = url;
      });
      URL.revokeObjectURL(url);

      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth || img.width;
      canvas.height = img.naturalHeight || img.height;
      const ctx = canvas.getContext('2d');
      ctx.drawImage(img, 0, 0);
      canvases.push({ canvas, originalWidth: canvas.width, originalHeight: canvas.height, pageNum: 1 });
    }

    return canvases;
  }

  // --- 1. Ảnh/Scan → Word/Excel/PDF/Text có OCR ---
  async function performOcr(file, options, onProgress) {
    const targetFormat = options.targetFormat || 'docx'; // docx | xlsx | pdf | txt
    const rotateDeg = parseInt(options.rotate || '0', 10);

    onProgress(5, 'Đang chuẩn bị thư viện OCR...');
    await window.CDNLoader.require(['tesseract']);
    if (targetFormat === 'docx') await window.CDNLoader.require(['docx']);
    if (targetFormat === 'xlsx') await window.CDNLoader.require(['xlsx']);
    if (targetFormat === 'pdf') await window.CDNLoader.require(['pdfLib']);

    onProgress(10, 'Đang phân tích hình ảnh đầu vào...');
    const pageCanvases = await loadImagesAsCanvases(file, onProgress);

    onProgress(30, 'Đang khởi tạo bộ nhận diện tiếng Việt & tiếng Anh...');
    // Tạo Tesseract worker với ngôn ngữ vie + eng
    const worker = await window.Tesseract.createWorker(['vie', 'eng'], 1, {
      logger: m => {
        if (m.status === 'recognizing text') {
          const pct = Math.round((m.progress || 0) * 100);
          onProgress(35 + Math.floor(pct * 0.45), `Nhận diện chữ: ${pct}%`);
        }
      }
    });

    const ocrResults = [];
    for (let i = 0; i < pageCanvases.length; i++) {
      const item = pageCanvases[i];
      onProgress(35, `Đang xử lý tương phản & OCR trang ${i + 1}/${pageCanvases.length}...`);
      const preprocessed = preprocessImageCanvas(item.canvas, rotateDeg);
      const res = await worker.recognize(preprocessed);
      ocrResults.push({
        rawCanvas: item.canvas,
        pageData: res.data,
        width: preprocessed.width,
        height: preprocessed.height
      });
    }

    await worker.terminate();
    onProgress(85, 'Đang định dạng kết quả xuất...');

    const baseName = file.name.replace(/\.[^.]+$/, '');

    // 1. Xuất Word (.docx) - giữ đoạn văn
    if (targetFormat === 'docx') {
      const paragraphs = [];
      ocrResults.forEach((res, pIdx) => {
        if (pIdx > 0) {
          paragraphs.push(new window.docx.Paragraph({ children: [], pageBreakBefore: true }));
        }
        const paras = res.pageData.paragraphs || [];
        if (paras.length > 0) {
          paras.forEach(p => {
            const text = p.text.trim();
            if (text) {
              paragraphs.push(new window.docx.Paragraph({
                children: [new window.docx.TextRun(text)],
                spacing: { after: 120 }
              }));
            }
          });
        } else {
          // Fallback sang text dòng
          const lines = (res.pageData.text || '').split('\n');
          lines.forEach(l => {
            if (l.trim()) {
              paragraphs.push(new window.docx.Paragraph({
                children: [new window.docx.TextRun(l.trim())]
              }));
            }
          });
        }
      });

      const doc = new window.docx.Document({
        sections: [{ properties: {}, children: paragraphs }]
      });
      const blob = await window.docx.Packer.toBlob(doc);
      onProgress(100, 'Xong!');
      return { blob, filename: `${baseName}_ocr.docx`, server: false };
    }

    // 2. Xuất Excel (.xlsx) - gom từ thành hàng & cột theo tọa độ bbox
    if (targetFormat === 'xlsx') {
      const wb = window.XLSX.utils.book_new();

      ocrResults.forEach((res, pIdx) => {
        const words = res.pageData.words || [];
        // Lọc từ rỗng
        const validWords = words.filter(w => w.text && w.text.trim());

        // Nhóm từ theo dòng dựa vào tọa độ y0
        const ROW_THRESHOLD = 12; // khoảng cách pixel để xem là cùng hàng
        const rows = []; // mảng các nhóm từ

        // Sắp xếp từ theo y0 từ trên xuống
        validWords.sort((a, b) => a.bbox.y0 - b.bbox.y0);

        validWords.forEach(w => {
          let placed = false;
          for (const row of rows) {
            const avgY = row.reduce((sum, item) => sum + item.bbox.y0, 0) / row.length;
            if (Math.abs(w.bbox.y0 - avgY) <= ROW_THRESHOLD) {
              row.push(w);
              placed = true;
              break;
            }
          }
          if (!placed) {
            rows.push([w]);
          }
        });

        // Với mỗi hàng, sắp xếp từ trái sang phải theo x0 và gom vào các cột
        const tableRows = [];
        const GAP_THRESHOLD = 35; // nếu khoảng cách từ > 35px thì tách cột

        rows.forEach(rowWords => {
          rowWords.sort((a, b) => a.bbox.x0 - b.bbox.x0);
          const cols = [];
          let currentCell = '';
          let lastX1 = null;

          rowWords.forEach(w => {
            if (lastX1 === null) {
              currentCell = w.text;
            } else if (w.bbox.x0 - lastX1 > GAP_THRESHOLD) {
              cols.push(currentCell.trim());
              currentCell = w.text;
            } else {
              currentCell += ' ' + w.text;
            }
            lastX1 = w.bbox.x1;
          });
          if (currentCell.trim()) {
            cols.push(currentCell.trim());
          }
          if (cols.length) tableRows.push(cols);
        });

        // Áp dụng SmartValue cho bảng
        const processedRows = window.SmartValue.processTableRows(tableRows);
        const ws = window.XLSX.utils.aoa_to_sheet(processedRows);
        window.XLSX.utils.book_append_sheet(wb, ws, `Trang_${pIdx + 1}`);
      });

      const out = window.XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      onProgress(100, 'Hoàn thành!');
      return {
        blob: new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        filename: `${baseName}_ocr.xlsx`,
        server: false
      };
    }

    // 3. Xuất Searchable PDF (Ảnh gốc + lớp chữ ẩn bằng pdf-lib)
    if (targetFormat === 'pdf') {
      const pdfDoc = await window.PDFLib.PDFDocument.create();

      for (let i = 0; i < ocrResults.length; i++) {
        const item = ocrResults[i];
        // Đưa canvas gốc thành JPEG để nhúng vào PDF
        const imgDataUrl = item.rawCanvas.toDataURL('image/jpeg', 0.9);
        const embeddedImg = await pdfDoc.embedJpg(imgDataUrl);

        const imgW = embeddedImg.width;
        const imgH = embeddedImg.height;

        // Tạo trang PDF đúng kích thước ảnh
        const page = pdfDoc.addPage([imgW, imgH]);

        // Vẽ ảnh nền
        page.drawImage(embeddedImg, {
          x: 0,
          y: 0,
          width: imgW,
          height: imgH
        });

        // Vẽ lớp chữ ẩn có thể bôi đen/tìm kiếm
        const words = item.pageData.words || [];
        const scaleX = imgW / item.width;
        const scaleY = imgH / item.height;

        for (const w of words) {
          const text = w.text.trim();
          if (!text) continue;
          const bbox = w.bbox;
          // Tọa độ PDF: (0,0) ở góc dưới bên trái
          const pdfX = bbox.x0 * scaleX;
          const pdfY = imgH - (bbox.y1 * scaleY);
          const fontSize = Math.max(6, Math.min(24, (bbox.y1 - bbox.y0) * scaleY));

          try {
            page.drawText(text, {
              x: pdfX,
              y: pdfY,
              size: fontSize,
              opacity: 0.0 // Chữ ẩn 100% trong suốt, vẫn bôi đen và tìm kiếm được!
            });
          } catch (_) {
            // bỏ qua ký tự font đặc biệt nếu pdf-lib không hỗ trợ trực tiếp
          }
        }
      }

      const pdfBytes = await pdfDoc.save();
      onProgress(100, 'Tạo PDF tìm kiếm được thành công!');
      return {
        blob: new Blob([pdfBytes], { type: 'application/pdf' }),
        filename: `${baseName}_searchable.pdf`,
        server: false
      };
    }

    // 4. Xuất Text (.txt)
    const fullText = ocrResults.map((r, idx) => `--- Trang ${idx + 1} ---\n` + r.pageData.text).join('\n\n');
    onProgress(100, 'Xong!');
    return {
      blob: new Blob([fullText], { type: 'text/plain;charset=utf-8' }),
      filename: `${baseName}_ocr.txt`,
      server: false
    };
  }

  // --- 2. PDF → Ảnh (pdf.js) ---
  async function convertPdfToImages(file, options, onProgress) {
    const format = options.imgFormat || 'png'; // png | jpeg
    const dpi = parseInt(options.dpi || '150', 10); // 96 | 150 | 300
    const pageSpec = options.pages ? options.pages.trim() : '';

    onProgress(10, 'Đang nạp thư viện đọc PDF...');
    await window.CDNLoader.require(['pdfjs', 'jszip']);

    const buffer = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: buffer }).promise;
    const totalPages = pdf.numPages;

    // Tính tỷ lệ scale theo DPI (mặc định PDF 72 DPI)
    const scale = dpi / 72.0;

    // Phân tích trang cần xuất
    let targetPages = [];
    if (!pageSpec || pageSpec.toLowerCase() === 'all') {
      for (let i = 1; i <= totalPages; i++) targetPages.push(i);
    } else {
      pageSpec.split(',').forEach(part => {
        part = part.trim();
        if (part.includes('-')) {
          const [a, b] = part.split('-').map(n => parseInt(n, 10));
          if (!isNaN(a) && !isNaN(b)) {
            for (let i = Math.max(1, a); i <= Math.min(b, totalPages); i++) targetPages.push(i);
          }
        } else {
          const n = parseInt(part, 10);
          if (!isNaN(n) && n >= 1 && n <= totalPages) targetPages.push(n);
        }
      });
      targetPages = Array.from(new Set(targetPages)).sort((a, b) => a - b);
    }

    if (!targetPages.length) {
      throw new Error('Không có trang nào hợp lệ để xuất ảnh.');
    }

    const baseName = file.name.replace(/\.[^.]+$/, '');
    const renderedImages = [];

    for (let i = 0; i < targetPages.length; i++) {
      const pNum = targetPages[i];
      onProgress(
        15 + Math.floor((i / targetPages.length) * 75),
        `Đang kết xuất ảnh trang ${pNum}/${totalPages} (${dpi} DPI)...`
      );

      const page = await pdf.getPage(pNum);
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement('canvas');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      const ctx = canvas.getContext('2d');

      await page.render({ canvasContext: ctx, viewport }).promise;

      const mimeType = format === 'jpeg' ? 'image/jpeg' : 'image/png';
      const ext = format === 'jpeg' ? 'jpg' : 'png';
      const blob = await new Promise(res => canvas.toBlob(res, mimeType, 0.92));
      renderedImages.push({
        name: `${baseName}_trang_${pNum}.${ext}`,
        blob
      });
    }

    // Nếu chỉ có 1 ảnh: tải trực tiếp file ảnh
    if (renderedImages.length === 1) {
      onProgress(100, 'Xong!');
      return {
        blob: renderedImages[0].blob,
        filename: renderedImages[0].name,
        server: false
      };
    }

    // Nhiều trang: nén thành file ZIP bằng JSZip
    onProgress(92, 'Đang nén các ảnh vào file ZIP...');
    const zip = new window.JSZip();
    renderedImages.forEach(img => {
      zip.file(img.name, img.blob);
    });

    const zipBlob = await zip.generateAsync({ type: 'blob' });
    onProgress(100, 'Nén file ZIP hoàn tất!');
    return {
      blob: zipBlob,
      filename: `${baseName}_images.zip`,
      server: false
    };
  }

  // --- 3. Ảnh → PDF (pdf-lib) ---
  async function convertImagesToPdf(files, options, onProgress) {
    if (!files || !files.length) {
      throw new Error('Vui lòng chọn ít nhất một file ảnh.');
    }

    const pageSizeMode = options.pageSize || 'a4'; // a4 | fit
    const quality = options.quality || 'high'; // high (0.92) | medium (0.75) | low (0.55)

    let qValue = 0.92;
    if (quality === 'medium') qValue = 0.75;
    if (quality === 'low') qValue = 0.55;

    onProgress(10, 'Đang nạp thư viện tạo PDF...');
    await window.CDNLoader.require(['pdfLib']);

    const pdfDoc = await window.PDFLib.PDFDocument.create();

    // Chuẩn A4 tính theo points (72 DPI): 595.28 x 841.89
    const A4_W = 595.28;
    const A4_H = 841.89;

    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      onProgress(15 + Math.floor((i / files.length) * 75), `Đang xử lý ảnh ${i + 1}/${files.length}...`);

      const img = new Image();
      const url = URL.createObjectURL(f);
      await new Promise((res, rej) => {
        img.onload = res;
        img.onerror = rej;
        img.src = url;
      });
      URL.revokeObjectURL(url);

      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth || img.width;
      canvas.height = img.naturalHeight || img.height;
      const ctx = canvas.getContext('2d');

      // Tùy chọn xoay nếu có gán cho file
      if (f.rotateDeg) {
        if (f.rotateDeg === 90 || f.rotateDeg === 270) {
          canvas.width = img.naturalHeight || img.height;
          canvas.height = img.naturalWidth || img.width;
        }
        ctx.translate(canvas.width / 2, canvas.height / 2);
        ctx.rotate((f.rotateDeg * Math.PI) / 180);
        ctx.drawImage(img, -img.naturalWidth / 2, -img.naturalHeight / 2);
      } else {
        ctx.drawImage(img, 0, 0);
      }

      // Nén ảnh sang JPEG
      const jpegDataUrl = canvas.toDataURL('image/jpeg', qValue);
      const embedded = await pdfDoc.embedJpg(jpegDataUrl);

      if (pageSizeMode === 'a4') {
        const isLandscape = embedded.width > embedded.height;
        const pageW = isLandscape ? A4_H : A4_W;
        const pageH = isLandscape ? A4_W : A4_H;
        const page = pdfDoc.addPage([pageW, pageH]);

        // Căn giữa và fit trong A4 (chừa lề 20pt)
        const margin = 20;
        const maxW = pageW - margin * 2;
        const maxH = pageH - margin * 2;
        const scale = Math.min(maxW / embedded.width, maxH / embedded.height);
        const w = embedded.width * scale;
        const h = embedded.height * scale;
        const x = (pageW - w) / 2;
        const y = (pageH - h) / 2;

        page.drawImage(embedded, { x, y, width: w, height: h });
      } else {
        // Theo kích thước gốc của ảnh
        const page = pdfDoc.addPage([embedded.width, embedded.height]);
        page.drawImage(embedded, { x: 0, y: 0, width: embedded.width, height: embedded.height });
      }
    }

    onProgress(95, 'Đang lưu file PDF hoàn chỉnh...');
    const pdfBytes = await pdfDoc.save();
    onProgress(100, 'Xong!');

    const baseName = files[0].name.replace(/\.[^.]+$/, '');
    return {
      blob: new Blob([pdfBytes], { type: 'application/pdf' }),
      filename: `${baseName}_ghep_anh.pdf`,
      server: false
    };
  }

  window.Group3Tools = {
    performOcr,
    convertPdfToImages,
    convertImagesToPdf
  };
})(typeof window !== 'undefined' ? window : this);

