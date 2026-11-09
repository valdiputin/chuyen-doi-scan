/**
 * group2_office.js
 * Nhóm 2: Office nội bộ [Trình duyệt]
 * - Word → Excel (mammoth: mỗi bảng 1 sheet, form "Nhãn: giá trị" gom thành sheet 2 cột)
 * - Excel → Word (SheetJS + docx: mỗi sheet thành bảng Word, tự động khổ ngang khi rộng)
 * - PowerPoint → Word (JSZip đọc slide XML + notes XML: dàn ý tiêu đề, gạch đầu dòng giữ cấp độ thụt lề, ghi chú)
 */

(function (window) {
  'use strict';

  // --- 1. Word → Excel (mammoth) ---
  async function convertWordToExcel(file, options, onProgress) {
    onProgress(15, 'Đang chuẩn bị thư viện đọc Word và xuất Excel...');
    await window.CDNLoader.require(['mammoth', 'xlsx']);

    onProgress(30, 'Đang phân tích cấu trúc tài liệu Word...');
    const buffer = await file.arrayBuffer();
    const result = await window.mammoth.convertToHtml({ arrayBuffer: buffer });
    const html = result.value || '';

    const parser = new DOMParser();
    const doc = parser.parseFromString(html, 'text/html');

    const wb = window.XLSX.utils.book_new();
    let sheetCount = 0;

    // 1. Tìm tất cả các bảng <table> trong Word
    const tables = doc.querySelectorAll('table');
    tables.forEach((table, tIdx) => {
      const rows = [];
      table.querySelectorAll('tr').forEach(tr => {
        const row = [];
        tr.querySelectorAll('th, td').forEach(td => {
          row.push(td.textContent.trim());
        });
        if (row.length && row.some(c => c !== '')) {
          rows.push(row);
        }
      });

      if (rows.length > 0) {
        sheetCount++;
        const processed = window.SmartValue.processTableRows(rows);
        const ws = window.XLSX.utils.aoa_to_sheet(processed);
        window.XLSX.utils.book_append_sheet(wb, ws, `Bang_${tIdx + 1}`.slice(0, 31));
      }
    });

    // 2. Tìm văn bản dạng form "Nhãn: giá trị" ngoài bảng
    onProgress(60, 'Đang quét dữ liệu biểu mẫu (Nhãn: giá trị)...');
    const formRows = [['Nhãn / Trường', 'Giá trị']];
    const paragraphs = doc.querySelectorAll('p, li, div');

    paragraphs.forEach(p => {
      // Bỏ qua nếu nằm trong table
      if (p.closest('table')) return;
      const text = p.textContent.trim();
      if (!text) return;

      // Regex nhận diện "Nhãn: giá trị"
      const match = text.match(/^([^:\r\n]{2,50})\s*:\s*(.+)$/);
      if (match) {
        const label = match[1].trim();
        const value = match[2].trim();
        formRows.push([label, value]);
      }
    });

    if (formRows.length > 1) {
      sheetCount++;
      const processedForm = window.SmartValue.processTableRows(formRows);
      const wsForm = window.XLSX.utils.aoa_to_sheet(processedForm);
      window.XLSX.utils.book_append_sheet(wb, wsForm, 'BieuMau_ThongTin');
    }

    // Nếu không có bảng hay form nào, xuất nội dung văn bản từng dòng
    if (sheetCount === 0) {
      const textRows = [['Nội dung văn bản']];
      paragraphs.forEach(p => {
        const t = p.textContent.trim();
        if (t) textRows.push([t]);
      });
      const wsText = window.XLSX.utils.aoa_to_sheet(textRows);
      window.XLSX.utils.book_append_sheet(wb, wsText, 'VanBan');
    }

    onProgress(90, 'Đang tạo file Excel...');
    const outBytes = window.XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    onProgress(100, 'Xong!');

    const baseName = file.name.replace(/\.[^.]+$/, '');
    return {
      blob: new Blob([outBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      filename: `${baseName}.xlsx`,
      server: false
    };
  }

  // --- 2. Excel → Word (SheetJS + docx) ---
  async function convertExcelToWord(file, options, onProgress) {
    onProgress(15, 'Đang nạp thư viện đọc bảng tính và tạo Word...');
    await window.CDNLoader.require(['xlsx', 'docx']);

    onProgress(30, 'Đang đọc các sheet từ file Excel...');
    const buffer = await file.arrayBuffer();
    const wb = window.XLSX.read(buffer, { type: 'array' });

    const sections = [];

    wb.SheetNames.forEach((sheetName, sIdx) => {
      onProgress(30 + Math.floor((sIdx / wb.SheetNames.length) * 55), `Đang xử lý sheet: ${sheetName}...`);
      const ws = wb.Sheets[sheetName];
      const rawRows = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
      if (!rawRows || !rawRows.length) return;

      const processed = window.SmartValue.processTableRows(rawRows);
      if (!processed.length) return;

      // Tiêu đề sheet
      const children = [
        new window.docx.Paragraph({
          text: `Sheet: ${sheetName}`,
          heading: window.docx.HeadingLevel.HEADING_2,
          spacing: { after: 200 }
        })
      ];

      // Xác định số cột lớn nhất để kiểm tra khổ ngang
      let maxCols = 0;
      processed.forEach(r => { if (r.length > maxCols) maxCols = r.length; });
      const isWide = maxCols >= 6;

      // Xây dựng Table trong Word
      const tableRows = [];

      processed.forEach((r, rIdx) => {
        const isHeader = rIdx === 0;
        const cells = [];

        // Đảm bảo mỗi hàng đủ số cột
        for (let c = 0; c < maxCols; c++) {
          const val = r[c] !== undefined && r[c] !== null ? String(r[c]) : '';
          cells.push(
            new window.docx.TableCell({
              children: [
                new window.docx.Paragraph({
                  children: [
                    new window.docx.TextRun({
                      text: val,
                      bold: isHeader,
                      size: isHeader ? 20 : 18 // 10pt hoặc 9pt
                    })
                  ]
                })
              ],
              shading: isHeader ? { fill: 'DDEBF7' } : undefined,
              margins: { top: 100, bottom: 100, left: 120, right: 120 }
            })
          );
        }

        tableRows.push(new window.docx.TableRow({ children: cells, tableHeader: isHeader }));
      });

      const wordTable = new window.docx.Table({
        rows: tableRows,
        width: { size: 100, type: window.docx.WidthType.PERCENTAGE }
      });
      children.push(wordTable);

      // Section riêng cho từng sheet
      sections.push({
        properties: {
          page: {
            size: {
              orientation: isWide
                ? window.docx.PageOrientation.LANDSCAPE
                : window.docx.PageOrientation.PORTRAIT
            }
          }
        },
        children
      });
    });

    if (!sections.length) {
      sections.push({
        properties: {},
        children: [new window.docx.Paragraph({ text: 'File Excel không có dữ liệu bảng.' })]
      });
    }

    onProgress(90, 'Đang kết xuất tài liệu Word (.docx)...');
    const doc = new window.docx.Document({ sections });
    const blob = await window.docx.Packer.toBlob(doc);
    onProgress(100, 'Xong!');

    const baseName = file.name.replace(/\.[^.]+$/, '');
    return {
      blob,
      filename: `${baseName}.docx`,
      server: false
    };
  }

  // --- 3. PowerPoint → Word (JSZip đọc XML) ---
  async function convertPptToWord(file, options, onProgress) {
    const onlyNotes = !!options.onlyNotes;

    onProgress(15, 'Đang chuẩn bị thư viện giải nén PPTX và tạo Word...');
    await window.CDNLoader.require(['jszip', 'docx']);

    onProgress(30, 'Đang giải nén cấu trúc file trình chiếu...');
    const buffer = await file.arrayBuffer();
    const zip = await window.JSZip.loadAsync(buffer);

    // Tìm tất cả các file slide: ppt/slides/slide1.xml, slide2.xml...
    const slideFiles = [];
    zip.forEach((relativePath) => {
      if (/^ppt\/slides\/slide\d+\.xml$/i.test(relativePath)) {
        slideFiles.push(relativePath);
      }
    });

    // Sắp xếp thứ tự tự nhiên của slide: slide1, slide2, ..., slide10
    slideFiles.sort((a, b) => {
      const numA = parseInt(a.match(/\d+/)[0], 10);
      const numB = parseInt(b.match(/\d+/)[0], 10);
      return numA - numB;
    });

    const parser = new DOMParser();
    const paragraphs = [];

    for (let sIdx = 0; sIdx < slideFiles.length; sIdx++) {
      const slidePath = slideFiles[sIdx];
      const slideNum = sIdx + 1;
      onProgress(30 + Math.floor((slideNum / slideFiles.length) * 55), `Đang phân tích slide ${slideNum}/${slideFiles.length}...`);

      const slideXmlText = await zip.file(slidePath).async('text');
      const slideDoc = parser.parseFromString(slideXmlText, 'text/xml');

      // Tìm tiêu đề slide
      let slideTitle = '';
      const shapes = slideDoc.getElementsByTagName('p:sp');
      const bodyItems = [];

      for (let i = 0; i < shapes.length; i++) {
        const sp = shapes[i];
        const isTitle = !!sp.querySelector('p\\:ph[type="title"], p\\:ph[type="ctrTitle"], ph[type="title"], ph[type="ctrTitle"]');
        const pList = sp.getElementsByTagName('a:p');

        for (let j = 0; j < pList.length; j++) {
          const pEl = pList[j];
          // Lấy chữ trong paragraph
          const tList = pEl.getElementsByTagName('a:t');
          let text = '';
          for (let k = 0; k < tList.length; k++) {
            text += tList[k].textContent;
          }
          text = text.trim();
          if (!text) continue;

          if (isTitle && !slideTitle) {
            slideTitle = text;
          } else {
            // Lấy cấp độ thụt lề (indent level: 0, 1, 2...)
            const pPr = pEl.getElementsByTagName('a:pPr')[0];
            const lvl = pPr && pPr.getAttribute('lvl') ? parseInt(pPr.getAttribute('lvl'), 10) : 0;
            bodyItems.push({ text, lvl });
          }
        }
      }

      // Đọc ghi chú (slide notes)
      let notesText = '';
      const notesPath = `ppt/notesSlides/notesSlide${slideNum}.xml`;
      const notesFile = zip.file(notesPath);
      if (notesFile) {
        const notesXmlText = await notesFile.async('text');
        const notesDoc = parser.parseFromString(notesXmlText, 'text/xml');
        const nShapes = notesDoc.getElementsByTagName('p:sp');
        for (let ni = 0; ni < nShapes.length; ni++) {
          const nSp = nShapes[ni];
          // Bỏ qua header/footer/slideNum trong notes
          const isHeaderFooter = !!nSp.querySelector('p\\:ph[type="hdr"], p\\:ph[type="ftr"], p\\:ph[type="sldNum"], ph[type="hdr"], ph[type="ftr"], ph[type="sldNum"]');
          if (isHeaderFooter) continue;

          const nTextEls = nSp.getElementsByTagName('a:t');
          let nChunk = '';
          for (let ti = 0; ti < nTextEls.length; ti++) {
            nChunk += nTextEls[ti].textContent + ' ';
          }
          if (nChunk.trim() && !nChunk.includes('Slide ' + slideNum)) {
            notesText += nChunk.trim() + '\n';
          }
        }
      }

      // Nếu tùy chọn "chỉ xuất ghi chú", bỏ qua slide không có ghi chú
      if (onlyNotes && !notesText.trim()) {
        continue;
      }

      // Tiêu đề Slide
      paragraphs.push(
        new window.docx.Paragraph({
          text: `Slide ${slideNum}: ${slideTitle || '(Không có tiêu đề)'}`,
          heading: window.docx.HeadingLevel.HEADING_2,
          spacing: { before: 240, after: 120 }
        })
      );

      // Nếu không phải "chỉ xuất ghi chú", thêm nội dung dàn ý slide
      if (!onlyNotes) {
        bodyItems.forEach(item => {
          paragraphs.push(
            new window.docx.Paragraph({
              children: [new window.docx.TextRun(item.text)],
              bullet: { level: Math.min(item.lvl, 4) },
              spacing: { after: 60 }
            })
          );
        });
      }

      // Thêm phần ghi chú (notes) nếu có
      if (notesText.trim()) {
        paragraphs.push(
          new window.docx.Paragraph({
            children: [
              new window.docx.TextRun({ text: '📝 Ghi chú (Notes): ', bold: true, color: '2563EB' }),
              new window.docx.TextRun({ text: notesText.trim(), italics: true })
            ],
            spacing: { before: 80, after: 160 }
          })
        );
      }
    }

    if (!paragraphs.length) {
      paragraphs.push(new window.docx.Paragraph({ text: 'Không tìm thấy nội dung phù hợp trong bài thuyết trình.' }));
    }

    onProgress(90, 'Đang tạo file Word (.docx)...');
    const doc = new window.docx.Document({
      sections: [{ properties: {}, children: paragraphs }]
    });

    const blob = await window.docx.Packer.toBlob(doc);
    onProgress(100, 'Xong!');

    const baseName = file.name.replace(/\.[^.]+$/, '');
    return {
      blob,
      filename: `${baseName}_dan_y.docx`,
      server: false
    };
  }

  window.Group2Tools = {
    convertWordToExcel,
    convertExcelToWord,
    convertPptToWord
  };
})(typeof window !== 'undefined' ? window : this);
