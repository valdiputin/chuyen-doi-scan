/**
 * group5_ebook_web.js
 * Nhóm 5: Sách điện tử & Web
 * - HTML / Web URL → PDF (Gotenberg Chromium / Dự phòng html2pdf cho file HTML)
 * - PDF → EPUB (pdf.js nhận diện tiêu đề chương cỡ chữ lớn, dựng EPUB 3 bằng JSZip)
 * - EPUB → PDF (JSZip đọc manifest & spine, ghép chương, kết xuất PDF bằng html2pdf)
 */

(function (window) {
  'use strict';

  // --- 1. HTML / Web URL → PDF ---
  async function convertHtmlToPdf(file, options, onProgress) {
    const sourceType = options.sourceType || 'file'; // file | url
    const isLandscape = !!options.landscape;
    const urlInput = (options.urlInput || '').trim();

    // Trường hợp 1: Chuyển đổi từ URL trang web
    if (sourceType === 'url') {
      if (!urlInput) {
        throw new Error('Vui lòng nhập đường dẫn URL trang web cần chuyển đổi (ví dụ: https://example.com).');
      }
      onProgress(20, 'Đang gửi yêu cầu in trang web tới máy chủ Chromium...');
      const res = await fetch(`/api/office-convert?type=url&landscape=${isLandscape}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ url: urlInput })
      });

      if (!res.ok) {
        let errMsg = 'Không thể kết nối đến máy chủ Chromium để in trang web.';
        try {
          const j = await res.json();
          errMsg = j.message || errMsg;
        } catch (_) {}
        throw new Error(errMsg);
      }

      onProgress(100, 'In trang web sang PDF thành công!');
      const blob = await res.blob();
      let domainName = 'webpage';
      try {
        domainName = new URL(urlInput).hostname.replace(/[^a-zA-Z0-9]/g, '_');
      } catch (_) {}
      return {
        blob,
        filename: `${domainName}.pdf`,
        server: true
      };
    }

    // Trường hợp 2: Chuyển đổi từ file .html tải lên
    if (!file) {
      throw new Error('Vui lòng chọn một file HTML để chuyển đổi.');
    }

    onProgress(15, 'Đang đọc nội dung file HTML...');
    const htmlText = await file.text();

    // Thử gửi Gotenberg trước
    try {
      onProgress(30, 'Đang gửi file HTML lên server Gotenberg (Chromium)...');
      const res = await fetch(`/api/office-convert?type=html&landscape=${isLandscape}`, {
        method: 'POST',
        headers: { 'Content-Type': 'text/html; charset=utf-8' },
        body: htmlText
      });

      if (res.ok) {
        onProgress(100, 'Chuyển đổi thành công!');
        const blob = await res.blob();
        return {
          blob,
          filename: file.name.replace(/\.[^.]+$/, '') + '.pdf',
          server: true
        };
      }
    } catch (e) {
      console.warn('Gotenberg Chromium không khả dụng:', e);
    }

    // Dự phòng trình duyệt: html2pdf.js
    onProgress(50, 'Đang chạy phương án dự phòng trên trình duyệt (html2pdf)...');
    await window.CDNLoader.require(['html2pdf']);

    const container = document.createElement('div');
    container.innerHTML = htmlText;

    onProgress(75, 'Đang kết xuất PDF...');
    const pdfBlob = await window.html2pdf().from(container).set({
      margin: 10,
      filename: file.name.replace(/\.[^.]+$/, '') + '.pdf',
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: isLandscape ? 'landscape' : 'portrait' }
    }).output('blob');

    onProgress(100, 'Hoàn thành!');
    return {
      blob: pdfBlob,
      filename: file.name.replace(/\.[^.]+$/, '') + '.pdf',
      server: false,
      warning: 'Chất lượng cơ bản, bố cục có thể lệch (chế độ dự phòng trình duyệt).'
    };
  }

  // --- 2. PDF → EPUB [Trình duyệt: pdf.js + JSZip] ---
  async function convertPdfToEpub(file, options, onProgress) {
    onProgress(10, 'Đang chuẩn bị thư viện đọc PDF và tạo EPUB...');
    await window.CDNLoader.require(['pdfjs', 'jszip']);

    onProgress(20, 'Đang đọc tài liệu PDF...');
    const buffer = await file.arrayBuffer();
    const pdf = await window.pdfjsLib.getDocument({ data: buffer }).promise;
    const numPages = pdf.numPages;

    const baseTitle = file.name.replace(/\.[^.]+$/, '');
    const chapters = [];
    let currentChapter = {
      title: 'Mở đầu',
      paragraphs: []
    };

    // Ngưỡng cỡ chữ phát hiện tiêu đề chương
    const TITLE_FONT_SIZE_MIN = 15;

    for (let p = 1; p <= numPages; p++) {
      onProgress(20 + Math.floor((p / numPages) * 60), `Đang phân tích cấu trúc chữ trang ${p}/${numPages}...`);
      const page = await pdf.getPage(p);
      const content = await page.getTextContent();

      // Gom dòng theo tọa độ Y
      const linesMap = new Map();
      const Y_TOLERANCE = 4.0;

      for (const item of content.items) {
        const text = item.str;
        if (!text && !item.hasEOL) continue;
        const tx = item.transform; // [scaleX, skewY, skewX, scaleY, x, y]
        const y = tx[5];
        const fontSize = Math.hypot(tx[0], tx[1]); // Độ cao font chữ tính theo biến đổi

        let matchedY = null;
        for (const existingY of linesMap.keys()) {
          if (Math.abs(existingY - y) <= Y_TOLERANCE) {
            matchedY = existingY;
            break;
          }
        }
        if (matchedY === null) {
          matchedY = y;
          linesMap.set(matchedY, []);
        }
        linesMap.get(matchedY).push({ x: tx[4], text, fontSize });
      }

      // Sắp xếp các dòng từ trên xuống
      const sortedYs = Array.from(linesMap.keys()).sort((a, b) => b - a);

      for (const y of sortedYs) {
        const items = linesMap.get(y);
        items.sort((a, b) => a.x - b.x);
        const lineText = items.map(it => it.text).join(' ').trim();
        if (!lineText) continue;

        const maxFontSize = Math.max(...items.map(it => it.fontSize || 10));

        // Nếu dòng chữ có cỡ font lớn -> Nhận diện là tiêu đề chương mới
        const isChapterHeader = maxFontSize >= TITLE_FONT_SIZE_MIN && lineText.length < 100;
        if (isChapterHeader) {
          if (currentChapter.paragraphs.length > 0) {
            chapters.push(currentChapter);
          }
          currentChapter = {
            title: lineText,
            paragraphs: []
          };
        } else {
          currentChapter.paragraphs.push(lineText);
        }
      }
    }

    if (currentChapter.paragraphs.length > 0 || chapters.length === 0) {
      chapters.push(currentChapter);
    }

    onProgress(85, 'Đang đóng gói cấu trúc EPUB 3...');
    const zip = new window.JSZip();

    // 1. mimetype: PHẢI là file đầu tiên và không nén (compression: STORE)
    zip.file('mimetype', 'application/epub+zip', { compression: 'STORE' });

    // 2. META-INF/container.xml
    zip.folder('META-INF').file('container.xml', `<?xml version="1.0" encoding="UTF-8"?>
<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container">
  <rootfiles>
    <rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/>
  </rootfiles>
</container>`);

    const oebps = zip.folder('OEBPS');

    // 3. Các file chapter XHTML
    const manifestItems = [];
    const spineItems = [];
    const navTocItems = [];

    chapters.forEach((chap, cIdx) => {
      const chapId = `chap_${cIdx + 1}`;
      const chapFileName = `${chapId}.xhtml`;
      manifestItems.push(`<item id="${chapId}" href="${chapFileName}" media-type="application/xhtml+xml"/>`);
      spineItems.push(`<itemref idref="${chapId}"/>`);
      navTocItems.push(`<li><a href="${chapFileName}">${escapeXml(chap.title)}</a></li>`);

      const parasHtml = chap.paragraphs.map(p => `<p>${escapeXml(p)}</p>`).join('\n    ');
      const chapXhtml = `<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="vi">
<head>
  <meta charset="utf-8"/>
  <title>${escapeXml(chap.title)}</title>
  <style>
    body { font-family: sans-serif; line-height: 1.6; margin: 5%; }
    h1 { color: #1e3a8a; border-bottom: 1px solid #cbd5e1; padding-bottom: 0.3em; }
    p { margin-bottom: 1em; text-indent: 1.5em; text-align: justify; }
  </style>
</head>
<body>
  <h1>${escapeXml(chap.title)}</h1>
  <div class="chapter-content">
    ${parasHtml}
  </div>
</body>
</html>`;
      oebps.file(chapFileName, chapXhtml);
    });

    // 4. Navigation nav.xhtml
    manifestItems.push(`<item id="nav" href="nav.xhtml" media-type="application/xhtml+xml" properties="nav"/>`);
    const navXhtml = `<?xml version="1.0" encoding="utf-8"?>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml" xmlns:epub="http://www.idpf.org/2007/ops" xml:lang="vi">
<head>
  <meta charset="utf-8"/>
  <title>Mục lục</title>
</head>
<body>
  <nav epub:type="toc" id="toc">
    <h1>Mục lục</h1>
    <ol>
      ${navTocItems.join('\n      ')}
    </ol>
  </nav>
</body>
</html>`;
    oebps.file('nav.xhtml', navXhtml);

    // 5. content.opf
    const bookId = 'urn:uuid:' + generateUUID();
    const contentOpf = `<?xml version="1.0" encoding="utf-8"?>
<package xmlns="http://www.idpf.org/2007/opf" version="3.0" unique-identifier="BookId">
  <metadata xmlns:dc="http://purl.org/dc/elements/1.1/">
    <dc:title>${escapeXml(baseTitle)}</dc:title>
    <dc:language>vi</dc:language>
    <dc:identifier id="BookId">${bookId}</dc:identifier>
    <meta property="dcterms:modified">${new Date().toISOString().replace(/\.\d+Z$/, 'Z')}</meta>
  </metadata>
  <manifest>
    ${manifestItems.join('\n    ')}
  </manifest>
  <spine>
    ${spineItems.join('\n    ')}
  </spine>
</package>`;
    oebps.file('content.opf', contentOpf);

    onProgress(95, 'Đang xuất file sách EPUB 3...');
    const epubBlob = await zip.generateAsync({ type: 'blob', mimeType: 'application/epub+zip' });
    onProgress(100, 'Xong!');

    return {
      blob: epubBlob,
      filename: `${baseTitle}.epub`,
      server: false
    };
  }

  // --- 3. EPUB → PDF [Trình duyệt: JSZip + html2pdf] ---
  async function convertEpubToPdf(file, options, onProgress) {
    onProgress(10, 'Đang nạp thư viện giải nén EPUB và kết xuất PDF...');
    await window.CDNLoader.require(['jszip', 'html2pdf']);

    onProgress(25, 'Đang đọc và phân tích cấu trúc sách EPUB...');
    const buffer = await file.arrayBuffer();
    const zip = await window.JSZip.loadAsync(buffer);

    // 1. Đọc META-INF/container.xml để tìm file OPF
    const containerXml = await zip.file('META-INF/container.xml').async('text');
    const parser = new DOMParser();
    const containerDoc = parser.parseFromString(containerXml, 'text/xml');
    const rootfileEl = containerDoc.getElementsByTagName('rootfile')[0];
    const opfPath = rootfileEl ? rootfileEl.getAttribute('full-path') : 'OEBPS/content.opf';

    const opfDir = opfPath.includes('/') ? opfPath.substring(0, opfPath.lastIndexOf('/') + 1) : '';
    const opfXml = await zip.file(opfPath).async('text');
    const opfDoc = parser.parseFromString(opfXml, 'text/xml');

    // 2. Đọc manifest để ánh xạ item ID -> href
    const manifestMap = new Map();
    const itemEls = opfDoc.getElementsByTagName('item');
    for (let i = 0; i < itemEls.length; i++) {
      manifestMap.set(itemEls[i].getAttribute('id'), itemEls[i].getAttribute('href'));
    }

    // 3. Đọc spine để lấy thứ tự các chương cần ghép
    const spineItemRefs = opfDoc.getElementsByTagName('itemref');
    const chapterHrefs = [];
    for (let i = 0; i < spineItemRefs.length; i++) {
      const idref = spineItemRefs[i].getAttribute('idref');
      const href = manifestMap.get(idref);
      if (href) chapterHrefs.push(href);
    }

    onProgress(45, 'Đang ghép các chương theo đúng thứ tự đọc (spine)...');
    let combinedHtml = `<style>
      body { font-family: 'Times New Roman', serif; font-size: 12pt; line-height: 1.6; color: #111; }
      .epub-chapter { page-break-before: always; padding: 15px 0; }
      .epub-chapter:first-child { page-break-before: avoid; }
      h1, h2, h3 { font-family: Arial, sans-serif; }
      p { margin-bottom: 0.8em; text-indent: 1.5em; text-align: justify; }
      img { max-width: 100%; height: auto; }
    </style>`;

    for (let i = 0; i < chapterHrefs.length; i++) {
      const href = chapterHrefs[i];
      onProgress(45 + Math.floor((i / chapterHrefs.length) * 35), `Đang tải chương ${i + 1}/${chapterHrefs.length}...`);
      const fullPath = opfDir + href;
      const chapFile = zip.file(fullPath);
      if (chapFile) {
        const chapHtml = await chapFile.async('text');
        const chapDoc = parser.parseFromString(chapHtml, 'text/html');
        const bodyContent = chapDoc.body ? chapDoc.body.innerHTML : chapHtml;
        combinedHtml += `<div class="epub-chapter">${bodyContent}</div>`;
      }
    }

    onProgress(85, 'Đang kết xuất PDF...');
    const container = document.createElement('div');
    container.innerHTML = combinedHtml;

    const baseName = file.name.replace(/\.[^.]+$/, '');
    const pdfBlob = await window.html2pdf().from(container).set({
      margin: 15,
      filename: `${baseName}.pdf`,
      html2canvas: { scale: 2, useCORS: true },
      jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    }).output('blob');

    onProgress(100, 'Xong!');
    return {
      blob: pdfBlob,
      filename: `${baseName}.pdf`,
      server: false
    };
  }

  function escapeXml(str) {
    return String(str || '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function generateUUID() {
    return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) {
      const r = (Math.random() * 16) | 0;
      const v = c === 'x' ? r : (r & 0x3) | 0x8;
      return v.toString(16);
    });
  }

  window.Group5Tools = {
    convertHtmlToPdf,
    convertPdfToEpub,
    convertEpubToPdf
  };
})(typeof window !== 'undefined' ? window : this);
