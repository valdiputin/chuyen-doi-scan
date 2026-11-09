/**
 * group4_data.js
 * Nhóm 4: Dữ liệu [Trình duyệt]
 * - Excel/CSV → JSON (áp dụng smartValue, cleanAccountValue, multi-sheet, pretty print)
 * - JSON → Excel/CSV (làm phẳng a.b.c, mảng lồng tách sheet có khóa liên kết)
 * - Excel ⇄ CSV (BOM UTF-8, dấu phân tách , ; tab tự nhận & chọn, multi-sheet ZIP)
 * - XML → Excel/CSV (ưu tiên hóa đơn điện tử TT78: TongHop + ChiTiet; XML thường: lặp nhiều nhất)
 */

(function (window) {
  'use strict';

  // --- 1. Excel/CSV → JSON ---
  async function convertExcelCsvToJson(file, options, onProgress) {
    const prettyPrint = !!options.prettyPrint;

    onProgress(15, 'Đang đọc file dữ liệu...');
    await window.CDNLoader.require(['xlsx']);

    const buffer = await file.arrayBuffer();
    const wb = window.XLSX.read(buffer, { type: 'array', raw: true, cellText: true });

    onProgress(50, 'Đang phân tích bảng và chuẩn hóa dữ liệu...');
    const resultObj = {};

    wb.SheetNames.forEach(sheetName => {
      const ws = wb.Sheets[sheetName];
      // Đọc bảng dạng 2D mảng string/raw
      const rawRows = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
      if (rawRows && rawRows.length > 0) {
        resultObj[sheetName] = window.SmartValue.tableToObjects(rawRows);
      } else {
        resultObj[sheetName] = [];
      }
    });

    onProgress(85, 'Đang định dạng JSON...');
    // Nếu chỉ có 1 sheet duy nhất, xuất thẳng mảng
    let finalData = resultObj;
    if (wb.SheetNames.length === 1) {
      finalData = resultObj[wb.SheetNames[0]];
    }

    const jsonText = prettyPrint ? JSON.stringify(finalData, null, 2) : JSON.stringify(finalData);
    const blob = new Blob([jsonText], { type: 'application/json;charset=utf-8' });
    onProgress(100, 'Xong!');

    return {
      blob,
      text: jsonText,
      filename: file.name.replace(/\.[^.]+$/, '') + '.json',
      server: false
    };
  }

  // --- 2. JSON → Excel/CSV ---
  // Hàm làm phẳng object: { a: { b: 1 } } -> { "a.b": 1 }
  function flattenObject(obj, prefix = '') {
    const flattened = {};
    for (const key in obj) {
      if (!Object.prototype.hasOwnProperty.call(obj, key)) continue;
      const propName = prefix ? `${prefix}.${key}` : key;
      const val = obj[key];
      if (val !== null && typeof val === 'object' && !Array.isArray(val)) {
        Object.assign(flattened, flattenObject(val, propName));
      } else if (!Array.isArray(val)) {
        flattened[propName] = val;
      }
    }
    return flattened;
  }

  async function convertJsonToExcelCsv(file, options, onProgress) {
    const outFormat = options.outFormat || 'xlsx'; // xlsx | csv
    onProgress(15, 'Đang đọc và phân tích cấu trúc JSON...');
    await window.CDNLoader.require(['xlsx']);

    const text = await file.text();
    let data;
    try {
      data = JSON.parse(text);
    } catch (e) {
      throw new Error(`Cú pháp JSON không hợp lệ: ${e.message}`);
    }

    onProgress(40, 'Đang trích xuất mảng dữ liệu và làm phẳng object...');
    const wb = window.XLSX.utils.book_new();

    // Tìm các mảng dữ liệu
    let mainList = [];
    const nestedSheets = {}; // sheetName -> array of items

    if (Array.isArray(data)) {
      mainList = data;
    } else if (typeof data === 'object' && data !== null) {
      // Tìm xem có property nào là mảng chính không
      const arrayKeys = Object.keys(data).filter(k => Array.isArray(data[k]));
      if (arrayKeys.length > 0) {
        arrayKeys.forEach(k => {
          nestedSheets[k] = data[k];
        });
        mainList = data[arrayKeys[0]];
      } else {
        // Nếu là single object, biến thành mảng 1 phần tử
        mainList = [data];
      }
    }

    // Xử lý mainList: làm phẳng object và tách mảng con
    const processedMain = [];
    mainList.forEach((item, rowIdx) => {
      if (typeof item !== 'object' || item === null) {
        processedMain.push({ GiaTri: item });
        return;
      }
      const flat = flattenObject(item);
      const parentId = flat.id || flat.ID || flat.Id || (rowIdx + 1);

      // Tìm mảng con lồng bên trong
      for (const k in item) {
        if (Array.isArray(item[k])) {
          const childSheetName = k.slice(0, 31);
          if (!nestedSheets[childSheetName]) nestedSheets[childSheetName] = [];
          item[k].forEach((childItem, cIdx) => {
            if (typeof childItem === 'object' && childItem !== null) {
              const childFlat = flattenObject(childItem);
              childFlat['_parent_id'] = parentId;
              nestedSheets[childSheetName].push(childFlat);
            } else {
              nestedSheets[childSheetName].push({
                _parent_id: parentId,
                GiaTri: childItem
              });
            }
          });
        }
      }
      processedMain.push(flat);
    });

    // Tạo sheet chính
    if (processedMain.length > 0) {
      const wsMain = window.XLSX.utils.json_to_sheet(processedMain);
      window.XLSX.utils.book_append_sheet(wb, wsMain, 'DuLieuChinh');
    }

    // Thêm các sheet con
    for (const sheetName in nestedSheets) {
      if (nestedSheets[sheetName].length > 0 && sheetName !== 'DuLieuChinh') {
        const wsSub = window.XLSX.utils.json_to_sheet(nestedSheets[sheetName]);
        window.XLSX.utils.book_append_sheet(wb, wsSub, sheetName.slice(0, 31));
      }
    }

    onProgress(85, 'Đang xuất file...');
    const baseName = file.name.replace(/\.[^.]+$/, '');

    if (outFormat === 'csv') {
      const firstSheetName = wb.SheetNames[0] || 'DuLieuChinh';
      const ws = wb.Sheets[firstSheetName];
      const csvContent = '\uFEFF' + window.XLSX.utils.sheet_to_csv(ws);
      onProgress(100, 'Xong!');
      return {
        blob: new Blob([csvContent], { type: 'text/csv;charset=utf-8' }),
        filename: `${baseName}.csv`,
        server: false
      };
    }

    const outBytes = window.XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    onProgress(100, 'Xong!');
    return {
      blob: new Blob([outBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      filename: `${baseName}.xlsx`,
      server: false
    };
  }

  // --- 3. Excel ⇄ CSV ---
  // Tự nhận dấu phân cách của CSV
  function detectCsvSeparator(text) {
    const lines = text.slice(0, 5000).split('\n').filter(l => l.trim());
    if (!lines.length) return ',';
    const firstLine = lines[0];
    const commas = (firstLine.match(/,/g) || []).length;
    const semicolons = (firstLine.match(/;/g) || []).length;
    const tabs = (firstLine.match(/\t/g) || []).length;
    if (tabs > commas && tabs > semicolons) return '\t';
    if (semicolons > commas) return ';';
    return ',';
  }

  async function convertExcelCsvSwap(file, options, onProgress) {
    const isCsvInput = file.name.toLowerCase().endsWith('.csv');
    const chosenSep = options.separator || ','; // , | ; | \t
    onProgress(15, 'Đang đọc file...');
    await window.CDNLoader.require(['xlsx', 'jszip']);

    const baseName = file.name.replace(/\.[^.]+$/, '');

    // Nếu đầu vào là CSV → Chuyển sang Excel (.xlsx)
    if (isCsvInput) {
      onProgress(35, 'Đang tự nhận diện định dạng CSV và áp dụng SmartValue...');
      const text = await file.text();
      const detectedSep = detectCsvSeparator(text);

      // Phân tích dòng CSV
      const lines = text.split(/\r?\n/).filter(l => l.trim());
      const rawRows = lines.map(line => {
        // Tách theo dấu phân cách cơ bản có hỗ trợ trích dẫn kép
        const re = new RegExp(`(?:^|${detectedSep === '\t' ? '\\t' : detectedSep})(?:"([^"]*(?:""[^"]*)*)"|([^"${detectedSep === '\t' ? '\\t' : detectedSep}]*))`, 'g');
        const row = [];
        let m;
        while ((m = re.exec(line)) !== null) {
          if (m.index === re.lastIndex) re.lastIndex++;
          const val = m[1] !== undefined ? m[1].replace(/""/g, '"') : (m[2] !== undefined ? m[2] : '');
          row.push(val);
        }
        return row;
      });

      // Áp dụng SmartValue và chuẩn hóa tài khoản
      const processed = window.SmartValue.processTableRows(rawRows);
      const wb = window.XLSX.utils.book_new();
      const ws = window.XLSX.utils.aoa_to_sheet(processed);
      window.XLSX.utils.book_append_sheet(wb, ws, 'DuLieu');

      onProgress(85, 'Đang xuất file Excel...');
      const out = window.XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
      onProgress(100, 'Xong!');
      return {
        blob: new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
        filename: `${baseName}.xlsx`,
        server: false
      };
    }

    // Nếu đầu vào là Excel → Chuyển sang CSV (UTF-8 BOM, chọn dấu phân cách)
    onProgress(35, 'Đang đọc các sheet trong file Excel...');
    const buffer = await file.arrayBuffer();
    const wb = window.XLSX.read(buffer, { type: 'array' });

    const csvFiles = [];
    wb.SheetNames.forEach(name => {
      const ws = wb.Sheets[name];
      const rawRows = window.XLSX.utils.sheet_to_json(ws, { header: 1, raw: false, defval: '' });
      const processed = window.SmartValue.processTableRows(rawRows);

      // Xuất CSV thủ công theo dấu phân cách đã chọn
      const csvLines = processed.map(row => {
        return row.map(val => {
          let s = val === null || val === undefined ? '' : String(val);
          if (s.includes('"') || s.includes(chosenSep) || s.includes('\n') || s.includes('\r')) {
            s = `"${s.replace(/"/g, '""')}"`;
          }
          return s;
        }).join(chosenSep);
      });

      // Thêm BOM UTF-8 (\uFEFF)
      const csvStr = '\uFEFF' + csvLines.join('\r\n');
      csvFiles.push({
        name: `${baseName}_${name}.csv`,
        content: csvStr
      });
    });

    if (csvFiles.length === 1) {
      onProgress(100, 'Xong!');
      return {
        blob: new Blob([csvFiles[0].content], { type: 'text/csv;charset=utf-8' }),
        filename: `${baseName}.csv`,
        server: false
      };
    }

    // Nhiều sheet: nén thành file ZIP
    onProgress(80, 'Excel có nhiều sheet, đang đóng gói file ZIP...');
    const zip = new window.JSZip();
    csvFiles.forEach(f => {
      zip.file(f.name, f.content);
    });
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    onProgress(100, 'Hoàn thành!');
    return {
      blob: zipBlob,
      filename: `${baseName}_csv_sheets.zip`,
      server: false
    };
  }

  // --- 4. XML → Excel/CSV (Hóa đơn điện tử TT78 & XML thường) ---
  function getXmlTagText(parent, tagName) {
    if (!parent) return '';
    const el = parent.getElementsByTagName(tagName)[0];
    return el && el.textContent ? el.textContent.trim() : '';
  }

  async function convertXmlToExcel(files, options, onProgress) {
    if (!files || !files.length) {
      throw new Error('Vui lòng chọn ít nhất một file XML.');
    }

    onProgress(15, 'Đang chuẩn bị thư viện xuất Excel...');
    await window.CDNLoader.require(['xlsx']);

    const parser = new DOMParser();
    const tongHopRows = [];
    const chiTietRows = [];

    // Header sheet TongHop
    tongHopRows.push([
      'File XML',
      'Ký hiệu mẫu số',
      'Ký hiệu HĐ',
      'Số hóa đơn',
      'Ngày lập',
      'MST Người Bán',
      'Tên Người Bán',
      'MST Người Mua',
      'Tên Người Mua',
      'Tiền trước thuế',
      'Tiền thuế',
      'Tổng tiền TT'
    ]);

    // Header sheet ChiTiet
    chiTietRows.push([
      'File XML',
      'Số hóa đơn',
      'Ký hiệu HĐ',
      'STT',
      'Tên hàng hóa, dịch vụ',
      'Đơn vị tính',
      'Số lượng',
      'Đơn giá',
      'Thành tiền',
      'Thuế suất',
      'Tiền thuế',
      'Tổng tiền'
    ]);

    let isAnyInvoice = false;
    const genericItems = [];

    for (let i = 0; i < files.length; i++) {
      const f = files[i];
      onProgress(20 + Math.floor((i / files.length) * 60), `Đang phân tích XML: ${f.name}...`);
      const xmlText = await f.text();
      const xmlDoc = parser.parseFromString(xmlText, 'text/xml');

      // Kiểm tra có phải Hóa đơn TT78 không (thẻ HDon, DLHDon, TTChung, KHHDon, SHDon, ...)
      const isInvoice = !!(
        xmlDoc.getElementsByTagName('HDon').length ||
        xmlDoc.getElementsByTagName('DLHDon').length ||
        xmlDoc.getElementsByTagName('KHHDon').length ||
        xmlDoc.getElementsByTagName('SHDon').length
      );

      if (isInvoice) {
        isAnyInvoice = true;
        // Trích xuất thông tin chung
        const khmsh = getXmlTagText(xmlDoc, 'KHMSHDon') || getXmlTagText(xmlDoc, 'KHDon');
        const khhd = getXmlTagText(xmlDoc, 'KHHDon');
        const shd = getXmlTagText(xmlDoc, 'SHDon');
        const nlap = getXmlTagText(xmlDoc, 'NLap') || getXmlTagText(xmlDoc, 'TDLap');

        // Người bán
        const nban = xmlDoc.getElementsByTagName('NBan')[0] || xmlDoc;
        const mstBan = getXmlTagText(nban, 'MST') || getXmlTagText(xmlDoc, 'MST');
        const tenBan = getXmlTagText(nban, 'Ten') || getXmlTagText(xmlDoc, 'Ten');

        // Người mua
        const nmua = xmlDoc.getElementsByTagName('NMua')[0];
        const mstMua = getXmlTagText(nmua, 'MST');
        const tenMua = getXmlTagText(nmua, 'Ten');

        // Tiền
        const ttoan = xmlDoc.getElementsByTagName('TToan')[0] || xmlDoc;
        const tgtcthue = getXmlTagText(ttoan, 'TgTCThue');
        const tgtthue = getXmlTagText(ttoan, 'TgTThue');
        const tgtttbso = getXmlTagText(ttoan, 'TgTTTBSo');

        tongHopRows.push([
          f.name,
          khmsh,
          khhd,
          shd,
          nlap,
          mstBan,
          tenBan,
          mstMua,
          tenMua,
          tgtcthue,
          tgtthue,
          tgtttbso
        ]);

        // Chi tiết từng dòng hàng hóa dịch vụ (HHDVu)
        const hhdvuList = xmlDoc.getElementsByTagName('HHDVu');
        for (let j = 0; j < hhdvuList.length; j++) {
          const item = hhdvuList[j];
          chiTietRows.push([
            f.name,
            shd,
            khhd,
            getXmlTagText(item, 'STT') || (j + 1),
            getXmlTagText(item, 'THHVu'),
            getXmlTagText(item, 'DVTinh'),
            getXmlTagText(item, 'SLuong'),
            getXmlTagText(item, 'DGia'),
            getXmlTagText(item, 'ThTien'),
            getXmlTagText(item, 'TSuat'),
            getXmlTagText(item, 'TThue'),
            getXmlTagText(item, 'TongTien') || getXmlTagText(item, 'ThTien')
          ]);
        }
      } else {
        // XML thông thường: Tìm thẻ lặp nhiều nhất ở cùng cấp làm dòng dữ liệu
        const tagCounts = {};
        const allElements = xmlDoc.getElementsByTagName('*');
        for (let e = 0; e < allElements.length; e++) {
          const tag = allElements[e].tagName;
          tagCounts[tag] = (tagCounts[tag] || 0) + 1;
        }

        let bestTag = '';
        let maxCount = 0;
        for (const tag in tagCounts) {
          if (tagCounts[tag] > maxCount && tagCounts[tag] > 1) {
            maxCount = tagCounts[tag];
            bestTag = tag;
          }
        }

        if (bestTag) {
          const repeated = xmlDoc.getElementsByTagName(bestTag);
          for (let r = 0; r < repeated.length; r++) {
            const el = repeated[r];
            const obj = { 'File XML': f.name };
            for (let c = 0; c < el.children.length; c++) {
              const child = el.children[c];
              obj[child.tagName] = child.textContent.trim();
            }
            genericItems.push(obj);
          }
        }
      }
    }

    onProgress(85, 'Đang áp dụng SmartValue và tạo file Excel...');
    const wb = window.XLSX.utils.book_new();

    if (isAnyInvoice) {
      // Áp dụng SmartValue cho cả 2 sheet hóa đơn
      const cleanTongHop = window.SmartValue.processTableRows(tongHopRows);
      const cleanChiTiet = window.SmartValue.processTableRows(chiTietRows);

      const wsTongHop = window.XLSX.utils.aoa_to_sheet(cleanTongHop);
      const wsChiTiet = window.XLSX.utils.aoa_to_sheet(cleanChiTiet);

      window.XLSX.utils.book_append_sheet(wb, wsTongHop, 'TongHop');
      window.XLSX.utils.book_append_sheet(wb, wsChiTiet, 'ChiTiet');
    } else if (genericItems.length > 0) {
      const wsGeneric = window.XLSX.utils.json_to_sheet(genericItems);
      window.XLSX.utils.book_append_sheet(wb, wsGeneric, 'DuLieuXML');
    } else {
      const emptyWs = window.XLSX.utils.aoa_to_sheet([['Thông báo'], ['Không tìm thấy dữ liệu có cấu trúc trong file XML.']]);
      window.XLSX.utils.book_append_sheet(wb, emptyWs, 'ThongBao');
    }

    const outBytes = window.XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
    onProgress(100, 'Xuất file thành công!');

    const baseName = files.length === 1 ? files[0].name.replace(/\.[^.]+$/, '') : 'hoa_don_tong_hop';
    return {
      blob: new Blob([outBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
      filename: `${baseName}.xlsx`,
      server: false
    };
  }

  window.Group4Tools = {
    convertExcelCsvToJson,
    convertJsonToExcelCsv,
    convertExcelCsvSwap,
    convertXmlToExcel
  };
})(typeof window !== 'undefined' ? window : this);

