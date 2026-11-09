/**
 * smart_value.js
 * Cổng trực tiếp từ logic _to_value, _account_cols, _account_value trong converter.py
 * Áp dụng thống nhất cho tất cả các luồng xử lý Excel, CSV, JSON.
 */

(function (window) {
  'use strict';

  /**
   * Kiểm tra xem tiêu đề cột có phải là cột Tài khoản, STK hoặc MST hay không.
   */
  function isAccountHeader(header) {
    if (!header) return false;
    const h = String(header)
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, ''); // Bỏ dấu tiếng Việt để đối chiếu an toàn
    return h.includes('tai khoan') || h.includes('stk') || h.includes('mst') || h.includes('account');
  }

  /**
   * Xử lý giá trị cột tài khoản / STK / MST:
   * Luôn là text; bỏ dấu +/- và khoảng trắng ở đầu, vd '- 0123' -> '0123'.
   * Không bao giờ biến thành số âm hay mất số 0 ở đầu.
   */
  function cleanAccountValue(val) {
    if (val === null || val === undefined) return '';
    const text = String(val).trim();
    if (!text) return '';
    const t = text.replace(/\s+/g, '');
    const digits = t.replace(/^[+-]+/, '');
    if (/^\d+$/.test(digits)) {
      return digits;
    }
    return text;
  }

  /**
   * Chuyển số kiểu '1.234,5' hoặc '1,234.5' thành số nếu được.
   * Cổng chính xác từ _to_value trong converter.py:
   * - Hiểu cả '1.234,5' lẫn '1,234.5'.
   * - Chuỗi toàn chữ số bắt đầu bằng 0 (dài > 1), hoặc dài từ 10 chữ số trở lên, giữ nguyên text.
   * - Có thể kèm dấu +/- ở đầu, vd '-0201245001' (không được thành số âm).
   */
  function smartValue(val) {
    if (val === null || val === undefined) return '';
    if (typeof val === 'number') {
      return Number.isFinite(val) ? val : '';
    }
    const raw = String(val).trim();
    if (!raw) return '';

    const t = raw.replace(/\s+/g, '');
    if (!t) return '';

    // Nếu chứa chữ cái (bao gồm chữ có dấu tiếng Việt) -> giữ nguyên text
    if (/[a-zA-Z\u00C0-\u024F\u1EA0-\u1EF9]/.test(t)) {
      return raw;
    }

    // Mã/số tài khoản: bắt đầu bằng 0 hoặc dãy số dài -> giữ dạng text
    const digits = t.replace(/^[+-]+/, '');
    if (/^\d+$/.test(digits)) {
      if (digits.length >= 10 || (digits.startsWith('0') && digits.length > 1)) {
        return raw;
      }
    }

    // Các ứng viên chuyển đổi
    const candidates = [];
    if (t.includes('.') && t.includes(',')) {
      if (t.indexOf('.') < t.indexOf(',')) {
        candidates.push(t.replace(/\./g, '').replace(/,/g, '.'));
      } else {
        candidates.push(t.replace(/,/g, ''));
      }
    }
    candidates.push(t, t.replace(/,/g, ''), t.replace(/\./g, '').replace(/,/g, '.'));

    for (const c of candidates) {
      // Chỉ parse khi c đúng định dạng số hợp lệ (có thể kèm dấu +/-)
      if (/^[+-]?\d+(\.\d+)?$/.test(c)) {
        const v = parseFloat(c);
        if (!isNaN(v) && isFinite(v)) {
          if (Number.isInteger(v) && !c.includes('.')) {
            return v;
          }
          return v;
        }
      }
    }
    return raw;
  }

  /**
   * Xử lý bảng 2 chiều (mảng các hàng):
   * Hàng đầu tiên (index 0) là tiêu đề, từ hàng 1 trở đi áp dụng cleanAccountValue hoặc smartValue.
   */
  function processTableRows(rows) {
    if (!rows || !rows.length) return [];
    const header = rows[0] || [];
    const accountCols = new Set();
    header.forEach((col, idx) => {
      if (isAccountHeader(col)) {
        accountCols.add(idx);
      }
    });

    const result = [header.map(h => (h === null || h === undefined ? '' : String(h).trim()))];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !row.length) continue;
      const newRow = [];
      for (let c = 0; c < row.length; c++) {
        const cell = row[c];
        if (accountCols.has(c)) {
          newRow.push(cleanAccountValue(cell));
        } else {
          newRow.push(smartValue(cell));
        }
      }
      result.push(newRow);
    }
    return result;
  }

  /**
   * Chuyển mảng bảng 2 chiều thành mảng Object dựa trên header,
   * đã áp dụng đầy đủ quy tắc smartValue và cleanAccountValue.
   */
  function tableToObjects(rows) {
    if (!rows || rows.length < 2) return [];
    const processed = processTableRows(rows);
    const headers = processed[0];
    const data = [];
    for (let i = 1; i < processed.length; i++) {
      const row = processed[i];
      if (!row.some(c => c !== '' && c !== null && c !== undefined)) continue;
      const obj = {};
      headers.forEach((h, colIdx) => {
        const key = h || `Cot_${colIdx + 1}`;
        obj[key] = row[colIdx] !== undefined ? row[colIdx] : '';
      });
      data.push(obj);
    }
    return data;
  }

  window.SmartValue = {
    isAccountHeader,
    cleanAccountValue,
    smartValue,
    processTableRows,
    tableToObjects
  };
})(typeof window !== 'undefined' ? window : this);

