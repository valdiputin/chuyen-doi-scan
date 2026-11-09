/**
 * app.js
 * Quản trị giao diện chính, điều hướng công cụ, kiểm tra sức khỏe server,
 * chế độ sáng/tối và xử lý luồng chuyển đổi.
 */

(function () {
  'use strict';

  // Định nghĩa danh sách các công cụ
  const TOOLS = [
    // Nhóm 1: PDF ⇄ Office
    {
      id: 'pdf-to-excel',
      name: 'PDF → Excel',
      group: 'group1',
      badge: 'server',
      desc: 'Trích xuất bảng biểu trong PDF sang Excel chuẩn xác. Tự động nhận diện số, tài khoản, chia phần nhỏ khi file lớn.',
      accept: '.pdf',
      multi: true,
      options: [
        {
          id: 'mode', label: 'Kiểu bảng', type: 'select', options: [
            { value: 'auto', text: 'Tự động' },
            { value: 'lines', text: 'Có kẻ viền' },
            { value: 'text', text: 'Không viền' }
          ], default: 'auto'
        },
        { id: 'pages', label: 'Trang cần chuyển', type: 'text', placeholder: 'Vd: 1-3,5 hoặc để trống (toàn bộ)', default: '' },
        { id: 'oneSheet', label: 'Gộp tất cả bảng vào 1 sheet (TatCa)', type: 'checkbox', default: true },
        { id: 'includeText', label: 'Xuất cả chữ ở trang không có bảng', type: 'checkbox', default: false }
      ]
    },
    {
      id: 'word-to-pdf',
      name: 'Word → PDF',
      group: 'group1',
      badge: 'server',
      desc: 'Giữ nguyên định dạng, phông chữ và bố cục. Tự động dự phòng trên trình duyệt nếu chưa cấu hình Gotenberg.',
      accept: '.docx,.doc',
      multi: false,
      options: []
    },
    {
      id: 'excel-to-pdf',
      name: 'Excel → PDF',
      group: 'group1',
      badge: 'server',
      desc: 'Chuyển đổi bảng tính sang PDF. Tùy chọn khổ ngang/dọc và khóa chỉnh sửa tài liệu.',
      accept: '.xlsx,.xls',
      multi: false,
      options: [
        {
          id: 'orientation', label: 'Khổ giấy', type: 'select', options: [
            { value: 'portrait', text: 'Khổ dọc (Portrait)' },
            { value: 'landscape', text: 'Khổ ngang (Landscape)' }
          ], default: 'landscape'
        },
        { id: 'lockEdit', label: 'Khóa chỉnh sửa (Read-Only)', type: 'checkbox', default: false }
      ]
    },
    {
      id: 'ppt-to-pdf',
      name: 'PowerPoint → PDF',
      group: 'group1',
      badge: 'server',
      desc: 'Chuyển đổi bài thuyết trình PowerPoint sang PDF. Tùy chọn xuất Handout 1, 2, 4, 6 slide mỗi trang.',
      accept: '.pptx,.ppt',
      multi: false,
      options: [
        {
          id: 'handout', label: 'Bố cục Handout', type: 'select', options: [
            { value: '1', text: '1 slide mỗi trang (Tiêu chuẩn)' },
            { value: '2', text: '2 slide mỗi trang' },
            { value: '4', text: '4 slide mỗi trang' },
            { value: '6', text: '6 slide mỗi trang' }
          ], default: '1'
        }
      ]
    },
    {
      id: 'pdf-to-word',
      name: 'PDF → Word',
      group: 'group1',
      badge: 'browser',
      desc: 'Trích chữ và gom dòng theo tọa độ, mỗi trang PDF sang một trang Word mới (.docx).',
      accept: '.pdf',
      multi: false,
      options: []
    },
    {
      id: 'pdf-to-ppt',
      name: 'PDF → PowerPoint',
      group: 'group1',
      badge: 'browser',
      desc: 'Mỗi trang PDF thành một slide ảnh sắc nét, kích thước slide tự điều chỉnh theo trang PDF.',
      accept: '.pdf',
      multi: false,
      options: []
    },

    // Nhóm 2: Office nội bộ (Đợt 2)
    {
      id: 'word-to-excel',
      name: 'Word → Excel',
      group: 'group2',
      badge: 'browser',
      desc: 'Mỗi bảng trong Word thành một sheet riêng. Văn bản dạng form "Nhãn: giá trị" gom thành sheet 2 cột.',
      accept: '.docx',
      multi: false,
      options: []
    },
    {
      id: 'excel-to-word',
      name: 'Excel → Word',
      group: 'group2',
      badge: 'browser',
      desc: 'Mỗi sheet thành một bảng Word có tiêu đề định dạng chuẩn, tự động xoay khổ ngang khi bảng rộng.',
      accept: '.xlsx,.xls',
      multi: false,
      options: []
    },
    {
      id: 'ppt-to-word',
      name: 'PowerPoint → Word',
      group: 'group2',
      badge: 'browser',
      desc: 'Dàn ý gồm tiêu đề slide, gạch đầu dòng giữ cấp độ thụt lề và ghi chú (notes).',
      accept: '.pptx',
      multi: false,
      options: [
        { id: 'onlyNotes', label: 'Chỉ xuất ghi chú thuyết trình (notes)', type: 'checkbox', default: false }
      ]
    },

    // Nhóm 3: Ảnh & OCR
    {
      id: 'ocr',
      name: 'Ảnh/Scan → Word / Excel / PDF / Text',
      group: 'group3',
      badge: 'browser',
      desc: 'Nhận diện chữ (OCR) tiếng Việt + tiếng Anh. Tự động tăng tương phản, xoay ảnh, xuất Word, Excel (gom cột theo tọa độ), hoặc PDF tìm kiếm được.',
      accept: '.jpg,.jpeg,.png,.webp,.pdf',
      multi: false,
      options: [
        {
          id: 'targetFormat', label: 'Định dạng đầu ra', type: 'select', options: [
            { value: 'docx', text: 'Microsoft Word (.docx) - Giữ đoạn văn' },
            { value: 'xlsx', text: 'Microsoft Excel (.xlsx) - Gom hàng cột' },
            { value: 'pdf', text: 'PDF tìm kiếm được (Ảnh + chữ ẩn)' },
            { value: 'txt', text: 'Văn bản thuần (.txt)' }
          ], default: 'docx'
        },
        {
          id: 'rotate', label: 'Xoay ảnh trước khi nhận diện', type: 'select', options: [
            { value: '0', text: 'Không xoay (0°)' },
            { value: '90', text: 'Xoay 90° theo chiều kim đồng hồ' },
            { value: '180', text: 'Xoay ngược 180°' },
            { value: '270', text: 'Xoay 270°' }
          ], default: '0'
        }
      ]
    },
    {
      id: 'pdf-to-image',
      name: 'PDF → Ảnh (PNG/JPG)',
      group: 'group3',
      badge: 'browser',
      desc: 'Chuyển từng trang PDF thành hình ảnh độ nét cao (96, 150 hoặc 300 DPI). Nhiều trang tự nén ZIP.',
      accept: '.pdf',
      multi: false,
      options: [
        {
          id: 'imgFormat', label: 'Định dạng ảnh', type: 'select', options: [
            { value: 'png', text: 'PNG (Không giảm chất lượng)' },
            { value: 'jpeg', text: 'JPG / JPEG (Dung lượng nhẹ)' }
          ], default: 'png'
        },
        {
          id: 'dpi', label: 'Độ phân giải DPI', type: 'select', options: [
            { value: '96', text: '96 DPI (Web / xem nhanh)' },
            { value: '150', text: '150 DPI (Cân bằng - Khuyên dùng)' },
            { value: '300', text: '300 DPI (Chất lượng in ấn)' }
          ], default: '150'
        },
        { id: 'pages', label: 'Khoảng trang', type: 'text', placeholder: 'vd: 1-5 hoặc all (tất cả)', default: 'all' }
      ]
    },
    {
      id: 'image-to-pdf',
      name: 'Ảnh → PDF',
      group: 'group3',
      badge: 'browser',
      desc: 'Ghép nhiều ảnh thành file PDF, cho phép xem trước, xoay từng ảnh, chọn khổ A4 hoặc khổ ảnh gốc.',
      accept: '.jpg,.jpeg,.png,.webp',
      multi: true,
      options: [
        {
          id: 'pageSize', label: 'Khổ trang PDF', type: 'select', options: [
            { value: 'a4', text: 'Khổ A4 tiêu chuẩn (Căn giữa)' },
            { value: 'fit', text: 'Theo kích thước ảnh gốc' }
          ], default: 'a4'
        },
        {
          id: 'quality', label: 'Chất lượng nén ảnh', type: 'select', options: [
            { value: 'high', text: 'Cao (Ảnh sắc nét)' },
            { value: 'medium', text: 'Vừa phải' },
            { value: 'low', text: 'Nén tối đa (Dung lượng nhỏ)' }
          ], default: 'high'
        }
      ]
    },

    // Nhóm 4: Dữ liệu
    {
      id: 'excel-to-json',
      name: 'Excel/CSV → JSON',
      group: 'group4',
      badge: 'browser',
      desc: 'Chuyển bảng thành mảng JSON objects. Tự động định dạng số SmartValue, cột Tài khoản/STK/MST giữ nguyên text.',
      accept: '.xlsx,.xls,.csv',
      multi: false,
      options: [
        { id: 'prettyPrint', label: 'In đẹp có thụt lề (Pretty Print)', type: 'checkbox', default: true }
      ]
    },
    {
      id: 'json-to-excel',
      name: 'JSON → Excel/CSV',
      group: 'group4',
      badge: 'browser',
      desc: 'Tự tìm mảng dữ liệu, làm phẳng object lồng nhau (a.b.c). Mảng lồng bên trong tách sheet riêng kèm khóa liên kết.',
      accept: '.json',
      multi: false,
      options: [
        {
          id: 'outFormat', label: 'Định dạng xuất', type: 'select', options: [
            { value: 'xlsx', text: 'Microsoft Excel (.xlsx)' },
            { value: 'csv', text: 'CSV UTF-8 (.csv)' }
          ], default: 'xlsx'
        }
      ]
    },
    {
      id: 'excel-csv-swap',
      name: 'Excel ⇄ CSV',
      group: 'group4',
      badge: 'browser',
      desc: 'Chuyển đổi 2 chiều Excel và CSV. CSV có BOM UTF-8, tùy chọn dấu phân cách (phẩy, chấm phẩy, tab). Nhiều sheet nén ZIP.',
      accept: '.xlsx,.xls,.csv',
      multi: false,
      options: [
        {
          id: 'separator', label: 'Dấu phân cách khi xuất CSV', type: 'select', options: [
            { value: ',', text: 'Dấu phẩy (,)' },
            { value: ';', text: 'Dấu chấm phẩy (;)' },
            { value: '\t', text: 'Dấu Tab (\\t)' }
          ], default: ','
        }
      ]
    },
    {
      id: 'xml-to-excel',
      name: 'XML → Excel/CSV (Hóa đơn TT78)',
      group: 'group4',
      badge: 'browser',
      desc: 'Chọn nhiều file XML. Xuất sheet TongHop và ChiTiet (HHDVu) chuẩn hóa đơn điện tử TT78. Với XML thường, tự tìm thẻ lặp làm dòng dữ liệu.',
      accept: '.xml',
      multi: true,
      options: []
    },

    // Nhóm 5: Sách điện tử & Web (Đợt 3)
    {
      id: 'html-to-pdf',
      name: 'HTML / Web URL → PDF',
      group: 'group5',
      badge: 'server',
      desc: 'Chuyển file .html hoặc nhập đường link trang web để in sang PDF qua Chromium. Có dự phòng cho file .html.',
      accept: '.html,.htm',
      multi: false,
      options: [
        {
          id: 'sourceType', label: 'Nguồn nội dung', type: 'select', options: [
            { value: 'file', text: 'Tải lên file HTML (.html)' },
            { value: 'url', text: 'Nhập đường dẫn trang web (URL)' }
          ], default: 'file'
        },
        { id: 'urlInput', label: 'URL trang web', type: 'text', placeholder: 'https://example.com', default: '' },
        { id: 'landscape', label: 'Khổ ngang', type: 'checkbox', default: false }
      ]
    },
    {
      id: 'pdf-to-epub',
      name: 'PDF → EPUB',
      group: 'group5',
      badge: 'browser',
      desc: 'Trích chữ, tự nhận tiêu đề chương theo cỡ chữ lớn và đóng gói EPUB 3 hoàn chỉnh.',
      accept: '.pdf',
      multi: false,
      options: []
    },
    {
      id: 'epub-to-pdf',
      name: 'EPUB → PDF',
      group: 'group5',
      badge: 'browser',
      desc: 'Đọc file sách điện tử EPUB, ghép các chương theo thứ tự spine và kết xuất PDF.',
      accept: '.epub',
      multi: false,
      options: []
    },
    {
      id: 'mobi-info',
      name: 'Hướng dẫn cho Kindle / MOBI',
      group: 'group5',
      badge: 'browser',
      desc: 'Hướng dẫn gửi file sách sang thiết bị Kindle qua tính năng chính thức Send to Kindle.',
      accept: '',
      multi: false,
      options: []
    }
  ];

  // Trạng thái ứng dụng
  let currentTool = null;
  let selectedFiles = [];
  let serverHealth = { status: 'checking', message: 'Đang kết nối...' };

  // DOM Elements
  const toolsGridEl = document.getElementById('tools-grid');
  const searchInputEl = document.getElementById('search-input');
  const categoryTabsEl = document.getElementById('category-tabs');
  const themeToggleBtn = document.getElementById('theme-toggle-btn');
  const serverPillEl = document.getElementById('server-status-pill');
  const toolViewEl = document.getElementById('tool-view');
  const homeViewEl = document.getElementById('home-view');

  // Khởi tạo theme
  function initTheme() {
    const saved = localStorage.getItem('doc_conv_theme') || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
    document.documentElement.setAttribute('data-theme', saved);
    updateThemeIcon(saved);
  }

  function toggleTheme() {
    const current = document.documentElement.getAttribute('data-theme') || 'light';
    const next = current === 'dark' ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('doc_conv_theme', next);
    updateThemeIcon(next);
  }

  function updateThemeIcon(theme) {
    if (themeToggleBtn) {
      themeToggleBtn.innerHTML = theme === 'dark' ? '☀️ Sáng' : '🌙 Tối';
    }
  }

  // Kiểm tra sức khỏe Server
  async function checkServerHealth() {
    try {
      const res = await fetch('/api/health');
      if (res.ok) {
        const data = await res.json();
        serverHealth = data.gotenberg || { status: 'not_configured', message: 'Chưa cấu hình Gotenberg' };
      } else {
        serverHealth = { status: 'error', message: 'API health trả về lỗi HTTP' };
      }
    } catch (e) {
      serverHealth = { status: 'unreachable', message: 'Không thể kết nối API máy chủ' };
    }
    renderServerStatus();
  }

  function renderServerStatus() {
    if (!serverPillEl) return;
    const dot = serverPillEl.querySelector('.status-dot');
    const text = serverPillEl.querySelector('.status-text');

    dot.className = `status-dot ${serverHealth.status}`;
    if (serverHealth.status === 'ok') {
      text.textContent = 'Server: Sẵn sàng';
    } else if (serverHealth.status === 'sleeping') {
      text.textContent = 'Server: Đang ngủ (chờ 1p)';
    } else if (serverHealth.status === 'not_configured') {
      text.textContent = 'Server: Chế độ duyệt web';
    } else {
      text.textContent = 'Server: Gián đoạn';
    }
    serverPillEl.title = serverHealth.message || '';
  }

  // Render danh sách công cụ
  function renderTools(filterGroup = 'all', searchQuery = '') {
    toolsGridEl.innerHTML = '';
    const query = searchQuery.trim().toLowerCase();

    const filtered = TOOLS.filter(t => {
      const matchGroup = filterGroup === 'all' || t.group === filterGroup;
      const matchSearch = !query || t.name.toLowerCase().includes(query) || t.desc.toLowerCase().includes(query);
      return matchGroup && matchSearch;
    });

    if (!filtered.length) {
      toolsGridEl.innerHTML = `
        <div style="grid-column: 1/-1; text-align: center; padding: 3rem; color: var(--text-muted);">
          <p>Không tìm thấy công cụ phù hợp với từ khóa "${searchQuery}".</p>
        </div>
      `;
      return;
    }

    filtered.forEach(tool => {
      const card = document.createElement('div');
      card.className = 'tool-card';
      card.innerHTML = `
        <div class="tool-card-header">
          <div class="tool-title">${tool.name}</div>
          <span class="badge ${tool.badge === 'browser' ? 'badge-browser' : 'badge-server'}">
            ${tool.badge === 'browser' ? '💻 [Trình duyệt]' : '☁️ [Server]'}
          </span>
        </div>
        <div class="tool-desc">${tool.desc}</div>
        <div class="tool-meta">
          <span>${tool.badge === 'browser' ? 'Xử lý trên máy bạn (bảo mật tuyệt đối)' : 'Xử lý qua server / có dự phòng'}</span>
          <span>Bấm để mở &rarr;</span>
        </div>
      `;
      card.addEventListener('click', () => openTool(tool.id));
      toolsGridEl.appendChild(card);
    });
  }

  // Mở giao diện công cụ
  function openTool(toolId) {
    const tool = TOOLS.find(t => t.id === toolId);
    if (!tool) return;
    currentTool = tool;
    selectedFiles = [];

    homeViewEl.style.display = 'none';
    toolViewEl.style.display = 'block';
    window.scrollTo({ top: 0, behavior: 'smooth' });

    renderToolWorkspace();
  }

  function closeTool() {
    currentTool = null;
    selectedFiles = [];
    toolViewEl.style.display = 'none';
    homeViewEl.style.display = 'block';
  }

  // Render workspace của từng công cụ
  function renderToolWorkspace() {
    const tool = currentTool;
    const isBrowser = tool.badge === 'browser';
    const isServerUnconfigured = !isBrowser && serverHealth.status === 'not_configured';

    let warningHtml = '';
    if (isServerUnconfigured) {
      warningHtml = `
        <div class="alert alert-warning">
          <span>⚠️ <strong>Lưu ý:</strong> Server Gotenberg chưa được cấu hình biến môi trường GOTENBERG_URL. Công cụ sẽ tự động chạy phương án dự phòng trên trình duyệt (chất lượng cơ bản, bố cục có thể lệch).</span>
        </div>
      `;
    } else if (!isBrowser && serverHealth.status === 'sleeping') {
      warningHtml = `
        <div class="alert alert-info">
          <span>ℹ️ <strong>Thông báo:</strong> Server chuyển đổi (Hugging Face Space) đang khởi động từ chế độ ngủ, quá trình chuyển đổi lần đầu có thể mất khoảng 1 phút. Vui lòng kiên nhẫn.</span>
        </div>
      `;
    }

    if (tool.id === 'mobi-info') {
      renderMobiGuideWorkspace();
      return;
    }

    // Các tùy chọn của công cụ
    let optionsHtml = '';
    if (tool.options && tool.options.length > 0) {
      optionsHtml = `
        <div class="options-panel">
          <div class="options-title">⚙️ Tùy chọn chuyển đổi</div>
          ${tool.options.map(opt => {
        if (opt.type === 'select') {
          return `
                <div class="option-row">
                  <label class="option-label" for="opt-${opt.id}">${opt.label}:</label>
                  <select id="opt-${opt.id}" class="option-input">
                    ${opt.options.map(o => `<option value="${o.value}" ${o.value === opt.default ? 'selected' : ''}>${o.text}</option>`).join('')}
                  </select>
                </div>
              `;
        } else if (opt.type === 'checkbox') {
          return `
                <div class="option-row">
                  <label class="option-label" for="opt-${opt.id}">${opt.label}:</label>
                  <input type="checkbox" id="opt-${opt.id}" ${opt.default ? 'checked' : ''} style="cursor: pointer; width: 18px; height: 18px;">
                </div>
              `;
        } else if (opt.type === 'text') {
          return `
                <div class="option-row">
                  <label class="option-label" for="opt-${opt.id}">${opt.label}:</label>
                  <input type="text" id="opt-${opt.id}" class="option-input" placeholder="${opt.placeholder || ''}" value="${opt.default || ''}" style="flex: 1; max-width: 350px;">
                </div>
              `;
        }
        return '';
      }).join('')}
        </div>
      `;
    }

    toolViewEl.innerHTML = `
      <div class="tool-view-header">
        <div class="tool-view-title-wrap">
          <button class="btn-back" id="btn-back-home">&larr; Danh sách công cụ</button>
          <h2 style="font-size: 1.4rem; font-weight: 700;">${tool.name}</h2>
          <span class="badge ${isBrowser ? 'badge-browser' : 'badge-server'}">
            ${isBrowser ? '💻 Xử lý trên máy bạn' : '☁️ Gửi lên server'}
          </span>
        </div>
      </div>

      ${warningHtml}

      <div style="color: var(--text-secondary); margin-bottom: 1.25rem; font-size: 0.95rem;">
        ${tool.desc}
      </div>

      <!-- Dropzone -->
      <div class="dropzone" id="file-dropzone">
        <input type="file" id="file-input" ${tool.multi ? 'multiple' : ''} accept="${tool.accept}" style="display: none;">
        <div class="dropzone-icon">📁</div>
        <div class="dropzone-text">Kéo thả file vào đây hoặc <span style="color: var(--accent-primary); text-decoration: underline;">bấm để chọn file</span></div>
        <div class="dropzone-hint">Định dạng hỗ trợ: ${tool.accept || 'Mọi định dạng'} ${tool.multi ? '(Cho phép chọn nhiều file)' : '(Chọn 1 file)'}</div>
      </div>

      <!-- File list -->
      <div class="selected-files-list" id="selected-files-list"></div>

      <!-- Image previews if tool is image-to-pdf -->
      <div id="image-thumbs-wrap"></div>

      <!-- Options -->
      ${optionsHtml}

      <!-- Actions -->
      <div style="margin-top: 1.25rem;">
        <button class="btn-primary" id="btn-start-convert" disabled>
          ⚡ Bắt đầu chuyển đổi
        </button>
      </div>

      <!-- Progress -->
      <div class="progress-wrap" id="progress-wrap" style="display: none;">
        <div class="progress-bar-bg">
          <div class="progress-bar-fill" id="progress-bar-fill"></div>
        </div>
        <div class="progress-status" id="progress-status">Đang chuẩn bị...</div>
      </div>

      <!-- Result Card -->
      <div id="result-container"></div>
    `;

    document.getElementById('btn-back-home').addEventListener('click', closeTool);

    setupDropzone();
    document.getElementById('btn-start-convert').addEventListener('click', handleConversion);

    // Xử lý riêng khi công cụ là HTML/URL
    if (tool.id === 'html-to-pdf') {
      const srcTypeEl = document.getElementById('opt-sourceType');
      const urlInputEl = document.getElementById('opt-urlInput');
      const dropzoneEl = document.getElementById('file-dropzone');
      const convertBtn = document.getElementById('btn-start-convert');

      function syncHtmlSourceType() {
        const isUrl = srcTypeEl && srcTypeEl.value === 'url';
        if (dropzoneEl) dropzoneEl.style.display = isUrl ? 'none' : 'block';
        if (urlInputEl) {
          const row = urlInputEl.closest('.option-row');
          if (row) row.style.display = isUrl ? 'flex' : 'none';
        }
        if (isUrl) {
          convertBtn.disabled = !urlInputEl || !urlInputEl.value.trim();
        } else {
          convertBtn.disabled = selectedFiles.length === 0;
        }
      }

      if (srcTypeEl) srcTypeEl.addEventListener('change', syncHtmlSourceType);
      if (urlInputEl) urlInputEl.addEventListener('input', syncHtmlSourceType);
      syncHtmlSourceType();
    }
  }

  // Giao diện hướng dẫn riêng cho MOBI (Kindle)
  function renderMobiGuideWorkspace() {
    toolViewEl.innerHTML = `
      <div class="tool-view-header">
        <div class="tool-view-title-wrap">
          <button class="btn-back" id="btn-back-home">&larr; Danh sách công cụ</button>
          <h2 style="font-size: 1.4rem; font-weight: 700;">Hướng dẫn cho Kindle / MOBI</h2>
        </div>
      </div>

      <div class="alert alert-info">
        <span>📖 <strong>Lưu ý quan trọng từ Amazon:</strong> Amazon Kindle đã chính thức ngừng hỗ trợ file .MOBI cũ và chuyển sang chuẩn <strong>EPUB</strong> cho tính năng "Send to Kindle".</span>
      </div>

      <div style="line-height: 1.8; color: var(--text-primary); margin-top: 1rem;">
        <h3 style="margin-bottom: 0.5rem;">Cách đọc sách tốt nhất trên Kindle hiện nay:</h3>
        <ol style="padding-left: 1.5rem; margin-bottom: 1.5rem;">
          <li>Sử dụng công cụ <strong>PDF → EPUB</strong> trên web app này để xuất file sách chuẩn EPUB.</li>
          <li>Truy cập trang gửi sách chính thức của Amazon: <a href="https://www.amazon.com/sendtokindle" target="_blank" style="color: var(--accent-primary);">amazon.com/sendtokindle</a>.</li>
          <li>Kéo thả file <code>.epub</code> đã tạo và bấm <strong>Send</strong>. Sách sẽ tự động đồng bộ về máy đọc sách Kindle của bạn với đầy đủ mục lục, đổi cỡ chữ và font chữ tùy chỉnh!</li>
        </ol>

        <button class="btn-primary" id="btn-goto-epub">
          👉 Mở công cụ PDF → EPUB ngay
        </button>
      </div>
    `;

    document.getElementById('btn-back-home').addEventListener('click', closeTool);
    document.getElementById('btn-goto-epub').addEventListener('click', () => openTool('pdf-to-epub'));
  }

  // Cấu hình kéo thả file
  function setupDropzone() {
    const dropzone = document.getElementById('file-dropzone');
    const input = document.getElementById('file-input');

    dropzone.addEventListener('click', () => input.click());

    ['dragenter', 'dragover'].forEach(eventName => {
      dropzone.addEventListener(eventName, e => {
        e.preventDefault();
        dropzone.classList.add('dragover');
      });
    });

    ['dragleave', 'drop'].forEach(eventName => {
      dropzone.addEventListener(eventName, e => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
      });
    });

    dropzone.addEventListener('drop', e => {
      const files = Array.from(e.dataTransfer.files);
      handleFilesAdded(files);
    });

    input.addEventListener('change', () => {
      const files = Array.from(input.files);
      handleFilesAdded(files);
      input.value = '';
    });
  }

  function handleFilesAdded(files) {
    if (!files.length) return;
    const tool = currentTool;

    if (tool.multi) {
      selectedFiles.push(...files);
    } else {
      selectedFiles = [files[0]];
    }

    renderSelectedFiles();
  }

  function renderSelectedFiles() {
    const listEl = document.getElementById('selected-files-list');
    const convertBtn = document.getElementById('btn-start-convert');
    const thumbsWrap = document.getElementById('image-thumbs-wrap');

    listEl.innerHTML = '';
    thumbsWrap.innerHTML = '';

    if (!selectedFiles.length) {
      convertBtn.disabled = true;
      return;
    }
    convertBtn.disabled = false;

    selectedFiles.forEach((file, index) => {
      const item = document.createElement('div');
      item.className = 'file-item';
      const sizeStr = file.size > 1048576 ? (file.size / 1048576).toFixed(2) + ' MB' : (file.size / 1024).toFixed(1) + ' KB';
      item.innerHTML = `
        <div class="file-item-info">
          <span>📄</span>
          <strong>${file.name}</strong>
          <span style="color: var(--text-muted);">(${sizeStr})</span>
        </div>
        <button class="file-item-remove" title="Xóa file">&times;</button>
      `;

      item.querySelector('.file-item-remove').addEventListener('click', () => {
        selectedFiles.splice(index, 1);
        renderSelectedFiles();
      });
      listEl.appendChild(item);
    });

    // Nếu công cụ là image-to-pdf, hiển thị thumbnail để sắp xếp và xoay
    if (currentTool.id === 'image-to-pdf' && selectedFiles.length > 0) {
      renderImageThumbnails(thumbsWrap);
    }
  }

  function renderImageThumbnails(container) {
    container.innerHTML = '<div class="options-title" style="margin-top: 1rem;">🖼️ Xem trước & Sắp xếp thứ tự ảnh (Kéo thả hoặc dùng mũi tên)</div>';
    const grid = document.createElement('div');
    grid.className = 'image-thumbs-grid';

    selectedFiles.forEach((f, idx) => {
      if (!f.rotateDeg) f.rotateDeg = 0;
      const card = document.createElement('div');
      card.className = 'thumb-card';
      const imgUrl = URL.createObjectURL(f);

      card.innerHTML = `
        <img src="${imgUrl}" style="transform: rotate(${f.rotateDeg}deg);">
        <div style="font-size: 0.72rem; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${f.name}</div>
        <div class="thumb-tools">
          <button class="thumb-btn btn-up" title="Lên trên" ${idx === 0 ? 'disabled' : ''}>&uarr;</button>
          <button class="thumb-btn btn-down" title="Xuống dưới" ${idx === selectedFiles.length - 1 ? 'disabled' : ''}>&darr;</button>
          <button class="thumb-btn btn-rot" title="Xoay 90°">↻</button>
        </div>
      `;

      card.querySelector('.btn-up').addEventListener('click', () => {
        const temp = selectedFiles[idx - 1];
        selectedFiles[idx - 1] = selectedFiles[idx];
        selectedFiles[idx] = temp;
        renderSelectedFiles();
      });

      card.querySelector('.btn-down').addEventListener('click', () => {
        const temp = selectedFiles[idx + 1];
        selectedFiles[idx + 1] = selectedFiles[idx];
        selectedFiles[idx] = temp;
        renderSelectedFiles();
      });

      card.querySelector('.btn-rot').addEventListener('click', () => {
        f.rotateDeg = (f.rotateDeg + 90) % 360;
        card.querySelector('img').style.transform = `rotate(${f.rotateDeg}deg)`;
      });

      grid.appendChild(card);
    });

    container.appendChild(grid);
  }

  // Thu thập các options từ UI
  function collectOptions() {
    const opts = {};
    if (!currentTool || !currentTool.options) return opts;
    currentTool.options.forEach(opt => {
      const el = document.getElementById(`opt-${opt.id}`);
      if (!el) return;
      if (opt.type === 'checkbox') {
        opts[opt.id] = el.checked;
      } else {
        opts[opt.id] = el.value;
      }
    });
    return opts;
  }

  // Cập nhật thanh tiến trình
  function updateProgress(percent, text) {
    const wrap = document.getElementById('progress-wrap');
    const fill = document.getElementById('progress-bar-fill');
    const status = document.getElementById('progress-status');
    if (!wrap) return;

    wrap.style.display = 'block';
    fill.style.width = `${Math.min(100, Math.max(0, percent))}%`;
    status.textContent = text || '';
  }

  // Xử lý chuyển đổi
  async function handleConversion() {
    const tool = currentTool;
    const options = collectOptions();
    if (!selectedFiles.length && !(tool.id === 'html-to-pdf' && options.sourceType === 'url')) {
      alert('Vui lòng chọn file trước khi chuyển đổi.');
      return;
    }
    const convertBtn = document.getElementById('btn-start-convert');
    const resultContainer = document.getElementById('result-container');
    resultContainer.innerHTML = '';
    convertBtn.disabled = true;

    try {
      updateProgress(5, 'Đang chuẩn bị tiến trình...');

      let result;

      // Điều hướng theo công cụ
      switch (tool.id) {
        // Nhóm 1
        case 'pdf-to-excel': {
          // Xử lý nhiều file tuần tự
          const outputFiles = [];
          for (let i = 0; i < selectedFiles.length; i++) {
            const f = selectedFiles[i];
            updateProgress(
              Math.floor((i / selectedFiles.length) * 90),
              `[${i + 1}/${selectedFiles.length}] Đang xử lý: ${f.name}...`
            );
            const res = await window.Group1Tools.convertPdfToExcel(f, options, (pct, msg) => {
              updateProgress(Math.floor((i / selectedFiles.length) * 90 + (pct * 0.9 / selectedFiles.length)), `[${i + 1}/${selectedFiles.length}] ${msg}`);
            });
            outputFiles.push(res);
          }
          if (outputFiles.length === 1) {
            result = outputFiles[0];
          } else {
            // Nén nhiều file thành ZIP
            await window.CDNLoader.require(['jszip']);
            const zip = new window.JSZip();
            outputFiles.forEach(o => zip.file(o.filename, o.blob));
            const zipBlob = await zip.generateAsync({ type: 'blob' });
            result = { blob: zipBlob, filename: 'excel_converted_files.zip', server: true };
          }
          break;
        }

        case 'word-to-pdf':
          result = await window.Group1Tools.convertWordToPdf(selectedFiles[0], options, updateProgress);
          break;

        case 'excel-to-pdf':
          result = await window.Group1Tools.convertExcelToPdf(selectedFiles[0], options, updateProgress);
          break;

        case 'ppt-to-pdf':
          result = await window.Group1Tools.convertPptToPdf(selectedFiles[0], options, updateProgress);
          break;

        case 'pdf-to-word':
          result = await window.Group1Tools.convertPdfToWord(selectedFiles[0], options, updateProgress);
          break;

        case 'pdf-to-ppt':
          result = await window.Group1Tools.convertPdfToPpt(selectedFiles[0], options, updateProgress);
          break;

        // Nhóm 2: Office nội bộ
        case 'word-to-excel':
          result = await window.Group2Tools.convertWordToExcel(selectedFiles[0], options, updateProgress);
          break;

        case 'excel-to-word':
          result = await window.Group2Tools.convertExcelToWord(selectedFiles[0], options, updateProgress);
          break;

        case 'ppt-to-word':
          result = await window.Group2Tools.convertPptToWord(selectedFiles[0], options, updateProgress);
          break;

        // Nhóm 3
        case 'ocr':
          result = await window.Group3Tools.performOcr(selectedFiles[0], options, updateProgress);
          break;

        case 'pdf-to-image':
          result = await window.Group3Tools.convertPdfToImages(selectedFiles[0], options, updateProgress);
          break;

        case 'image-to-pdf':
          result = await window.Group3Tools.convertImagesToPdf(selectedFiles, options, updateProgress);
          break;

        // Nhóm 4
        case 'excel-to-json':
          result = await window.Group4Tools.convertExcelCsvToJson(selectedFiles[0], options, updateProgress);
          break;

        case 'json-to-excel':
          result = await window.Group4Tools.convertJsonToExcelCsv(selectedFiles[0], options, updateProgress);
          break;

        case 'excel-csv-swap':
          result = await window.Group4Tools.convertExcelCsvSwap(selectedFiles[0], options, updateProgress);
          break;

        case 'xml-to-excel':
          result = await window.Group4Tools.convertXmlToExcel(selectedFiles, options, updateProgress);
          break;

        // Nhóm 5: Sách điện tử & Web
        case 'html-to-pdf':
          result = await window.Group5Tools.convertHtmlToPdf(selectedFiles[0], options, updateProgress);
          break;

        case 'pdf-to-epub':
          result = await window.Group5Tools.convertPdfToEpub(selectedFiles[0], options, updateProgress);
          break;

        case 'epub-to-pdf':
          result = await window.Group5Tools.convertEpubToPdf(selectedFiles[0], options, updateProgress);
          break;

        default:
          throw new Error(`Công cụ chưa được kích hoạt trong đợt này: ${tool.id}`);
      }

      updateProgress(100, 'Chuyển đổi hoàn tất thành công!');
      renderResult(result);

    } catch (err) {
      console.error(err);
      updateProgress(0, 'Đã xảy ra lỗi!');
      resultContainer.innerHTML = `
        <div class="alert alert-error" style="margin-top: 1rem;">
          <div>
            <strong>❌ Lỗi chuyển đổi:</strong><br>
            ${escapeHtml(err.message || 'Không thể hoàn thành tác vụ. Vui lòng thử lại.')}
          </div>
        </div>
      `;
    } finally {
      convertBtn.disabled = false;
    }
  }

  function renderResult(result) {
    const container = document.getElementById('result-container');
    if (!container || !result) return;

    let warningHtml = '';
    if (result.warning) {
      warningHtml = `
        <div class="alert alert-warning" style="margin-bottom: 0.75rem;">
          ⚠️ ${escapeHtml(result.warning)}
        </div>
      `;
    }

    let extraPreview = '';
    if (result.text && result.text.length < 5000) {
      extraPreview = `
        <div style="margin: 0.75rem 0;">
          <textarea readonly style="width: 100%; height: 140px; font-family: monospace; font-size: 0.8rem; background: var(--bg-surface); color: var(--text-primary); border: 1px solid var(--border-color); border-radius: 4px; padding: 0.5rem;">${escapeHtml(result.text)}</textarea>
        </div>
      `;
    }

    container.innerHTML = `
      <div class="result-card">
        ${warningHtml}
        <div class="result-header">
          <span>🎉 Kết quả chuyển đổi sẵn sàng!</span>
        </div>
        <div style="font-size: 0.9rem; margin-bottom: 0.75rem;">
          File đã tạo: <strong>${escapeHtml(result.filename)}</strong>
        </div>
        ${extraPreview}
        <div class="result-actions">
          <button class="btn-primary" id="btn-download-result">
            💾 Tải về (${escapeHtml(result.filename)})
          </button>
          ${result.text ? `<button class="btn-secondary" id="btn-copy-result">📋 Sao chép nội dung</button>` : ''}
        </div>
      </div>
    `;

    document.getElementById('btn-download-result').addEventListener('click', () => {
      downloadBlob(result.blob, result.filename);
    });

    if (result.text) {
      const copyBtn = document.getElementById('btn-copy-result');
      if (copyBtn) {
        copyBtn.addEventListener('click', () => {
          navigator.clipboard.writeText(result.text).then(() => {
            copyBtn.textContent = '✅ Đã sao chép!';
            setTimeout(() => { copyBtn.textContent = '📋 Sao chép nội dung'; }, 2000);
          });
        });
      }
    }
  }

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

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  // Khởi động sự kiện
  function initEvents() {
    initTheme();
    themeToggleBtn.addEventListener('click', toggleTheme);

    searchInputEl.addEventListener('input', e => {
      const activeTab = categoryTabsEl.querySelector('.category-tab.active');
      const group = activeTab ? activeTab.getAttribute('data-group') : 'all';
      renderTools(group, e.target.value);
    });

    categoryTabsEl.querySelectorAll('.category-tab').forEach(tab => {
      tab.addEventListener('click', () => {
        categoryTabsEl.querySelectorAll('.category-tab').forEach(t => t.classList.remove('active'));
        tab.classList.add('active');
        const group = tab.getAttribute('data-group');
        renderTools(group, searchInputEl.value);
      });
    });

    serverPillEl.addEventListener('click', () => {
      checkServerHealth();
    });

    // Check health ngay khi tải trang và lặp lại mỗi 30s
    checkServerHealth();
    setInterval(checkServerHealth, 30000);

    renderTools('all', '');
  }

  window.addEventListener('DOMContentLoaded', initEvents);
})();