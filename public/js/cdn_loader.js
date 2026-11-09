/**
 * cdn_loader.js
 * Quản lý nạp động thư viện từ CDN với version cố định, chỉ nạp khi mở công cụ cần đến.
 */

(function (window) {
  'use strict';

  const CDN_DEFS = {
    pdfjs: {
      url: 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',
      check: () => window.pdfjsLib !== undefined,
      init: () => {
        if (window.pdfjsLib && !window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
          window.pdfjsLib.GlobalWorkerOptions.workerSrc =
            'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
        }
      }
    },
    pdfLib: {
      url: 'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js',
      check: () => window.PDFLib !== undefined
    },
    xlsx: {
      url: 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js',
      check: () => window.XLSX !== undefined
    },
    docx: {
      url: 'https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js',
      check: () => window.docx !== undefined
    },
    pptxgen: {
      url: 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js',
      check: () => window.PptxGenJS !== undefined
    },
    jszip: {
      url: 'https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js',
      check: () => window.JSZip !== undefined
    },
    tesseract: {
      url: 'https://cdnjs.cloudflare.com/ajax/libs/tesseract.js/5.1.0/tesseract.min.js',
      check: () => window.Tesseract !== undefined
    },
    html2pdf: {
      url: 'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js',
      check: () => window.html2pdf !== undefined
    },
    mammoth: {
      url: 'https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js',
      check: () => window.mammoth !== undefined
    }
  };

  const loadingPromises = {};

  function loadLibrary(name) {
    const def = CDN_DEFS[name];
    if (!def) {
      return Promise.reject(new Error(`Thư viện chưa đăng ký: ${name}`));
    }
    if (def.check()) {
      if (def.init) def.init();
      return Promise.resolve();
    }
    if (loadingPromises[name]) {
      return loadingPromises[name];
    }

    loadingPromises[name] = new Promise((resolve, reject) => {
      const script = document.createElement('script');
      script.src = def.url;
      script.async = true;
      script.onload = () => {
        if (def.init) def.init();
        resolve();
      };
      script.onerror = () => {
        delete loadingPromises[name];
        reject(new Error(`Không thể nạp thư viện ${name} từ CDN: ${def.url}`));
      };
      document.head.appendChild(script);
    });

    return loadingPromises[name];
  }

  function requireLibs(libNames) {
    return Promise.all(libNames.map(name => loadLibrary(name)));
  }

  window.CDNLoader = {
    load: loadLibrary,
    require: requireLibs
  };
})(typeof window !== 'undefined' ? window : this);

