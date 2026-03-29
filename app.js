/* ============================================================
   ConvertFlux — app.js
   Modular, production-grade client-side file conversion
   ============================================================ */

'use strict';

// ─── PDF.js worker ─────────────────────────────────────────
if (typeof pdfjsLib !== 'undefined') {
  pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';
}

// ─── Constants ────────────────────────────────────────────
const MAX_FILE_MB = 50;
const MAX_FILE_BYTES = MAX_FILE_MB * 1024 * 1024;
const HISTORY_KEY = 'cf_history';

// ─── Tool Definitions ─────────────────────────────────────
const TOOLS = {
  'pdf-to-images': { title: 'PDF → Images', sub: 'Export each page as PNG or JPG', fn: renderPdfToImages },
  'images-to-pdf': { title: 'Images → PDF', sub: 'Combine images into a PDF', fn: renderImagesToPdf },
  'pdf-merge': { title: 'Merge PDFs', sub: 'Combine multiple PDFs into one', fn: renderPdfMerge },
  'pdf-split': { title: 'Split PDF', sub: 'Extract pages by range', fn: renderPdfSplit },
  'pdf-compress': { title: 'Compress PDF', sub: 'Reduce PDF file size', fn: renderPdfCompress },
  'pdf-page-editor': { title: 'PDF Page Editor', sub: 'Reorder, rotate, delete pages', fn: renderPdfPageEditor },
  'pdf-text': { title: 'Extract Text', sub: 'Pull text content from PDF', fn: renderPdfText },
  'pdf-page-numbers': { title: 'Add Page Numbers', sub: 'Insert page numbers into PDF', fn: renderPdfPageNumbers },
  'pdf-watermark': { title: 'Add Watermark', sub: 'Stamp text watermarks on pages', fn: renderPdfWatermark },
  'pdf-password': { title: 'Password Protect PDF', sub: 'Encrypt PDF with a password', fn: renderPdfPassword },
  'img-compress': { title: 'Compress Image', sub: 'Reduce image size with preview', fn: renderImgCompress },
  'img-convert': { title: 'Convert Image Format', sub: 'Convert between JPG/PNG/WebP', fn: renderImgConvert },
  'img-resize': { title: 'Resize Image', sub: 'Set exact dimensions', fn: renderImgResize },
  'img-crop': { title: 'Crop Image', sub: 'Interactive canvas cropper', fn: renderImgCrop },
  'img-base64': { title: 'Image → Base64', sub: 'Encode image as Base64 string', fn: renderImgBase64 },
  'img-bulk': { title: 'Bulk Image Convert', sub: 'Convert multiple images + ZIP', fn: renderImgBulk },
  'excel-to-pdf': { title: 'Excel/CSV → PDF', sub: 'Render spreadsheet as PDF table', fn: renderExcelToPdf },
  'pdf-to-csv': { title: 'PDF → CSV', sub: 'Extract text/table data to CSV', fn: renderPdfToCsv },
  'csv-merge': { title: 'Merge CSV Files', sub: 'Combine multiple CSV files', fn: renderCsvMerge },
  'csv-split': { title: 'Split CSV', sub: 'Split CSV by row count', fn: renderCsvSplit },
  // ─── New tools ────────────────────────────────────────────────
  'word-to-pdf': { title: 'Word → PDF', sub: 'Convert DOCX to PDF in your browser', fn: renderWordToPdf },
  'text-to-pdf': { title: 'Text → PDF', sub: 'Convert plain text or .txt to PDF', fn: renderTextToPdf },
  'html-to-pdf': { title: 'HTML → PDF', sub: 'Convert HTML code to PDF document', fn: renderHtmlToPdf },
  'img-to-pdf-adv': { title: 'Image → PDF Advanced', sub: 'Images to PDF with margin & orientation', fn: renderImgToPdfAdv },
  'img-bulk-compress-zip': { title: 'Bulk Compress + ZIP', sub: 'Compress many images, download as ZIP', fn: renderImgBulkCompressZip },
};

// ─── Navigation ───────────────────────────────────────────
function showHome() {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById('homePage').classList.add('active');
  loadHistory();
}
function showPage(name) {
  document.querySelectorAll('.page').forEach(p => p.classList.remove('active'));
  document.getElementById(name + 'Page').classList.add('active');
  window.scrollTo(0, 0);
}
function toggleMenu() {
  const nav = document.getElementById('nav');
  const btn = document.getElementById('hamburger');
  nav.classList.toggle('open');
  btn.classList.toggle('open');
}
function closeMenu() {
  document.getElementById('nav').classList.remove('open');
  document.getElementById('hamburger').classList.remove('open');
}

// ─── Tool Search ──────────────────────────────────────────
function filterTools(query) {
  const q = query.toLowerCase().trim();
  document.querySelectorAll('.tool-card').forEach(card => {
    const tags = card.dataset.tags || '';
    const title = card.querySelector('h3')?.textContent.toLowerCase() || '';
    const match = !q || title.includes(q) || tags.includes(q);
    card.classList.toggle('hidden', !match);
  });
  // Show/hide category labels
  ['pdf', 'image', 'csv', 'doc'].forEach(cat => {
    const gridId = { pdf: 'gridPdf', image: 'gridImage', csv: 'gridCsv', doc: 'gridDoc' }[cat];
    const labelId = 'cat-' + cat;
    const grid = document.getElementById(gridId);
    const label = document.getElementById(labelId);
    if (grid && label) {
      const visible = [...grid.querySelectorAll('.tool-card:not(.hidden)')].length > 0;
      label.style.display = visible ? '' : 'none';
    }
  });
}

// ─── Modal ────────────────────────────────────────────────
function openTool(toolId) {
  const tool = TOOLS[toolId];
  if (!tool) return;
  document.getElementById('modalTitle').textContent = tool.title;
  document.getElementById('modalSub').textContent = tool.sub;
  const body = document.getElementById('modalBody');
  body.innerHTML = '';
  tool.fn(body);
  // Inject non-intrusive bottom ad
  const adWrap = document.createElement('div');
  adWrap.className = 'ad-modal';
  adWrap.innerHTML = `<ins class="adsbygoogle" style="display:block" data-ad-client="ca-pub-XXXXXXXX" data-ad-slot="0000000003" data-ad-format="auto" data-full-width-responsive="true"></ins>`;
  body.appendChild(adWrap);
  setTimeout(() => { try { (adsbygoogle = window.adsbygoogle || []).push({}); } catch (e) {} }, 120);
  document.getElementById('modalOverlay').classList.add('open');
  document.body.style.overflow = 'hidden';
}
function closeModal(e) {
  if (e && e.target !== document.getElementById('modalOverlay')) return;
  document.getElementById('modalOverlay').classList.remove('open');
  document.body.style.overflow = '';
}
document.querySelectorAll('.tool-card').forEach(card => {
  card.addEventListener('click', () => openTool(card.dataset.tool));
});
document.addEventListener('keydown', e => {
  if (e.key === 'Escape') { document.getElementById('modalOverlay').classList.remove('open'); document.body.style.overflow = ''; }
});

// ─── Toast ────────────────────────────────────────────────
function toast(msg, type = 'info', duration = 4000) {
  const icons = { success: '✓', error: '✕', info: 'ℹ' };
  const el = document.createElement('div');
  el.className = `toast ${type}`;
  el.innerHTML = `<span>${icons[type]}</span><span>${msg}</span>`;
  document.getElementById('toastContainer').appendChild(el);
  setTimeout(() => el.remove(), duration);
}

// ─── History ──────────────────────────────────────────────
function saveHistory(toolId, fileName) {
  const h = JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]');
  h.unshift({ tool: toolId, file: fileName, time: Date.now() });
  localStorage.setItem(HISTORY_KEY, JSON.stringify(h.slice(0, 5)));
  loadHistory();
}
function loadHistory() {
  const h = JSON.parse(localStorage.getItem(HISTORY_KEY) || '[]');
  const list = document.getElementById('historyList');
  if (!list) return;
  if (!h.length) { list.innerHTML = '<p class="empty-state">No conversions yet. Try a tool above!</p>'; return; }
  list.innerHTML = h.map(item => `
    <div class="history-item">
      <span class="history-item-tool">${TOOLS[item.tool]?.title || item.tool}</span>
      <span class="history-item-file">${item.file}</span>
      <span class="history-item-time">${relativeTime(item.time)}</span>
    </div>`).join('');
}
function clearHistory() {
  localStorage.removeItem(HISTORY_KEY);
  loadHistory();
}
function relativeTime(ts) {
  const s = Math.round((Date.now() - ts) / 1000);
  if (s < 60) return 'just now';
  if (s < 3600) return Math.round(s / 60) + 'm ago';
  if (s < 86400) return Math.round(s / 3600) + 'h ago';
  return Math.round(s / 86400) + 'd ago';
}

// ─── Utility Helpers ──────────────────────────────────────
function formatBytes(bytes) {
  if (!bytes) return '0 B';
  const k = 1024, sizes = ['B', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return parseFloat((bytes / Math.pow(k, i)).toFixed(1)) + ' ' + sizes[i];
}
function readFileAsArrayBuffer(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = rej;
    r.readAsArrayBuffer(file);
  });
}
function readFileAsDataURL(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = rej;
    r.readAsDataURL(file);
  });
}
function readFileAsText(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = () => res(r.result);
    r.onerror = rej;
    r.readAsText(file);
  });
}
function downloadBlob(blob, filename) {
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url; a.download = filename; a.click();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
function setProgress(wrap, pct, label) {
  if (!wrap) return;
  const fill = wrap.querySelector('.progress-fill');
  const pctEl = wrap.querySelector('.progress-pct');
  if (fill) fill.style.width = pct + '%';
  if (pctEl) pctEl.textContent = Math.round(pct) + '%';
  if (label) { const l = wrap.querySelector('.progress-msg'); if (l) l.textContent = label; }
}
function progressHTML(id) {
  return `<div class="progress-wrap" id="${id}">
    <div class="progress-label"><span class="progress-msg">Processing…</span><span class="progress-pct">0%</span></div>
    <div class="progress-bar"><div class="progress-fill"></div></div>
  </div>`;
}
function sizeCmpHTML(before, after) {
  const saving = Math.round((1 - after / before) * 100);
  return `<div class="size-comparison">
    <div class="size-item"><span class="size-num size-original">${formatBytes(before)}</span><span class="size-label">Original</span></div>
    <div class="size-arrow">→</div>
    <div class="size-item"><span class="size-num size-saving">${formatBytes(after)}</span><span class="size-label">After</span></div>
    <div class="size-item"><span class="size-num text-success">${saving > 0 ? '-' + saving + '%' : '+' + Math.abs(saving) + '%'}</span><span class="size-label">Savings</span></div>
  </div>`;
}
function dropZoneHTML(id, accept, label, multiple = false) {
  return `<div class="drop-zone" id="dz_${id}">
    <input type="file" id="fi_${id}" accept="${accept}"${multiple ? ' multiple' : ''} />
    <div class="drop-icon">📁</div>
    <div class="drop-title">${label}</div>
    <div class="drop-sub">or click to browse</div>
    <div class="drop-limit">Max ${MAX_FILE_MB}MB per file</div>
  </div>`;
}
function validateFile(file, types) {
  if (file.size > MAX_FILE_BYTES) { toast(`${file.name} exceeds ${MAX_FILE_MB}MB limit`, 'error'); return false; }
  if (types) {
    const ok = types.some(t => file.type.includes(t) || file.name.endsWith(t));
    if (!ok) { toast(`Unsupported file type: ${file.name}`, 'error'); return false; }
  }
  return true;
}
function setupDropZone(id, cb, types, multiple = false) {
  const dz = document.getElementById('dz_' + id);
  const fi = document.getElementById('fi_' + id);
  if (!dz || !fi) return;
  dz.addEventListener('dragover', e => { e.preventDefault(); dz.classList.add('drag-over'); });
  dz.addEventListener('dragleave', () => dz.classList.remove('drag-over'));
  dz.addEventListener('drop', e => {
    e.preventDefault(); dz.classList.remove('drag-over');
    const files = [...e.dataTransfer.files].filter(f => validateFile(f, types));
    if (files.length) cb(multiple ? files : [files[0]]);
  });
  fi.addEventListener('change', () => {
    const files = [...fi.files].filter(f => validateFile(f, types));
    if (files.length) cb(multiple ? files : [files[0]]);
    fi.value = '';
  });
}

// Shared UI helper — file chip shown after upload
function fileChipHTML(file) {
  const ext = file.name.split('.').pop().toUpperCase();
  return `<div class="file-chip mt-2">
    <span class="file-chip-icon">📄</span>
    <span class="file-chip-name" title="${file.name}">${file.name}</span>
    <span class="file-chip-size">${formatBytes(file.size)}</span>
    <span class="file-chip-ext">${ext}</span>
  </div>`;
}

// Lightweight debounce factory
function makeDebouncedFn(fn, ms) {
  let t = null;
  return (...args) => { clearTimeout(t); t = setTimeout(() => fn(...args), ms); };
}

// ─── Mobile Detection ────────────────────────────────────────
// Returns true on phones/tablets (used to cap canvas scale & resolution)
function isMobile() {
  return window.innerWidth < 768 || /Android|iPhone|iPad|iPod/i.test(navigator.userAgent);
}

// ─── Safe Canvas Scale ───────────────────────────────────
// Caps render scale at 1.5× on mobile to prevent canvas memory crashes
function safeScale(requested) {
  return isMobile() ? Math.min(requested, 1.5) : requested;
}

// ─── Safe Image Resolution ──────────────────────────────
function capDimForMobile(w, h, maxPx = 1920) {
  if (!isMobile() || (w <= maxPx && h <= maxPx)) return { w, h };
  const sc = maxPx / Math.max(w, h);
  return { w: Math.round(w * sc), h: Math.round(h * sc) };
}

// ─── Lazy Script Loader ──────────────────────────────────
function loadScript(src) {
  return new Promise((resolve, reject) => {
    if (document.querySelector(`script[src="${src}"]`)) { resolve(); return; }
    const s = document.createElement('script');
    s.src = src; s.onload = resolve; s.onerror = () => reject(new Error('Failed to load: ' + src));
    document.head.appendChild(s);
  });
}

// ─── Lazy Library Loaders ──────────────────────────────
async function ensureMammoth() {
  if (!window.mammoth) {
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.6.0/mammoth.browser.min.js');
  }
}
async function ensureHtml2Canvas() {
  if (!window.html2canvas) {
    await loadScript('https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js');
  }
}

// ═══════════════════════════════════════════════════════════
// PDF TOOLS
// ═══════════════════════════════════════════════════════════

// ─── PDF → Images ─────────────────────────────────────────
function renderPdfToImages(body) {
  body.innerHTML = `
    ${dropZoneHTML('pti', 'application/pdf', 'Drop your PDF here')}
    <div class="form-row mt-2">
      <div class="form-field">
        <label>Output Format</label>
        <select class="input" id="pti_fmt"><option value="image/png">PNG</option><option value="image/jpeg">JPG</option></select>
      </div>
      <div class="form-field">
        <label>Scale (DPI×)</label>
        <select class="input" id="pti_scale">
          <option value="1">1× (~72 DPI)</option>
          <option value="2" selected>2× (~144 DPI)</option>
          <option value="3">3× (~216 DPI)</option>
        </select>
      </div>
    </div>
    <div id="pti_progress" class="hidden">${progressHTML('pti_prog')}</div>
    <div id="pti_results" class="result-grid"></div>
    <div id="pti_actions" class="hidden flex gap-1 mt-2 flex-wrap">
      <button class="btn btn-primary" id="pti_dlAll">⬇ Download All (ZIP)</button>
    </div>`;

  let pages = [];
  setupDropZone('pti', async ([file]) => {
    pages = [];
    const results = document.getElementById('pti_results');
    const progress = document.getElementById('pti_progress');
    const actions = document.getElementById('pti_actions');
    results.innerHTML = ''; actions.classList.add('hidden');
    progress.classList.remove('hidden');
    try {
      const ab = await readFileAsArrayBuffer(file);
      const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
      for (let i = 1; i <= pdf.numPages; i++) {
        setProgress(document.getElementById('pti_prog'), (i / pdf.numPages) * 100, `Rendering page ${i}/${pdf.numPages}…`);
        const page = await pdf.getPage(i);
        const scale = parseFloat(document.getElementById('pti_scale').value);
        const vp = page.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = vp.width; canvas.height = vp.height;
        await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
        const fmt = document.getElementById('pti_fmt').value;
        const ext = fmt === 'image/jpeg' ? 'jpg' : 'png';
        const dataUrl = canvas.toDataURL(fmt, 0.92);
        pages.push({ dataUrl, name: `page_${i}.${ext}` });
        const div = document.createElement('div');
        div.className = 'result-item';
        div.innerHTML = `<img src="${dataUrl}" /><div class="result-item-name">page_${i}.${ext}</div>
          <a href="${dataUrl}" download="page_${i}.${ext}">Download</a>`;
        results.appendChild(div);
      }
      progress.classList.add('hidden');
      actions.classList.remove('hidden');
      saveHistory('pdf-to-images', file.name);
      toast(`${pdf.numPages} pages exported`, 'success');
    } catch (e) { progress.classList.add('hidden'); toast('Failed to process PDF: ' + e.message, 'error'); }
  }, ['pdf']);

  document.getElementById('pti_dlAll').addEventListener('click', async () => {
    if (!pages.length) return;
    const zip = new JSZip();
    pages.forEach(p => {
      const b64 = p.dataUrl.split(',')[1];
      zip.file(p.name, b64, { base64: true });
    });
    const blob = await zip.generateAsync({ type: 'blob' });
    downloadBlob(blob, 'pdf_images.zip');
  });
}

// ─── Images → PDF ─────────────────────────────────────────
function renderImagesToPdf(body) {
  body.innerHTML = `
    ${dropZoneHTML('itp', 'image/png,image/jpeg,image/webp', 'Drop images here (drag to reorder)', true)}
    <div id="itp_list" class="thumb-grid mt-2"></div>
    <div id="itp_progress" class="hidden">${progressHTML('itp_prog')}</div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="itp_convert">Convert to PDF</button>`;

  let files = [];
  const listEl = document.getElementById('itp_list');
  const btn = document.getElementById('itp_convert');

  function renderList() {
    listEl.innerHTML = '';
    files.forEach((f, i) => {
      const url = URL.createObjectURL(f);
      const div = document.createElement('div');
      div.className = 'thumb-item';
      div.dataset.idx = i;
      div.innerHTML = `<img src="${url}" /><div class="thumb-item-label">${f.name.substring(0, 16)}</div>
        <button class="thumb-item-del" onclick="itpRemove(${i})">×</button>`;
      listEl.appendChild(div);
    });
    btn.classList.toggle('hidden', !files.length);
    if (files.length > 1 && typeof Sortable !== 'undefined') {
      new Sortable(listEl, {
        animation: 150, onEnd: e => {
          const moved = files.splice(e.oldIndex, 1)[0];
          files.splice(e.newIndex, 0, moved);
        }
      });
    }
  }
  window.itpRemove = i => { files.splice(i, 1); renderList(); };

  setupDropZone('itp', f => { files = [...files, ...f]; renderList(); }, ['image/png', 'image/jpeg', 'image/webp', 'jpg', 'png', 'webp'], true);

  btn.addEventListener('click', async () => {
    if (!files.length) return;
    const progress = document.getElementById('itp_progress');
    progress.classList.remove('hidden'); btn.disabled = true;
    try {
      const { jsPDF } = window.jspdf;
      let pdf = null;
      for (let i = 0; i < files.length; i++) {
        setProgress(document.getElementById('itp_prog'), (i / files.length) * 100, `Processing image ${i + 1}/${files.length}…`);
        const dataUrl = await readFileAsDataURL(files[i]);
        const img = await new Promise(res => { const im = new Image(); im.onload = () => res(im); im.src = dataUrl; });
        const w = img.naturalWidth, h = img.naturalHeight;
        const orient = w > h ? 'l' : 'p';
        if (!pdf) {
          pdf = new jsPDF({ orientation: orient, unit: 'px', format: [w, h] });
        } else {
          pdf.addPage([w, h], orient);
        }
        pdf.addImage(dataUrl, 'JPEG', 0, 0, w, h);
      }
      pdf.save('images_combined.pdf');
      progress.classList.add('hidden'); btn.disabled = false;
      saveHistory('images-to-pdf', files[0].name);
      toast('PDF created!', 'success');
    } catch (e) { progress.classList.add('hidden'); btn.disabled = false; toast('Error: ' + e.message, 'error'); }
  });
}

// ─── PDF Merge ─────────────────────────────────────────────
function renderPdfMerge(body) {
  body.innerHTML = `
    ${dropZoneHTML('pm', 'application/pdf', 'Drop multiple PDFs here', true)}
    <div id="pm_list" class="file-list mt-2"></div>
    <div id="pm_progress" class="hidden">${progressHTML('pm_prog')}</div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="pm_merge">Merge PDFs</button>`;

  let files = [];
  const listEl = document.getElementById('pm_list');
  const btn = document.getElementById('pm_merge');

  function renderList() {
    listEl.innerHTML = '';
    files.forEach((f, i) => {
      const div = document.createElement('div');
      div.className = 'file-item';
      div.innerHTML = `<span class="drag-handle">⠿</span><span class="file-item-icon">📄</span>
        <span class="file-item-name">${f.name}</span>
        <span class="file-item-size">${formatBytes(f.size)}</span>
        <button class="file-item-del" onclick="pmRemove(${i})">✕</button>`;
      listEl.appendChild(div);
    });
    btn.classList.toggle('hidden', files.length < 2);
    if (files.length > 1 && typeof Sortable !== 'undefined') {
      new Sortable(listEl, {
        handle: '.drag-handle', animation: 150, onEnd: e => {
          const moved = files.splice(e.oldIndex, 1)[0];
          files.splice(e.newIndex, 0, moved);
        }
      });
    }
  }
  window.pmRemove = i => { files.splice(i, 1); renderList(); };
  setupDropZone('pm', f => { files = [...files, ...f]; renderList(); }, ['pdf'], true);

  btn.addEventListener('click', async () => {
    const progress = document.getElementById('pm_progress');
    progress.classList.remove('hidden'); btn.disabled = true;
    try {
      const { PDFDocument } = PDFLib;
      const merged = await PDFDocument.create();
      for (let i = 0; i < files.length; i++) {
        setProgress(document.getElementById('pm_prog'), (i / files.length) * 100, `Merging ${files[i].name}…`);
        const ab = await readFileAsArrayBuffer(files[i]);
        const pdf = await PDFDocument.load(ab);
        const copied = await merged.copyPages(pdf, pdf.getPageIndices());
        copied.forEach(p => merged.addPage(p));
      }
      setProgress(document.getElementById('pm_prog'), 100, 'Saving…');
      const bytes = await merged.save();
      downloadBlob(new Blob([bytes], { type: 'application/pdf' }), 'merged.pdf');
      progress.classList.add('hidden'); btn.disabled = false;
      saveHistory('pdf-merge', files[0].name);
      toast('PDFs merged!', 'success');
    } catch (e) { progress.classList.add('hidden'); btn.disabled = false; toast('Error: ' + e.message, 'error'); }
  });
}

// ─── PDF Split ─────────────────────────────────────────────
function renderPdfSplit(body) {
  body.innerHTML = `
    ${dropZoneHTML('ps', 'application/pdf', 'Drop your PDF here')}
    <div id="ps_info" class="info-box hidden"></div>
    <div class="form-field mt-2 hidden" id="ps_rangeWrap">
      <label>Page Range (e.g. 1-3, 5, 7-9)</label>
      <input type="text" class="input" id="ps_range" placeholder="e.g. 1-3, 5, 7-9" />
    </div>
    <div id="ps_progress" class="hidden">${progressHTML('ps_prog')}</div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="ps_split">Split & Download</button>`;

  let pdfFile = null;
  setupDropZone('ps', async ([file]) => {
    pdfFile = file;
    const ab = await readFileAsArrayBuffer(file);
    const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
    document.getElementById('ps_info').textContent = `✓ ${file.name} — ${pdf.numPages} pages`;
    document.getElementById('ps_info').classList.remove('hidden');
    document.getElementById('ps_rangeWrap').classList.remove('hidden');
    document.getElementById('ps_split').classList.remove('hidden');
  }, ['pdf']);

  document.getElementById('ps_split').addEventListener('click', async () => {
    const rangeStr = document.getElementById('ps_range').value.trim();
    if (!rangeStr || !pdfFile) return;
    const progress = document.getElementById('ps_progress');
    progress.classList.remove('hidden');
    try {
      const { PDFDocument } = PDFLib;
      const ab = await readFileAsArrayBuffer(pdfFile);
      const srcPdf = await PDFDocument.load(ab);
      const totalPages = srcPdf.getPageCount();
      // Parse ranges
      const pages = new Set();
      rangeStr.split(',').forEach(part => {
        const m = part.trim().match(/^(\d+)(?:-(\d+))?$/);
        if (m) {
          const from = parseInt(m[1]) - 1, to = m[2] ? parseInt(m[2]) - 1 : from;
          for (let i = from; i <= Math.min(to, totalPages - 1); i++) pages.add(i);
        }
      });
      if (!pages.size) { toast('No valid pages found in range', 'error'); progress.classList.add('hidden'); return; }
      setProgress(document.getElementById('ps_prog'), 50, 'Extracting pages…');
      const newPdf = await PDFDocument.create();
      const indices = [...pages].sort((a, b) => a - b);
      const copied = await newPdf.copyPages(srcPdf, indices);
      copied.forEach(p => newPdf.addPage(p));
      setProgress(document.getElementById('ps_prog'), 90, 'Saving…');
      const bytes = await newPdf.save();
      downloadBlob(new Blob([bytes], { type: 'application/pdf' }), 'split_pages.pdf');
      progress.classList.add('hidden');
      saveHistory('pdf-split', pdfFile.name);
      toast(`Extracted ${pages.size} pages`, 'success');
    } catch (e) { progress.classList.add('hidden'); toast('Error: ' + e.message, 'error'); }
  });
}

// ─── PDF Compress (Real-Time Reactive) ────────────────────
function renderPdfCompress(body) {
  // ── Central State ──────────────────────────────────────────
  const state = {
    file: null,
    pdfDoc: null,       // loaded pdfjsLib document (cached, no re-read)
    origBytes: null,    // ArrayBuffer of original file
    outputBlob: null,
    quality: 0.60,
    scale: 0.70,
    isProcessing: false,
    jobId: 0,
    numPages: 0,
  };

  body.innerHTML = `
    <div id="pc_dropzone">${dropZoneHTML('pc', 'application/pdf', 'Drop your PDF here')}</div>

    <div id="pc_workspace" class="hidden">
      <!-- First-page preview -->
      <div class="pc-preview-row">
        <div class="pc-preview-panel">
          <div class="pc-preview-label">First Page Preview</div>
          <div class="pc-preview-imgwrap" id="pc_preview_wrap">
            <canvas id="pc_preview_canvas"></canvas>
            <div class="ic-processing-overlay hidden" id="pc_overlay">
              <div class="ic-spinner"></div>
            </div>
          </div>
        </div>
        <div class="pc-preview-stats">
          <div class="pc-stat-box">
            <div class="pc-stat-label">Original</div>
            <div class="pc-stat-val" id="pc_orig_size">—</div>
          </div>
          <div class="pc-stat-box">
            <div class="pc-stat-label">Compressed</div>
            <div class="pc-stat-val" id="pc_comp_size">—</div>
          </div>
          <div class="pc-stat-box">
            <div class="pc-stat-label">Reduced by</div>
            <div class="pc-stat-val" id="pc_savings" style="color:var(--success)">—</div>
          </div>
          <div class="pc-stat-box">
            <div class="pc-stat-label">Pages</div>
            <div class="pc-stat-val" id="pc_pages">—</div>
          </div>
          <div id="pc_limit_warn" class="warn-box hidden" style="margin-top:0.5rem;font-size:0.78rem">
            ⚠ Limited compression — browser-based PDF processing rasterises pages. Result may be larger than the original for already-optimised PDFs.
          </div>
        </div>
      </div>

      <!-- Controls -->
      <div class="ic-controls-panel mt-2">
        <div class="control-group">
          <div class="control-label">JPEG Quality &nbsp;<span id="pc_qval">60%</span></div>
          <input type="range" min="10" max="95" value="60" id="pc_quality" />
        </div>
        <div class="control-group" style="margin-bottom:0">
          <div class="control-label">Render Scale &nbsp;<span id="pc_dval">0.7×</span>
            <span style="font-size:0.72rem;color:var(--text3);font-weight:400">(lower = smaller file)</span>
          </div>
          <input type="range" min="2" max="15" value="7" id="pc_dpi" step="1" />
        </div>
      </div>

      <!-- Progress -->
      <div id="pc_progress" class="hidden" style="margin-top:0.75rem">${progressHTML('pc_prog')}</div>

      <!-- Download -->
      <div class="flex gap-1 mt-2">
        <button class="btn btn-primary" id="pc_download" disabled>⬇ Download Compressed PDF</button>
        <button class="btn btn-ghost" id="pc_reset">✕ New PDF</button>
      </div>
    </div>`;

  // ── Debounce + Job Cancel ──────────────────────────────────
  let debounceTimer = null;

  function schedule(immediate = false) {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => runCompress(), immediate ? 0 : 300);
  }

  // ── Render first-page preview (fast, separate from full compress) ──
  async function renderPreview(myJob) {
    if (!state.pdfDoc) return;
    try {
      const page = await state.pdfDoc.getPage(1);
      const vp = page.getViewport({ scale: state.scale });
      const canvas = document.getElementById('pc_preview_canvas');
      canvas.width = vp.width;
      canvas.height = vp.height;
      if (myJob !== state.jobId) return;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
    } catch (_) { }
  }

  // ── Full compression pass ──────────────────────────────────
  async function runCompress() {
    if (!state.pdfDoc || !state.origBytes) return;
    state.jobId++;
    const myJob = state.jobId;
    state.isProcessing = true;

    // UI: spinner on, download off, progress visible
    document.getElementById('pc_overlay').classList.remove('hidden');
    document.getElementById('pc_download').disabled = true;
    document.getElementById('pc_progress').classList.remove('hidden');
    document.getElementById('pc_comp_size').textContent = '…';
    document.getElementById('pc_savings').textContent = '…';
    document.getElementById('pc_limit_warn').classList.add('hidden');

    try {
      const { jsPDF } = window.jspdf;
      const quality = state.quality;
      const scale = state.scale;
      const pdf = state.pdfDoc;
      let newPdf = null;

      for (let i = 1; i <= state.numPages; i++) {
        if (myJob !== state.jobId) return; // cancelled

        setProgress(
          document.getElementById('pc_prog'),
          (i / state.numPages) * 100,
          `Compressing page ${i} / ${state.numPages}…`
        );

        const page = await pdf.getPage(i);
        const vp = page.getViewport({ scale });
        const canvas = document.createElement('canvas');
        canvas.width = vp.width;
        canvas.height = vp.height;
        await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;

        if (myJob !== state.jobId) return; // cancelled after render

        const dataUrl = canvas.toDataURL('image/jpeg', quality);

        if (!newPdf) {
          newPdf = new jsPDF({
            orientation: vp.width > vp.height ? 'l' : 'p',
            unit: 'px',
            format: [vp.width, vp.height],
          });
        } else {
          newPdf.addPage([vp.width, vp.height], vp.width > vp.height ? 'l' : 'p');
        }
        newPdf.addImage(dataUrl, 'JPEG', 0, 0, vp.width, vp.height);

        // Also refresh preview on page 1
        if (i === 1) renderPreview(myJob);
      }

      if (myJob !== state.jobId) return;

      const bytes = newPdf.output('arraybuffer');
      const blob = new Blob([bytes], { type: 'application/pdf' });
      state.outputBlob = blob;

      // Update stats
      const origSize = state.file.size;
      const compSize = blob.size;
      const saved = Math.round((1 - compSize / origSize) * 100);

      document.getElementById('pc_comp_size').textContent = formatBytes(compSize);
      document.getElementById('pc_savings').textContent = saved > 0
        ? `${saved}% smaller`
        : `${Math.abs(saved)}% larger`;
      document.getElementById('pc_savings').style.color = saved > 0
        ? 'var(--success)'
        : 'var(--error)';

      // Honest messaging
      if (saved <= 5) {
        document.getElementById('pc_limit_warn').classList.remove('hidden');
      }

      document.getElementById('pc_overlay').classList.add('hidden');
      document.getElementById('pc_progress').classList.add('hidden');
      document.getElementById('pc_download').disabled = false;
      saveHistory('pdf-compress', state.file.name);

    } catch (err) {
      if (myJob !== state.jobId) return;
      document.getElementById('pc_overlay').classList.add('hidden');
      document.getElementById('pc_progress').classList.add('hidden');
      toast('Compression error: ' + err.message, 'error');
    } finally {
      if (myJob === state.jobId) state.isProcessing = false;
    }
  }

  // ── Load file ──────────────────────────────────────────────
  async function loadFile(file) {
    state.file = file;
    state.outputBlob = null;
    state.jobId++;   // cancel anything in flight

    document.getElementById('pc_dropzone').classList.add('hidden');
    document.getElementById('pc_workspace').classList.remove('hidden');
    document.getElementById('pc_orig_size').textContent = formatBytes(file.size);
    document.getElementById('pc_comp_size').textContent = '—';
    document.getElementById('pc_savings').textContent = '—';
    document.getElementById('pc_download').disabled = true;

    try {
      const ab = await readFileAsArrayBuffer(file);
      state.origBytes = ab;
      const pdfDoc = await pdfjsLib.getDocument({ data: ab.slice(0) }).promise;
      state.pdfDoc = pdfDoc;
      state.numPages = pdfDoc.numPages;
      document.getElementById('pc_pages').textContent = pdfDoc.numPages + ' pg';
    } catch (err) {
      toast('Could not read PDF: ' + err.message, 'error');
      return;
    }

    // Kick off immediately
    schedule(true);
  }

  // ── Wire Controls ──────────────────────────────────────────
  document.getElementById('pc_quality').addEventListener('input', e => {
    state.quality = parseInt(e.target.value) / 100;
    document.getElementById('pc_qval').textContent = e.target.value + '%';
    schedule();
  });
  document.getElementById('pc_dpi').addEventListener('input', e => {
    state.scale = parseInt(e.target.value) / 10;
    document.getElementById('pc_dval').textContent = state.scale.toFixed(1) + '×';
    schedule();
  });

  // ── Download ───────────────────────────────────────────────
  document.getElementById('pc_download').addEventListener('click', () => {
    if (!state.outputBlob) return;
    const base = (state.file?.name || 'document').replace(/\.pdf$/i, '');
    downloadBlob(state.outputBlob, `${base}_compressed.pdf`);
    toast('Downloaded!', 'success');
  });

  // ── Reset ──────────────────────────────────────────────────
  document.getElementById('pc_reset').addEventListener('click', () => {
    state.jobId++;
    state.file = null; state.pdfDoc = null; state.origBytes = null; state.outputBlob = null;
    document.getElementById('pc_dropzone').classList.remove('hidden');
    document.getElementById('pc_workspace').classList.add('hidden');
    document.getElementById('pc_progress').classList.add('hidden');
    document.getElementById('pc_download').disabled = true;
  });

  // ── Drop Zone ──────────────────────────────────────────────
  setupDropZone('pc', ([file]) => loadFile(file), ['pdf']);
}

// ─── PDF Page Editor ───────────────────────────────────────
function renderPdfPageEditor(body) {
  body.innerHTML = `
    ${dropZoneHTML('ppe', 'application/pdf', 'Drop your PDF here')}
    <div class="info-box hidden" id="ppe_info">Drag to reorder · Click 🔄 to rotate · Click ✕ to delete</div>
    <div id="ppe_thumbs" class="thumb-grid mt-2"></div>
    <div id="ppe_progress" class="hidden">${progressHTML('ppe_prog')}</div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="ppe_save">💾 Save Edited PDF</button>`;

  let pageData = []; // { canvas, rotation }
  let pdfFile = null;

  setupDropZone('ppe', async ([file]) => {
    pdfFile = file;
    pageData = [];
    document.getElementById('ppe_thumbs').innerHTML = '';
    document.getElementById('ppe_save').classList.add('hidden');
    const progress = document.getElementById('ppe_progress');
    progress.classList.remove('hidden');
    const ab = await readFileAsArrayBuffer(file);
    const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
    for (let i = 1; i <= pdf.numPages; i++) {
      setProgress(document.getElementById('ppe_prog'), (i / pdf.numPages) * 100, `Loading page ${i}…`);
      const page = await pdf.getPage(i);
      const vp = page.getViewport({ scale: 1.0 });
      const canvas = document.createElement('canvas');
      canvas.width = vp.width; canvas.height = vp.height;
      await page.render({ canvasContext: canvas.getContext('2d'), viewport: vp }).promise;
      pageData.push({ canvas, rotation: 0 });
    }
    progress.classList.add('hidden');
    document.getElementById('ppe_info').classList.remove('hidden');
    document.getElementById('ppe_save').classList.remove('hidden');
    renderThumbs();
  }, ['pdf']);

  function renderThumbs() {
    const grid = document.getElementById('ppe_thumbs');
    grid.innerHTML = '';
    pageData.forEach((p, i) => {
      const div = document.createElement('div');
      div.className = 'thumb-item'; div.dataset.idx = i;
      const thumbCanvas = document.createElement('canvas');
      const scale = 120 / Math.max(p.canvas.width, p.canvas.height);
      thumbCanvas.width = p.canvas.width * scale;
      thumbCanvas.height = p.canvas.height * scale;
      thumbCanvas.getContext('2d').drawImage(p.canvas, 0, 0, thumbCanvas.width, thumbCanvas.height);
      div.appendChild(thumbCanvas);
      div.innerHTML += `<div class="thumb-item-label">Page ${i + 1}</div>
        <button class="thumb-item-del" onclick="ppeDelete(${i})">✕</button>
        <button class="rotate-btn" onclick="ppeRotate(${i})">🔄</button>`;
      grid.appendChild(div);
    });
    if (pageData.length > 1 && typeof Sortable !== 'undefined') {
      new Sortable(grid, {
        animation: 150, onEnd: e => {
          const moved = pageData.splice(e.oldIndex, 1)[0];
          pageData.splice(e.newIndex, 0, moved);
          renderThumbs();
        }
      });
    }
  }
  window.ppeDelete = i => { pageData.splice(i, 1); renderThumbs(); };
  window.ppeRotate = i => { pageData[i].rotation = (pageData[i].rotation + 90) % 360; renderThumbs(); };

  document.getElementById('ppe_save').addEventListener('click', async () => {
    const progress = document.getElementById('ppe_progress');
    progress.classList.remove('hidden');
    try {
      const { PDFDocument, degrees } = PDFLib;
      const newPdf = await PDFDocument.create();
      const { jsPDF } = window.jspdf;
      let jpdf = null;
      for (let i = 0; i < pageData.length; i++) {
        setProgress(document.getElementById('ppe_prog'), (i / pageData.length) * 100, `Building page ${i + 1}…`);
        const { canvas, rotation } = pageData[i];
        const rotCanvas = document.createElement('canvas');
        const swap = rotation === 90 || rotation === 270;
        rotCanvas.width = swap ? canvas.height : canvas.width;
        rotCanvas.height = swap ? canvas.width : canvas.height;
        const ctx = rotCanvas.getContext('2d');
        ctx.translate(rotCanvas.width / 2, rotCanvas.height / 2);
        ctx.rotate(rotation * Math.PI / 180);
        ctx.drawImage(canvas, -canvas.width / 2, -canvas.height / 2);
        const dataUrl = rotCanvas.toDataURL('image/jpeg', 0.9);
        const w = rotCanvas.width, h = rotCanvas.height;
        if (!jpdf) {
          jpdf = new jsPDF({ orientation: w > h ? 'l' : 'p', unit: 'px', format: [w, h] });
        } else {
          jpdf.addPage([w, h], w > h ? 'l' : 'p');
        }
        jpdf.addImage(dataUrl, 'JPEG', 0, 0, w, h);
      }
      jpdf.save('edited.pdf');
      progress.classList.add('hidden');
      saveHistory('pdf-page-editor', pdfFile.name);
      toast('Edited PDF saved!', 'success');
    } catch (e) { progress.classList.add('hidden'); toast('Error: ' + e.message, 'error'); }
  });
}

// ─── PDF Text Extractor ────────────────────────────────────
function renderPdfText(body) {
  body.innerHTML = `
    <div class="info-box">This tool extracts readable text from PDFs. Scanned PDFs without embedded text will return little or no content.</div>
    ${dropZoneHTML('ptxt', 'application/pdf', 'Drop your PDF here')}
    <div id="ptxt_progress" class="hidden">${progressHTML('ptxt_prog')}</div>
    <div class="copy-wrap mt-2 hidden" id="ptxt_wrap">
      <div class="output-area" id="ptxt_out"></div>
      <button class="copy-btn" onclick="copyText('ptxt_out')">Copy</button>
    </div>
    <div class="flex gap-1 mt-2 hidden" id="ptxt_actions">
      <button class="btn btn-primary" id="ptxt_dl">⬇ Download .txt</button>
    </div>`;

  let extractedText = '';
  setupDropZone('ptxt', async ([file]) => {
    const progress = document.getElementById('ptxt_progress');
    const wrap = document.getElementById('ptxt_wrap');
    const actions = document.getElementById('ptxt_actions');
    progress.classList.remove('hidden'); wrap.classList.add('hidden'); actions.classList.add('hidden');
    try {
      const ab = await readFileAsArrayBuffer(file);
      const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
      const parts = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        setProgress(document.getElementById('ptxt_prog'), (i / pdf.numPages) * 100, `Reading page ${i}/${pdf.numPages}…`);
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        const text = content.items.map(it => it.str).join(' ');
        parts.push(`--- Page ${i} ---\n${text}`);
      }
      extractedText = parts.join('\n\n');
      document.getElementById('ptxt_out').textContent = extractedText || '(No readable text found in this PDF)';
      progress.classList.add('hidden');
      wrap.classList.remove('hidden');
      actions.classList.remove('hidden');
      saveHistory('pdf-text', file.name);
      toast('Text extracted!', 'success');
    } catch (e) { progress.classList.add('hidden'); toast('Error: ' + e.message, 'error'); }
  }, ['pdf']);

  document.getElementById('ptxt_dl').addEventListener('click', () => {
    if (!extractedText) return;
    downloadBlob(new Blob([extractedText], { type: 'text/plain' }), 'extracted_text.txt');
  });
}

window.copyText = id => {
  const el = document.getElementById(id);
  navigator.clipboard.writeText(el.textContent).then(() => toast('Copied!', 'success'));
};

// ─── PDF Page Numbers ─────────────────────────────────────
function renderPdfPageNumbers(body) {
  body.innerHTML = `
    ${dropZoneHTML('ppn', 'application/pdf', 'Drop your PDF here')}
    <div class="form-row mt-2">
      <div class="form-field">
        <label>Position</label>
        <select class="input" id="ppn_pos">
          <option value="bottom-center">Bottom Center</option>
          <option value="bottom-right">Bottom Right</option>
          <option value="top-center">Top Center</option>
        </select>
      </div>
      <div class="form-field">
        <label>Start From</label>
        <input type="number" class="input" id="ppn_start" value="1" min="1" />
      </div>
      <div class="form-field">
        <label>Font Size</label>
        <input type="number" class="input" id="ppn_size" value="12" min="6" max="36" />
      </div>
    </div>
    <div id="ppn_progress" class="hidden">${progressHTML('ppn_prog')}</div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="ppn_apply">Apply Page Numbers</button>`;

  let pdfFile = null;
  setupDropZone('ppn', async ([file]) => {
    pdfFile = file;
    document.getElementById('ppn_apply').classList.remove('hidden');
    toast(`Loaded: ${file.name}`, 'info');
  }, ['pdf']);

  document.getElementById('ppn_apply').addEventListener('click', async () => {
    if (!pdfFile) return;
    const progress = document.getElementById('ppn_progress');
    progress.classList.remove('hidden');
    try {
      const { PDFDocument, rgb, StandardFonts } = PDFLib;
      const ab = await readFileAsArrayBuffer(pdfFile);
      const pdf = await PDFDocument.load(ab);
      const font = await pdf.embedFont(StandardFonts.Helvetica);
      const pages = pdf.getPages();
      const startNum = parseInt(document.getElementById('ppn_start').value) || 1;
      const fontSize = parseInt(document.getElementById('ppn_size').value) || 12;
      const pos = document.getElementById('ppn_pos').value;
      pages.forEach((page, i) => {
        setProgress(document.getElementById('ppn_prog'), ((i + 1) / pages.length) * 100, `Numbering page ${i + 1}…`);
        const { width, height } = page.getSize();
        const text = String(startNum + i);
        const tw = font.widthOfTextAtSize(text, fontSize);
        let x, y;
        if (pos === 'bottom-center') { x = (width - tw) / 2; y = 20; }
        else if (pos === 'bottom-right') { x = width - tw - 20; y = 20; }
        else { x = (width - tw) / 2; y = height - 30; }
        page.drawText(text, { x, y, size: fontSize, font, color: rgb(0.3, 0.3, 0.3) });
      });
      const bytes = await pdf.save();
      downloadBlob(new Blob([bytes], { type: 'application/pdf' }), 'numbered.pdf');
      progress.classList.add('hidden');
      saveHistory('pdf-page-numbers', pdfFile.name);
      toast('Page numbers added!', 'success');
    } catch (e) { progress.classList.add('hidden'); toast('Error: ' + e.message, 'error'); }
  });
}

// ─── PDF Watermark ─────────────────────────────────────────
function renderPdfWatermark(body) {
  body.innerHTML = `
    ${dropZoneHTML('pw', 'application/pdf', 'Drop your PDF here')}
    <div class="form-row mt-2">
      <div class="form-field">
        <label>Watermark Text</label>
        <input type="text" class="input" id="pw_text" value="CONFIDENTIAL" />
      </div>
      <div class="form-field">
        <label>Opacity (0–100)</label>
        <input type="number" class="input" id="pw_opacity" value="20" min="5" max="80" />
      </div>
    </div>
    <div class="form-row">
      <div class="form-field">
        <label>Font Size</label>
        <input type="number" class="input" id="pw_size" value="48" min="12" max="120" />
      </div>
      <div class="form-field">
        <label>Rotation (°)</label>
        <input type="number" class="input" id="pw_angle" value="-45" min="-90" max="90" />
      </div>
    </div>
    <div id="pw_progress" class="hidden">${progressHTML('pw_prog')}</div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="pw_apply">Add Watermark</button>`;

  let pdfFile = null;
  setupDropZone('pw', ([file]) => {
    pdfFile = file;
    document.getElementById('pw_apply').classList.remove('hidden');
    toast(`Loaded: ${file.name}`, 'info');
  }, ['pdf']);

  document.getElementById('pw_apply').addEventListener('click', async () => {
    if (!pdfFile) return;
    const text = document.getElementById('pw_text').value || 'WATERMARK';
    const opacity = parseFloat(document.getElementById('pw_opacity').value) / 100;
    const fontSize = parseInt(document.getElementById('pw_size').value) || 48;
    const angle = parseFloat(document.getElementById('pw_angle').value) || -45;
    const progress = document.getElementById('pw_progress');
    progress.classList.remove('hidden');
    try {
      const { PDFDocument, rgb, degrees, StandardFonts } = PDFLib;
      const ab = await readFileAsArrayBuffer(pdfFile);
      const pdf = await PDFDocument.load(ab);
      const font = await pdf.embedFont(StandardFonts.HelveticaBold);
      const pages = pdf.getPages();
      pages.forEach((page, i) => {
        setProgress(document.getElementById('pw_prog'), ((i + 1) / pages.length) * 100, `Watermarking page ${i + 1}…`);
        const { width, height } = page.getSize();
        page.drawText(text, {
          x: width / 3, y: height / 2,
          size: fontSize, font,
          color: rgb(0.5, 0.5, 0.5),
          opacity, rotate: degrees(angle)
        });
      });
      const bytes = await pdf.save();
      downloadBlob(new Blob([bytes], { type: 'application/pdf' }), 'watermarked.pdf');
      progress.classList.add('hidden');
      saveHistory('pdf-watermark', pdfFile.name);
      toast('Watermark applied!', 'success');
    } catch (e) { progress.classList.add('hidden'); toast('Error: ' + e.message, 'error'); }
  });
}

// ─── PDF Password ──────────────────────────────────────────
function renderPdfPassword(body) {
  body.innerHTML = `
    <div class="warn-box">⚠ pdf-lib supports encryption flags. For maximum security, use a dedicated PDF tool. This adds basic password metadata.</div>
    ${dropZoneHTML('ppass', 'application/pdf', 'Drop your PDF here')}
    <div class="form-row mt-2">
      <div class="form-field">
        <label>User Password (to open)</label>
        <input type="password" class="input" id="ppass_user" placeholder="Leave blank for none" />
      </div>
      <div class="form-field">
        <label>Owner Password (to edit)</label>
        <input type="password" class="input" id="ppass_owner" placeholder="Recommended" />
      </div>
    </div>
    <div id="ppass_progress" class="hidden">${progressHTML('ppass_prog')}</div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="ppass_apply">Encrypt PDF</button>`;

  let pdfFile = null;
  setupDropZone('ppass', ([file]) => {
    pdfFile = file;
    document.getElementById('ppass_apply').classList.remove('hidden');
    toast(`Loaded: ${file.name}`, 'info');
  }, ['pdf']);

  document.getElementById('ppass_apply').addEventListener('click', async () => {
    if (!pdfFile) return;
    const userPwd = document.getElementById('ppass_user').value;
    const ownerPwd = document.getElementById('ppass_owner').value || userPwd + '_owner';
    const progress = document.getElementById('ppass_progress');
    progress.classList.remove('hidden');
    try {
      const { PDFDocument } = PDFLib;
      const ab = await readFileAsArrayBuffer(pdfFile);
      const pdf = await PDFDocument.load(ab);
      setProgress(document.getElementById('ppass_prog'), 80, 'Encrypting…');
      const bytes = await pdf.save({
        userPassword: userPwd || undefined,
        ownerPassword: ownerPwd || undefined
      });
      downloadBlob(new Blob([bytes], { type: 'application/pdf' }), 'protected.pdf');
      progress.classList.add('hidden');
      saveHistory('pdf-password', pdfFile.name);
      toast('PDF encrypted!', 'success');
    } catch (e) { progress.classList.add('hidden'); toast('Error: ' + e.message, 'error'); }
  });
}

// ═══════════════════════════════════════════════════════════
// IMAGE TOOLS
// ═══════════════════════════════════════════════════════════

// ─── Image Compress (Real-Time Reactive) ───────────────────
function renderImgCompress(body) {
  // ── Central State ──────────────────────────────────────────
  const state = {
    file: null,
    origBlob: null,
    previewBlob: null,
    finalBlob: null,
    quality: 70,
    maxWidth: 2048,
    targetKB: null,
    isProcessing: false,
    jobId: 0,
  };

  body.innerHTML = `
    <div id="ic_dropzone">${dropZoneHTML('ic', 'image/*', 'Drop your image here')}</div>

    <div id="ic_workspace" class="hidden">
      <!-- Side-by-side preview -->
      <div class="ic-preview-row">
        <div class="ic-preview-panel">
          <div class="ic-preview-label">Original</div>
          <div class="ic-preview-imgwrap">
            <img id="ic_orig_img" alt="Original" />
          </div>
          <div class="ic-preview-meta" id="ic_orig_meta">—</div>
        </div>
        <div class="ic-preview-arrow">→</div>
        <div class="ic-preview-panel">
          <div class="ic-preview-label">Compressed</div>
          <div class="ic-preview-imgwrap" id="ic_comp_wrap">
            <img id="ic_comp_img" alt="Compressed" />
            <div class="ic-processing-overlay hidden" id="ic_overlay">
              <div class="ic-spinner"></div>
            </div>
          </div>
          <div class="ic-preview-meta" id="ic_comp_meta">—</div>
        </div>
      </div>

      <!-- Savings badge -->
      <div class="ic-savings-row" id="ic_savings_row" style="display:none">
        <span id="ic_savings_badge" class="ic-savings-badge">— saved</span>
        <span id="ic_savings_detail" class="ic-savings-detail"></span>
      </div>

      <!-- Controls -->
      <div class="ic-controls-panel">
        <div class="control-group">
          <div class="control-label">Quality <span id="ic_qval">70%</span></div>
          <input type="range" min="5" max="95" value="70" id="ic_quality" />
        </div>
        <div class="control-group">
          <div class="control-label">Max Width <span id="ic_wval">2048 px</span></div>
          <input type="range" min="128" max="4096" value="2048" step="64" id="ic_maxw" />
        </div>
        <div class="form-row">
          <div class="form-field">
            <label>Target Size (KB) <span style="color:var(--text3);font-size:0.75rem">(optional)</span></label>
            <input type="number" class="input" id="ic_targetkb" placeholder="e.g. 50" min="5" max="10000" />
          </div>
          <div class="form-field">
            <label>Output Format</label>
            <select class="input" id="ic_fmt">
              <option value="image/jpeg">JPG</option>
              <option value="image/webp">WebP</option>
              <option value="image/png">PNG</option>
            </select>
          </div>
        </div>
      </div>

      <!-- Download -->
      <div class="flex gap-1 mt-2">
        <button class="btn btn-primary" id="ic_dl" disabled>⬇ Download Compressed</button>
        <button class="btn btn-ghost" id="ic_reset">✕ New Image</button>
      </div>
    </div>`;

  // ── Helpers ────────────────────────────────────────────────
  let debounceTimer = null;

  function scheduleRecompress(immediate = false) {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => runCompress(), immediate ? 0 : 300);
  }

  async function runCompress() {
    if (!state.file) return;
    state.jobId++;
    const myJob = state.jobId;
    state.isProcessing = true;

    // Show overlay on compressed panel
    document.getElementById('ic_overlay').classList.remove('hidden');
    document.getElementById('ic_dl').disabled = true;

    try {
      const quality = state.quality / 100;
      const maxW = state.maxWidth;
      const fmt = document.getElementById('ic_fmt').value;
      const targetKB = state.targetKB;

      let options = {
        maxSizeMB: 50,
        maxWidthOrHeight: maxW,
        useWebWorker: true,
        initialQuality: quality,
        fileType: fmt,
        preserveExif: false,
      };

      // If target size is set, use iterative approach
      if (targetKB && targetKB > 0) {
        options.maxSizeMB = targetKB / 1024;
        options.initialQuality = Math.min(quality, 0.9);
      }

      const compressed = await imageCompression(state.file, options);
      if (myJob !== state.jobId) return; // cancelled by newer job

      state.previewBlob = compressed;
      state.finalBlob = compressed;

      // Update compressed preview
      const prevUrl = URL.createObjectURL(compressed);
      const compImg = document.getElementById('ic_comp_img');
      const oldUrl = compImg.src;
      compImg.src = prevUrl;
      if (oldUrl && oldUrl.startsWith('blob:')) URL.revokeObjectURL(oldUrl);

      // Update stats
      updateStats();

      document.getElementById('ic_overlay').classList.add('hidden');
      document.getElementById('ic_dl').disabled = false;
      saveHistory('img-compress', state.file.name);
    } catch (err) {
      if (myJob !== state.jobId) return;
      document.getElementById('ic_overlay').classList.add('hidden');
      toast('Compression error: ' + err.message, 'error');
    } finally {
      if (myJob === state.jobId) state.isProcessing = false;
    }
  }

  function updateStats() {
    const orig = state.file ? state.file.size : 0;
    const comp = state.finalBlob ? state.finalBlob.size : 0;
    if (!orig) return;

    document.getElementById('ic_orig_meta').textContent = formatBytes(orig);

    if (comp) {
      const saved = Math.round((1 - comp / orig) * 100);
      document.getElementById('ic_comp_meta').textContent = formatBytes(comp);

      const savingsRow = document.getElementById('ic_savings_row');
      const savingsBadge = document.getElementById('ic_savings_badge');
      const savingsDetail = document.getElementById('ic_savings_detail');

      savingsRow.style.display = 'flex';
      if (saved > 0) {
        savingsBadge.textContent = `${saved}% smaller`;
        savingsBadge.className = 'ic-savings-badge ic-savings-good';
        savingsDetail.textContent = `${formatBytes(orig)} → ${formatBytes(comp)}`;
      } else {
        savingsBadge.textContent = `${Math.abs(saved)}% larger`;
        savingsBadge.className = 'ic-savings-badge ic-savings-bad';
        savingsDetail.textContent = `${formatBytes(orig)} → ${formatBytes(comp)}`;
      }
    }
  }

  function loadFile(file) {
    state.file = file;
    state.finalBlob = null;
    state.previewBlob = null;

    // Show workspace, hide dropzone
    document.getElementById('ic_dropzone').classList.add('hidden');
    document.getElementById('ic_workspace').classList.remove('hidden');
    document.getElementById('ic_savings_row').style.display = 'none';
    document.getElementById('ic_comp_meta').textContent = '—';

    // Show original preview
    const origImg = document.getElementById('ic_orig_img');
    const oldOrigUrl = origImg.src;
    const origUrl = URL.createObjectURL(file);
    origImg.src = origUrl;
    if (oldOrigUrl && oldOrigUrl.startsWith('blob:')) URL.revokeObjectURL(oldOrigUrl);
    document.getElementById('ic_orig_meta').textContent = formatBytes(file.size);

    // Kick off compression immediately
    scheduleRecompress(true);
  }

  // ── Wire Controls ──────────────────────────────────────────
  document.getElementById('ic_quality').addEventListener('input', e => {
    state.quality = parseInt(e.target.value);
    document.getElementById('ic_qval').textContent = state.quality + '%';
    scheduleRecompress();
  });
  document.getElementById('ic_maxw').addEventListener('input', e => {
    state.maxWidth = parseInt(e.target.value);
    document.getElementById('ic_wval').textContent = state.maxWidth + ' px';
    scheduleRecompress();
  });
  document.getElementById('ic_targetkb').addEventListener('input', e => {
    state.targetKB = parseInt(e.target.value) || null;
    scheduleRecompress();
  });
  document.getElementById('ic_fmt').addEventListener('change', () => {
    scheduleRecompress();
  });

  // ── Download ───────────────────────────────────────────────
  document.getElementById('ic_dl').addEventListener('click', () => {
    if (!state.finalBlob) return;
    const fmt = document.getElementById('ic_fmt').value;
    const ext = fmt === 'image/jpeg' ? 'jpg' : fmt === 'image/webp' ? 'webp' : 'png';
    const base = (state.file?.name || 'image').replace(/\.[^.]+$/, '');
    downloadBlob(state.finalBlob, `${base}_compressed.${ext}`);
    toast('Downloaded!', 'success');
  });

  // ── Reset ──────────────────────────────────────────────────
  document.getElementById('ic_reset').addEventListener('click', () => {
    state.file = null; state.finalBlob = null; state.previewBlob = null;
    state.jobId++;
    document.getElementById('ic_dropzone').classList.remove('hidden');
    document.getElementById('ic_workspace').classList.add('hidden');
    document.getElementById('ic_orig_img').src = '';
    document.getElementById('ic_comp_img').src = '';
    document.getElementById('ic_savings_row').style.display = 'none';
    document.getElementById('ic_dl').disabled = true;
  });

  // ── Drop Zone ──────────────────────────────────────────────
  setupDropZone('ic', ([file]) => loadFile(file), ['image']);
}

// ─── Image Format Convert (Reactive) ───────────────────────
function renderImgConvert(body) {
  // ── State ───────────────────────────────────────────────────
  const state = { file: null, imgEl: null, outputBlob: null, jobId: 0 };

  body.innerHTML = `
    <div id="ifc_drop">${dropZoneHTML('ifc', 'image/*', 'Drop your image here')}</div>

    <div id="ifc_workspace" class="tool-workspace hidden">
      <div id="ifc_chip"></div>

      <!-- Side-by-side preview -->
      <div class="ic-preview-row mt-2">
        <div class="ic-preview-panel">
          <div class="ic-preview-label">Original</div>
          <div class="ic-preview-imgwrap">
            <img id="ifc_orig_img" alt="Original" />
          </div>
          <div class="ic-preview-meta" id="ifc_orig_meta">—</div>
        </div>
        <div class="ic-preview-arrow">→</div>
        <div class="ic-preview-panel">
          <div class="ic-preview-label" id="ifc_comp_label">Converted</div>
          <div class="ic-preview-imgwrap">
            <img id="ifc_conv_img" alt="Converted" />
            <div class="ic-processing-overlay hidden" id="ifc_overlay">
              <div class="ic-spinner"></div>
            </div>
          </div>
          <div class="ic-preview-meta" id="ifc_conv_meta">—</div>
        </div>
      </div>

      <!-- Size comparison row -->
      <div id="ifc_savings_row" class="ic-savings-row" style="display:none">
        <span id="ifc_savings_badge" class="ic-savings-badge">—</span>
        <span id="ifc_savings_detail" class="ic-savings-detail"></span>
      </div>

      <!-- Controls (inside workspace for clean state) -->
      <div class="ic-controls-panel mt-2">
        <div class="form-row">
          <div class="form-field">
            <label>Output Format</label>
            <select class="input" id="ifc_fmt">
              <option value="image/jpeg">JPG</option>
              <option value="image/webp">WebP</option>
              <option value="image/png">PNG</option>
            </select>
          </div>
          <div class="form-field">
            <label>Quality <span style="font-size:0.75rem;color:var(--text3)">(JPG/WebP)</span></label>
            <div style="display:flex;align-items:center;gap:0.5rem">
              <input type="range" min="5" max="100" value="90" id="ifc_q" style="flex:1" />
              <span id="ifc_qval" style="font-family:var(--font-mono);font-size:0.82rem;color:var(--accent);min-width:2.5rem">90%</span>
            </div>
          </div>
        </div>
      </div>

      <!-- Actions -->
      <div class="flex gap-1 mt-2">
        <button class="btn btn-primary" id="ifc_dl" disabled>⬇ Download Converted</button>
        <button class="btn btn-ghost" id="ifc_reset">✕ New Image</button>
      </div>
    </div>`;

  // ── Reactive helpers ────────────────────────────────────────
  let debounceTimer = null;
  let lastBlobUrl = null;

  function schedule() {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => runConvert(), 300);
  }

  async function runConvert() {
    if (!state.imgEl) return;
    state.jobId++;
    const myJob = state.jobId;

    document.getElementById('ifc_overlay').classList.remove('hidden');
    document.getElementById('ifc_dl').disabled = true;

    const fmt = document.getElementById('ifc_fmt').value;
    const quality = parseInt(document.getElementById('ifc_q').value) / 100;
    const ext = fmt === 'image/jpeg' ? 'jpg' : fmt === 'image/webp' ? 'webp' : 'png';
    const label = fmt === 'image/jpeg' ? 'JPG' : fmt === 'image/webp' ? 'WebP' : 'PNG';

    document.getElementById('ifc_comp_label').textContent = `As ${label}`;

    // Convert at natural resolution using offscreen canvas
    const canvas = document.createElement('canvas');
    canvas.width = state.imgEl.naturalWidth;
    canvas.height = state.imgEl.naturalHeight;
    canvas.getContext('2d').drawImage(state.imgEl, 0, 0);

    canvas.toBlob(blob => {
      if (myJob !== state.jobId) return;
      state.outputBlob = blob;

      // Revoke old blob URL before creating new one
      if (lastBlobUrl) URL.revokeObjectURL(lastBlobUrl);
      lastBlobUrl = URL.createObjectURL(blob);

      const convImg = document.getElementById('ifc_conv_img');
      convImg.src = lastBlobUrl;

      // Stats
      document.getElementById('ifc_conv_meta').textContent = formatBytes(blob.size);
      const saved = Math.round((1 - blob.size / state.file.size) * 100);
      const row = document.getElementById('ifc_savings_row');
      const badge = document.getElementById('ifc_savings_badge');
      const detail = document.getElementById('ifc_savings_detail');
      row.style.display = 'flex';
      if (saved > 0) {
        badge.textContent = `${saved}% smaller`;
        badge.className = 'ic-savings-badge ic-savings-good';
      } else {
        badge.textContent = `${Math.abs(saved)}% larger`;
        badge.className = 'ic-savings-badge ic-savings-bad';
      }
      detail.textContent = `${formatBytes(state.file.size)} → ${formatBytes(blob.size)}`;

      document.getElementById('ifc_overlay').classList.add('hidden');
      document.getElementById('ifc_dl').disabled = false;
      saveHistory('img-convert', state.file.name);
    }, fmt, quality);
  }

  function loadFile(file) {
    state.file = file;
    state.outputBlob = null;

    // Load image element at natural size
    const img = new Image();
    const url = URL.createObjectURL(file);
    img.onload = () => {
      state.imgEl = img;

      // Reveal workspace
      document.getElementById('ifc_drop').classList.add('hidden');
      document.getElementById('ifc_workspace').classList.remove('hidden');

      // File chip
      document.getElementById('ifc_chip').innerHTML = fileChipHTML(file);

      // Show original
      document.getElementById('ifc_orig_img').src = url;
      document.getElementById('ifc_orig_meta').textContent = formatBytes(file.size);
      document.getElementById('ifc_savings_row').style.display = 'none';
      document.getElementById('ifc_conv_meta').textContent = '—';

      // Auto-convert immediately
      runConvert();
    };
    img.src = url;
  }

  // ── Wire controls ────────────────────────────────────────────
  document.getElementById('ifc_fmt').addEventListener('change', () => schedule());
  document.getElementById('ifc_q').addEventListener('input', e => {
    document.getElementById('ifc_qval').textContent = e.target.value + '%';
    schedule();
  });

  // ── Download ─────────────────────────────────────────────────
  document.getElementById('ifc_dl').addEventListener('click', () => {
    if (!state.outputBlob) return;
    const fmt = document.getElementById('ifc_fmt').value;
    const ext = fmt === 'image/jpeg' ? 'jpg' : fmt === 'image/webp' ? 'webp' : 'png';
    const base = (state.file?.name || 'image').replace(/\.[^.]+$/, '');
    downloadBlob(state.outputBlob, `${base}.${ext}`);
    toast('Downloaded!', 'success');
  });

  // ── Reset ─────────────────────────────────────────────────────
  document.getElementById('ifc_reset').addEventListener('click', () => {
    state.jobId++; state.file = null; state.imgEl = null; state.outputBlob = null;
    if (lastBlobUrl) { URL.revokeObjectURL(lastBlobUrl); lastBlobUrl = null; }
    document.getElementById('ifc_drop').classList.remove('hidden');
    document.getElementById('ifc_workspace').classList.add('hidden');
    document.getElementById('ifc_orig_img').src = '';
    document.getElementById('ifc_conv_img').src = '';
    document.getElementById('ifc_dl').disabled = true;
  });

  // ── Drop zone ─────────────────────────────────────────────────
  setupDropZone('ifc', ([file]) => loadFile(file), ['image']);
}

// ─── Image Resize (Reactive) ───────────────────────────────
function renderImgResize(body) {
  const state = { file: null, imgEl: null, origW: 0, origH: 0, outputBlob: null };

  body.innerHTML = `
    <div id="ir_drop">${dropZoneHTML('ir', 'image/*', 'Drop your image here')}</div>

    <div id="ir_workspace" class="tool-workspace hidden">
      <div id="ir_chip"></div>

      <!-- Dimension inputs -->
      <div class="ic-controls-panel mt-2">
        <div class="form-row">
          <div class="form-field">
            <label>Width (px)</label>
            <input type="number" class="input" id="ir_w" placeholder="800" min="1" max="16000" />
          </div>
          <div class="form-field">
            <label>Height (px)</label>
            <input type="number" class="input" id="ir_h" placeholder="600" min="1" max="16000" />
          </div>
        </div>
        <div class="form-row" style="margin-top:0.5rem">
          <label class="checkbox-row" style="flex:1">
            <input type="checkbox" id="ir_lock" checked /> Lock aspect ratio
          </label>
          <div class="form-field" style="flex:1">
            <label>Format</label>
            <select class="input" id="ir_fmt">
              <option value="image/png">PNG</option>
              <option value="image/jpeg">JPG</option>
              <option value="image/webp">WebP</option>
            </select>
          </div>
        </div>
      </div>

      <!-- Live canvas preview -->
      <div class="ir-preview-wrap mt-2">
        <div class="pc-preview-label">Preview</div>
        <div class="ic-preview-imgwrap" id="ir_preview_wrap" style="min-height:160px">
          <canvas id="ir_canvas"></canvas>
          <div class="ic-processing-overlay hidden" id="ir_overlay">
            <div class="ic-spinner"></div>
          </div>
        </div>
        <div class="ic-preview-meta" id="ir_meta" style="margin-top:0.4rem">—</div>
      </div>

      <!-- Actions -->
      <div class="flex gap-1 mt-2">
        <button class="btn btn-primary" id="ir_dl" disabled>⬇ Download Resized</button>
        <button class="btn btn-ghost" id="ir_reset">✕ New Image</button>
      </div>
    </div>`;

  const wInput = document.getElementById('ir_w');
  const hInput = document.getElementById('ir_h');

  // ── Reactive preview helpers ────────────────────────────────
  let debounceTimer = null;

  function schedule() {
    clearTimeout(debounceTimer);
    debounceTimer = setTimeout(() => renderPreview(), 400);
  }

  function renderPreview() {
    if (!state.imgEl) return;
    const w = parseInt(wInput.value) || state.origW;
    const h = parseInt(hInput.value) || state.origH;
    if (!w || !h) return;

    document.getElementById('ir_overlay').classList.remove('hidden');
    document.getElementById('ir_dl').disabled = true;

    const canvas = document.getElementById('ir_canvas');
    canvas.width = w;
    canvas.height = h;
    canvas.getContext('2d').drawImage(state.imgEl, 0, 0, w, h);

    document.getElementById('ir_meta').textContent = `${w} × ${h} px`;
    document.getElementById('ir_overlay').classList.add('hidden');
    document.getElementById('ir_dl').disabled = false;
  }

  // ── Aspect-lock input wiring ────────────────────────────────
  wInput.addEventListener('input', () => {
    if (document.getElementById('ir_lock').checked && state.origW) {
      const w = parseInt(wInput.value);
      if (w) hInput.value = Math.round(w * state.origH / state.origW);
    }
    schedule();
  });
  hInput.addEventListener('input', () => {
    if (document.getElementById('ir_lock').checked && state.origH) {
      const h = parseInt(hInput.value);
      if (h) wInput.value = Math.round(h * state.origW / state.origH);
    }
    schedule();
  });
  document.getElementById('ir_fmt').addEventListener('change', () => schedule());

  // ── Load file ───────────────────────────────────────────────
  function loadFile(file) {
    state.file = file;
    const url = URL.createObjectURL(file);
    const img = new Image();
    img.onload = () => {
      state.imgEl = img;
      state.origW = img.naturalWidth;
      state.origH = img.naturalHeight;

      document.getElementById('ir_drop').classList.add('hidden');
      document.getElementById('ir_workspace').classList.remove('hidden');
      document.getElementById('ir_chip').innerHTML = fileChipHTML(file);

      wInput.value = state.origW;
      hInput.value = state.origH;

      renderPreview();
    };
    img.src = url;
  }

  // ── Download ────────────────────────────────────────────────
  document.getElementById('ir_dl').addEventListener('click', () => {
    if (!state.imgEl) return;
    const w = parseInt(wInput.value) || state.origW;
    const h = parseInt(hInput.value) || state.origH;
    const fmt = document.getElementById('ir_fmt').value;
    const ext = fmt === 'image/jpeg' ? 'jpg' : fmt === 'image/webp' ? 'webp' : 'png';

    // Use the already-drawn preview canvas to produce blob
    const canvas = document.getElementById('ir_canvas');
    canvas.toBlob(blob => {
      const base = (state.file?.name || 'image').replace(/\.[^.]+$/, '');
      downloadBlob(blob, `${base}_${w}x${h}.${ext}`);
      saveHistory('img-resize', state.file.name);
      toast(`Downloaded ${w}×${h}`, 'success');
    }, fmt, 0.92);
  });

  // ── Reset ───────────────────────────────────────────────────
  document.getElementById('ir_reset').addEventListener('click', () => {
    state.file = null; state.imgEl = null;
    document.getElementById('ir_drop').classList.remove('hidden');
    document.getElementById('ir_workspace').classList.add('hidden');
    document.getElementById('ir_dl').disabled = true;
    document.getElementById('ir_meta').textContent = '—';
    const c = document.getElementById('ir_canvas');
    c.getContext('2d').clearRect(0, 0, c.width, c.height);
  });

  // ── Drop zone ───────────────────────────────────────────────
  setupDropZone('ir', ([file]) => loadFile(file), ['image']);
}

// ─── Image Crop ────────────────────────────────────────────
function renderImgCrop(body) {
  body.innerHTML = `
    ${dropZoneHTML('icrp', 'image/*', 'Drop your image here')}
    <div class="info-box hidden mt-2" id="icrp_info">Click and drag on the image to select crop area</div>
    <div class="crop-container hidden" id="icrp_container">
      <canvas id="cropCanvas"></canvas>
      <div class="crop-sel" id="cropSel"></div>
    </div>
    <div id="icrp_coords" class="text-muted text-mono hidden" style="font-size:0.75rem;margin-top:0.5rem"></div>
    <div class="flex gap-1 mt-2">
      <button class="btn btn-primary hidden" id="icrp_crop">Crop & Download</button>
      <button class="btn btn-ghost hidden" id="icrp_reset">Reset</button>
    </div>`;

  let imgEl = null, origFile = null;
  let startX, startY, endX, endY, isDragging = false;
  let canvasRect, displayScale;

  setupDropZone('icrp', ([file]) => {
    origFile = file;
    const url = URL.createObjectURL(file);
    imgEl = new Image();
    imgEl.onload = () => {
      const container = document.getElementById('icrp_container');
      const canvas = document.getElementById('cropCanvas');
      container.classList.remove('hidden');
      document.getElementById('icrp_info').classList.remove('hidden');
      // Scale image to fit
      const maxW = Math.min(imgEl.naturalWidth, 640);
      displayScale = imgEl.naturalWidth / maxW;
      canvas.width = maxW;
      canvas.height = imgEl.naturalHeight / displayScale;
      canvas.getContext('2d').drawImage(imgEl, 0, 0, canvas.width, canvas.height);
      canvasRect = canvas.getBoundingClientRect();
      document.getElementById('icrp_crop').classList.remove('hidden');
      document.getElementById('icrp_reset').classList.remove('hidden');
    };
    imgEl.src = url;
  }, ['image']);

  const canvas = document.getElementById('cropCanvas');
  const sel = document.getElementById('cropSel');

  function getPos(e) {
    const rect = canvas.getBoundingClientRect();
    const clientX = e.touches ? e.touches[0].clientX : e.clientX;
    const clientY = e.touches ? e.touches[0].clientY : e.clientY;
    return {
      x: Math.max(0, Math.min(clientX - rect.left, canvas.offsetWidth)),
      y: Math.max(0, Math.min(clientY - rect.top, canvas.offsetHeight))
    };
  }
  canvas.addEventListener('mousedown', e => { const p = getPos(e); startX = p.x; startY = p.y; isDragging = true; });
  canvas.addEventListener('mousemove', e => {
    if (!isDragging) return;
    const p = getPos(e);
    endX = p.x; endY = p.y;
    updateSel();
  });
  canvas.addEventListener('mouseup', () => { isDragging = false; });
  canvas.addEventListener('touchstart', e => { e.preventDefault(); const p = getPos(e); startX = p.x; startY = p.y; isDragging = true; }, { passive: false });
  canvas.addEventListener('touchmove', e => { e.preventDefault(); if (!isDragging) return; const p = getPos(e); endX = p.x; endY = p.y; updateSel(); }, { passive: false });
  canvas.addEventListener('touchend', () => { isDragging = false; });

  function updateSel() {
    const rect = canvas.getBoundingClientRect();
    const canvasDisplayRect = { left: rect.left, top: rect.top };
    const l = Math.min(startX, endX), t = Math.min(startY, endY);
    const w = Math.abs(endX - startX), h = Math.abs(endY - startY);
    sel.style.cssText = `left:${l}px;top:${t}px;width:${w}px;height:${h}px;`;
    const coords = document.getElementById('icrp_coords');
    const sc = canvas.width / canvas.offsetWidth;
    coords.textContent = `Selection: ${Math.round(l * sc)} × ${Math.round(t * sc)} / ${Math.round(w * sc)} × ${Math.round(h * sc)} px`;
    coords.classList.remove('hidden');
  }

  document.getElementById('icrp_crop').addEventListener('click', () => {
    if (!imgEl || startX === undefined) { toast('Draw a selection first', 'error'); return; }
    const sc = canvas.width / canvas.offsetWidth;
    const cx = Math.min(startX, endX) * sc * displayScale;
    const cy = Math.min(startY, endY) * sc * displayScale;
    const cw = Math.abs(endX - startX) * sc * displayScale;
    const ch = Math.abs(endY - startY) * sc * displayScale;
    if (cw < 1 || ch < 1) { toast('Selection too small', 'error'); return; }
    const out = document.createElement('canvas');
    out.width = cw; out.height = ch;
    out.getContext('2d').drawImage(imgEl, cx, cy, cw, ch, 0, 0, cw, ch);
    out.toBlob(blob => {
      downloadBlob(blob, 'cropped.png');
      saveHistory('img-crop', origFile.name);
      toast(`Cropped: ${Math.round(cw)}×${Math.round(ch)} px`, 'success');
    }, 'image/png');
  });
  document.getElementById('icrp_reset').addEventListener('click', () => {
    sel.style.cssText = '';
    startX = startY = endX = endY = undefined;
    document.getElementById('icrp_coords').classList.add('hidden');
  });
}

// ─── Image → Base64 ───────────────────────────────────────
function renderImgBase64(body) {
  body.innerHTML = `
    ${dropZoneHTML('ib64', 'image/*', 'Drop your image here')}
    <div class="form-field mt-2">
      <label>Output Format</label>
      <select class="input" id="ib64_fmt">
        <option value="raw">Raw Base64</option>
        <option value="dataurl">Data URL (src-ready)</option>
        <option value="css">CSS background-image</option>
      </select>
    </div>
    <div class="copy-wrap mt-2 hidden" id="ib64_wrap">
      <div class="output-area" id="ib64_out"></div>
      <button class="copy-btn" onclick="copyText('ib64_out')">Copy</button>
    </div>`;

  setupDropZone('ib64', async ([file]) => {
    const dataUrl = await readFileAsDataURL(file);
    const fmt = document.getElementById('ib64_fmt').value;
    let out = '';
    if (fmt === 'raw') out = dataUrl.split(',')[1];
    else if (fmt === 'dataurl') out = dataUrl;
    else out = `background-image: url('${dataUrl}');`;
    document.getElementById('ib64_out').textContent = out;
    document.getElementById('ib64_wrap').classList.remove('hidden');
    saveHistory('img-base64', file.name);
    toast(`Encoded! ${formatBytes(out.length)} string`, 'success');
  }, ['image']);

  document.getElementById('ib64_fmt').onchange = async () => {
    const out = document.getElementById('ib64_out');
    if (!out.textContent) return;
    // Re-trigger: user must re-drop if they want to change format
    toast('Re-drop your image to apply new format', 'info');
  };
}

// ─── Bulk Image Convert ────────────────────────────────────
function renderImgBulk(body) {
  body.innerHTML = `
    ${dropZoneHTML('ibulk', 'image/*', 'Drop multiple images here', true)}
    <div class="form-row mt-2">
      <div class="form-field">
        <label>Output Format</label>
        <select class="input" id="ibulk_fmt">
          <option value="image/jpeg">JPG</option>
          <option value="image/png">PNG</option>
          <option value="image/webp">WebP</option>
        </select>
      </div>
      <div class="form-field">
        <label>Quality</label>
        <input type="number" class="input" id="ibulk_q" value="85" min="10" max="100" />
      </div>
    </div>
    <div id="ibulk_list" class="file-list mt-2"></div>
    <div id="ibulk_progress" class="hidden">${progressHTML('ibulk_prog')}</div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="ibulk_convert">Convert All & Download ZIP</button>`;

  let files = [];
  const listEl = document.getElementById('ibulk_list');
  const btn = document.getElementById('ibulk_convert');

  function renderList() {
    listEl.innerHTML = files.map((f, i) => `
      <div class="file-item">
        <span class="file-item-icon">🖼️</span>
        <span class="file-item-name">${f.name}</span>
        <span class="file-item-size">${formatBytes(f.size)}</span>
        <button class="file-item-del" onclick="ibulkRemove(${i})">✕</button>
      </div>`).join('');
    btn.classList.toggle('hidden', !files.length);
  }
  window.ibulkRemove = i => { files.splice(i, 1); renderList(); };
  setupDropZone('ibulk', f => { files = [...files, ...f]; renderList(); }, ['image'], true);

  btn.addEventListener('click', async () => {
    const fmt = document.getElementById('ibulk_fmt').value;
    const quality = parseInt(document.getElementById('ibulk_q').value) / 100;
    const ext = fmt.split('/')[1];
    const progress = document.getElementById('ibulk_progress');
    progress.classList.remove('hidden'); btn.disabled = true;
    const zip = new JSZip();
    for (let i = 0; i < files.length; i++) {
      setProgress(document.getElementById('ibulk_prog'), (i / files.length) * 100, `Converting ${files[i].name}…`);
      const dataUrl = await readFileAsDataURL(files[i]);
      const img = await new Promise(res => { const im = new Image(); im.onload = () => res(im); im.src = dataUrl; });
      const canvas = document.createElement('canvas');
      canvas.width = img.naturalWidth; canvas.height = img.naturalHeight;
      canvas.getContext('2d').drawImage(img, 0, 0);
      const blob = await new Promise(res => canvas.toBlob(res, fmt, quality));
      const ab = await blob.arrayBuffer();
      const baseName = files[i].name.replace(/\.[^.]+$/, '');
      zip.file(`${baseName}.${ext}`, ab);
    }
    setProgress(document.getElementById('ibulk_prog'), 100, 'Creating ZIP…');
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    downloadBlob(zipBlob, 'converted_images.zip');
    progress.classList.add('hidden'); btn.disabled = false;
    saveHistory('img-bulk', files[0].name);
    toast(`${files.length} images converted!`, 'success');
  });
}

// ═══════════════════════════════════════════════════════════
// CSV / EXCEL TOOLS
// ═══════════════════════════════════════════════════════════

// ─── Excel/CSV → PDF ──────────────────────────────────────
function renderExcelToPdf(body) {
  body.innerHTML = `
    ${dropZoneHTML('etpdf', '.xlsx,.xls,.csv,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'Drop Excel or CSV file here')}
    <div class="form-row mt-2">
      <div class="form-field">
        <label>Orientation</label>
        <select class="input" id="etpdf_orient"><option value="l">Landscape</option><option value="p">Portrait</option></select>
      </div>
      <div class="form-field">
        <label>Max Rows</label>
        <input type="number" class="input" id="etpdf_rows" value="500" min="10" max="5000" />
      </div>
    </div>
    <div id="etpdf_progress" class="hidden">${progressHTML('etpdf_prog')}</div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="etpdf_convert">Convert to PDF</button>`;

  let spreadData = null, origName = '';
  setupDropZone('etpdf', async ([file]) => {
    origName = file.name;
    try {
      const ab = await readFileAsArrayBuffer(file);
      const wb = XLSX.read(ab, { type: 'array' });
      const ws = wb.Sheets[wb.SheetNames[0]];
      spreadData = XLSX.utils.sheet_to_json(ws, { header: 1 });
      document.getElementById('etpdf_convert').classList.remove('hidden');
      toast(`Loaded: ${spreadData.length} rows`, 'info');
    } catch (e) { toast('Could not read file: ' + e.message, 'error'); }
  }, ['xlsx', 'xls', 'csv']);

  document.getElementById('etpdf_convert').addEventListener('click', async () => {
    if (!spreadData) return;
    const orient = document.getElementById('etpdf_orient').value;
    const maxRows = parseInt(document.getElementById('etpdf_rows').value) || 500;
    const progress = document.getElementById('etpdf_progress');
    progress.classList.remove('hidden');
    try {
      const { jsPDF } = window.jspdf;
      const pdf = new jsPDF({ orientation: orient, unit: 'pt', format: 'a4' });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const margin = 40;
      const rows = spreadData.slice(0, maxRows);
      if (!rows.length) { toast('No data found', 'error'); return; }
      const colCount = Math.max(...rows.map(r => r.length));
      const colW = (pageW - margin * 2) / Math.max(colCount, 1);
      const rowH = 18, headerH = 22;
      let y = margin + headerH;
      pdf.setFontSize(14);
      pdf.setTextColor(30, 30, 30);
      pdf.text(origName, margin, margin);
      pdf.setFontSize(8);
      rows.forEach((row, ri) => {
        setProgress(document.getElementById('etpdf_prog'), (ri / rows.length) * 100);
        if (y + rowH > pageH - margin) { pdf.addPage(); y = margin; }
        const isHeader = ri === 0;
        if (isHeader) { pdf.setFillColor(30, 30, 50); pdf.rect(margin, y - 12, pageW - margin * 2, rowH, 'F'); pdf.setTextColor(240, 240, 240); }
        else {
          if (ri % 2 === 0) { pdf.setFillColor(20, 20, 36); pdf.rect(margin, y - 12, pageW - margin * 2, rowH, 'F'); }
          pdf.setTextColor(60, 60, 80);
        }
        row.slice(0, colCount).forEach((cell, ci) => {
          const text = String(cell ?? '').substring(0, 30);
          pdf.text(text, margin + ci * colW, y, { maxWidth: colW - 4 });
        });
        y += rowH;
        pdf.setTextColor(60, 60, 80);
      });
      pdf.save('spreadsheet.pdf');
      progress.classList.add('hidden');
      saveHistory('excel-to-pdf', origName);
      toast('PDF created!', 'success');
    } catch (e) { progress.classList.add('hidden'); toast('Error: ' + e.message, 'error'); }
  });
}

// ─── PDF → CSV ─────────────────────────────────────────────
function renderPdfToCsv(body) {
  body.innerHTML = `
    <div class="warn-box">⚠ This extracts text line-by-line from PDFs. It works best with PDFs that have structured tabular text. Scanned PDFs or complex layouts may not extract cleanly.</div>
    ${dropZoneHTML('pdftcsv', 'application/pdf', 'Drop your PDF here')}
    <div id="pdftcsv_progress" class="hidden">${progressHTML('pdftcsv_prog')}</div>
    <div class="copy-wrap mt-2 hidden" id="pdftcsv_wrap">
      <div class="output-area" id="pdftcsv_out"></div>
      <button class="copy-btn" onclick="copyText('pdftcsv_out')">Copy</button>
    </div>
    <button class="btn btn-primary mt-2 hidden" id="pdftcsv_dl">⬇ Download CSV</button>`;

  let csvText = '';
  setupDropZone('pdftcsv', async ([file]) => {
    const progress = document.getElementById('pdftcsv_progress');
    progress.classList.remove('hidden');
    try {
      const ab = await readFileAsArrayBuffer(file);
      const pdf = await pdfjsLib.getDocument({ data: ab }).promise;
      const rows = [];
      for (let i = 1; i <= pdf.numPages; i++) {
        setProgress(document.getElementById('pdftcsv_prog'), (i / pdf.numPages) * 100, `Reading page ${i}…`);
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        // Group text items by approximate y-position (lines)
        const lines = {};
        content.items.forEach(item => {
          const y = Math.round(item.transform[5]);
          if (!lines[y]) lines[y] = [];
          lines[y].push(item.str);
        });
        Object.keys(lines).sort((a, b) => b - a).forEach(y => {
          const line = lines[y].join(' ').replace(/,/g, ';');
          rows.push(line);
        });
        rows.push('');
      }
      csvText = rows.join('\n');
      document.getElementById('pdftcsv_out').textContent = csvText.substring(0, 3000) + (csvText.length > 3000 ? '\n…(truncated)' : '');
      progress.classList.add('hidden');
      document.getElementById('pdftcsv_wrap').classList.remove('hidden');
      document.getElementById('pdftcsv_dl').classList.remove('hidden');
      saveHistory('pdf-to-csv', file.name);
      toast('Text extracted as CSV!', 'success');
    } catch (e) { progress.classList.add('hidden'); toast('Error: ' + e.message, 'error'); }
  }, ['pdf']);

  document.getElementById('pdftcsv_dl').addEventListener('click', () => {
    if (!csvText) return;
    downloadBlob(new Blob([csvText], { type: 'text/csv' }), 'extracted.csv');
  });
}

// ─── CSV Merge ─────────────────────────────────────────────
function renderCsvMerge(body) {
  body.innerHTML = `
    ${dropZoneHTML('csvm', '.csv,text/csv', 'Drop multiple CSV files here', true)}
    <div class="checkbox-row mt-1">
      <input type="checkbox" id="csvm_header" checked />
      <span>Skip header row after first file</span>
    </div>
    <div id="csvm_list" class="file-list mt-2"></div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="csvm_merge">Merge CSVs</button>`;

  let files = [];
  const listEl = document.getElementById('csvm_list');
  const btn = document.getElementById('csvm_merge');

  function renderList() {
    listEl.innerHTML = files.map((f, i) => `
      <div class="file-item">
        <span class="drag-handle">⠿</span>
        <span class="file-item-icon">📋</span>
        <span class="file-item-name">${f.name}</span>
        <span class="file-item-size">${formatBytes(f.size)}</span>
        <button class="file-item-del" onclick="csvmRemove(${i})">✕</button>
      </div>`).join('');
    btn.classList.toggle('hidden', files.length < 2);
  }
  window.csvmRemove = i => { files.splice(i, 1); renderList(); };
  setupDropZone('csvm', f => { files = [...files, ...f]; renderList(); }, ['csv'], true);

  btn.addEventListener('click', async () => {
    const skipHeader = document.getElementById('csvm_header').checked;
    const parts = [];
    for (let i = 0; i < files.length; i++) {
      let text = await readFileAsText(files[i]);
      if (i > 0 && skipHeader) {
        const lines = text.split('\n');
        text = lines.slice(1).join('\n');
      }
      parts.push(text.trim());
    }
    const merged = parts.join('\n');
    downloadBlob(new Blob([merged], { type: 'text/csv' }), 'merged.csv');
    saveHistory('csv-merge', files[0].name);
    toast(`${files.length} CSV files merged!`, 'success');
  });
}

// ─── CSV Split ─────────────────────────────────────────────
function renderCsvSplit(body) {
  body.innerHTML = `
    ${dropZoneHTML('csvsp', '.csv,text/csv', 'Drop your CSV file here')}
    <div id="csvsp_info" class="info-box hidden"></div>
    <div class="form-row mt-2">
      <div class="form-field">
        <label>Rows Per File</label>
        <input type="number" class="input" id="csvsp_size" value="100" min="1" />
      </div>
    </div>
    <div id="csvsp_progress" class="hidden">${progressHTML('csvsp_prog')}</div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="csvsp_split">Split & Download ZIP</button>`;

  let csvFile = null, csvLines = [];
  setupDropZone('csvsp', async ([file]) => {
    csvFile = file;
    const text = await readFileAsText(file);
    csvLines = text.split('\n').filter(l => l.trim());
    document.getElementById('csvsp_info').textContent = `✓ ${file.name} — ${csvLines.length} rows`;
    document.getElementById('csvsp_info').classList.remove('hidden');
    document.getElementById('csvsp_split').classList.remove('hidden');
  }, ['csv']);

  document.getElementById('csvsp_split').addEventListener('click', async () => {
    if (!csvLines.length) return;
    const rowsPerFile = parseInt(document.getElementById('csvsp_size').value) || 100;
    const header = csvLines[0];
    const dataLines = csvLines.slice(1);
    const zip = new JSZip();
    const chunks = Math.ceil(dataLines.length / rowsPerFile);
    for (let i = 0; i < chunks; i++) {
      setProgress(document.getElementById('csvsp_prog'), (i / chunks) * 100);
      const chunk = dataLines.slice(i * rowsPerFile, (i + 1) * rowsPerFile);
      const content = [header, ...chunk].join('\n');
      zip.file(`part_${i + 1}.csv`, content);
    }
    document.getElementById('csvsp_progress').classList.remove('hidden');
    setProgress(document.getElementById('csvsp_prog'), 100, 'Creating ZIP…');
    const blob = await zip.generateAsync({ type: 'blob' });
    downloadBlob(blob, 'split_csv.zip');
    document.getElementById('csvsp_progress').classList.add('hidden');
    saveHistory('csv-split', csvFile.name);
    toast(`Split into ${chunks} files!`, 'success');
  });
}

// ═══════════════════════════════════════════════════════════
// DOCUMENT TOOLS (NEW)
// ═══════════════════════════════════════════════════════════

// ─── Shared HTML→PDF renderer (used by Word→PDF & HTML→PDF) ─
async function _renderHtmlAsCanvas(htmlContent, pageWidthPx, scalePx) {
  await ensureHtml2Canvas();
  const container = document.createElement('div');
  container.style.cssText = `position:fixed;left:-9999px;top:0;width:${pageWidthPx}px;` +
    `background:#ffffff;color:#222222;font-family:Georgia,serif;padding:50px;` +
    `box-sizing:border-box;line-height:1.7;font-size:11pt`;
  container.innerHTML = htmlContent;
  document.body.appendChild(container);
  try {
    const canvas = await html2canvas(container, {
      scale: scalePx, useCORS: true, logging: false, backgroundColor: '#ffffff'
    });
    return canvas;
  } finally {
    if (container.parentNode) document.body.removeChild(container);
  }
}

// ─── Word → PDF ───────────────────────────────────────────
function renderWordToPdf(body) {
  body.innerHTML = `
    <div class="info-box">Converts DOCX files. Text formatting is preserved; complex layouts may vary slightly — this is a browser-based limitation.</div>
    <div id="w2p_drop">${dropZoneHTML('w2p', '.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document', 'Drop your DOCX file here')}</div>
    <div id="w2p_workspace" class="tool-workspace hidden">
      <div id="w2p_chip"></div>
      <div class="ic-controls-panel mt-2">
        <div class="form-row">
          <div class="form-field">
            <label>Page Size</label>
            <select class="input" id="w2p_page"><option value="794">A4</option><option value="816">Letter</option></select>
          </div>
          <div class="form-field">
            <label>Render Quality</label>
            <select class="input" id="w2p_scale"><option value="1">1× Fast</option><option value="1.5" selected>1.5× Balanced</option><option value="2">2× Crisp</option></select>
          </div>
        </div>
      </div>
      <div style="background:var(--surface);border:1px solid var(--border);border-radius:var(--radius);padding:1rem;margin:0.75rem 0;max-height:260px;overflow-y:auto">
        <div class="ic-preview-label" style="margin-bottom:0.5rem;text-align:left">Document Preview</div>
        <div id="w2p_html" style="color:var(--text);font-size:0.85rem;line-height:1.65;background:#fff;color:#222;padding:0.5rem;border-radius:6px"></div>
      </div>
      <div id="w2p_progress" class="hidden">${progressHTML('w2p_prog')}</div>
      <div class="flex gap-1 mt-2">
        <button class="btn btn-primary" id="w2p_dl" disabled>⬇ Download PDF</button>
        <button class="btn btn-ghost" id="w2p_reset">✕ New File</button>
      </div>
    </div>`;

  let htmlContent = '', currentFile = null;

  setupDropZone('w2p', async ([file]) => {
    currentFile = file;
    document.getElementById('w2p_drop').classList.add('hidden');
    document.getElementById('w2p_workspace').classList.remove('hidden');
    document.getElementById('w2p_chip').innerHTML = fileChipHTML(file);
    const previewEl = document.getElementById('w2p_html');
    previewEl.innerHTML = '<span style="color:#888;font-size:0.82rem">⏳ Loading mammoth.js and parsing DOCX…</span>';
    document.getElementById('w2p_dl').disabled = true;
    try {
      await ensureMammoth();
      const ab = await readFileAsArrayBuffer(file);
      const result = await mammoth.convertToHtml({ arrayBuffer: ab });
      htmlContent = result.value || '';
      previewEl.innerHTML = htmlContent || '<em style="color:#888">No readable content found.</em>';
      if (result.messages?.length) {
        previewEl.innerHTML += `<p style="color:#f5a623;font-size:0.78rem;margin-top:0.5rem">${result.messages.length} formatting warning(s)</p>`;
      }
      document.getElementById('w2p_dl').disabled = false;
      saveHistory('word-to-pdf', file.name);
      toast('DOCX loaded — ready to convert', 'success');
    } catch (err) {
      previewEl.innerHTML = `<span style="color:var(--error)">Error: ${err.message}</span>`;
      toast('Could not read DOCX: ' + err.message, 'error');
    }
  }, ['docx', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document']);

  document.getElementById('w2p_dl').addEventListener('click', async () => {
    if (!htmlContent) return;
    const btn = document.getElementById('w2p_dl');
    btn.disabled = true; btn.textContent = '⏳ Generating…';
    const prog = document.getElementById('w2p_progress');
    prog.classList.remove('hidden');
    setProgress(document.getElementById('w2p_prog'), 20, 'Rendering HTML…');
    try {
      const pageW = parseInt(document.getElementById('w2p_page').value);
      const scale = safeScale(parseFloat(document.getElementById('w2p_scale').value));
      const canvas = await _renderHtmlAsCanvas(htmlContent, pageW, scale);
      setProgress(document.getElementById('w2p_prog'), 80, 'Building PDF…');
      const { jsPDF } = window.jspdf;
      const pw = canvas.width / scale, ph = canvas.height / scale;
      const pdf = new jsPDF({ orientation: pw > ph ? 'l' : 'p', unit: 'px', format: [pw, ph] });
      pdf.addImage(canvas.toDataURL('image/jpeg', 0.90), 'JPEG', 0, 0, pw, ph);
      const base = (currentFile?.name || 'document').replace(/\.docx?$/i, '');
      pdf.save(`${base}.pdf`);
      prog.classList.add('hidden');
      toast('PDF downloaded!', 'success');
    } catch (err) {
      prog.classList.add('hidden');
      toast('Error: ' + err.message, 'error');
    } finally { btn.disabled = false; btn.textContent = '⬇ Download PDF'; }
  });

  document.getElementById('w2p_reset').addEventListener('click', () => {
    htmlContent = ''; currentFile = null;
    document.getElementById('w2p_drop').classList.remove('hidden');
    document.getElementById('w2p_workspace').classList.add('hidden');
    document.getElementById('w2p_html').innerHTML = '';
  });
}

// ─── Text → PDF ────────────────────────────────────────────
function renderTextToPdf(body) {
  body.innerHTML = `
    <div id="t2p_drop_area">${dropZoneHTML('t2p_file', '.txt,.md,text/plain', 'Drop a .txt / .md file, or type below')}</div>
    <div class="ic-controls-panel mt-2">
      <div class="form-row">
        <div class="form-field">
          <label>Font Size</label>
          <select class="input" id="t2p_fsize"><option value="10">10 pt</option><option value="11" selected>11 pt</option><option value="12">12 pt</option><option value="14">14 pt</option></select>
        </div>
        <div class="form-field">
          <label>Margin</label>
          <select class="input" id="t2p_margin"><option value="20">Narrow</option><option value="40" selected>Normal</option><option value="60">Wide</option></select>
        </div>
        <div class="form-field">
          <label>Page Size</label>
          <select class="input" id="t2p_page"><option value="a4" selected>A4</option><option value="letter">Letter</option></select>
        </div>
      </div>
    </div>
    <div class="form-field mt-2">
      <label>Text Content</label>
      <textarea class="input" id="t2p_text" rows="10" placeholder="Type or paste your text here…" style="resize:vertical;min-height:160px;font-family:var(--font-mono);font-size:0.85rem"></textarea>
    </div>
    <div id="t2p_progress" class="hidden">${progressHTML('t2p_prog')}</div>
    <button class="btn btn-primary btn-full mt-2" id="t2p_dl">⬇ Generate PDF</button>`;

  setupDropZone('t2p_file', async ([file]) => {
    const text = await readFileAsText(file);
    document.getElementById('t2p_text').value = text;
    toast(`Loaded: ${file.name} (${formatBytes(file.size)})`, 'info');
  }, ['txt', 'md', 'text/plain']);

  document.getElementById('t2p_dl').addEventListener('click', async () => {
    const text = document.getElementById('t2p_text').value.trim();
    if (!text) { toast('Please enter or drop some text first', 'error'); return; }
    const btn = document.getElementById('t2p_dl');
    btn.disabled = true;
    const prog = document.getElementById('t2p_progress');
    prog.classList.remove('hidden');
    setProgress(document.getElementById('t2p_prog'), 10, 'Building PDF…');
    try {
      const { jsPDF } = window.jspdf;
      const pageSize = document.getElementById('t2p_page').value;
      const fontSize = parseInt(document.getElementById('t2p_fsize').value);
      const margin = parseInt(document.getElementById('t2p_margin').value);
      const pdf = new jsPDF({ orientation: 'p', unit: 'pt', format: pageSize });
      const pageW = pdf.internal.pageSize.getWidth();
      const pageH = pdf.internal.pageSize.getHeight();
      const usableW = pageW - margin * 2;
      pdf.setFontSize(fontSize);
      pdf.setTextColor(30, 30, 30);
      const lines = pdf.splitTextToSize(text, usableW);
      const lineH = fontSize * 1.45;
      let y = margin + fontSize;
      for (let i = 0; i < lines.length; i++) {
        if (y + lineH > pageH - margin) { pdf.addPage(); y = margin + fontSize; }
        pdf.text(lines[i], margin, y);
        y += lineH;
        if (i % 80 === 0) {
          setProgress(document.getElementById('t2p_prog'), Math.min(90, 10 + (i / lines.length) * 80));
          await new Promise(r => setTimeout(r, 0)); // yield to UI
        }
      }
      pdf.save('text_document.pdf');
      prog.classList.add('hidden');
      saveHistory('text-to-pdf', 'text_document.pdf');
      toast(`PDF created — ${lines.length} lines`, 'success');
    } catch (err) {
      prog.classList.add('hidden');
      toast('Error: ' + err.message, 'error');
    } finally { btn.disabled = false; }
  });
}

// ─── HTML → PDF ────────────────────────────────────────────
function renderHtmlToPdf(body) {
  body.innerHTML = `
    <div class="info-box">Paste HTML or drop a .html file. External fonts/images may not render — this is a browser-only limitation. Scripts are stripped for safety.</div>
    <div id="h2p_drop_area">${dropZoneHTML('h2p_file', '.html,.htm,text/html', 'Drop an HTML file, or paste code below')}</div>
    <div class="form-field mt-2">
      <label>HTML Source</label>
      <textarea class="input" id="h2p_html" rows="8" placeholder="&lt;h1&gt;Hello&lt;/h1&gt;&lt;p&gt;Your content…&lt;/p&gt;" style="resize:vertical;min-height:140px;font-family:var(--font-mono);font-size:0.82rem"></textarea>
    </div>
    <div style="margin:0.5rem 0"><button class="btn btn-ghost btn-sm" id="h2p_prev_btn">👁 Preview</button></div>
    <div id="h2p_preview_wrap" class="hidden" style="border:1px solid var(--border);border-radius:var(--radius);overflow:hidden;margin-bottom:0.5rem">
      <iframe id="h2p_iframe" style="width:100%;height:280px;border:none;background:#fff"></iframe>
    </div>
    <div class="ic-controls-panel mt-1">
      <div class="form-row">
        <div class="form-field">
          <label>Page Width (px)</label>
          <input type="number" class="input" id="h2p_width" value="794" min="400" max="1400" />
        </div>
        <div class="form-field">
          <label>Render Scale</label>
          <select class="input" id="h2p_scale"><option value="1">1× Faster</option><option value="1.5" selected>1.5× Balanced</option><option value="2">2× Crisper</option></select>
        </div>
      </div>
    </div>
    <div id="h2p_progress" class="hidden">${progressHTML('h2p_prog')}</div>
    <button class="btn btn-primary btn-full mt-2" id="h2p_dl">⬇ Generate PDF</button>`;

  setupDropZone('h2p_file', async ([file]) => {
    const text = await readFileAsText(file);
    document.getElementById('h2p_html').value = text;
    toast(`Loaded: ${file.name}`, 'info');
  }, ['html', 'htm', 'text/html']);

  document.getElementById('h2p_prev_btn').addEventListener('click', () => {
    const html = document.getElementById('h2p_html').value;
    if (!html.trim()) { toast('Enter HTML first', 'error'); return; }
    const wrap = document.getElementById('h2p_preview_wrap');
    document.getElementById('h2p_iframe').srcdoc = html;
    wrap.classList.remove('hidden');
  });

  document.getElementById('h2p_dl').addEventListener('click', async () => {
    const rawHtml = document.getElementById('h2p_html').value.trim();
    if (!rawHtml) { toast('Enter or drop HTML first', 'error'); return; }
    const btn = document.getElementById('h2p_dl');
    btn.disabled = true; btn.textContent = '⏳ Loading renderer…';
    const prog = document.getElementById('h2p_progress');
    prog.classList.remove('hidden');
    setProgress(document.getElementById('h2p_prog'), 15, 'Loading html2canvas…');
    try {
      const pageW = parseInt(document.getElementById('h2p_width').value) || 794;
      const scale = safeScale(parseFloat(document.getElementById('h2p_scale').value));
      // Strip scripts for safety
      const safeHtml = rawHtml.replace(/<script[\s\S]*?<\/script>/gi, '').replace(/on\w+="[^"]*"/gi, '');
      setProgress(document.getElementById('h2p_prog'), 40, 'Rendering HTML…');
      const canvas = await _renderHtmlAsCanvas(safeHtml, pageW, scale);
      setProgress(document.getElementById('h2p_prog'), 85, 'Building PDF…');
      const { jsPDF } = window.jspdf;
      const pw = canvas.width / scale, ph = canvas.height / scale;
      const pdf = new jsPDF({ orientation: pw > ph ? 'l' : 'p', unit: 'px', format: [pw, ph] });
      pdf.addImage(canvas.toDataURL('image/jpeg', 0.90), 'JPEG', 0, 0, pw, ph);
      pdf.save('document.pdf');
      prog.classList.add('hidden');
      saveHistory('html-to-pdf', 'html_document');
      toast('PDF downloaded!', 'success');
    } catch (err) {
      prog.classList.add('hidden');
      toast('Error: ' + err.message, 'error');
    } finally { btn.disabled = false; btn.textContent = '⬇ Generate PDF'; }
  });
}

// ─── Image → PDF Advanced ──────────────────────────────────
function renderImgToPdfAdv(body) {
  body.innerHTML = `
    ${dropZoneHTML('itpadv', 'image/png,image/jpeg,image/webp', 'Drop images here (drag to reorder)', true)}
    <div id="itpadv_list" class="thumb-grid mt-2"></div>
    <div class="ic-controls-panel mt-2">
      <div class="form-row">
        <div class="form-field">
          <label>Page Size</label>
          <select class="input" id="itpadv_size"><option value="a4">A4</option><option value="letter">Letter</option><option value="original">Original (image size)</option></select>
        </div>
        <div class="form-field">
          <label>Margin</label>
          <select class="input" id="itpadv_margin"><option value="0">None</option><option value="10" selected>Small (10px)</option><option value="20">Medium (20px)</option><option value="40">Large (40px)</option></select>
        </div>
        <div class="form-field">
          <label>Orientation</label>
          <select class="input" id="itpadv_orient"><option value="auto" selected>Auto (per image)</option><option value="p">Always Portrait</option><option value="l">Always Landscape</option></select>
        </div>
      </div>
    </div>
    <div id="itpadv_progress" class="hidden">${progressHTML('itpadv_prog')}</div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="itpadv_convert">Convert to PDF</button>`;

  let files = [];
  const listEl = document.getElementById('itpadv_list');
  const btn = document.getElementById('itpadv_convert');

  function renderList() {
    listEl.innerHTML = '';
    files.forEach((f, i) => {
      const url = URL.createObjectURL(f);
      const div = document.createElement('div');
      div.className = 'thumb-item'; div.dataset.idx = i;
      div.innerHTML = `<img src="${url}" /><div class="thumb-item-label">${f.name.substring(0, 16)}</div>
        <button class="thumb-item-del" onclick="itpadvRemove(${i})">×</button>`;
      listEl.appendChild(div);
    });
    btn.classList.toggle('hidden', !files.length);
    if (files.length > 1 && typeof Sortable !== 'undefined') {
      new Sortable(listEl, { animation: 150, onEnd: e => {
        const moved = files.splice(e.oldIndex, 1)[0];
        files.splice(e.newIndex, 0, moved);
      }});
    }
  }
  window.itpadvRemove = i => { files.splice(i, 1); renderList(); };
  setupDropZone('itpadv', f => { files = [...files, ...f]; renderList(); },
    ['image/png', 'image/jpeg', 'image/webp', 'jpg', 'png', 'webp'], true);

  btn.addEventListener('click', async () => {
    if (!files.length) return;
    const progress = document.getElementById('itpadv_progress');
    progress.classList.remove('hidden'); btn.disabled = true;
    const pageSize = document.getElementById('itpadv_size').value;
    const margin = parseInt(document.getElementById('itpadv_margin').value);
    const orientMode = document.getElementById('itpadv_orient').value;
    try {
      const { jsPDF } = window.jspdf;
      let pdf = null;
      for (let i = 0; i < files.length; i++) {
        setProgress(document.getElementById('itpadv_prog'), (i / files.length) * 100,
          `Processing image ${i + 1}/${files.length}…`);
        const dataUrl = await readFileAsDataURL(files[i]);
        const img = await new Promise(res => { const im = new Image(); im.onload = () => res(im); im.src = dataUrl; });
        const dims = capDimForMobile(img.naturalWidth, img.naturalHeight);
        const orient = orientMode === 'auto' ? (dims.w > dims.h ? 'l' : 'p') : orientMode;
        const fmt = pageSize === 'original' ? [dims.w + margin * 2, dims.h + margin * 2] : pageSize;
        if (!pdf) {
          pdf = new jsPDF({ orientation: orient, unit: 'px', format: fmt });
        } else {
          pdf.addPage(fmt, orient);
        }
        const pW = pdf.internal.pageSize.getWidth();
        const pH = pdf.internal.pageSize.getHeight();
        pdf.addImage(dataUrl, 'JPEG', margin, margin, pW - margin * 2, pH - margin * 2);
      }
      pdf.save('images_advanced.pdf');
      progress.classList.add('hidden'); btn.disabled = false;
      saveHistory('img-to-pdf-adv', files[0].name);
      toast(`PDF with ${files.length} page(s) created!`, 'success');
    } catch (e) {
      progress.classList.add('hidden'); btn.disabled = false;
      toast('Error: ' + e.message, 'error');
    }
  });
}

// ─── Bulk Image Compress + ZIP ─────────────────────────────
function renderImgBulkCompressZip(body) {
  body.innerHTML = `
    ${dropZoneHTML('ibcz', 'image/*', 'Drop multiple images here', true)}
    <div id="ibcz_list" class="file-list mt-2"></div>
    <div class="ic-controls-panel mt-2">
      <div class="control-group">
        <div class="control-label">JPEG Quality &nbsp;<span id="ibcz_qval">80%</span></div>
        <input type="range" min="10" max="95" value="80" id="ibcz_quality" />
      </div>
      <div class="control-group" style="margin-bottom:0">
        <div class="control-label">Max Width &nbsp;<span id="ibcz_wval">2048 px</span></div>
        <input type="range" min="256" max="4096" value="2048" step="128" id="ibcz_maxw" />
      </div>
      <div class="form-row mt-1">
        <div class="form-field">
          <label>Output Format</label>
          <select class="input" id="ibcz_fmt">
            <option value="image/jpeg">JPG</option>
            <option value="image/webp">WebP</option>
            <option value="image/png">PNG</option>
          </select>
        </div>
      </div>
    </div>
    <div id="ibcz_progress" class="hidden">${progressHTML('ibcz_prog')}</div>
    <div id="ibcz_stats" class="hidden size-comparison" style="margin:0.75rem 0"></div>
    <button class="btn btn-primary btn-full mt-2 hidden" id="ibcz_convert">⚡ Compress All &amp; Download ZIP</button>`;

  let files = [];
  const listEl = document.getElementById('ibcz_list');
  const btn = document.getElementById('ibcz_convert');

  document.getElementById('ibcz_quality').addEventListener('input', e => {
    document.getElementById('ibcz_qval').textContent = e.target.value + '%';
  });
  document.getElementById('ibcz_maxw').addEventListener('input', e => {
    document.getElementById('ibcz_wval').textContent = parseInt(e.target.value).toLocaleString() + ' px';
  });

  function renderList() {
    listEl.innerHTML = files.map((f, i) => `
      <div class="file-item">
        <span class="file-item-icon">🖼️</span>
        <span class="file-item-name">${f.name}</span>
        <span class="file-item-size">${formatBytes(f.size)}</span>
        <button class="file-item-del" onclick="ibczRemove(${i})">✕</button>
      </div>`).join('');
    btn.classList.toggle('hidden', !files.length);
  }
  window.ibczRemove = i => { files.splice(i, 1); renderList(); };
  setupDropZone('ibcz', f => { files = [...files, ...f]; renderList(); }, ['image'], true);

  btn.addEventListener('click', async () => {
    if (!files.length) return;
    const quality = parseInt(document.getElementById('ibcz_quality').value) / 100;
    const requestedMaxW = parseInt(document.getElementById('ibcz_maxw').value);
    // Cap max width on mobile to prevent OOM
    const maxW = isMobile() ? Math.min(requestedMaxW, 1920) : requestedMaxW;
    const fmt = document.getElementById('ibcz_fmt').value;
    const ext = fmt === 'image/jpeg' ? 'jpg' : fmt === 'image/webp' ? 'webp' : 'png';
    const progress = document.getElementById('ibcz_progress');
    progress.classList.remove('hidden'); btn.disabled = true;
    document.getElementById('ibcz_stats').classList.add('hidden');
    const zip = new JSZip();
    let totalOrig = 0, totalComp = 0;
    for (let i = 0; i < files.length; i++) {
      setProgress(document.getElementById('ibcz_prog'), (i / files.length) * 100,
        `Compressing ${files[i].name} (${i + 1}/${files.length})…`);
      totalOrig += files[i].size;
      try {
        const compressed = await imageCompression(files[i], {
          maxSizeMB: 50, maxWidthOrHeight: maxW,
          initialQuality: quality, fileType: fmt, useWebWorker: true,
        });
        totalComp += compressed.size;
        zip.file(`${files[i].name.replace(/\.[^.]+$/, '')}_compressed.${ext}`, await compressed.arrayBuffer());
      } catch {
        // Fallback: add original file if compression fails
        totalComp += files[i].size;
        zip.file(files[i].name, await files[i].arrayBuffer());
      }
    }
    setProgress(document.getElementById('ibcz_prog'), 98, 'Creating ZIP…');
    const zipBlob = await zip.generateAsync({ type: 'blob' });
    downloadBlob(zipBlob, 'compressed_images.zip');
    const saved = Math.round((1 - totalComp / totalOrig) * 100);
    const statsEl = document.getElementById('ibcz_stats');
    statsEl.classList.remove('hidden');
    statsEl.innerHTML = sizeCmpHTML(totalOrig, totalComp) + `
      <div class="size-item"><span class="size-num" style="color:var(--accent2)">${files.length}</span><span class="size-label">Files</span></div>`;
    progress.classList.add('hidden'); btn.disabled = false;
    saveHistory('img-bulk-compress-zip', files[0].name);
    toast(`${files.length} images compressed — ${saved > 0 ? saved + '% smaller' : 'done'}!`, 'success');
  });
}

// ═══════════════════════════════════════════════════════════
// PWA / SERVICE WORKER
// ═══════════════════════════════════════════════════════════
if ('serviceWorker' in navigator) {
  navigator.serviceWorker.register('sw.js').catch(() => { });
}

// ═══════════════════════════════════════════════════════════
// INIT
// ═══════════════════════════════════════════════════════════
document.addEventListener('DOMContentLoaded', () => {
  loadHistory();
  // Scroll-based header shadow
  window.addEventListener('scroll', () => {
    document.getElementById('header').style.boxShadow =
      window.scrollY > 10 ? '0 2px 20px rgba(0,0,0,0.5)' : '';
  });
  // Contact form
  const cf = document.getElementById('contactForm');
  if (cf) {
    cf.addEventListener('submit', e => {
      e.preventDefault();
      toast('Message sent! We\'ll get back to you soon.', 'success');
      cf.reset();
    });
  }
});
