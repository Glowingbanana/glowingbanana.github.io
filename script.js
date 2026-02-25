/* ========= Config ========= */
const CONFIG = {
  NORMALIZE_CURRENCY_TO: null, // set to 'SGD' to map "Singapore Dollar" -> "SGD"
  DATE_FORMAT: 'dd/mm/yyyy',   // exported display format (we store true dates)
  EXCLUDE_DRAFT_DEFAULT: false,
  MIN_DIGITAL_TEXT_LEN: 20,    // fallback to OCR if fewer chars than this
  OCR_SCALE: 2,                // canvas scale for OCR rendering
  NUMBER_TOLERANCE: 0.02       // reconciliation tolerance when computing totals
};

/* ========= Globals ========= */
let selectedFile = null;
let extractedDataRaw = ''; // for preview only
let exportRows = [];       // final rows for Option A (one row per line item)
let ocrReady = false;

/* ========= PDF.js Worker Configuration ========= */
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/* ========= DOM Elements ========= */
const uploadArea = document.getElementById('uploadArea');
const fileInput  = document.getElementById('fileInput');
const forceOCRCb = document.getElementById('forceOCR');
const excludeDraftCb = document.getElementById('excludeDraft');

const statusEl   = document.getElementById('status');
const fileInfo   = document.getElementById('fileInfo');
const fileNameEl = document.getElementById('fileName');
const fileSizeEl = document.getElementById('fileSize');

const preview    = document.getElementById('preview');
const previewContent = document.getElementById('previewContent');

const btnConvert = document.getElementById('btnConvert');
const btnDownload = document.getElementById('btnDownload');

const textInput  = document.getElementById('textInput');

/* ========= Tesseract OCR Worker ========= */
const ocrWorker = Tesseract.createWorker({
  workerPath: 'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/worker.min.js',
  corePath:   'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.0.0/tesseract-core.wasm.js',
  langPath:   'https://tessdata.projectnaptha.com/4.0.0',
  logger: m => {
    if (m && m.status && typeof m.progress === 'number') {
      showStatus(`<span class="spinner"></span>${m.status} ${(m.progress * 100).toFixed(0)}%`, 'loading');
    }
  }
});

async function ensureOCR() {
  if (ocrReady) return;
  await ocrWorker.load();
  await ocrWorker.loadLanguage('eng');
  await ocrWorker.initialize('eng');
  ocrReady = true;
}

/* ========= UI: Upload / Drag & Drop ========= */
uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('dragover', (e) => {
  e.preventDefault();
  uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', () => {
  uploadArea.classList.remove('dragover');
});
uploadArea.addEventListener('drop', (e) => {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
  const files = e.dataTransfer.files;
  if (files && files.length > 0) handleFile(files[0]);
});
fileInput.addEventListener('change', (e) => {
  if (e.target.files && e.target.files.length > 0) handleFile(e.target.files[0]);
});

function handleFile(file) {
  const isPdfMime = file.type === 'application/pdf';
  const isPdfName = /\.pdf$/i.test(file.name);
  if (!isPdfMime && !isPdfName) {
    showStatus('Please select a PDF file', 'error');
    return;
  }

  selectedFile = file;

  fileNameEl.textContent = `File: ${file.name}`;
  const sizeMB = (file.size / 1024 / 1024);
  fileSizeEl.textContent = `Size: ${sizeMB >= 0.1 ? sizeMB.toFixed(2) + ' MB' : (file.size/1024).toFixed(0) + ' KB'}`;
  fileInfo.classList.add('show');

  btnConvert.disabled = false;
  btnDownload.style.display = 'none';
  preview.classList.remove('show');
  extractedDataRaw = '';
  exportRows = [];

  excludeDraftCb.checked = CONFIG.EXCLUDE_DRAFT_DEFAULT;

  showStatus(`Ready to convert: ${file.name}`, 'success');
}

function showStatus(message, type = 'loading') {
  statusEl.innerHTML = message;
  statusEl.className = `status show ${type}`;
}
function clearStatus() {
  statusEl.textContent = '';
  statusEl.className = 'status';
}

function clearFile() {
  selectedFile = null;
  extractedDataRaw = '';
  exportRows = [];

  fileInput.value = '';
  fileInfo.classList.remove('show');

  btnConvert.disabled = true;
  btnDownload.style.display = 'none';

  preview.classList.remove('show');
  clearStatus();
}
window.clearFile = clearFile;

/* ========= Convert ========= */
async function convertPDF() {
  if (selectedFile) {
    await convertSelectedPDF();
    return;
  }

  const pasted = (textInput.value || '').trim();
  if (pasted.length > 0) {
    const rows = await extractRowsFromText(pasted, { excludeDraft: excludeDraftCb.checked });
    exportRows = rows;
    extractedDataRaw = pasted;

    const previewText = pasted.substring(0, 800) + (pasted.length > 800 ? '…' : '');
    previewContent.textContent = previewText;
    preview.classList.add('show');

    btnDownload.style.display = 'block';
    showStatus('Text is ready to export. Click "Download Excel File".', 'success');
    return;
  }

  showStatus('Please select a PDF or paste text, then click Convert.', 'error');
}
window.convertPDF = convertPDF;

async function convertSelectedPDF() {
  if (!selectedFile) {
    showStatus('Please select a PDF file first', 'error');
    return;
  }

  btnConvert.disabled = true;
  showStatus('<span class="spinner"></span>Opening PDF…', 'loading');

  try {
    const arrayBuffer = await selectedFile.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    const invoices = new Map(); // invoiceNo -> { header:{...}, items:[], totals:{...}, gstRate }
    let currentInvoiceNo = null;

    extractedDataRaw = '';
    const useForceOCR = !!forceOCRCb.checked;

    for (let i = 1; i <= pdf.numPages; i++) {
      showStatus(`<span class="spinner"></span>Processing page ${i} of ${pdf.numPages}…`, 'loading');

      const page = await pdf.getPage(i);

      // Try digital text first
      let pageText = '';
      if (!useForceOCR) {
        const textContent = await page.getTextContent();
        pageText = textContent.items.map(it => it.str).join(' ').replace(/\s+/g, ' ').trim();
      }

      const needsOCR = useForceOCR || !pageText || pageText.length < CONFIG.MIN_DIGITAL_TEXT_LEN;

      if (needsOCR) {
        const viewport = page.getViewport({ scale: CONFIG.OCR_SCALE });
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d', { willReadFrequently: true });
        canvas.width  = Math.floor(viewport.width);
        canvas.height = Math.floor(viewport.height);
        await page.render({ canvasContext: context, viewport }).promise;

        await ensureOCR();
        const { data: { text } } = await ocrWorker.recognize(canvas);
        pageText = (text || '').replace(/\s+/g, ' ').trim();
      }

      // Keep a human preview
      extractedDataRaw += `\n\n--- Page ${i} ---\n` + pageText;

      // Normalize
      const norm = normalize(pageText);

      // 1) Detect invoice number (start/switch context)
      const invNo = findInvoiceNo(norm);
      if (invNo) currentInvoiceNo = invNo;

      if (!currentInvoiceNo) {
        continue;
      }

      const inv = getOrCreateInvoice(invoices, currentInvoiceNo);

      // 2) Header fields
      const hdr = extractHeaderFields(norm);
      assignHeader(inv.header, hdr);

      // 3) GST rate
      const gstRate = findGSTRate(norm);
      if (gstRate != null) inv.gstRate = gstRate;

      // 4) Line items
      const items = extractLineItems(norm);
      if (items.length) inv.items.push(...items);

      // 5) Totals
      const totals = extractTotals(norm);
      assignTotals(inv.totals, totals);
    }

    // Build rows
    const excludeDraft = !!excludeDraftCb.checked;
    exportRows = [];

    invoices.forEach((inv, invNo) => {
      const st = (inv.header.invoiceStatus || '').toLowerCase();
      if (excludeDraft && st === 'draft') return;

      if (!hasAllTotals(inv.totals)) computeTotals(inv);

      if (CONFIG.NORMALIZE_CURRENCY_TO === 'SGD' && inv.totals.currency) {
        inv.totals.currency = 'SGD';
      }

      // Emit option A rows
      for (const li of inv.items) {
        exportRows.push([
          inv.header.vendorId || '',
          inv.header.attentionTo || '',
          toExcelDate(inv.header.invoiceDate) || inv.header.invoiceDate || '',
          inv.header.creditTerm || '',
          inv.header.invoiceNo || invNo,
          inv.header.relatedInvoiceNo || '',
          inv.header.invoiceStatus || '',
          inv.header.instructionId || '',
          inv.header.headerDescription || '',

          li.lineNo ?? '',
          li.description || '',
          toNumber(li.quantity),
          toNumber(li.unitPrice),
          toNumber(li.grossEx),
          li.gstRate != null ? toNumber(li.gstRate) : (inv.gstRate != null ? inv.gstRate : ''),
          toNumber(li.grossInc),

          inv.totals.currency || '',
          toNumber(inv.totals.subtotal),
          toNumber(inv.totals.gst),
          toNumber(inv.totals.freight),
          toNumber(inv.totals.total)
        ]);
      }
    });

    const previewText = extractedDataRaw.substring(0, 800) + (extractedDataRaw.length > 800 ? '…' : '');
    previewContent.textContent = previewText;
    preview.classList.add('show');

    btnDownload.style.display = exportRows.length ? 'block' : 'none';
    if (exportRows.length) {
      showStatus(`Parsed ${exportRows.length} line item row(s). Click "Download Excel File".`, 'success');
    } else {
      showStatus('No line items were found. Try enabling Force OCR or check the PDF content.', 'error');
    }

  } catch (err) {
    console.error(err);
    showStatus(`Error: ${err.message}`, 'error');
  } finally {
    btnConvert.disabled = false;
  }
}

/* ========= Download Excel (Option A: one sheet, one row per line item) ========= */
function downloadExcel() {
  if (!exportRows || !exportRows.length) {
    showStatus('No data to download', 'error');
    return;
  }

  try {
    const headers = [
      'Vendor ID',
      'Attention To',
      'Invoice Date',
      'Credit Term',
      'Invoice No',
      'Related Invoice No',
      'Invoice Status',
      'Invoicing Instruction ID',
      'Description',                    // header/top description
      'No.',                             // line number
      'Description',                     // line-item description (beside No.)
      'Quantity',
      'Unit Price',
      'Gross Amt (Ex. GST)',
      'GST @ %',
      'Gross Amt (Inc. GST)',
      'Currency',
      'Sub Total (Excluding GST)',
      'Total GST Payable',
      'Freight Amount',
      'Total Invoice Amount'
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...exportRows]);

    ws['!cols'] = [
      { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 10 }, { wch: 18 },
      { wch: 18 }, { wch: 12 }, { wch: 20 }, { wch: 70 }, { wch: 6  },
      { wch: 70 }, { wch: 10 }, { wch: 12 }, { wch: 16 }, { wch: 8  },
      { wch: 18 }, { wch: 16 }, { wch: 20 }, { wch: 18 }, { wch: 16 }, { wch: 20 }
    ];
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };

    XLSX.utils.book_append_sheet(wb, ws, 'Invoice Lines');

    const base = selectedFile
      ? selectedFile.name.replace(/\.pdf$/i, '')
      : 'pasted_text';
    const outName = `${base}_OptionA.xlsx`;

    XLSX.writeFile(wb, outName);
    showStatus(`Downloaded: ${outName}`, 'success');
  } catch (err) {
    console.error(err);
    showStatus(`Error downloading: ${err.message}`, 'error');
  }
}
window.downloadExcel = downloadExcel;

/* ========= Helpers: parsing & normalization ========= */

function normalize(s) {
  if (!s) return '';
  let t = s.replace(/\u00A0/g, ' ');
  t = t.replace(/\s*:\s*/g, ' : ');
  t = t.replace(/\s+/g, ' ').trim();
  return t;
}

function toNumber(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') return v;
  const s = String(v).replace(/,/g, '');
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : '';
}

function toExcelDate(dmy) {
  // expects dd/mm/yyyy or d/m/yyyy; returns JS Date or ''
  if (!dmy || typeof dmy !== 'string') return '';
  const m = dmy.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (!m) return '';
  let [, d, mth, y] = m;
  if (y.length === 2) y = '20' + y;
  const date = new Date(parseInt(y, 10), parseInt(mth, 10) - 1, parseInt(d, 10));
  return isNaN(date.getTime()) ? '' : date;
}

function getOrCreateInvoice(map, invoiceNo) {
  if (!map.has(invoiceNo)) {
    map.set(invoiceNo, {
      header: { invoiceNo },
      items: [],
      totals: {},
      gstRate: null
    });
  }
  return map.get(invoiceNo);
}

/* ---- Header extraction ---- */
function findInvoiceNo(text) {
  const m = text.match(/Invoice\s*No\s*:\s*([A-Z0-9-]+)/i);
  return m ? m[1] : null;
}

function extractHeaderFields(text) {
  const out = {};

  const id = text.match(/Vendor\s*ID\s*:\s*([A-Z0-9]+)/i);
  if (id) out.vendorId = id[1];

  const att = text.match(/Attention\s*To\s*:\s*([A-Za-z0-9 ,.'()/-]+)/i);
  if (att) out.attentionTo = att[1].trim();

  const dt = text.match(/Invoice\s*Date\s*:\s*([0-3]?\d[\/\-][01]?\d[\/\-]\d{2,4})/i);
  if (dt) out.invoiceDate = dt[1];

  const term = text.match(/Credit\s*Term\s*:\s*([A-Za-z0-9 ]+)/i);
  if (term) out.creditTerm = term[1].trim();

  const inv = text.match(/Invoice\s*No\s*:\s*([A-Z0-9-]+)/i);
  if (inv) out.invoiceNo = inv[1];

  const rel = text.match(/Related\s*Invoice\s*No\s*:\s*([A-Z0-9-]+)/i);
  if (rel) out.relatedInvoiceNo = rel[1];

  const st = text.match(/Invoice\s*Status\s*:\s*([A-Za-z]+)/i);
  if (st) out.invoiceStatus = st[1];

  const instr = text.match(/Invoicing\s*Instruction\s*ID\s*:\s*([A-Z0-9-]+)/i);
  if (instr) out.instructionId = instr[1];

  const desc = text.match(/Description\s*:\s*(.+?)(?=\s(?:No\.\s|Currency\s*:|Invoice\s*Amount\s*Summary|Sub\s*Total|Total\s*GST|Freight\s*Amount|Total\s*Invoice\s*Amount|$))/i);
  if (desc) out.headerDescription = desc[1].trim();

  return out;
}

function assignHeader(target, src) {
  if (!src) return;
  for (const k of Object.keys(src)) {
    if (target[k] == null || target[k] === '') {
      target[k] = src[k];
    }
  }
}

/* ---- GST rate ---- */
function findGSTRate(text) {
  const m = text.match(/GST\s*@\s*(\d{1,2})\s*%/i);
  if (!m) return null;
  return parseFloat(m[1]);
}

/* ---- Line items extraction ---- */
function extractLineItems(text) {
  const items = [];
  const rx = /(?:^|\s)(\d{1,3})\s+(.+?)\s+(?:0\s+)?(\d+(?:\.\d{1,5})?)\s+([0-9][0-9,]*\.\d{2})\s+([0-9][0-9,]*\.\d{2})\s+([0-9][0-9,]*\.\d{2})\s+([0-9][0-9,]*\.\d{2})(?=\s|$)/gi;

  let m;
  while ((m = rx.exec(text)) !== null) {
    const lineNo = parseInt(m[1], 10);
    const description = m[2].trim();
    const quantity = m[3];
    const unitPrice = m[4];
    const grossEx = m[5];
    const gstAmt  = m[6];
    const grossInc = m[7];

    if (description.length < 5) continue;

    let gstRate = null;
    const ex = toNumber(grossEx);
    const g  = toNumber(gstAmt);
    if (Number.isFinite(ex) && ex > 0 && Number.isFinite(g)) {
      gstRate = Math.round((g / ex) * 1000) / 10; // 1 decimal
    }

    items.push({
      lineNo,
      description,
      quantity,
      unitPrice,
      grossEx,
      gstAmount: gstAmt,
      grossInc,
      gstRate
    });
  }

  return items;
}

/* ---- Totals extraction ---- */
function extractTotals(text) {
  const out = {};
  const cur = text.match(/Currency\s*:\s*([A-Za-z ]+)/i);
  if (cur) out.currency = cur[1].trim();

  const sub = text.match(/Sub\s*Total\s*\(Excluding\s*GST\)\s*:\s*([0-9][0-9,]*\.\d{2})/i);
  if (sub) out.subtotal = sub[1];

  const gst = text.match(/Total\s*GST\s*Payable\s*:\s*([0-9][0-9,]*\.\d{2})/i);
  if (gst) out.gst = gst[1];

  const fr = text.match(/Freight\s*Amount\s*:\s*([0-9][0-9,]*\.\d{2})/i);
  if (fr) out.freight = fr[1];

  const tot = text.match(/Total\s*Invoice\s*Amount\s*:\s*([0-9][0-9,]*\.\d{2})/i);
  if (tot) out.total = tot[1];

  return out;
}

function assignTotals(target, src) {
  if (!src) return;
  for (const k of Object.keys(src)) {
    if (!target[k]) target[k] = src[k];
  }
}

function hasAllTotals(t) {
  return t && t.currency && t.subtotal && t.gst && (t.freight != null) && t.total;
}

/* ---- Compute totals when missing ---- */
function computeTotals(inv) {
  let subtotal = 0;
  let gst = 0;
  let total = 0;

  for (const li of inv.items) {
    const ex = toNumber(li.grossEx);
    const g  = toNumber(li.gstAmount);
    const inc = toNumber(li.grossInc);

    if (Number.isFinite(ex)) subtotal += ex;
    if (Number.isFinite(g))  gst += g;
    if (Number.isFinite(inc)) total += inc;
  }

  inv.totals.subtotal = subtotal.toFixed(2);
  inv.totals.gst      = gst.toFixed(2);
  inv.totals.freight  = inv.totals.freight ? inv.totals.freight : '0.00';
  inv.totals.total    = total > 0 ? total.toFixed(2)
                                  : (subtotal + gst + toNumber(inv.totals.freight || 0)).toFixed(2);

  if (!inv.totals.currency) inv.totals.currency = 'Singapore Dollar'; // default
}

/* ========= Pasted text handler ========= */
async function extractRowsFromText(text, { excludeDraft }) {
  const invoices = new Map();
  let currentInvoiceNo = null;

  const norm = normalize(text);

  // Split into pseudo-pages on headers or totals markers to reduce accidental cross-matching
  const chunks = norm.split(/\s(?:Tax\s*Invoice|Invoice\s*Amount\s*Summary)\s/i);

  for (const chunk of chunks) {
    if (!chunk || chunk.length < 10) continue;

    const invNo = findInvoiceNo(chunk);
    if (invNo) currentInvoiceNo = invNo;
    if (!currentInvoiceNo) continue;

    const inv = getOrCreateInvoice(invoices, currentInvoiceNo);

    const hdr = extractHeaderFields(chunk);
    assignHeader(inv.header, hdr);

    const gstRate = findGSTRate(chunk);
    if (gstRate != null) inv.gstRate = gstRate;

    const items = extractLineItems(chunk);
    if (items.length) inv.items.push(...items);

    const totals = extractTotals(chunk);
    assignTotals(inv.totals, totals);
  }

  const rows = [];
  invoices.forEach((inv, invNo) => {
    const st = (inv.header.invoiceStatus || '').toLowerCase();
    if (excludeDraft && st === 'draft') return;

    if (!hasAllTotals(inv.totals)) computeTotals(inv);
    if (CONFIG.NORMALIZE_CURRENCY_TO === 'SGD' && inv.totals.currency) inv.totals.currency = 'SGD';

    for (const li of inv.items) {
      rows.push([
        inv.header.vendorId || '',
        inv.header.attentionTo || '',
        toExcelDate(inv.header.invoiceDate) || inv.header.invoiceDate || '',
        inv.header.creditTerm || '',
        inv.header.invoiceNo || invNo,
        inv.header.relatedInvoiceNo || '',
        inv.header.invoiceStatus || '',
        inv.header.instructionId || '',
        inv.header.headerDescription || '',

        li.lineNo ?? '',
        li.description || '',
        toNumber(li.quantity),
        toNumber(li.unitPrice),
        toNumber(li.grossEx),
        li.gstRate != null ? toNumber(li.gstRate) : (inv.gstRate != null ? inv.gstRate : ''),
        toNumber(li.grossInc),

        inv.totals.currency || '',
        toNumber(inv.totals.subtotal),
        toNumber(inv.totals.gst),
        toNumber(inv.totals.freight),
        toNumber(inv.totals.total)
      ]);
    }
  });

  return rows;
}
