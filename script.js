/* ========= Config ========= */
const CONFIG = {
  NORMALIZE_CURRENCY_TO: null, // set to 'SGD' if you want "Singapore Dollar" normalized
  DATE_FORMAT: 'dd/mm/yyyy',
  EXCLUDE_DRAFT_DEFAULT: false,
  MIN_DIGITAL_TEXT_LEN: 20,
  OCR_SCALE: 2,
  NUMBER_TOLERANCE: 0.02
};

/* ========= Globals ========= */
let selectedFile = null;
let extractedDataRaw = ''; // preview only
let exportRows = []; // final rows (one per line item)
let ocrReady = false;

/* ========= PDF.js Worker ========= */
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/* ========= DOM ========= */
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const forceOCRCb = document.getElementById('forceOCR');
const excludeDraftCb = document.getElementById('excludeDraft');
const statusEl = document.getElementById('status');
const fileInfo = document.getElementById('fileInfo');
const fileNameEl = document.getElementById('fileName');
const fileSizeEl = document.getElementById('fileSize');
const preview = document.getElementById('preview');
const previewContent = document.getElementById('previewContent');
const btnConvert = document.getElementById('btnConvert');
const btnDownload = document.getElementById('btnDownload');
const textInput = document.getElementById('textInput');

/* ========= Tesseract OCR Worker ========= */
const ocrWorker = Tesseract.createWorker({
  workerPath: 'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/worker.min.js',
  corePath: 'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.0.0/tesseract-core.wasm.js',
  langPath: 'https://tessdata.projectnaptha.com/4.0.0',
  logger: m => {
    if (m && m.status && typeof m.progress === 'number') {
      showStatus(`${m.status} ${(m.progress * 100).toFixed(0)}%`, 'loading');
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

/* ========= UI ========= */
uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('dragover', e => {
  e.preventDefault();
  uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', () =>
  uploadArea.classList.remove('dragover')
);
uploadArea.addEventListener('drop', e => {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
  const f = e.dataTransfer.files?.[0];
  if (f) handleFile(f);
});

fileInput.addEventListener('change', e => {
  const f = e.target.files?.[0];
  if (f) handleFile(f);
});

function handleFile(file) {
  const isPdf =
    file.type === 'application/pdf' || /\.pdf$/i.test(file.name);
  if (!isPdf) return showStatus('Please select a PDF file', 'error');

  selectedFile = file;
  fileNameEl.textContent = `File: ${file.name}`;
  const sizeMB = file.size / 1024 / 1024;
  fileSizeEl.textContent =
    sizeMB >= 0.1
      ? `Size: ${sizeMB.toFixed(2)} MB`
      : `Size: ${(file.size / 1024).toFixed(0)} KB`;

  fileInfo.classList.add('show');
  btnConvert.disabled = false;
  btnDownload.style.display = 'none';
  preview.classList.remove('show');
  extractedDataRaw = '';
  exportRows = [];
  excludeDraftCb.checked = CONFIG.EXCLUDE_DRAFT_DEFAULT;
  showStatus(`Ready to convert: ${file.name}`, 'success');
}

function showStatus(msg, type = 'loading') {
  statusEl.innerHTML = msg;
  statusEl.className = `status show ${type}`;
}

/* ========= Convert ========= */
async function convertPDF() {
  if (!selectedFile) return;

  btnConvert.disabled = true;
  showStatus('Opening PDF…', 'loading');

  try {
    const arrayBuffer = await selectedFile.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    extractedDataRaw = '';
    exportRows = [];

    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      let pageText = '';

      if (!forceOCRCb.checked) {
        const tc = await page.getTextContent();
        pageText = tc.items.map(x => x.str).join(' ').replace(/\s+/g, ' ');
      }

      if (
        forceOCRCb.checked ||
        !pageText ||
        pageText.length < CONFIG.MIN_DIGITAL_TEXT_LEN
      ) {
        const viewport = page.getViewport({ scale: CONFIG.OCR_SCALE });
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d', { willReadFrequently: true });
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        await page.render({ canvasContext: ctx, viewport }).promise;

        await ensureOCR();
        const {
          data: { text }
        } = await ocrWorker.recognize(canvas);
        pageText = text.replace(/\s+/g, ' ');
      }

      extractedDataRaw += `\n\n--- Page ${i} ---\n${pageText}`;

      const items = extractLineItems(pageText);
      for (const li of items) {
        exportRows.push([
          '', '', '', '', '', '', '', '', '',        // header columns untouched
          li.lineNo,                               // No.
          li.description,                          // Description (line)
          toNumber(li.quantity),
          toNumber(li.unitPrice),
          toNumber(li.grossEx),
          toNumber(li.gstAmount),
          toNumber(li.grossInc),
          '', '', ''
        ]);
      }
    }

    previewContent.textContent =
      extractedDataRaw.substring(0, 800) +
      (extractedDataRaw.length > 800 ? '…' : '');
    preview.classList.add('show');
    btnDownload.style.display = exportRows.length ? 'block' : 'none';

    showStatus(
      exportRows.length
        ? `Parsed ${exportRows.length} line item row(s).`
        : 'No line items found.',
      exportRows.length ? 'success' : 'error'
    );
  } catch (err) {
    console.error(err);
    showStatus(`Error: ${err.message}`, 'error');
  } finally {
    btnConvert.disabled = false;
  }
}

/* ========= Download ========= */
function downloadExcel() {
  if (!exportRows.length)
    return showStatus('No data to download', 'error');

  const headers = [
    'Vendor ID',
    'Attention To',
    'Invoice Date',
    'Credit Term',
    'Invoice No',
    'Related Invoice No',
    'Invoice Status',
    'Invoicing Instruction ID',
    'Description',
    'No.',
    'Description',
    'Quantity',
    'Unit Price',
    'Gross Amt (EX. GST)',
    'GST @ 9%',
    'Gross Amt (Inc. GST)',
    'Currency',
    'Sub Total (Excluding GST)',
    'Total GST Payable'
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...exportRows]);
  XLSX.utils.book_append_sheet(wb, ws, 'Invoice Lines');

  const outName = selectedFile.name.replace(/\.pdf$/i, '') + '.xlsx';
  XLSX.writeFile(wb, outName);
}

/* ========= Helpers ========= */
function toNumber(v) {
  if (v === '' || v == null) return '';
  const n = parseFloat(String(v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : '';
}

/* ===========================================================
   ✅ ONLY FIXED FUNCTION – NOTHING ELSE CHANGED
   =========================================================== */
function extractLineItems(text) {
  const items = [];
  const rx =
    /(?:^|\s)(\d{1,3})\s+(.+?)\s+([0-9,]+\.\d{2,5})\s+0\s+([0-9.]+)\s+([0-9,]+\.\d{2})\s+([0-9,]+\.\d{2})\s+([0-9,]+\.\d{2})(?=\s|$)/g;

  let m;
  while ((m = rx.exec(text)) !== null) {
    items.push({
      lineNo: parseInt(m[1], 10),
      description: m[2].trim(),
      unitPrice: m[3],
      quantity: m[4],
      grossEx: m[5],
      gstAmount: m[6],
      grossInc: m[7]
    });
  }
  return items;
}
