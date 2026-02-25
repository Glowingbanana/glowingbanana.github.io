/* ========= Globals ========= */
let selectedFile = null;
let extractedData = null; // string of all extracted text (joined by \n)
let ocrReady = false;

/* ========= PDF.js Worker Configuration ========= */
// Must match the version in index.html
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/* ========= DOM Elements ========= */
const uploadArea = document.getElementById('uploadArea');
const fileInput  = document.getElementById('fileInput');
const forceOCRCb = document.getElementById('forceOCR');

const statusEl   = document.getElementById('status');
const fileInfo   = document.getElementById('fileInfo');
const fileNameEl = document.getElementById('fileName');
const fileSizeEl = document.getElementById('fileSize');

const preview    = document.getElementById('preview');
const previewContent = document.getElementById('previewContent');

const btnConvert = document.getElementById('btnConvert');
const btnDownload = document.getElementById('btnDownload');

const textInput  = document.getElementById('textInput');

/* ========= Tesseract OCR Worker =========
   We configure explicit paths for GitHub Pages stability.
   Language path hosts *.traineddata(.gz) files.
*/
const ocrWorker = Tesseract.createWorker({
  workerPath: 'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/worker.min.js',
  corePath:   'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.0.0/tesseract-core.wasm.js',
  langPath:   'https://tessdata.projectnaptha.com/4.0.0',
  logger: m => {
    // Show progress during OCR
    if (m && m.status && typeof m.progress === 'number') {
      showStatus(`<span class="spinner"></span>${m.status} ${(m.progress * 100).toFixed(0)}%`, 'loading');
    }
  }
});

async function ensureOCR() {
  if (ocrReady) return;
  await ocrWorker.load();
  await ocrWorker.loadLanguage('eng');   // If you need additional languages, tell me which ones.
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

/* ========= Helpers ========= */
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
  extractedData = null;

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

/* ========= Clear ========= */
function clearFile() {
  selectedFile = null;
  extractedData = null;

  fileInput.value = '';
  fileInfo.classList.remove('show');

  btnConvert.disabled = true;
  btnDownload.style.display = 'none';

  preview.classList.remove('show');
  clearStatus();

  // Keep pasted text as-is; user may still want to convert it.
}
window.clearFile = clearFile; // Expose to inline onclick

/* ========= Convert ========= */
async function convertPDF() {
  // 1) If a PDF is selected -> process PDF
  if (selectedFile) {
    await convertSelectedPDF();
    return;
  }

  // 2) Else, if pasted text exists -> convert that directly
  const pasted = (textInput.value || '').trim();
  if (pasted.length > 0) {
    extractedData = pasted;
    const previewText = pasted.substring(0, 800) + (pasted.length > 800 ? '…' : '');
    previewContent.textContent = previewText;
    preview.classList.add('show');

    btnDownload.style.display = 'block';
    showStatus('Text is ready to export. Click "Download Excel File".', 'success');
    return;
  }

  // 3) Nothing to convert
  showStatus('Please select a PDF or paste text, then click Convert.', 'error');
}
window.convertPDF = convertPDF; // Expose to inline onclick

async function convertSelectedPDF() {
  if (!selectedFile) {
    showStatus('Please select a PDF file first', 'error');
    return;
  }

  btnConvert.disabled = true;
  showStatus('<span class="spinner"></span>Opening PDF…', 'loading');

  try {
    const arrayBuffer = await selectedFile.arrayBuffer();

    // IMPORTANT: pass as {data} or a Uint8Array
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    let fullText = '';
    const useForceOCR = !!forceOCRCb.checked;

    for (let i = 1; i <= pdf.numPages; i++) {
      showStatus(`<span class="spinner"></span>Processing page ${i} of ${pdf.numPages}…`, 'loading');

      const page = await pdf.getPage(i);

      let pageText = '';
      let needsOCR = useForceOCR;

      if (!useForceOCR) {
        // Try digital extraction first
        const textContent = await page.getTextContent();
        pageText = textContent.items.map(it => it.str).join(' ').replace(/\s+/g, ' ').trim();

        // Heuristic: if too little text, treat as scanned
        if (!pageText || pageText.length < 20) needsOCR = true;
      }

      if (needsOCR) {
        showStatus(`<span class="spinner"></span>OCR page ${i} of ${pdf.numPages}… (may take a moment)`, 'loading');

        // Render to canvas for OCR
        const viewport = page.getViewport({ scale: 2 }); // 2~2.5 is a good balance
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d', { willReadFrequently: true });

        canvas.width  = Math.floor(viewport.width);
        canvas.height = Math.floor(viewport.height);

        await page.render({ canvasContext: context, viewport }).promise;

        // Ensure OCR worker is ready once
        await ensureOCR();

        const { data: { text } } = await ocrWorker.recognize(canvas);
        pageText = (text || '').replace(/\s+\n/g, '\n').trim();
      }

      if (pageText) {
        fullText += pageText + '\n';
      }
    }

    extractedData = fullText;

    // Show preview (first ~800 chars)
    const previewText = fullText.substring(0, 800) + (fullText.length > 800 ? '…' : '');
    previewContent.textContent = previewText;
    preview.classList.add('show');

    btnDownload.style.display = 'block';
    showStatus('PDF converted successfully! Click "Download Excel File".', 'success');

  } catch (err) {
    console.error(err);
    showStatus(`Error: ${err.message}`, 'error');
  } finally {
    btnConvert.disabled = false;
  }
}

/* ========= Download Excel ========= */
function downloadExcel() {
  if (!extractedData) {
    showStatus('No data to download', 'error');
    return;
  }

  try {
    // Split text by lines and map to rows (one column)
    const lines = extractedData.split('\n').map(s => s.trim()).filter(Boolean);
    const data = lines.map(line => [line]);

    // Create workbook & sheet
    const wb = XLSX.utils.book_new();
    const header = [['Extracted Text from PDF'], []]; // title + blank row
    const ws = XLSX.utils.aoa_to_sheet([...header, ...data]);

    // Column width
    ws['!cols'] = [{ wch: 60 }];

    // Append sheet
    XLSX.utils.book_append_sheet(wb, ws, 'PDF Data');

    // Filename
    const base = selectedFile
      ? selectedFile.name.replace(/\.pdf$/i, '')
      : 'pasted_text';
    const outName = `${base}_converted.xlsx`;

    // Download
    XLSX.writeFile(wb, outName);
    showStatus(`Downloaded: ${outName}`, 'success');
  } catch (err) {
    console.error(err);
    showStatus(`Error downloading: ${err.message}`, 'error');
  }
}
window.downloadExcel = downloadExcel; // Expose to inline onclick

/* ========= Optional: Clean up OCR worker on unload ========= */
window.addEventListener('beforeunload', async () => {
  try {
    if (ocrReady) {
      await ocrWorker.terminate();
      ocrReady = false;
    }
  } catch {}
});
