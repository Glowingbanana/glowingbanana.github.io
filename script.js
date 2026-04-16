/* ========= Config ========= */
const CONFIG = {
    DATE_FORMAT: 'dd/mm/yyyy',
    EXCLUDE_DRAFT_DEFAULT: false,
    MIN_DIGITAL_TEXT_LEN: 30,
    OCR_SCALE: 2.5
};

/* ========= Globals ========= */
let selectedFile = null;
let extractedDataRaw = '';
let exportRows = [];
let ocrReady = false;

/* ========= PDF.js Worker ========= */
pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/* ========= DOM ========= */
const uploadArea    = document.getElementById('uploadArea');
const fileInput     = document.getElementById('fileInput');
const forceOCRCb    = document.getElementById('forceOCR');
const excludeDraftCb= document.getElementById('excludeDraft');
const statusEl      = document.getElementById('status');
const fileInfo      = document.getElementById('fileInfo');
const fileNameEl    = document.getElementById('fileName');
const fileSizeEl    = document.getElementById('fileSize');
const preview       = document.getElementById('preview');
const previewContent= document.getElementById('previewContent');
const btnConvert    = document.getElementById('btnConvert');
const btnDownload   = document.getElementById('btnDownload');

/* ========= OCR ========= */
const ocrWorker = Tesseract.createWorker({
    workerPath : 'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/worker.min.js',
    corePath   : 'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.0.0/tesseract-core.wasm.js',
    langPath   : 'https://tessdata.projectnaptha.com/4.0.0',
    logger: m => {
        if (m?.status && typeof m.progress === 'number') {
            showStatus(`OCR: ${m.status} ${(m.progress * 100).toFixed(0)}%`, 'loading');
        }
    }
});

async function ensureOCR() {
    if (ocrReady) return;
    showStatus('Loading OCR engine…', 'loading');
    await ocrWorker.load();
    await ocrWorker.loadLanguage('eng');
    await ocrWorker.initialize('eng');
    ocrReady = true;
}

/* ========= UI Events ========= */
uploadArea.addEventListener('click', () => fileInput.click());
uploadArea.addEventListener('dragover', e => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
});
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
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
    if (!(file.type === 'application/pdf' || /\.pdf$/i.test(file.name))) {
        return showStatus('Please select a PDF file.', 'error');
    }
    selectedFile = file;
    fileNameEl.textContent = `File: ${file.name}`;
    fileSizeEl.textContent = `Size: ${(file.size / 1024 / 1024).toFixed(2)} MB`;
    fileInfo.classList.add('show');
    btnConvert.disabled = false;
    btnDownload.style.display = 'none';
    preview.classList.remove('show');
    extractedDataRaw = '';
    exportRows = [];
    excludeDraftCb.checked = CONFIG.EXCLUDE_DRAFT_DEFAULT;
    showStatus(`Ready: ${file.name}`, 'success');
}

function showStatus(msg, type = 'loading') {
    statusEl.innerHTML = msg;
    statusEl.className = `status show ${type}`;
}

function clearFile() {
    selectedFile = null;
    extractedDataRaw = '';
    exportRows = [];
    fileInfo.classList.remove('show');
    preview.classList.remove('show');
    btnConvert.disabled = true;
    btnDownload.style.display = 'none';
    fileInput.value = '';
    const ta = document.getElementById('textInput');
    if (ta) ta.value = '';
    statusEl.className = 'status';
}

/* ========= MAIN CONVERT ========= */
async function convertPDF() {
    if (!selectedFile) return showStatus('No file selected.', 'error');

    btnConvert.disabled = true;
    exportRows = [];
    extractedDataRaw = '';
    showStatus('Opening PDF…', 'loading');

    try {
        const arrayBuffer = await selectedFile.arrayBuffer();
        const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
        const total = pdf.numPages;

        for (let i = 1; i <= total; i++) {
            showStatus(`Processing page ${i} of ${total}…`, 'loading');
            const page = await pdf.getPage(i);
            let pageText = '';

            /* ── Try digital text first ── */
            if (!forceOCRCb.checked) {
                const tc = await page.getTextContent();
                pageText = buildPageText(tc.items);
            }

            /* ── Fall back to OCR if needed ── */
            const needsOCR = forceOCRCb.checked ||
                             !pageText ||
                             pageText.length < CONFIG.MIN_DIGITAL_TEXT_LEN;

            if (needsOCR) {
                const vp      = page.getViewport({ scale: CONFIG.OCR_SCALE });
                const canvas  = document.createElement('canvas');
                const ctx     = canvas.getContext('2d', { willReadFrequently: true });
                canvas.width  = vp.width;
                canvas.height = vp.height;
                await page.render({ canvasContext: ctx, viewport: vp }).promise;
                await ensureOCR();
                const { data } = await ocrWorker.recognize(canvas);
                pageText = data.text;
            }

            extractedDataRaw += `\n\n--- Page ${i} ---\n${pageText}`;

            /* ── Skip non-invoice pages ── */
            if (!looksLikeInvoice(pageText)) {
                console.info(`Page ${i}: skipped (no invoice markers found)`);
                continue;
            }

            /* ── Exclude drafts if requested ── */
            const status = grab(pageText, /Invoice\s*Status\s*[:\-]\s*([A-Za-z]+)/i);
            if (excludeDraftCb.checked && /draft/i.test(status || '')) continue;

            /* ── Parse & push row ── */
            const row = parsePage(pageText);
            exportRows.push(row);
        }

        /* ── Preview ── */
        previewContent.textContent =
            extractedDataRaw.substring(0, 800) +
            (extractedDataRaw.length > 800 ? '…' : '');
        preview.classList.add('show');

        btnDownload.style.display = exportRows.length ? 'block' : 'none';
        showStatus(
            exportRows.length
                ? `✅ Done — ${exportRows.length} invoice row(s) extracted from ${total} page(s).`
                : `⚠️ No invoice data found in ${total} page(s). Try enabling Force OCR.`,
            exportRows.length ? 'success' : 'error'
        );

    } catch (err) {
        console.error(err);
        showStatus(`Error: ${err.message}`, 'error');
    } finally {
        btnConvert.disabled = false;
    }
}

/* ========= BUILD PAGE TEXT ========= */
/*
 * PDF.js returns items with x/y positions. We reconstruct lines by grouping
 * items with similar Y coordinates (within 5px), then sorting each line by X.
 * This handles broken/split-column layouts far better than a flat join.
 */
function buildPageText(items) {
    if (!items || !items.length) return '';

    const lines = [];
    const YTOL  = 5; // px tolerance for same-line grouping

    for (const item of items) {
        const str = (item.str || '').replace(/\u00A0/g, ' ');
        if (!str.trim()) continue;
        const y = Math.round(item.transform[5]);
        const x = Math.round(item.transform[4]);

        let matched = false;
        for (const line of lines) {
            if (Math.abs(line.y - y) <= YTOL) {
                line.parts.push({ x, str });
                matched = true;
                break;
            }
        }
        if (!matched) lines.push({ y, parts: [{ x, str }] });
    }

    /* Sort lines top→bottom, parts left→right */
    lines.sort((a, b) => b.y - a.y);
    return lines
        .map(l => l.parts.sort((a, b) => a.x - b.x).map(p => p.str).join(' '))
        .join('\n');
}

/* ========= INVOICE DETECTION ========= */
function looksLikeInvoice(text) {
    return /invoice/i.test(text) &&
           (/vendor\s*id/i.test(text) || /invoice\s*no/i.test(text));
}

/* ========= PARSE ONE PAGE → ONE ROW ========= */
function parsePage(text) {
    /* Normalise whitespace but preserve newlines for multi-line fields */
    const t = text.replace(/\u00A0/g, ' ').replace(/[ \t]+/g, ' ');

    return [
        /* 01 */ grab(t, /Vendor\s*ID\s*[:\-]\s*([^\n\r]+)/i),
        /* 02 */ grab(t, /Attention\s*To\s*[:\-]\s*([^\n\r]+)/i),
        /* 03 */ toExcelDate(grab(t, /Invoice\s*Date\s*[:\-]\s*([\d\/\-]+)/i)),
        /* 04 */ grab(t, /Credit\s*Term\s*[:\-]\s*([^\n\r]+)/i),
        /* 05 */ grab(t, /Invoice\s*No\s*[:\-]\s*([A-Z0-9\/\-]+)/i),
        /* 06 */ grab(t, /Related\s*Invoice\s*No\s*[:\-]\s*([^\n\r]*)/i) || '',
        /* 07 */ grab(t, /Invoice\s*Status\s*[:\-]\s*([A-Za-z]+)/i),
        /* 08 */ grab(t, /Invoicing\s*Instruction\s*(?:ID)?\s*[:\-]\s*([^\n\r]+)/i),
        /* 09 */ grabHeaderDescription(t),
        /* 10 */ grabLineNo(t),
        /* 11 */ grabLineDescription(t),
        /* 12 */ toNumber(grabTableCol(t, 'quantity')),
        /* 13 */ toNumber(grabTableCol(t, 'unitprice')),
        /* 14 */ toNumber(grabTableCol(t, 'grossex')),
        /* 15 */ toNumber(grabTableCol(t, 'gst')),
        /* 16 */ toNumber(grabTableCol(t, 'grossinc')),
        /* 17 */ grab(t, /Currency\s*[:\-]\s*([^\n\r]+)/i),
        /* 18 */ toNumber(grab(t, /Sub\s*Total\s*\(?Excluding\s*GST\)?\s*[:\-]\s*([0-9,\.]+)/i)),
        /* 19 */ toNumber(grab(t, /Total\s*GST\s*Payable\s*[:\-]\s*([0-9,\.]+)/i)),
        /* 20 */ toNumber(grab(t, /Freight\s*Amount\s*[:\-]\s*([0-9,\.]+)/i)),
        /* 21 */ toNumber(grab(t, /Total\s*Invoice\s*Amount\s*[:\-]\s*([0-9,\.]+)/i))
    ];
}

/* ========= FIELD GRABBERS ========= */

/** Generic: grab first capture group, trimmed */
function grab(text, regex) {
    const m = text.match(regex);
    return m ? m[1].trim() : '';
}

/**
 * Header Description — the label "Description :" that appears ABOVE the table.
 * We grab it by looking for Description before the line-item table header (No.).
 */
function grabHeaderDescription(text) {
    /* Split at the table header row so we only look in the header section */
    const headerSection = text.split(/\bNo\.?\s+Description\b/i)[0] || text;
    return grab(headerSection, /\bDescription\s*[:\-]\s*([^\n\r]+)/i);
}

/**
 * Line item No. — the row number in the table (e.g. "4")
 * Looks for a digit at the start of a table row after the header row.
 */
function grabLineNo(text) {
    /* Find table section after "No. Description Quantity …" header */
    const tableSection = getTableSection(text);
    if (!tableSection) return '';
    const m = tableSection.match(/^\s*(\d{1,3})\s/m);
    return m ? m[1] : '';
}

/**
 * Line item Description — the description cell in the table row.
 * Sits between the row number and Quantity.
 * We grab everything between the leading number and the first numeric column.
 */
function grabLineDescription(text) {
    const tableSection = getTableSection(text);
    if (!tableSection) return '';

    /*
     * Pattern: leading line-no, then text, then first number (qty/price).
     * We capture the text blob between them and clean it up.
     */
    const m = tableSection.match(
        /^\s*\d{1,3}\s+([\s\S]+?)\s+\d[\d,]*\.?\d*\s+\d[\d,]*\.?\d*/m
    );
    if (m) return m[1].replace(/\s+/g, ' ').trim();

    /* Fallback: grab everything after row number on the same line */
    const m2 = tableSection.match(/^\s*\d{1,3}\s+(.+)/m);
    return m2 ? m2[1].trim() : '';
}

/**
 * Returns the portion of text that is the line-item table body
 * (everything after the "No. Description Quantity …" header row,
 *  up to the "Invoice Amount Summary" section).
 */
function getTableSection(text) {
    const start = text.search(/\bNo\.?\s+Description\b/i);
    if (start === -1) return '';
    const end   = text.search(/Invoice\s*Amount\s*Summary/i);
    return end > start ? text.slice(start, end) : text.slice(start);
}

/**
 * Grab numeric columns from the table row.
 * col: 'quantity' | 'unitprice' | 'grossex' | 'gst' | 'grossinc'
 *
 * The table row pattern (from the invoice image):
 *   <No>  <description text>  <Qty>  <UnitPrice>  <GrossEx>  <GST>  <GrossInc>
 *
 * We extract the full row then pick columns by position.
 */
function grabTableCol(text, col) {
    const tableSection = getTableSection(text);
    if (!tableSection) return '';

    /*
     * Match a row that starts with a line number and contains
     * at least 4 numeric values (handles OCR spacing noise).
     */
    const rowMatch = tableSection.match(
        /^\s*\d{1,3}\s+[\s\S]+?([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s+([\d,]+\.?\d*)\s*$/m
    );

    if (!rowMatch) return '';

    /* rowMatch[1]=Qty, [2]=UnitPrice, [3]=GrossEx, [4]=GST, [5]=GrossInc */
    const map = {
        quantity : rowMatch[1],
        unitprice: rowMatch[2],
        grossex  : rowMatch[3],
        gst      : rowMatch[4],
        grossinc : rowMatch[5]
    };
    return map[col] || '';
}

/* ========= DOWNLOAD ========= */
function downloadExcel() {
    if (!exportRows.length) return showStatus('No data to download.', 'error');

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
        'Gross Amt (Ex. GST)',
        'GST @ 9%',
        'Gross Amt (Inc. GST)',
        'Currency',
        'Sub Total (Excluding GST)',
        'Total GST Payable',
        'Freight Amount',
        'Total Invoice Amount'
    ];

    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet([headers, ...exportRows]);

    /* ── Column widths ── */
    ws['!cols'] = [
        { wch: 12 }, { wch: 20 }, { wch: 14 }, { wch: 14 },
        { wch: 16 }, { wch: 18 }, { wch: 14 }, { wch: 24 },
        { wch: 30 }, { wch: 6  }, { wch: 30 }, { wch: 10 },
        { wch: 12 }, { wch: 18 }, { wch: 12 }, { wch: 18 },
        { wch: 18 }, { wch: 22 }, { wch: 18 }, { wch: 16 },
        { wch: 20 }
    ];

    XLSX.utils.book_append_sheet(wb, ws, 'Invoice Lines');
    XLSX.writeFile(wb, selectedFile.name.replace(/\.pdf$/i, '') + '.xlsx');
    showStatus('✅ Excel file downloaded!', 'success');
}

/* ========= HELPERS ========= */
function toNumber(v) {
    if (v == null || v === '') return '';
    const n = parseFloat(String(v).replace(/,/g, ''));
    return Number.isFinite(n) ? n : '';
}

function toExcelDate(s) {
    if (!s) return '';
    /* Support dd/mm/yyyy, dd-mm-yyyy, yyyy-mm-dd */
    let d, mth, y;
    let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (m) {
        [, d, mth, y] = m;
        if (y.length === 2) y = '20' + y;
        const dt = new Date(+y, +mth - 1, +d);
        return isNaN(dt) ? s : dt;
    }
    m = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
    if (m) {
        [, y, mth, d] = m;
        const dt = new Date(+y, +mth - 1, +d);
        return isNaN(dt) ? s : dt;
    }
    return s;
}
