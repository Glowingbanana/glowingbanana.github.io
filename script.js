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
let ocrWorker = null;

/* ========= PDF.js Worker ========= */
pdfjsLib.GlobalWorkerOptions.workerSrc =
    'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/* ========= DOM ========= */
const uploadArea    = document.getElementById('uploadArea');
const fileInput     = document.getElementById('fileInput');
const excludeDraftCb= document.getElementById('excludeDraft') || { checked: false };
const statusEl      = document.getElementById('status');
const fileInfo      = document.getElementById('fileInfo');
const fileNameEl    = document.getElementById('fileName');
const fileSizeEl    = document.getElementById('fileSize');
const preview       = document.getElementById('preview');
const previewContent= document.getElementById('previewContent');
const btnConvert    = document.getElementById('btnConvert');
const btnDownload   = document.getElementById('btnDownload');

/* ========= OCR (Tesseract v5) ========= */
async function ensureOCR() {
    if (ocrWorker) return;
    showStatus('Loading OCR engine…', 'loading');
    ocrWorker = await Tesseract.createWorker('eng', 1, {
        workerPath : 'https://cdn.jsdelivr.net/npm/tesseract.js@5/dist/worker.min.js',
        corePath   : 'https://cdn.jsdelivr.net/npm/tesseract.js-core@5.0.0/tesseract-core.wasm.js',
        langPath   : 'https://tessdata.projectnaptha.com/4.0.0',
        logger: m => {
            if (m?.status && typeof m.progress === 'number') {
                showStatus(`OCR: ${m.status} ${(m.progress * 100).toFixed(0)}%`, 'loading');
            }
        }
    });
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
            const tc = await page.getTextContent();
            pageText = buildPageText(tc.items);

            /* ── Fall back to OCR if page has no/little digital text ── */
            const needsOCR = !pageText || pageText.length < CONFIG.MIN_DIGITAL_TEXT_LEN;

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

            if (!pageText || pageText.length < CONFIG.MIN_DIGITAL_TEXT_LEN) {
                console.log(`Page ${i}: skipped (no text extracted)`);
                continue;
            }

            extractedDataRaw += `\n\n--- Page ${i} ---\n${pageText}`;

            /* ── Check if this is a continuation page ── */
            const hasTotals      = /Sub\s*Total|Total\s*GST|Freight\s*Amount|Total\s*Invoice\s*Amount/i.test(pageText);
            const hasHeader      = /Vendor\s*ID|Invoice\s*No/i.test(pageText);
            const isContinuation = hasTotals && !hasHeader && exportRows.length > 0;

            if (isContinuation) {
                const lastRow = exportRows[exportRows.length - 1];
                const currency = grab(pageText, /Currency\s*(?:\s*[:\-]\s*|\s+)([^\n\r]+)/i);
                const subtotal = toNumber(grab(pageText, /Sub\s*Total\s*\(?Excluding\s*GST\)?\s*(?:\s*[:\-]\s*|\s+)([\d\s,\.]+)/i));
                const gst      = toNumber(grab(pageText, /Total\s*GST\s*Payable\s*(?:\s*[:\-]\s*|\s+)([\d\s,\.]+)/i));
                const freight  = toNumber(grab(pageText, /Freight\s*Amount\s*(?:\s*[:\-]\s*|\s+)([\d\s,\.]+)/i));
                const total    = toNumber(grab(pageText, /Total\s*Invoice\s*Amount\s*(?:\s*[:\-]\s*|\s+)([\d\s,\.]+)/i));
                if (currency) lastRow[16] = currency;
                if (subtotal !== '') lastRow[17] = subtotal;
                if (gst      !== '') lastRow[18] = gst;
                if (freight  !== '') lastRow[19] = freight;
                if (total    !== '') lastRow[20] = total;
                console.log(`Page ${i}: continuation — totals merged into previous row`);
                continue;
            }

            /* ── Skip non-invoice pages ── */
            if (!looksLikeInvoice(pageText)) {
                console.log(`Page ${i}: skipped (no invoice markers found)`);
                continue;
            }

            /* ── Exclude drafts if requested ── */
            const status = grab(pageText, /Invoice\s*Status\s*(?:\s*[:\-]\s*|\s+)([A-Za-z]+)/i);
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
                : `⚠️ No invoice data found in ${total} page(s). Try a different PDF.`,
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
function buildPageText(items) {
    if (!items || !items.length) return '';

    const lines = [];
    const YTOL  = 5;

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
/*
 * SEP matches all separator styles:
 *   "Field: value"   — colon directly after label
 *   "Field : value"  — space then colon (native digital PDFs)
 *   "Field value"    — no colon (Tesseract OCR output)
 *
 * KEY INSIGHT: In digital PDFs, "Invoicing Instruction ID" and "Description"
 * are side by side on the same line, so PDF.js outputs them as:
 * "Invoicing Instruction ID : GVT000EPO25000378 Description : INTSG015..."
 * We use a lookahead to stop Invoicing Instruction ID before "Description :"
 */
const SEP = '(?:\\s*[:\\-]\\s*|\\s+)';

function parsePage(text) {
    const t = text.replace(/\u00A0/g, ' ').replace(/[ \t]+/g, ' ');

    return [
        /* 01 */ grabVendorID(t),
        /* 02 */ grab(t, new RegExp(`Attention\\s*To${SEP}([^\\n\\r]+?)(?=\\s*Invoice\\s*Date|\\s*\\n|\\s*$)`, 'i')),
        /* 03 */ toExcelDate(grab(t, new RegExp(`Invoice\\s*Date${SEP}(?:\\d\\s+)?(\\d{1,2}[\\.\\/\\-]\\d{1,2}[\\.\\/\\-]\\d{2,4})`, 'i'))),
        /* 04 */ grab(t, new RegExp(`Credit\\s*Term${SEP}([^\\n\\r]+?)(?=\\s*Invoice\\s*No|\\s*\\n|\\s*$)`, 'i')),
        /* 05 */ grab(t, new RegExp(`(?<!Related\\s*)Invoice\\s*No${SEP}([A-Z0-9\\/\\-]+)`, 'i')),
        /* 06 */ grabRelatedInvoiceNo(t),
        /* 07 */ grab(t, new RegExp(`Invoice\\s*Status${SEP}([A-Za-z]+)`, 'i')),
        /* 08 */ grabInvoicingInstructionID(t),
        /* 09 */ grabHeaderDescription(t),
        /* 10 */ grabLineNo(t),
        /* 11 */ grabLineDescription(t),
        /* 12 */ toNumber(grabTableCol(t, 'quantity')),
        /* 13 */ toNumber(grabTableCol(t, 'unitprice')),
        /* 14 */ toNumber(grabTableCol(t, 'grossex')),
        /* 15 */ toNumber(grabTableCol(t, 'gst')),
        /* 16 */ toNumber(grabTableCol(t, 'grossinc')),
        /* 17 */ grab(t, new RegExp(`Currency${SEP}([^\\n\\r]+)`, 'i')),
        /* 18 */ toNumber(grab(t, new RegExp(`Sub\\s*Total\\s*\\(?Excluding\\s*GST\\)?${SEP}([\\d\\s,\\.]+)`, 'i'))),
        /* 19 */ toNumber(grab(t, new RegExp(`Total\\s*GST\\s*Payable${SEP}([\\d\\s,\\.]+)`, 'i'))),
        /* 20 */ toNumber(grab(t, new RegExp(`Freight\\s*Amount${SEP}([\\d\\s,\\.]+)`, 'i'))),
        /* 21 */ toNumber(grab(t, new RegExp(`Total\\s*Invoice\\s*Amount${SEP}([\\d\\s,\\.]+)`, 'i')))
    ];
}

/* ========= FIELD GRABBERS ========= */

function grab(text, regex) {
    const m = text.match(regex);
    return m ? m[1].trim() : '';
}

/**
 * Vendor ID — OCR sometimes inserts a stray digit between the label and value.
 * e.g. "Vendor ID 1 200202851H" — skip any leading single digit/noise.
 */
function grabVendorID(text) {
    const m = text.match(/Vendor\s*ID\s*(?:\s*[:\-\xA9]\s*|\s+)(?:\d\s+)?([A-Z0-9]{5,})(?=\s+Attention|\s*\n|\s*$)/im);
    return m ? m[1].trim() : '';
}

/**
 * Invoicing Instruction ID — stops before "Description :" on the same line.
 * Handles both "Invoicing Instruction ID" and "Invoice Instruction ID".
 */
function grabInvoicingInstructionID(text) {
    const m = text.match(/Invoic(?:e|ing)?\s*Instruction\s*(?:ID)?\s*[:\-]\s*(.+?)(?=\s*\bDescription\s*(?:[:\-]|\s+\d)|\n|$)/i);
    return m ? m[1].trim() : '';
}

function grabHeaderDescription(text) {
    const m = text.match(/\bDescription\s*(?:[:\-]\s*|\s+\d+\s*)(.+?)(?=\s*(?:No\.?\s+Description\b|Quantity\s+Unit)|\n[^\n]*\||\s*$)/is);
    if (!m) return '';
    return m[1].replace(/\s+/g, ' ').trim();
}

function grabRelatedInvoiceNo(text) {
    const m = text.match(/Related\s*Invoice\s*No\s*(?:\s*[:\-]\s*|\s+)?(.*?)(?=Invoice\s*Status|Invoice\s*No\b|$)/is);
    if (!m) return '';
    const val = m[1].replace(/\s+/g, ' ').trim();
    if (!val || /Invoice/i.test(val)) return '';
    if (!/^[A-Za-z]{2}[A-Za-z0-9\-]+$/.test(val)) return '';
    return val;
}

function grabLineNo(text) {
    const tableSection = getTableSection(text);
    if (!tableSection) return '';
    const m = tableSection.match(/^\s*(\d{1,3})\s/m);
    return m ? m[1] : '';
}

function grabLineDescription(text) {
    const tableSection = getTableSection(text);
    if (!tableSection) return '';

    const m = tableSection.match(
        /^\s*\d{1,3}\s+([\s\S]+?)\s+[\d\s,]+\.?\d*\s+[\d\s,]+\.?\d*/m
    );
    if (m) return m[1].replace(/\s+/g, ' ').trim().replace(/^[\|\[\(©\s]+/, '').replace(/[\]\|©\s]+$/, '');

    const m2 = tableSection.match(/^\s*\d{1,3}\s+(.+)/m);
    return m2 ? m2[1].trim() : '';
}

/**
 * getTableSection — finds the line item table body.
 * Primary: looks for "No. Description" header row.
 * Fallback: when OCR garbles the header, find first line item row directly.
 */
function getTableSection(text) {
    const end = text.search(/Invoice\s*Amount\s*Summary/i);

    const start = text.search(/\bNo\.?\s+Description\b/i);
    if (start !== -1) {
        return end > start ? text.slice(start, end) : text.slice(start);
    }

    /* Fallback for garbled OCR headers */
    const fallback = text.search(/^\s*\d{1,3}\s+(?:[\|\[\(]+\s*)?[A-Za-z]/m);
    if (fallback !== -1) {
        return end > fallback ? text.slice(fallback, end) : text.slice(fallback);
    }

    return '';
}

function grabTableCol(text, col) {
    const tableSection = getTableSection(text);
    if (!tableSection) return '';

    const NUM = '([\\d][\\d,]*\\.?\\d*)';
    const rowMatch = tableSection.match(
        new RegExp(`^\\s*\\d{1,3}\\s+[\\s\\S]+?${NUM}\\s+${NUM}\\s+${NUM}\\s+${NUM}\\s+${NUM}\\s*$`, 'm')
    );

    if (!rowMatch) return '';

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
async function downloadExcel() {
    if (!exportRows.length) return showStatus('No data to download.', 'error');

    const headers = [
        'Vendor ID', 'Attention To', 'Invoice Date', 'Credit Term',
        'Invoice No', 'Related Invoice No', 'Invoice Status',
        'Invoicing Instruction ID', 'Description', 'No.', 'Description',
        'Quantity', 'Unit Price', 'Gross Amt (Ex. GST)', 'GST @ 9%',
        'Gross Amt (Inc. GST)', 'Currency', 'Sub Total (Excluding GST)',
        'Total GST Payable', 'Freight Amount', 'Total Invoice Amount'
    ];

    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Invoice Lines');

    ws.columns = headers.map((h, i) => {
        const dataMax = exportRows.reduce((max, row) => {
            const v = row[i] != null ? String(row[i]) : '';
            return Math.max(max, v.length);
        }, 0);
        return { key: 'col' + i, width: Math.max(h.length, dataMax) + 4 };
    });

    const headerRow = ws.addRow(headers);
    headerRow.height = 30;
    headerRow.eachCell({ includeEmpty: true }, cell => {
        cell.value = headers[cell.col - 1];
        cell.fill  = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFBDD7EE' } };
        cell.font  = { bold: true, color: { argb: 'FF1F3864' }, size: 11, name: 'Calibri' };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = {
            top:    { style: 'thin', color: { argb: 'FFA0C4E2' } },
            bottom: { style: 'medium', color: { argb: 'FF2E75B6' } },
            left:   { style: 'thin', color: { argb: 'FFA0C4E2' } },
            right:  { style: 'thin', color: { argb: 'FFA0C4E2' } }
        };
    });

    exportRows.forEach((row, idx) => {
        const exRow = ws.addRow(row);
        exRow.height = 20;
        const isAlt = idx % 2 !== 0;
        exRow.eachCell({ includeEmpty: true }, cell => {
            cell.alignment = { horizontal: 'center', vertical: 'middle' };
            if (isAlt) {
                cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFF2F8FD' } };
            }
            cell.border = {
                bottom: { style: 'hair', color: { argb: 'FFD9E1F2' } },
                right:  { style: 'hair', color: { argb: 'FFD9E1F2' } }
            };
        });
    });

    ws.views = [{ state: 'frozen', xSplit: 0, ySplit: 1, topLeftCell: 'A2', activeCell: 'A2' }];

    const buf  = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url  = URL.createObjectURL(blob);
    const a    = document.createElement('a');
    a.href     = url;
    a.download = selectedFile.name.replace(/\.pdf$/i, '') + '.xlsx';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
    showStatus('✅ Excel file downloaded!', 'success');
}

/* ========= HELPERS ========= */
function toNumber(v) {
    if (v == null || v === '') return '';
    const n = parseFloat(String(v).replace(/\s/g, '').replace(/,/g, ''));
    return Number.isFinite(n) ? n : '';
}

function toExcelDate(s) {
    if (!s) return '';
    s = s.replace(/\s/g, ''); // strip OCR-injected spaces e.g. '28/07/ 2025'
    let d, mth, y;
    let m = s.match(/^(\d{1,2})\s*[\/\-]\s*(\d{1,2})\s*[\/\-]\s*(\d{2,4})$/);
    if (m) {
        [, d, mth, y] = m;
        if (y.length === 2) y = '20' + y;
        const dt = new Date(+y, +mth - 1, +d);
        return isNaN(dt) ? s : dt;
    }
    m = s.match(/^(\d{4})\s*[\/\-]\s*(\d{1,2})\s*[\/\-]\s*(\d{1,2})$/);
    if (m) {
        [, y, mth, d] = m;
        const dt = new Date(+y, +mth - 1, +d);
        return isNaN(dt) ? s : dt;
    }
    return s;
}
