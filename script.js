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
let extractedDataRaw = '';   // preview only
let exportRows = [];         // final rows (one per line item)
let ocrReady = false;

/* ========= PDF.js Worker ========= */
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

/* ========= DOM ========= */
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
      showStatus(`<span class="spinner"></span>${m.status} ${(m.progress*100).toFixed(0)}%`, 'loading');
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
uploadArea.addEventListener('dragover', e => { e.preventDefault(); uploadArea.classList.add('dragover'); });
uploadArea.addEventListener('dragleave', () => uploadArea.classList.remove('dragover'));
uploadArea.addEventListener('drop', e => {
  e.preventDefault(); uploadArea.classList.remove('dragover');
  const f = e.dataTransfer.files?.[0]; if (f) handleFile(f);
});
fileInput.addEventListener('change', e => {
  const f = e.target.files?.[0]; if (f) handleFile(f);
});

function handleFile(file) {
  const isPdfMime = file.type === 'application/pdf';
  const isPdfName = /\.pdf$/i.test(file.name);
  if (!isPdfMime && !isPdfName) return showStatus('Please select a PDF file', 'error');

  selectedFile = file;
  fileNameEl.textContent = `File: ${file.name}`;
  const sizeMB = file.size / 1024 / 1024;
  fileSizeEl.textContent = `Size: ${sizeMB >= .1 ? sizeMB.toFixed(2) + ' MB' : (file.size/1024).toFixed(0) + ' KB'}`;
  fileInfo.classList.add('show');

  btnConvert.disabled = false;
  btnDownload.style.display = 'none';
  preview.classList.remove('show');
  extractedDataRaw = '';
  exportRows = [];

  excludeDraftCb.checked = CONFIG.EXCLUDE_DRAFT_DEFAULT;
  showStatus(`Ready to convert: ${file.name}`, 'success');
}

function showStatus(msg, type='loading'){ statusEl.innerHTML = msg; statusEl.className = `status show ${type}`; }
function clearStatus(){ statusEl.textContent=''; statusEl.className='status'; }

function clearFile(){
  selectedFile = null; extractedDataRaw=''; exportRows=[];
  fileInput.value=''; fileInfo.classList.remove('show');
  btnConvert.disabled=true; btnDownload.style.display='none';
  preview.classList.remove('show'); clearStatus();
}
window.clearFile = clearFile;

/* ========= Convert ========= */
async function convertPDF(){
  if (selectedFile) return convertSelectedPDF();

  const pasted = (textInput.value || '').trim();
  if (pasted.length){
    exportRows = await extractRowsFromText(pasted, { excludeDraft: excludeDraftCb.checked });
    extractedDataRaw = pasted;
    previewContent.textContent = pasted.substring(0,800) + (pasted.length>800 ? '…':'');
    preview.classList.add('show'); btnDownload.style.display='block';
    return showStatus('Text is ready. Click "Download Excel File".','success');
  }
  showStatus('Please select a PDF or paste text, then click Convert.','error');
}
window.convertPDF = convertPDF;

async function convertSelectedPDF(){
  btnConvert.disabled = true;
  showStatus('<span class="spinner"></span>Opening PDF…','loading');
  try{
    const arrayBuffer = await selectedFile.arrayBuffer();
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;

    const invoices = new Map();   // invoiceNo -> { header:{}, items:[], totals:{}, gstRate }
    let currentInvoiceNo = null;
    extractedDataRaw = '';
    const useForceOCR = !!forceOCRCb.checked;

    for(let i=1;i<=pdf.numPages;i++){
      showStatus(`<span class="spinner"></span>Processing page ${i} of ${pdf.numPages}…`,'loading');
      const page = await pdf.getPage(i);

      // Try digital text first
      let pageText = '';
      if (!useForceOCR){
        const textContent = await page.getTextContent();
        pageText = textContent.items.map(x=>x.str).join(' ').replace(/\s+/g,' ').trim();
      }
      const needsOCR = useForceOCR || !pageText || pageText.length < CONFIG.MIN_DIGITAL_TEXT_LEN;
      if (needsOCR){
        const viewport = page.getViewport({ scale: CONFIG.OCR_SCALE });
        const canvas = document.createElement('canvas');
        const ctx = canvas.getContext('2d',{ willReadFrequently:true });
        canvas.width = Math.floor(viewport.width); canvas.height = Math.floor(viewport.height);
        await page.render({ canvasContext: ctx, viewport }).promise;
        await ensureOCR();
        const { data:{ text } } = await ocrWorker.recognize(canvas);
        pageText = (text||'').replace(/\s+/g,' ').trim();
      }

      extractedDataRaw += `\n\n--- Page ${i} ---\n` + pageText;
      const norm = normalize(pageText);

      // Invoice No (defines grouping across pages)
      const invNo = findInvoiceNo(norm);
      if (invNo) currentInvoiceNo = invNo;
      if (!currentInvoiceNo) continue;

      const inv = getOrCreateInvoice(invoices, currentInvoiceNo);

      // Header
      assignHeader(inv.header, extractHeaderFields(norm));

      // Optional GST rate in text (not used for per-line calc, but kept)
      const gstRate = findGSTRate(norm);
      if (gstRate != null) inv.gstRate = gstRate;

      // Line items (your real pattern)
      const items = extractLineItems(norm);
      if (items.length) inv.items.push(...items);

      // Totals (currency, subtotal, total GST, freight if present)
      assignTotals(inv.totals, extractTotals(norm));
    }

    // Build export rows
    const excludeDraft = !!excludeDraftCb.checked;
    exportRows = [];
    invoices.forEach((inv, invNo) => {
      const st = (inv.header.invoiceStatus || '').toLowerCase();
      if (excludeDraft && st === 'draft') return;

      // Compute missing totals from items
      if (!inv.totals.currency || !inv.totals.subtotal || !inv.totals.gst){
        computeTotals(inv);
      }
      if (CONFIG.NORMALIZE_CURRENCY_TO === 'SGD' && inv.totals.currency) inv.totals.currency = 'SGD';

      for (const li of inv.items){
        // Header fields (repeated per line)
        const row = [
          inv.header.vendorId || '',
          inv.header.attentionTo || '',
          toExcelDate(inv.header.invoiceDate) || inv.header.invoiceDate || '',
          inv.header.creditTerm || '',
          inv.header.invoiceNo || invNo,
          inv.header.relatedInvoiceNo || '',
          inv.header.invoiceStatus || '',
          inv.header.instructionId || '',
          inv.header.headerDescription || '',

          // Line item fields
          li.lineNo ?? '',
          li.description || '',
          toNumber(li.quantity),
          toNumber(li.unitPrice),
          toNumber(li.grossEx),
          toNumber(li.gstAmount),              // GST @ 9% => amount per line (Inc − Ex)
          toNumber(li.grossInc),

          // Totals (invoice-level, repeated)
          inv.totals.currency || '',
          toNumber(inv.totals.subtotal),
          toNumber(inv.totals.gst)
        ];
        exportRows.push(row);
      }
    });

    // Preview & UI
    previewContent.textContent = extractedDataRaw.substring(0,800) + (extractedDataRaw.length>800 ? '…' : '');
    preview.classList.add('show');
    btnDownload.style.display = exportRows.length ? 'block' : 'none';
    showStatus(exportRows.length
      ? `Parsed ${exportRows.length} line item row(s). Click "Download Excel File".`
      : 'No line items were found. Try "Force OCR".', exportRows.length ? 'success' : 'error');

  }catch(err){
    console.error(err);
    showStatus(`Error: ${err.message}`,'error');
  }finally{
    btnConvert.disabled = false;
  }
}
window.convertPDF = convertPDF;

/* ========= Download (single sheet; headers from your sample) ========= */
function downloadExcel(){
  if (!exportRows.length) return showStatus('No data to download','error');

  const headers = [
    'Vendor ID','Attention To','Invoice Date','Credit Term','Invoice No','Related Invoice No',
    'Invoice Status','Invoicing Instruction ID','Description',  // header/top description
    'No.','Description',                                        // line #
    'Quantity','Unit Price','Gross Amt (EX. GST)','GST @ 9%','Gross Amt (inc. GST)',
    'Currency','Sub Total (Excluding GST)','Total GST Payable'
  ];

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet([headers, ...exportRows]);
  ws['!cols'] = [
    {wch:12},{wch:18},{wch:12},{wch:10},{wch:18},{wch:18},{wch:12},{wch:20},{wch:70},
    {wch:6},{wch:70},{wch:10},{wch:12},{wch:16},{wch:16},{wch:18},{wch:16},{wch:20},{wch:18}
  ];
  ws['!freeze'] = { xSplit:0, ySplit:1 };
  XLSX.utils.book_append_sheet(wb, ws, 'Invoice Lines');

  const base = selectedFile ? selectedFile.name.replace(/\.pdf$/i,'') : 'pasted_text';
  const outName = `${base}_OptionA.xlsx`;
  XLSX.writeFile(wb, outName);
  showStatus(`Downloaded: ${outName}`, 'success');
}
window.downloadExcel = downloadExcel;

/* ========= Parsing helpers ========= */
function normalize(s){
  if (!s) return '';
  let t = s.replace(/\u00A0/g,' ');
  t = t.replace(/\s*:\s*/g,' : ').replace(/\s+/g,' ').trim();
  return t;
}
function toNumber(v){
  if (v == null || v==='') return '';
  if (typeof v === 'number') return v;
  const s = String(v).replace(/,/g,'');
  const n = parseFloat(s);
  return Number.isFinite(n) ? n : '';
}
function toExcelDate(s){
  if (!s || typeof s !== 'string') return '';
  // Match dd/mm/yyyy or d/m/yyyy
  const m = s.match(/(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})/);
  if (!m) return '';
  let [, d, mo, y] = m;
  if (y.length === 2) y = '20' + y; // assume 20xx for 2-digit years
  const date = new Date(parseInt(y,10), parseInt(mo,10)-1, parseInt(d,10));
  return isNaN(date.getTime()) ? '' : date;
}
function getOrCreateInvoice(map, invoiceNo){
  if (!map.has(invoiceNo)){
    map.set(invoiceNo,{ header:{ invoiceNo }, items:[], totals:{}, gstRate:null });
  }
  return map.get(invoiceNo);
}

/* ---- Header extraction (tight boundaries) ---- */
function findInvoiceNo(text){
  const m = text.match(/Invoice\s*No\s*:\s*([A-Z0-9-]+)/i);
  return m ? m[1] : null;
}
function extractHeaderFields(text){
  const out = {};
  const NEXT = '(?:Invoice\\s*Date|Credit\\s*Term|Invoice\\s*No|Related\\s*Invoice\\s*No|Invoice\\s*Status|Invoicing\\s*Instruction\\s*ID|Description)\\s*:';

  const id  = text.match(/Vendor\s*ID\s*:\s*([A-Z0-9]+)/i); if (id) out.vendorId = id[1];

  const att = text.match(new RegExp(`Attention\\s*To\\s*:\\s*([^:]+?)(?=\\s${NEXT})`, 'i'));
  if (att) out.attentionTo = att[1].trim();

  const dt  = text.match(/Invoice\s*Date\s*:\s*([0-3]?\d[\/\-][01]?\d[\/\-]\d{2,4})/i);
  if (dt) out.invoiceDate = dt[1];

  const term= text.match(new RegExp(`Credit\\s*Term\\s*:\\s*([^:]+?)(?=\\s${NEXT})`, 'i'));
  if (term) out.creditTerm = term[1].trim();

  const inv = text.match(/Invoice\s*No\s*:\s*([A-Z0-9-]+)/i);
  if (inv) out.invoiceNo = inv[1];

  const rel = text.match(new RegExp(`Related\\s*Invoice\\s*No\\s*:\\s*([^:]+?)(?=\\s${NEXT})`, 'i'));
  if (rel) out.relatedInvoiceNo = rel[1]?.trim();

  const st  = text.match(/Invoice\s*Status\s*:\s*([A-Za-z]+)/i);
  if (st) out.invoiceStatus = st[1];

  const ins = text.match(/Invoicing\s*Instruction\s*ID\s*:\s*([A-Z0-9-]+)/i);
  if (ins) out.instructionId = ins[1];

  const desc= text.match(/Description\s*:\s*(.+?)(?=\s(?:No\.|Currency\s*:|Invoice\s*Amount\s*Summary|Sub\s*Total|Total\s*GST|Freight\s*Amount|Total\s*Invoice\s*Amount|Vendor\s*ID\s*:|Attention\s*To\s*:|$))/i);
  if (desc) out.headerDescription = desc[1].trim();

  return out;
}
function assignHeader(target, src){ if (!src) return; for (const k of Object.keys(src)){ if (!target[k]) target[k]=src[k]; } }
function findGSTRate(text){ const m = text.match(/GST\s*@\s*(\d{1,2})\s*%/i); return m ? parseFloat(m[1]) : null; }

/* ---- Line items extraction (fault-tolerant) ----
   Matches:
   <No> <Description> <UnitPrice> [0] <Quantity> <GrossEx> [GSTAmt] <GrossInc>
   GST amount per line = GrossInc - GrossEx (if GSTAmt absent)
*/
function extractLineItems(text){
  const items = [];
  const rx = /(?:^|\s)(\d{1,3})\s+(.+?)\s+([0-9][0-9,]*\.\d{2,5})\s+(?:0(?:\.0+)?\s+)?([0-9]+(?:\.\d{2,5})?)\s+([0-9][0-9,]*\.\d{2})(?:\s+([0-9][0-9,]*\.\d{2}))?\s+([0-9][0-9,]*\.\d{2})(?=\s|$)/gi;

  let m;
  while ((m = rx.exec(text)) !== null){
    const lineNo      = parseInt(m[1],10);
    const description = m[2].trim();
    const unitPrice   = m[3];
    const quantity    = m[4];
    const grossEx     = m[5];
    const maybeGST    = m[6]; // optional GST amount
    const grossInc    = m[7];
    if (description.length < 3) continue;

    let gstAmount = '';
    const ex  = toNumber(grossEx);
    const inc = toNumber(grossInc);
    if (maybeGST) {
      gstAmount = toNumber(maybeGST);
    } else if (Number.isFinite(ex) && Number.isFinite(inc)) {
      gstAmount = +(inc - ex).toFixed(2);
    }

    items.push({ lineNo, description, quantity, unitPrice, grossEx, gstAmount: gstAmount===''?'':gstAmount, grossInc });
  }
  return items;
}

/* ---- Totals extraction ---- */
function extractTotals(text){
  const out = {};
  const cur = text.match(/Currency\s*:\s*([A-Za-z ]+?)(?=\s+Sub\s*Total|\s*$)/i);
  if (cur) out.currency = cur[1].trim();

  const sub = text.match(/Sub\s*Total\s*\(Excluding\s*GST\)\s*:\s*([0-9][0-9,]*\.\d{2})/i);
  if (sub) out.subtotal = sub[1];

  const gst = text.match(/Total\s*GST\s*Payable\s*:\s*([0-9][0-9,]*\.\d{2})/i);
  if (gst) out.gst = gst[1];

  const fr  = text.match(/Freight\s*Amount\s*:\s*([0-9][0-9,]*\.\d{2})/i);
  if (fr) out.freight = fr[1];

  const tot = text.match(/Total\s*Invoice\s*Amount\s*:\s*([0-9][0-9,]*\.\d{2})/i);
  if (tot) out.total = tot[1];

  return out;
}
function assignTotals(target, src){ if (!src) return; for(const k of Object.keys(src)){ if (!target[k]) target[k]=src[k]; } }
function computeTotals(inv){
  let subtotal=0, gst=0;
  for(const li of inv.items){
    const ex = toNumber(li.grossEx);
    const inc= toNumber(li.grossInc);
    if (Number.isFinite(ex)) subtotal += ex;
    if (Number.isFinite(ex) && Number.isFinite(inc) && inc>=ex) gst += (inc - ex);
  }
  if (!inv.totals.currency) inv.totals.currency = 'Singapore Dollar';
  if (!inv.totals.subtotal) inv.totals.subtotal = subtotal.toFixed(2);
  if (!inv.totals.gst)      inv.totals.gst      = gst.toFixed(2);
}

/* ========= Pasted text path ========= */
async function extractRowsFromText(text, { excludeDraft }){
  const invoices = new Map(); let currentInvoiceNo = null;
  const norm = normalize(text);
  const chunks = norm.split(/\s(?:Tax\s*Invoice|Invoice\s*Amount\s*Summary)\s/i);
  for (const chunk of chunks){
    if (!chunk || chunk.length<10) continue;
    const invNo = findInvoiceNo(chunk); if (invNo) currentInvoiceNo = invNo;
    if (!currentInvoiceNo) continue;
    const inv = getOrCreateInvoice(invoices, currentInvoiceNo);
    assignHeader(inv.header, extractHeaderFields(chunk));
    const gstRate = findGSTRate(chunk); if (gstRate != null) inv.gstRate=gstRate;
    const items = extractLineItems(chunk); if (items.length) inv.items.push(...items);
    assignTotals(inv.totals, extractTotals(chunk));
  }
  const rows=[]; invoices.forEach((inv, invNo)=>{
    const st = (inv.header.invoiceStatus||'').toLowerCase();
    if (excludeDraft && st==='draft') return;
    if (!inv.totals.currency || !inv.totals.subtotal || !inv.totals.gst) computeTotals(inv);
    if (CONFIG.NORMALIZE_CURRENCY_TO==='SGD' && inv.totals.currency) inv.totals.currency='SGD';

    for(const li of inv.items){
      rows.push([
        inv.header.vendorId||'', inv.header.attentionTo||'',
        toExcelDate(inv.header.invoiceDate)||inv.header.invoiceDate||'',
        inv.header.creditTerm||'', inv.header.invoiceNo||invNo, inv.header.relatedInvoiceNo||'',
        inv.header.invoiceStatus||'', inv.header.instructionId||'', inv.header.headerDescription||'',
        li.lineNo??'', li.description||'',
        toNumber(li.quantity), toNumber(li.unitPrice), toNumber(li.grossEx),
        toNumber(li.gstAmount), toNumber(li.grossInc),
        inv.totals.currency||'', toNumber(inv.totals.subtotal), toNumber(inv.totals.gst)
      ]);
    }
  });
  return rows;
}
