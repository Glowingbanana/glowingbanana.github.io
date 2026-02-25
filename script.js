/* ========= Config ========= */
const CONFIG = {
  NORMALIZE_CURRENCY_TO: null, // set to 'SGD' to map "Singapore Dollar" -> "SGD"
  DATE_FORMAT: 'dd/mm/yyyy',
  EXCLUDE_DRAFT_DEFAULT: false,
  MIN_DIGITAL_TEXT_LEN: 20,
  OCR_SCALE: 2,
  NUMBER_TOLERANCE: 0.02
};

/* ========= Globals ========= */
let selectedFile = null;
let extractedDataRaw = '';   // for preview
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

/* ========= OCR Worker ========= */
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

      // Optional GST rate in text
      const gstRate = findGSTRate(norm);
      if (gstRate != null) inv.gstRate = gstRate;

      // Line items (using your true pattern)
      const items = extractLineItems(norm);
      if (items.length) inv.items.push(...items);

      // Totals (currency, subtotal, total GST)
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
          li.gstPercent != null ? toNumber(li.gstPercent) : 0,  // "GST @ 9%" column per your header
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
    {wch:6},{wch:70},{wch:10},{wch:12},{wch:16},{wch:8},{wch:18},{wch:16},{wch:20},{wch:18}
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
function toExcelDate(dmy){
  if (!dmy || typeof dmy !== 'string') return '';
  const m = dmy.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (!m) return '';
  let [, d, mo, y] = m;
  if (y.length === 2) y = '20' + y;
  const dt = new Date(parseInt(y,10), parseInt(mo,10)-1, parseInt(d,10));
  return isNaN(dt.getTime()) ? '' : dt;
}
function getOrCreateInvoice(map, invoiceNo){
  if (!map.has(invoiceNo)){
    map.set(invoiceNo,{ header:{ invoiceNo }, items:[], totals:{}, gstRate:null });
  }
  return map.get(invoiceNo);
}

/* ---- Header extraction (labels seen across your PDF) ---- */
function findInvoiceNo(text){ const m = text.match(/Invoice\s*No\s*:\s*([A-Z0-9-]+)/i); return m ? m[1] : null; }
function extractHeaderFields(text){
  const out = {};
  const id  = text.match(/Vendor\s*ID\s*:\s*([A-Z0-9]+)/i);             if (id)  out.vendorId = id[1];
  const att = text.match(/Attention\s*To\s*:\s*([A-Za-z0-9 ,.'()\/-]+)/i); if (att) out.attentionTo = att[1].trim();
  const dt  = text.match(/Invoice\s*Date\s*:\s*([0-3]?\d[\/\-][01]?\d[\/\-]\d{2,4})/i); if (dt) out.invoiceDate = dt[1];
  const term= text.match(/Credit\s*Term\s*:\s*([A-Za-z0-9 ]+)/i);       if (term) out.creditTerm = term[1].trim();
  const inv = text.match(/Invoice\s*No\s*:\s*([A-Z0-9-]+)/i);           if (inv) out.invoiceNo = inv[1];
  const rel = text.match(/Related\s*Invoice\s*No\s*:\s*([A-Z0-9-]+)/i); if (rel) out.relatedInvoiceNo = rel[1];
  const st  = text.match(/Invoice\s*Status\s*:\s*([A-Za-z]+)/i);        if (st)  out.invoiceStatus = st[1];
  const ins = text.match(/Invoicing\s*Instruction\s*ID\s*:\s*([A-Z0-9-]+)/i); if (ins) out.instructionId = ins[1];
  // Header Description until a strong boundary
  const desc= text.match(/Description\s*:\s*(.+?)(?=\s(?:No\.\s|Currency\s*:|Invoice\s*Amount\s*Summary|Sub\s*Total|Total\s*GST|Freight\s*Amount|Total\s*Invoice\s*Amount|$))/i);
  if (desc) out.headerDescription = desc[1].trim();
  return out;
}
function assignHeader(target, src){ if (!src) return; for (const k of Object.keys(src)){ if (!target[k]) target[k]=src[k]; } }
function findGSTRate(text){ const m = text.match(/GST\s*@\s*(\d{1,2})\s*%/i); return m ? parseFloat(m[1]) : null; }

/* ---- Line items extraction (your real pattern) ----
   Pattern derived from your sample workbook rows: lineNo, description,
   unitPrice, 0, quantity, grossEx, grossInc.  [1](https://cognizantonline-my.sharepoint.com/personal/2453742_cognizant_com/_layouts/15/Doc.aspx?sourcedoc=%7B3B882CD4-232D-46B1-B05A-F5913F2CFDE7%7D&file=Vendors@gov-(29%20Sep%2025)-1-25_converted.xlsx&action=default&mobileredirect=true)
   We also include a fallback that accepts a GST amount between grossEx and grossInc.
*/
function extractLineItems(text){
  const items = [];

  // Primary (unitPrice with 3–4 decimals, a literal 0, quantity with 5 decimals, ex, inc)
  const rxPrimary = /(?:^|\s)(\d{1,3})\s+(.+?)\s+([0-9][0-9,]*\.\d{3,4})\s+0(?:\.0+)?\s+([0-9]+\.\d{5})\s+([0-9][0-9,]*\.\d{2})\s+([0-9][0-9,]*\.\d{2})(?=\s|$)/gi;

  // Fallback (some layouts include GST amount; quantity still has 5 decimals)
  const rxFallback = /(?:^|\s)(\d{1,3})\s+(.+?)\s+([0-9][0-9,]*\.\d{2,4})\s+([0-9]+\.\d{5})\s+([0-9][0-9,]*\.\d{2})\s+([0-9][0-9,]*\.\d{2})\s+([0-9][0-9,]*\.\d{2})(?=\s|$)/gi;

  let m;
  while ((m = rxPrimary.exec(text)) !== null){
    const lineNo = parseInt(m[1],10);
    const description = m[2].trim();
    const unitPrice = m[3];
    const quantity  = m[4];
    const grossEx   = m[5];
    const grossInc  = m[6];

    // GST % column: your sheet header reads "GST @ 9%".
    // Keep 0 for this column (as seen in your sample rows) and compute if needed.
    const gstPercent = 0;
    if (description.length >= 3){
      items.push({ lineNo, description, quantity, unitPrice, grossEx, grossInc, gstPercent });
    }
  }

  // If nothing matched, try fallback
  if (!items.length){
    let f;
    while ((f = rxFallback.exec(text)) !== null){
      const lineNo = parseInt(f[1],10);
      const description = f[2].trim();
      const unitPrice = f[3];
      const quantity  = f[4];
      const grossEx   = f[5];
      const gstAmt    = f[6];
      const grossInc  = f[7];

      // Compute GST% ~= (gstAmt / grossEx)*100
      let gstPercent = null;
      const ex = toNumber(grossEx), g = toNumber(gstAmt);
      if (Number.isFinite(ex) && ex>0 && Number.isFinite(g)){
        gstPercent = Math.round((g/ex)*1000)/10;
      }
      items.push({ lineNo, description, quantity, unitPrice, grossEx, grossInc, gstPercent });
    }
  }

  return items;
}

/* ---- Totals extraction (invoice-level) ---- */
function extractTotals(text){
  const out = {};
  const cur = text.match(/Currency\s*:\s*([A-Za-z ]+)/i);                    if (cur) out.currency = cur[1].trim();
  const sub = text.match(/Sub\s*Total\s*\(Excluding\s*GST\)\s*:\s*([0-9][0-9,]*\.\d{2})/i); if (sub) out.subtotal = sub[1];
  const gst = text.match(/Total\s*GST\s*Payable\s*:\s*([0-9][0-9,]*\.\d{2})/i);             if (gst) out.gst = gst[1];
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
        li.gstPercent!=null ? toNumber(li.gstPercent) : 0,
        toNumber(li.grossInc),
        inv.totals.currency||'', toNumber(inv.totals.subtotal), toNumber(inv.totals.gst)
      ]);
    }
  });
  return rows;
}
