let selectedFile = null;
let extractedData = null;

// Set up PDF.js worker
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// File upload handling
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');

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
    if (files.length > 0) {
        handleFile(files[0]);
    }
});

fileInput.addEventListener('change', (e) => {
    if (e.target.files.length > 0) {
        handleFile(e.target.files[0]);
    }
});

function handleFile(file) {
    if (file.type !== 'application/pdf') {
        showStatus('Please select a PDF file', 'error');
        return;
    }

    selectedFile = file;
    
    // Show file info
    document.getElementById('fileName').textContent = `File: ${file.name}`;
    document.getElementById('fileSize').textContent = `Size: ${(file.size / 1024).toFixed(2)} KB`;
    document.getElementById('fileInfo').classList.add('show');
    
    // Enable convert button
    document.getElementById('btnConvert').disabled = false;
    document.getElementById('btnDownload').style.display = 'none';
    document.getElementById('preview').classList.remove('show');
    
    showStatus(`Ready to convert: ${file.name}`, 'success');
}

function clearFile() {
    selectedFile = null;
    extractedData = null;
    fileInput.value = '';
    document.getElementById('fileInfo').classList.remove('show');
    document.getElementById('btnConvert').disabled = true;
    document.getElementById('btnDownload').style.display = 'none';
    document.getElementById('preview').classList.remove('show');
    document.getElementById('status').classList.remove('show');
}

async function convertPDF() {
    if (!selectedFile) {
        showStatus('Please select a PDF file first', 'error');
        return;
    }

    showStatus('<span class="spinner"></span>Converting PDF...', 'loading');
    document.getElementById('btnConvert').disabled = true;

    try {
        const arrayBuffer = await selectedFile.arrayBuffer();
        const pdf = await pdfjsLib.getDocument(arrayBuffer).promise;
        
        let fullText = '';
        
        for (let i = 1; i <= pdf.numPages; i++) {
            const page = await pdf.getPage(i);
            const textContent = await page.getTextContent();
            const pageText = textContent.items.map(item => item.str).join(' ');
            fullText += pageText + '\n';
        }

        extractedData = fullText;
        
        // Show preview
        document.getElementById('previewContent').textContent = fullText.substring(0, 500) + (fullText.length > 500 ? '...' : '');
        document.getElementById('preview').classList.add('show');

        showStatus('PDF converted successfully! Click "Download Excel File" to save.', 'success');
        document.getElementById('btnDownload').style.display = 'block';

    } catch (error) {
        showStatus(`Error: ${error.message}`, 'error');
        console.error(error);
    }

    document.getElementById('btnConvert').disabled = false;
}

function downloadExcel() {
    if (!extractedData) {
        showStatus('No data to download', 'error');
        return;
    }

    try {
        // Create workbook
        const workbook = XLSX.utils.book_new();
        
        // Split text into lines and create array of arrays
        const lines = extractedData.split('\n').filter(line => line.trim());
        const data = lines.map(line => [line]);
        
        // Create worksheet
        const worksheet = XLSX.utils.aoa_to_sheet([['Extracted Text from PDF'], [], ...data]);
        worksheet['!cols'] = [{ wch: 50 }]; // Set column width
        
        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, 'PDF Data');
        
        // Generate filename
        const fileName = selectedFile.name.replace('.pdf', '') + '_converted.xlsx';
        
        // Download
        XLSX.writeFile(workbook, fileName);
        
        showStatus(`Downloaded: ${fileName}`, 'success');
    } catch (error) {
        showStatus(`Error downloading: ${error.message}`, 'error');
        console.error(error);
    }
}

function showStatus(message, type) {
    const status = document.getElementById('status');
    status.innerHTML = message;
    status.className = `status show ${type}`;
}