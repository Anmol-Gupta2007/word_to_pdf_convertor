const fileInput = document.getElementById('file-input');
const outputContainer = document.getElementById('output-container');
const uploadArea = document.getElementById('upload-area');
const actionsContainer = document.getElementById('actions-container');
const downloadAllBtn = document.getElementById('download-all-btn');
const downloadZipBtn = document.getElementById('download-zip-btn');
const loadingStatus = document.getElementById('loading-status');

// Array to keep track of all generated PDFs and their "removed" state
let generatedDocs = [];
let globalDocIdCounter = 0;

// --- Drag and Drop Logic ---
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
    handleFiles(e.dataTransfer.files);
});

// --- Click Upload Logic ---
fileInput.addEventListener('change', (e) => {
    handleFiles(e.target.files);
});

// --- File Processing ---
async function handleFiles(files) {
    if (!files || files.length === 0) return;

    loadingStatus.style.display = 'block';

    for (const file of files) {
        if (!file.name.toLowerCase().endsWith('.docx')) {
            alert(`"${file.name}" is not a .docx Word file. Skipping.`);
            continue;
        }

        try {
            const arrayBuffer = await file.arrayBuffer();
            await convertDocxToPdf(arrayBuffer, file.name);
        } catch (error) {
            console.error("Error converting Word file:", error);
            alert("An error occurred while processing " + file.name);
        }
    }

    loadingStatus.style.display = 'none';

    if (generatedDocs.length > 0) {
        actionsContainer.style.display = 'block';
    }

    fileInput.value = ''; // Reset input
}

// --- Convert Word (.docx) to PDF Blob ---
async function convertDocxToPdf(arrayBuffer, originalName) {
    // 1. Convert .docx to clean HTML using Mammoth
    const result = await mammoth.convertToHtml({ arrayBuffer: arrayBuffer });
    const htmlContent = result.value;

    // 2. Create a styled container for PDF rendering
    const pdfContainer = document.createElement('div');
    pdfContainer.innerHTML = `
        <div style="
            font-family: 'Segoe UI', Arial, sans-serif;
            color: #222;
            line-height: 1.6;
            font-size: 12pt;
            padding: 10px;
        ">
            <style>
                h1, h2, h3, h4 { color: #111; margin: 14px 0 8px; }
                p { margin-bottom: 10px; }
                ul, ol { margin: 0 0 10px 24px; }
                table { width: 100%; border-collapse: collapse; margin: 12px 0; }
                th, td { border: 1px solid #999; padding: 6px 8px; text-align: left; }
                img { max-width: 100%; height: auto; }
            </style>
            ${htmlContent || '<p><em>(Empty document)</em></p>'}
        </div>
    `;

    const baseName = originalName.replace(/\.docx$/i, '');
    const newFileName = `${baseName}.pdf`;

    // 3. Generate PDF Blob via html2pdf.js
    const opt = {
        margin: 15,
        filename: newFileName,
        image: { type: 'jpeg', quality: 0.98 },
        html2canvas: { scale: 2, useCORS: true },
        jsPDF: { unit: 'mm', format: 'a4', orientation: 'portrait' }
    };

    const pdfBlob = await html2pdf().set(opt).from(pdfContainer).outputPdf('blob');
    const pdfUrl = URL.createObjectURL(pdfBlob);

    const currentId = globalDocIdCounter++;

    generatedDocs.push({
        id: currentId,
        url: pdfUrl,
        blob: pdfBlob,
        filename: newFileName,
        removed: false
    });

    createDocCard(htmlContent, pdfUrl, newFileName, currentId);
}

// --- Build UI Card for Each Converted Document ---
function createDocCard(htmlPreview, pdfUrl, fileName, id) {
    const card = document.createElement('div');
    card.className = 'doc-card';
    card.id = `card-${id}`;

    // Document preview box
    const previewBox = document.createElement('div');
    previewBox.className = 'doc-preview';
    previewBox.innerHTML = htmlPreview || '<em>Empty Word Document</em>';

    // File name
    const nameObj = document.createElement('p');
    nameObj.className = 'file-name';
    nameObj.textContent = fileName;

    // Button wrapper
    const btnWrapper = document.createElement('div');
    btnWrapper.className = 'card-buttons';

    // Individual Download PDF Button
    const singleDownloadBtn = document.createElement('button');
    singleDownloadBtn.className = 'single-download-btn';
    singleDownloadBtn.textContent = '⬇️ Download PDF';
    singleDownloadBtn.onclick = () => {
        triggerDownload(pdfUrl, fileName);
    };

    // Toggle Remove Button
    const removeBtn = document.createElement('button');
    removeBtn.className = 'remove-btn';
    removeBtn.textContent = '❌ Remove this file';

    removeBtn.onclick = () => {
        const docObj = generatedDocs.find(doc => doc.id === id);
        if (!docObj) return;

        docObj.removed = !docObj.removed;

        if (docObj.removed) {
            card.classList.add('removed');
            removeBtn.textContent = '✔️ Removed';
        } else {
            card.classList.remove('removed');
            removeBtn.textContent = '❌ Remove this file';
        }
    };

    btnWrapper.appendChild(singleDownloadBtn);
    btnWrapper.appendChild(removeBtn);

    card.appendChild(previewBox);
    card.appendChild(nameObj);
    card.appendChild(btnWrapper);

    outputContainer.appendChild(card);
}

// --- Helper Function to Trigger File Download ---
function triggerDownload(url, filename) {
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
}

// --- Download All Logic (Skips removed files) ---
downloadAllBtn.addEventListener('click', async () => {
    const activeDocs = generatedDocs.filter(doc => !doc.removed);

    if (activeDocs.length === 0) {
        alert("No PDFs available to download. You removed them all!");
        return;
    }

    for (let i = 0; i < activeDocs.length; i++) {
        triggerDownload(activeDocs[i].url, activeDocs[i].filename);
        await new Promise(resolve => setTimeout(resolve, 300));
    }
});

// --- Download as ZIP Logic ---
downloadZipBtn.addEventListener('click', async () => {
    const activeDocs = generatedDocs.filter(doc => !doc.removed);

    if (activeDocs.length === 0) {
        alert("No PDFs available to zip. You removed them all!");
        return;
    }

    const originalText = downloadZipBtn.innerHTML;
    downloadZipBtn.innerHTML = "⏳ Zipping...";
    downloadZipBtn.disabled = true;

    try {
        const zip = new JSZip();

        activeDocs.forEach(docData => {
            zip.file(docData.filename, docData.blob);
        });

        const content = await zip.generateAsync({ type: "blob" });
        const zipUrl = URL.createObjectURL(content);

        triggerDownload(zipUrl, "Converted_PDFs.zip");
        URL.revokeObjectURL(zipUrl);

    } catch (error) {
        console.error("Error creating ZIP:", error);
        alert("An error occurred while creating the ZIP file.");
    } finally {
        downloadZipBtn.innerHTML = originalText;
        downloadZipBtn.disabled = false;
    }
});
