const { ipcRenderer } = require('electron');
const path = require('path');

let excelPath = '';
let pptxPath = '';

const excelStatus = document.getElementById('select-excel');
const pptxStatus = document.getElementById('select-pptx');
const fileList = document.getElementById('file-list');
const generateButton = document.getElementById('generate');
const generateButtonText = generateButton.querySelector(".button-text");
const loadingSpinner = document.getElementById('loading-spinner');

generateButton.disabled = true;
generateButton.classList.add('opacity-50');

// Button Click Event
document.getElementById('select-excel').addEventListener('click', () => {
    ipcRenderer.send('open-file-dialog-for-excel');
});
document.getElementById('select-pptx').addEventListener('click', () => {
    ipcRenderer.send('open-file-dialog-for-pptx');
});

document.getElementById('generate').addEventListener('click', () => {
    if (!excelPath || !pptxPath) {
        showCustomAlert("❌ 오류", "파일을 선택해주세요.");
    }

    generateButtonText.textContent = "처리중...";
    loadingSpinner.classList.remove("hidden");
    loadingSpinner.classList.add("inline-block");  
    generateButton.disabled = true;
    generateButton.classList.add("opacity-50");

    const marginX = parseFloat(document.getElementById('margin_x').value) || 0;
    const marginY = parseFloat(document.getElementById('margin_y').value) || 0;
    const paddingX = parseFloat(document.getElementById('padding_x').value) || 0;
    const paddingY = parseFloat(document.getElementById('padding_y').value) || 0;
    const perSlide = document.getElementById('per_slide').value === 'max' ? null : parseInt(document.getElementById('per_slide').value) || null;

    const requestData = {
        task: "generate_pptx",
        data: {
            pptx: pptxPath,
            excel: excelPath,
            margin_x: marginX,
            margin_y: marginY,
            padding_x: paddingX,
            padding_y: paddingY,
            per_slide: perSlide
        }
    };
    ipcRenderer.send("execute-task", requestData);

    ipcRenderer.once("generate-complete", () => {
        generateButtonText.textContent = "만들기";
        loadingSpinner.classList.add("hidden");
        loadingSpinner.classList.remove("inline-block");
        generateButton.disabled = false;
        generateButton.classList.remove("opacity-50");
    });     
});

ipcRenderer.on("task-result", (event, response) => {
    if (response.task === "generate_pptx") {
        ipcRenderer.emit("generate-complete");
    } else if (response.task === "get_excel_header") {
        ipcRenderer.emit("excel-header-complete", event, response);
    } else if (response.task === "get_pptx_slide_text") {
        ipcRenderer.emit("pptx-slide-text-complete", event, response);
    }

    if (response.status === "success" && response.message) {
        showCustomAlert("✅ 성공", `${response.message}`);
        console.log("Success details:", response);
    } 
    else if (response.status === "developer_error") {
        showCustomAlert("🛠️ 개발자 오류", `${response.message}`);
        console.error("Developer Error details:", response);
    }
    else if (response.status === "error") {
        showCustomAlert("❌ 오류", `${response.message}`);
        console.error("Error details:", response);
    }
});

ipcRenderer.on("excel-header-complete", (event, response) => {
    if (!response.headers) {
        console.error("❌ headers가 없음:", response);
        showCustomAlert("🛠️ 개발자 오류", "headers 속성이 누락되었습니다. Python 응답을 확인하세요.");
        return;
    }

    console.log("✅ Excel 헤더 목록:", response.headers);
    let headers = response.headers;
    headers = headers
        .filter(header => header.toLowerCase() !== "sample num")
        .sort((a, b) => a.localeCompare(b));
    headers.unshift("sample num");

    const statusElement = document.getElementById("header-status");
    let headerHtml =  `
        <span class="px-2 py-1 bg-green-200 text-green-800 rounded flex items-center font-bold">
            <i class="fa fa-file-excel mr-2"></i> Excel 헤더
        </span>
    `;
    
    const hasSampleNum = headers.includes("sample num");
    if (hasSampleNum) {
        headerHtml += `
            <span class="px-2 py-1 bg-gray-200 rounded flex items-center font-bold">
                sample num <i class="fa fa-check-circle text-green-500 ml-2"></i>
            </span>
        `;
    } else {
        headerHtml += `
            <span class="px-2 py-1 bg-gray-200 rounded flex items-center font-bold">
                sample num <i class="fa fa-times-circle text-red-500 ml-2"></i>
            </span>
        `;
    }

    headers.forEach(header => {
        if (header !== "sample num") {
            headerHtml += `<span class="px-2 py-1 bg-gray-200 rounded">${header}</span>`;
        }
    });

    statusElement.innerHTML = headerHtml;

    matchHeadersWithSlideText();
});

ipcRenderer.on("pptx-slide-text-complete", (event, response) => {
    if (!response.slides) {
        console.error("❌ slides가 없음:", response);
        showCustomAlert("🛠️ 개발자 오류", "slides 속성이 누락되었습니다. Python 응답을 확인하세요.");
        return;
    }

    console.log("✅ PPTX 슬라이드 텍스트:", response.slides);
    const slides = response.slides;

    const statusElement = document.getElementById("slide-text-status");
    statusElement.innerHTML = "";

    slides.forEach((slideTexts, index) => {
        slideTexts = slideTexts.sort((a, b) => a.localeCompare(b));
        let slideHtml = `
            <div class="flex items-center mb-2 gap-2">
                <span class="px-2 py-1 bg-red-200 text-red-700 rounded font-bold">
                    <i class="fas fa-file-powerpoint"></i> Sample ${index}
                </span>
        `;

        if (slideTexts.length > 0) {
            slideTexts.forEach(text => {
                slideHtml += `
                    <span class="px-2 py-1 bg-gray-200 rounded">${text}</span>
                `;
            });
        } else {
            slideHtml += `
                <span class="px-2 py-1 bg-yellow-200 rounded font-bold">⚠️ 텍스트 없음</span>
                <span class="text-gray-600 text-sm">(동일하게 복사됩니다.)</span>
            `;
        }

        slideHtml += `</div>`;
        statusElement.innerHTML += slideHtml;

        matchHeadersWithSlideText();
    });
});

function matchHeadersWithSlideText() {
    const headerElements = document.querySelectorAll("#header-status span");
    const slideTextElements = document.querySelectorAll("#slide-text-status span");

    slideTextElements.forEach(el => {
        if (el.textContent.includes("⚠️") || el.textContent.trim().startsWith("Sample")) {
            return;
        }
        el.innerHTML = el.innerHTML.replace(/<i class="fa fa-file-excel text-green-700"><\/i>/g, "");
        el.classList.remove("font-bold");
    });
    headerElements.forEach(el => {
        if (el.textContent.trim() === "sample num" || el.textContent.trim() === "Excel 헤더") {
            return;
        }
        el.classList.remove("font-bold");
    });
    
    const headers = new Map();
    headerElements.forEach(el => {
        const text = el.textContent.trim().toLowerCase();
        headers.set(text, el);
    });

    slideTextElements.forEach(el => {
        const text = el.textContent.trim().toLowerCase();
        if (headers.has(text)) {
            el.classList.add("font-bold");
            el.innerHTML += ' <i class="fa fa-file-excel text-green-700"></i>';

            const headerEl = headers.get(text);
            if (headerEl) {
                headerEl.classList.add("font-bold");
            }
        }
    });
}

function showCustomAlert(title, message) {
    document.getElementById("alert-title").textContent = title;
    document.getElementById("alert-message").textContent = message;
    document.getElementById("custom-alert").classList.remove("hidden");

    document.getElementById("alert-close").addEventListener("click", () => {
        document.getElementById("custom-alert").classList.add("hidden");
    });
}

// IPC Event
ipcRenderer.on('selected-excel', (event, filePath) => {
    updateFileSelection('excel', filePath);
});
ipcRenderer.on('selected-pptx', (event, filePath) => {
    updateFileSelection('pptx', filePath);
});

// Drag and Drop Event
const dropArea = document.getElementById('drop-area');

['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropArea.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

['dragenter', 'dragover'].forEach(eventName => {
    dropArea.classList.remove('border-gray-400');
    dropArea.classList.add('border-blue-500');
});

['dragleave', 'drop'].forEach(eventName => {
    dropArea.classList.remove('border-blue-500');
    dropArea.classList.add('border-gray-400');
});

dropArea.addEventListener('drop', handleDrop, false);

function handleDrop(e) {
    const dt = e.dataTransfer;
    const files = dt.files;

    for (let file of files) {
        if (file.name.endsWith('.xlsx')) {
            updateFileSelection('excel', file.path);
        } else if (file.name.endsWith('.pptx')) {
            updateFileSelection('pptx', file.path);
        } else {
            console.log('지원되지 않는 파일 형식: ' + file.name);
        }
    }
    updateFileList();
}

function updateFileSelection(type, filePath) {
    let statusElement;
    if (type === 'excel') {
        excelPath = filePath;
        statusElement = document.getElementById('select-excel').querySelector('.button-text');
        const requestData = {
            task: "get_excel_header",
            data: {
                excel: excelPath,
            }
        };
        ipcRenderer.send("execute-task", requestData);
    } else if (type === 'pptx') {
        pptxPath = filePath;
        statusElement = document.getElementById('select-pptx').querySelector('.button-text');
        const requestData = {
            task: "get_pptx_slide_text",
            data: {
                pptx: pptxPath,
            }
        };
        ipcRenderer.send("execute-task", requestData);
    }

    if (!statusElement.querySelector(".fa-check-circle")) {
        const checkIcon = document.createElement("i");
        checkIcon.classList.add("fa", "fa-check-circle", "text-white-500", "ml-2");
        statusElement.appendChild(checkIcon);
    }

    updateFileList();
    checkIfFilesSelected();
    console.log(`${type === 'excel' ? 'Excel' : 'PPTX'} 파일 선택됨: ${filePath}`);
}

function updateFileList() {
    fileList.innerHTML = '';  // 기존 목록 초기화

    if (excelPath) {
        const excelItem = document.createElement('li');
        excelItem.textContent = 'Excel 파일: ' + path.basename(excelPath);  // 파일 이름만 출력
        fileList.appendChild(excelItem);
    }

    if (pptxPath) {
        const pptxItem = document.createElement('li');
        pptxItem.textContent = 'PPTX 파일: ' + path.basename(pptxPath);  // 파일 이름만 출력
        fileList.appendChild(pptxItem);
    }
}

function checkIfFilesSelected() {
    if (excelPath && pptxPath) {
        generateButton.disabled = false;
        generateButton.classList.remove('opacity-50');
    } else {
        generateButton.disabled = true;
        generateButton.classList.add('opacity-50');
    }
}
