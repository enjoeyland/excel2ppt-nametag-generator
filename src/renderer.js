const { ipcRenderer } = require('electron');
const path = require('path');

let excelPath = '';
let pptxPath = '';

const excelStatus = document.getElementById('select-excel');
const pptxStatus = document.getElementById('select-pptx');
const fileList = document.getElementById('file-list');
const generateButton = document.getElementById('generate');

generateButton.disabled = true;
generateButton.classList.add('opacity-50', 'cursor-not-allowed');

// Button Click Event
document.getElementById('select-excel').addEventListener('click', () => {
    ipcRenderer.send('open-file-dialog-for-excel');
});
document.getElementById('select-pptx').addEventListener('click', () => {
    ipcRenderer.send('open-file-dialog-for-pptx');
});

document.getElementById('generate').addEventListener('click', () => {
    if (excelPath && pptxPath) {
        const paddingX = parseFloat(document.getElementById('padding_x').value) || 0;
        const paddingY = parseFloat(document.getElementById('padding_y').value) || 0;
        const marginX = parseFloat(document.getElementById('margin_x').value) || 0;
        const marginY = parseFloat(document.getElementById('margin_y').value) || 0;
        const perSlide = document.getElementById('per_slide').value === 'max' ? 'max' : parseInt(document.getElementById('per_slide').value) || 'max';

        ipcRenderer.send('execute-python', {
            excelPath,
            pptxPath,
            paddingX,
            paddingY,
            marginX,
            marginY,
            perSlide
        });
    } else {
        console.log('파일 경로가 필요합니다.');
    }
});

// IPC Event
ipcRenderer.on('selected-excel', (event, filePath) => {
    updateFileSelection('excel', filePath);
});
ipcRenderer.on('selected-pptx', (event, filePath) => {
    updateFileSelection('pptx', filePath);
});
ipcRenderer.on('python-output', (event, message) => {
    const alertBox = document.createElement('div');
    alertBox.innerHTML = message.replace(/\n/g, '<br>');
    alertBox.classList.add('alert-box');
    document.body.appendChild(alertBox);

    setTimeout(() => {
        alertBox.classList.add('fade-out');
    }, 5000);
    setTimeout(() => {
        document.body.removeChild(alertBox);
    }, 7000);
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
        statusElement = excelStatus;
    } else if (type === 'pptx') {
        pptxPath = filePath;
        statusElement = pptxStatus;
    }

    if (!statusElement.textContent.endsWith('O')) {
        statusElement.textContent += ' O';
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
        generateButton.classList.remove('opacity-50', 'cursor-not-allowed');
    } else {
        generateButton.disabled = true;
        generateButton.classList.add('opacity-50', 'cursor-not-allowed');
    }
}
