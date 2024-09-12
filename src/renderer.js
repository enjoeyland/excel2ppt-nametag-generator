const { ipcRenderer } = require('electron');
const path = require('path');  // Node.js의 path 모듈 사용

let excelPath = '';
let pptxPath = '';

const excelStatus = document.getElementById('select-excel');
const pptxStatus = document.getElementById('select-pptx');
const fileList = document.getElementById('file-list');
const generateButton = document.getElementById('generate');

// 처음에는 만들기 버튼을 비활성화
generateButton.disabled = true;
generateButton.classList.add('opacity-50', 'cursor-not-allowed');

// Excel 파일 선택 버튼
document.getElementById('select-excel').addEventListener('click', () => {
    ipcRenderer.send('open-file-dialog-for-excel');
});

// PPTX 파일 선택 버튼
document.getElementById('select-pptx').addEventListener('click', () => {
    ipcRenderer.send('open-file-dialog-for-pptx');
});

// Python 실행 버튼 클릭 시 이벤트
document.getElementById('generate').addEventListener('click', () => {
    if (excelPath && pptxPath) {
        // 폼에서 입력한 값 가져오기
        const paddingX = parseFloat(document.getElementById('padding_x').value) || 0;
        const paddingY = parseFloat(document.getElementById('padding_y').value) || 0;
        const marginX = parseFloat(document.getElementById('margin_x').value) || 0;
        const marginY = parseFloat(document.getElementById('margin_y').value) || 0;
        const perSlide = document.getElementById('per_slide').value === 'max' ? 'max' : parseInt(document.getElementById('per_slide').value) || 'max';

        // Python 스크립트에 인자 전달
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

// Excel 파일 선택
ipcRenderer.on('selected-excel', (event, filePath) => {
    updateFileSelection('excel', filePath);
});

// PPTX 파일 선택
ipcRenderer.on('selected-pptx', (event, filePath) => {
    updateFileSelection('pptx', filePath);
});



// Drag and Drop 기능 추가
const dropArea = document.getElementById('drop-area');

// 기본 브라우저 동작 방지
['dragenter', 'dragover', 'dragleave', 'drop'].forEach(eventName => {
    dropArea.addEventListener(eventName, preventDefaults, false);
});

function preventDefaults(e) {
    e.preventDefault();
    e.stopPropagation();
}

// 드래그 상태에 따른 시각적 변화 추가
['dragenter', 'dragover'].forEach(eventName => {
    dropArea.classList.remove('border-gray-400');
    dropArea.classList.add('border-blue-500');
});

['dragleave', 'drop'].forEach(eventName => {
    dropArea.classList.remove('border-blue-500');
    dropArea.classList.add('border-gray-400');
});

// 파일 드롭 처리
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

// 파일 선택 및 상태 업데이트 함수
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

// 파일 목록 업데이트
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

// 두 파일이 모두 선택되었는지 확인하여 버튼 활성화/비활성화
function checkIfFilesSelected() {
    if (excelPath && pptxPath) {
        generateButton.disabled = false;
        generateButton.classList.remove('opacity-50', 'cursor-not-allowed');
    } else {
        generateButton.disabled = true;
        generateButton.classList.add('opacity-50', 'cursor-not-allowed');
    }
}
