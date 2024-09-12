const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const { exec } = require('child_process');

let win;

function createWindow() {
  win = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      nodeIntegration: true,
      contextIsolation: false
    }
  });

  win.loadFile('src/index.html');
  win.on('closed', () => {
    win = null;
  });
}

app.whenReady().then(createWindow);

ipcMain.on('open-file-dialog-for-excel', (event) => {
  dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [{ name: 'Excel Files', extensions: ['xls', 'xlsx'] }]
  }).then(result => {
    if (!result.canceled && result.filePaths.length > 0) {
      const excelPath = result.filePaths[0];
      event.reply('selected-excel', excelPath);
    }
  }).catch(err => {
    console.log(err);
  });
});

ipcMain.on('open-file-dialog-for-pptx', (event) => {
  dialog.showOpenDialog({
    properties: ['openFile'],
    filters: [{ name: 'PowerPoint Files', extensions: ['ppt', 'pptx'] }]
  }).then(result => {
    if (!result.canceled && result.filePaths.length > 0) {
      const pptxPath = result.filePaths[0];
      event.reply('selected-pptx', pptxPath);
    }
  }).catch(err => {
    console.log(err);
  });
});

ipcMain.on('execute-python', (event, args) => {
  const { excelPath, pptxPath, paddingX, paddingY, marginX, marginY, perSlide } = args;

  // 배포 환경과 개발 환경 구분
  let isDevelopment = process.env.NODE_ENV !== 'production'; // NODE_ENV로 구분

  // 실행 파일 또는 스크립트 경로
  let scriptPath = "";

  // 운영체제에 따른 실행 파일 경로 설정
  if (isDevelopment) {
      // 개발 환경 (Python 스크립트 직접 실행)
      scriptPath = `python ${path.join(__dirname, 'main.py')}`;
  } else {
      // 배포 환경 (운영체제에 맞는 실행 파일 경로 설정)
      if (process.platform === "win32") {
          scriptPath = `"${path.join(__dirname, 'python-script', 'main.exe')}"`;
      } else if (process.platform === "darwin" || process.platform === "linux") {
          scriptPath = `"${path.join(__dirname, 'python-script', 'main')}"`;
      }
  }

  // 명령어 생성
  let command = `${scriptPath} --excel "${excelPath}" --pptx "${pptxPath}" --padding_x ${paddingX} --padding_y ${paddingY} --margin_x ${marginX} --margin_y ${marginY}`;

  // perSlide가 'max'가 아니면 추가
  if (perSlide !== 'max') {
    command += ` --per_slide ${perSlide}`;
  }

  // Python 명령 실행
  exec(command, (error, stdout, stderr) => {
    if (error) {
      console.error(`exec error: ${error.message}`);
      event.reply('python-output', `Error: ${error.message}`);
      return;
    }
    if (stderr) {
      console.error(`stderr: ${stderr}`);
      event.reply('python-output', `Error: ${stderr}`);
      return;
    }

    // 성공적인 결과 출력
    console.log(`stdout: ${stdout}`);
    event.reply('python-output', stdout);
  });
});
