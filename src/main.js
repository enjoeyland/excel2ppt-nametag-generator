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

  // 기본 명령 생성
  let command = `python main.py --excel "${excelPath}" --pptx "${pptxPath}" --padding_x ${paddingX} --padding_y ${paddingY} --margin_x ${marginX} --margin_y ${marginY}`;

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
