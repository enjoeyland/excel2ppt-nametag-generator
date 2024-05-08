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

ipcMain.on('execute-python', (event, excelPath, pptxPath) => {
  exec(`python main.py --excel "${excelPath}" --pptx "${pptxPath}"`, (error, stdout, stderr) => {
    if (error) {
      console.error(`exec error: ${error}`);
      return;
    }
    event.reply('python-output', stdout);
  });
});
