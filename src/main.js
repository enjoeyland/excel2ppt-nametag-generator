const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const { exec } = require('child_process');
const path = require('path');

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

  let scriptPath = getScriptPath();

  let command = `${scriptPath} --excel "${excelPath}" --pptx "${pptxPath}" --padding_x ${paddingX} --padding_y ${paddingY} --margin_x ${marginX} --margin_y ${marginY}`;

  if (perSlide !== 'max') {
    command += ` --per_slide ${perSlide}`;
  }

  exec(command, (error, stdout, stderr) => {
    if (error) {
      console.error(`exec error: ${error.message}`);
      event.reply('python-output', `❗Error❗\n ${error.message}`);
      return;
    }
    if (stderr) {
      console.error(`stderr: ${stderr}`);
      event.reply('python-output', `❗Error❗\n ${stderr}`);
      return;
    }
    console.log(`stdout: ${stdout}`);
    event.reply('python-output', stdout);
  });
});

function getScriptPath() {
  let isDevelopment = !app.isPackaged;

  let scriptPath = "";
  if (isDevelopment) {
    scriptPath = `python ${path.join(__dirname, '..', 'main.py')}`;
  } else {
    let root_dirname = "";
    if (__dirname.includes('app.asar')) {
      root_dirname = `"${path.join(__dirname, '..', '..')}"`
    } else { 
      root_dirname = `"${path.join(__dirname, '..')}"`
    }

    if (process.platform === "win32") {
      scriptPath = `"${path.join(root_dirname, 'python-script', 'main.exe')}"`;
    } else if (process.platform === "darwin" || process.platform === "linux") {
      scriptPath = `"${path.join(root_dirname, 'python-script', 'main')}"`;
    }
  }
  return scriptPath;
}

