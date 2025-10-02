const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const { spawn } = require("child_process");
const path = require('path');

let win;
let pythonProcess;
let pythonReady = false;

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
        if (pythonProcess) {
            pythonProcess.kill();
        }
    });

    win.webContents.once('did-finish-load', () => {
        startPythonIPCServer();
    });
}

app.whenReady().then(createWindow);

function startPythonIPCServer() {
    console.log("🚀 Python IPC 서버 시작...");

    const pythonCommand = getPythonScript();
    pythonProcess = spawn(pythonCommand[0], [...pythonCommand.slice(1), "--rpc"], { 
        stdio: ["pipe", "pipe", "pipe"],
        encoding: 'utf8'
    });

    pythonProcess.stdout.on("data", (data) => {
        const text = data.toString().trim();
        for (const line of text.split("\n")) {
            try {
                const response = JSON.parse(line);

                if (response && response.status) {
                    console.log(`Sending to renderer:`, response);
                    win.webContents.send("task-result", response);
                } else {
                    console.warn("Received JSON but no 'status' field:", response);
                }
            } catch (error) {
                if (text.includes("ready")) {
                    console.log("✅ Python 실행 완료! 이제 요청을 받을 수 있음.");
                    pythonReady = true;
                }
                console.log(`Python: ${line}`);
            }
        }
    });

    pythonProcess.stderr.on("data", (data) => {
        console.error(`Python Error: ${data}`);
    });

    pythonProcess.on("close", (code) => {
        console.log(`Python process exited with code ${code}`);
        pythonReady = false;
    });
}

ipcMain.on("execute-task", (event, requestData) => {
    if (!pythonReady) {
        console.warn("⏳ Python이 아직 실행되지 않았음. 요청 대기 중...");
        if (requestData.task === "generate_pptx") {
            win.webContents.send("python-waiting");
        }
        setTimeout(() => {
            ipcMain.emit("execute-task", event, requestData);
        }, 500);
        return;
    }
    if (requestData.task === "generate_pptx") {
        win.webContents.send("python-ready");
    }
    
    if (pythonProcess) {
        console.log("🚀 요청 전송...");
        pythonProcess.stdin.write(JSON.stringify(requestData, null, 0) + "\n", 'utf8');
    }
});

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

function getPythonScript() {
    let isDevelopment = !app.isPackaged;
    let scriptPath;
  
    if (isDevelopment) {
      const pythonCommand = process.platform === "win32" ? "python" : "python3";
      scriptPath = path.join(__dirname, "..", "main.py");
      return [pythonCommand, scriptPath];
    } else {
      let root_dirname = __dirname.includes("app.asar") 
        ? path.join(__dirname, "..", "..") 
        : path.join(__dirname, "..");
  
      if (process.platform === "win32") {
        scriptPath = path.join(root_dirname, "python-script", "main.exe");
      } else {
        scriptPath = path.join(root_dirname, "python-script", "main");
      }
      return [scriptPath];
    }
  }
