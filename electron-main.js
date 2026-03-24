const { app, BrowserWindow, Menu, ipcMain, shell, nativeImage } = require('electron');
const path = require('path');
const { exec } = require('child_process');
const fs = require('fs');

// Set App User Model ID for Windows taskbar icon consistency (Must be set early)
if (process.platform === 'win32') {
    app.setAppUserModelId('com.master.marks');
}

const DATA_DIR = process.env.PORTABLE_EXECUTABLE_DIR || app.getPath('userData');
const DB_PATH = path.join(DATA_DIR, 'scorebookdata.json');

// Helper to ensure database file exists
function ensureDb() {
    if (!fs.existsSync(DB_PATH)) {
        fs.writeFileSync(DB_PATH, JSON.stringify({ teacherInfo: {}, classes: [] }, null, 2));
    }
}

ipcMain.handle('save-app-state', async (event, state) => {
    try {
        fs.writeFileSync(DB_PATH, JSON.stringify(state, null, 2));
        return { success: true };
    } catch (error) {
        console.error('Failed to save state:', error);
        return { success: false, error: error.message };
    }
});

ipcMain.handle('load-app-state', async () => {
    try {
        ensureDb();
        const data = fs.readFileSync(DB_PATH, 'utf-8');
        return JSON.parse(data);
    } catch (error) {
        console.error('Failed to load state:', error);
        return null;
    }
});

ipcMain.handle('toggle-fullscreen', async (event) => {
    const win = BrowserWindow.fromWebContents(event.sender);
    if (win) {
        const isFullscreen = win.isFullScreen();
        win.setFullScreen(!isFullscreen);
        return !isFullscreen;
    }
    return false;
});

ipcMain.handle('close-app', async () => {
    app.quit();
});

function createWindow() {
    const iconPath = path.join(__dirname, 'assets', 'icon.ico');
    const win = new BrowserWindow({
        width: 1280,
        height: 800,
        icon: nativeImage.createFromPath(iconPath),
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            preload: path.resolve(__dirname, 'preload.js'),
            sandbox: false
        }
    });

    win.maximize();
    win.loadFile('index.html');

    // Open external links in default browser
    win.webContents.setWindowOpenHandler(({ url }) => {
        shell.openExternal(url);
        return { action: 'deny' };
    });

    Menu.setApplicationMenu(null);
}

// التعامل مع طلب معرف الجهاز (HWID)
ipcMain.handle('get-hwid', async () => {
    const crypto = require('crypto');
    const getCommandOutput = (cmd) => {
        return new Promise((resolve) => {
            exec(cmd, { timeout: 7000, shell: true }, (error, stdout) => {
                if (error || !stdout) {
                    resolve(null);
                } else {
                    resolve(stdout.trim());
                }
            });
        });
    };

    let rawId = await getCommandOutput('powershell.exe -Command "(Get-CimInstance Win32_ComputerSystemProduct).UUID"');

    if (!rawId || rawId === 'FFFFFFFF-FFFF-FFFF-FFFF-FFFFFFFFFFFF') {
        const wmicOut = await getCommandOutput('wmic csproduct get uuid');
        if (wmicOut) {
            const lines = wmicOut.split(/\r?\n/).map(l => l.trim()).filter(l => l.length > 0);
            if (lines.length > 1) rawId = lines[1];
        }
    }

    if (!rawId || rawId.includes('FFFFFFFF')) {
        const volOut = await getCommandOutput('vol c:');
        if (volOut) {
            const match = volOut.match(/Number is\s+([A-Z0-9-]+)/i);
            if (match) rawId = match[1];
        }
    }

    if (!rawId) {
        rawId = process.env.COMPUTERNAME || 'UNKNOWN-DEVICE';
    }

    const hash = crypto.createHash('sha256').update(rawId).digest('hex');
    let digits = '';
    for (let i = 0; i < hash.length && digits.length < 20; i++) {
        const char = hash[i];
        if (/[0-9]/.test(char)) {
            digits += char;
        } else {
            digits += char.charCodeAt(0).toString().slice(-1);
        }
    }

    while (digits.length < 15) {
        digits += (digits.length % 10).toString();
    }

    const prefix = "TLILI";
    const groups = [];
    for (let i = 0; i < 5; i++) {
        groups.push(prefix[i] + digits.substring(i * 3, i * 3 + 3));
    }
    const finalHWID = groups.join(' - ');

    return finalHWID;
});

// حفظ ملف Excel عبر نافذة الحفظ الأصلية
const { dialog } = require('electron');

ipcMain.handle('save-file', async (event, options) => {
    const { defaultPath, buffer, filters } = options;
    const { canceled, filePath } = await dialog.showSaveDialog({
        defaultPath: defaultPath || 'export.xlsx',
        filters: filters || [{ name: 'All Files', extensions: ['*'] }]
    });

    if (canceled || !filePath) return false;

    try {
        const uint8 = Buffer.from(buffer);
        fs.writeFileSync(filePath, uint8);
        return true;
    } catch (err) {
        console.error('save-file error:', err);
        return false;
    }
});

app.whenReady().then(() => {
    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    });
});

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});
