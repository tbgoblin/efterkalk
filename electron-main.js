const { app, BrowserWindow, shell, dialog } = require('electron');
const { autoUpdater } = require('electron-updater');
const os = require('os');
const path = require('path');

// RDS / virtual desktop compatibility: disable GPU acceleration
// These must be set BEFORE app is ready
app.commandLine.appendSwitch('disable-gpu');
app.commandLine.appendSwitch('disable-software-rasterizer');
app.commandLine.appendSwitch('disable-gpu-compositing');
app.commandLine.appendSwitch('no-sandbox');

const { ensureServerStarted } = require('./server');

// In RDS multiple users share the same machine — give each user their own port
// based on a hash of the username to avoid conflicts
function getUserPort() {
    if (process.env.PORT) return Number(process.env.PORT);
    if (process.env.EFTERKALK_PORT) return Number(process.env.EFTERKALK_PORT);
    const username = (process.env.USERNAME || process.env.USER || os.userInfo().username || 'default').toLowerCase();
    let hash = 0;
    for (let i = 0; i < username.length; i++) {
        hash = (hash * 31 + username.charCodeAt(i)) & 0xffff;
    }
    // Range 3000-3999 — one port per user
    return 3000 + (hash % 1000);
}

const USER_PORT = getUserPort();
process.env.PORT = String(USER_PORT);

const APP_URL = 'http://localhost:' + USER_PORT;
const APP_NAME = 'Gantech Efterkalk';
const SHOULD_AUTO_START = String(process.env.EFTERKALK_AUTO_START || '1') === '1';

// Detect RDS environment
const IS_RDS = !!process.env.SESSIONNAME && process.env.SESSIONNAME !== 'Console';

console.info('Desktop process booting... port=' + USER_PORT + ' rds=' + IS_RDS);

function waitForSingleUpdateResult(timeoutMs = 15000) {
    return new Promise((resolve) => {
        let settled = false;

        const cleanup = () => {
            autoUpdater.removeListener('update-available', onAvailable);
            autoUpdater.removeListener('update-not-available', onNotAvailable);
            autoUpdater.removeListener('error', onError);
        };

        const finish = (payload) => {
            if (settled) return;
            settled = true;
            cleanup();
            resolve(payload);
        };

        const onAvailable = (info) => {
            finish({ ok: true, status: 'available', version: info && info.version ? info.version : null, message: 'Opdatering fundet.' });
        };

        const onNotAvailable = () => {
            finish({ ok: true, status: 'up-to-date', message: 'App er allerede opdateret.' });
        };

        const onError = (err) => {
            finish({ ok: false, status: 'error', message: err && err.message ? err.message : 'Ukendt updater-fejl.' });
        };

        autoUpdater.on('update-available', onAvailable);
        autoUpdater.on('update-not-available', onNotAvailable);
        autoUpdater.on('error', onError);

        setTimeout(() => {
            finish({ ok: true, status: 'checking', message: 'Opdateringskontrol startet. Proev igen om lidt.' });
        }, timeoutMs);
    });
}

async function triggerManualUpdateCheck() {
    if (!app.isPackaged) {
        return { ok: false, status: 'unsupported', message: 'Manuel opdatering virker kun i installeret app.' };
    }

    if (manualUpdateCheckRunning) {
        return { ok: true, status: 'busy', message: 'Opdateringskontrol koerer allerede.' };
    }

    manualUpdateCheckRunning = true;
    try {
        const waitResultPromise = waitForSingleUpdateResult();
        autoUpdater.checkForUpdates().catch((err) => {
            console.warn('Manual update check failed:', err.message);
        });
        return await waitResultPromise;
    } finally {
        manualUpdateCheckRunning = false;
    }
}

function setupAutoUpdater() {
    autoUpdater.checkForUpdatesAndNotify().catch((err) => {
        console.warn('Update check failed:', err.message);
    });

    autoUpdater.on('update-available', (info) => {
        console.info('Update available:', info.version);
    });

    autoUpdater.on('update-not-available', () => {
        console.info('Application is up to date');
    });

    autoUpdater.on('update-downloaded', (info) => {
        console.info('Update downloaded, will apply on restart');
        if (mainWindow) {
            dialog.showMessageBox(mainWindow, {
                type: 'info',
                title: 'Opdatering tilgængelig',
                message: 'Gantech Efterkalk ' + info.version + ' er klar.',
                detail: 'Klik OK for at genstarte og installere opdateringen.',
                buttons: ['OK', 'Senere']
            }).then((result) => {
                if (result.response === 0) {
                    autoUpdater.quitAndInstall();
                }
            });
        }
    });

    autoUpdater.on('error', (err) => {
        console.error('Updater error:', err.message);
    });
}

function configureAutoStart() {
    // Skip auto-start in RDS environments (shared machines)
    if (IS_RDS) return;
    if (process.platform !== 'win32') return;

    try {
        app.setLoginItemSettings({
            openAtLogin: SHOULD_AUTO_START,
            openAsHidden: false,
            path: process.execPath,
            args: []
        });
    } catch (e) {
        console.warn('setLoginItemSettings failed (RDS?):', e.message);
    }
}

function createMainWindow() {
    mainWindow = new BrowserWindow({
        title: APP_NAME,
        width: 1440,
        height: 960,
        minWidth: 1100,
        minHeight: 760,
        autoHideMenuBar: true,
        backgroundColor: '#c0392b',
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            sandbox: !IS_RDS  // sandbox must be disabled on some RDS configurations
        }
    });

    mainWindow.removeMenu();

    mainWindow.webContents.setWindowOpenHandler(({ url }) => {
        shell.openExternal(url).catch(() => {});
        return { action: 'deny' };
    });

    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

async function bootDesktopApp() {
    configureAutoStart();

    // Show loading screen immediately — don't wait for server
    createMainWindow();
    const pkgVersion = (() => { try { return require('./package.json').version; } catch(e) { return ''; } })();
    const loadingPath = path.join(__dirname, 'loading.html');
    mainWindow.loadFile(loadingPath, { query: { v: pkgVersion } });

    // Start server in background while loading screen is visible
    // Add a 60-second timeout so user always gets feedback instead of stuck red screen
    const serverStartPromise = ensureServerStarted();
    const timeoutPromise = new Promise((_, reject) =>
        setTimeout(() => reject(new Error('Server startup timed out after 60 seconds')), 60000)
    );

    try {
        await Promise.race([serverStartPromise, timeoutPromise]);
    } catch (err) {
        console.error('Server start failed:', err.message);
        if (mainWindow) {
            const errHtml = `data:text/html,<!DOCTYPE html><html><head><meta charset="UTF-8">
<style>body{font-family:Arial;background:#c0392b;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;margin:0;}
.box{background:rgba(0,0,0,0.3);border-radius:12px;padding:40px;max-width:600px;text-align:center;}
h2{margin-bottom:16px;}p{opacity:0.85;font-size:13px;word-break:break-all;}</style></head>
<body><div class="box"><h2>⚠️ Server kunne ikke starte</h2>
<p>${err.message.replace(/</g,'&lt;').replace(/>/g,'&gt;')}</p>
<p style="margin-top:20px;font-size:12px;opacity:0.6;">Luk og prøv igen. Kontakt IT hvis problemet fortsætter.</p>
</div></body></html>`;
            mainWindow.loadURL(errHtml);
        }
        return;
    }

    global.__desktopManualUpdateCheck = triggerManualUpdateCheck;

    // Navigate to app once server is ready
    if (mainWindow) {
        mainWindow.loadURL(APP_URL);
    }

    setupAutoUpdater();
}

app.whenReady().then(bootDesktopApp).catch((err) => {
    console.error('Desktop startup failed:', err);
    app.quit();
});

app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
        createMainWindow();
    }
});

app.on('window-all-closed', () => {
    app.quit();
});

app.on('before-quit', () => {
    delete global.__desktopManualUpdateCheck;
});
