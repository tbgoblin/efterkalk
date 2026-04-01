const { app, BrowserWindow, shell, dialog } = require('electron');
const { autoUpdater } = require('electron-updater');
const os = require('os');
const fs = require('fs');
const path = require('path');

// RDS / virtual desktop compatibility: disable GPU acceleration
// These must be set BEFORE app is ready
app.commandLine.appendSwitch('disable-gpu');
app.commandLine.appendSwitch('disable-software-rasterizer');
app.commandLine.appendSwitch('disable-gpu-compositing');
app.commandLine.appendSwitch('no-sandbox');

// Keep logs in a predictable shared folder when possible.
if (!process.env.GANTECH_LOG_DIR) {
    const logCandidates = [
        'C:\\GantechCache',
        'C:\\cache\\Gantech',
        process.env.LOCALAPPDATA ? path.join(process.env.LOCALAPPDATA, 'Gantech Efterkalk') : null
    ].filter(Boolean);

    for (const dir of logCandidates) {
        try {
            fs.mkdirSync(dir, { recursive: true });
            fs.appendFileSync(path.join(dir, 'gantech.log'), '');
            process.env.GANTECH_LOG_DIR = dir;
            break;
        } catch (_) {
            // Try next candidate directory.
        }
    }
}

const MAIN_LOG_FILE = path.join(process.env.GANTECH_LOG_DIR || process.cwd(), 'gantech.log');

function writeDesktopLog(message) {
    const timestamp = new Date().toISOString();
    const line = '[' + timestamp + '] [desktop] ' + message + '\n';
    try {
        fs.appendFileSync(MAIN_LOG_FILE, line);
    } catch (_) {
        // Ignore logging failures.
    }
    try {
        console.log(line.trim());
    } catch (_) {
        // Ignore console failures.
    }
}

const { ensureServerStarted } = require('./server');

// In RDS multiple sessions can run under the same username.
// Include session identifiers in the hash to reduce port collisions.
function getUserPort() {
    if (process.env.PORT) return Number(process.env.PORT);
    if (process.env.EFTERKALK_PORT) return Number(process.env.EFTERKALK_PORT);
    const username = (process.env.USERNAME || process.env.USER || os.userInfo().username || 'default').toLowerCase();
    const sessionName = String(process.env.SESSIONNAME || '').toLowerCase();
    const clientName = String(process.env.CLIENTNAME || '').toLowerCase();
    const seed = username + '|' + sessionName + '|' + clientName;
    let hash = 0;
    for (let i = 0; i < seed.length; i++) {
        hash = (hash * 31 + seed.charCodeAt(i)) & 0xffff;
    }
    // Range 3000-3999
    return 3000 + (hash % 1000);
}

const USER_PORT = getUserPort();
process.env.PORT = String(USER_PORT);

const APP_URL = 'http://localhost:' + USER_PORT;
const APP_NAME = 'Gantech Efterkalk';
const SHOULD_AUTO_START = String(process.env.EFTERKALK_AUTO_START || '1') === '1';

// Detect RDS environment
const IS_RDS = !!process.env.SESSIONNAME && process.env.SESSIONNAME !== 'Console';

const logDirInfo = process.env.GANTECH_LOG_DIR || '(auto)';
console.info('Desktop process booting... port=' + USER_PORT + ' rds=' + IS_RDS + ' logDir=' + logDirInfo);
writeDesktopLog('boot port=' + USER_PORT + ' rds=' + IS_RDS + ' logDir=' + logDirInfo);

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

    mainWindow.webContents.on('did-start-loading', () => {
        writeDesktopLog('did-start-loading url=' + mainWindow.webContents.getURL());
    });

    mainWindow.webContents.on('did-finish-load', () => {
        writeDesktopLog('did-finish-load url=' + mainWindow.webContents.getURL());
    });

    mainWindow.webContents.on('did-fail-load', (_event, errorCode, errorDescription, validatedURL, isMainFrame) => {
        writeDesktopLog('did-fail-load code=' + errorCode + ' desc=' + errorDescription + ' url=' + validatedURL + ' mainFrame=' + isMainFrame);
    });

    mainWindow.webContents.on('render-process-gone', (_event, details) => {
        writeDesktopLog('render-process-gone reason=' + details.reason + ' exitCode=' + details.exitCode);
    });

    mainWindow.on('unresponsive', () => {
        writeDesktopLog('window-unresponsive');
    });

    mainWindow.webContents.setWindowOpenHandler(({ url }) => {
        shell.openExternal(url).catch(() => {});
        return { action: 'deny' };
    });

    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

function bootDesktopApp() {
    configureAutoStart();

    // Show loading screen immediately — don't wait for server
    createMainWindow();
    const pkgVersion = (() => { try { return require('./package.json').version; } catch(e) { return ''; } })();
    const loadingPath = path.join(__dirname, 'loading.html');
    writeDesktopLog('load loading screen path=' + loadingPath + ' port=' + USER_PORT);
    mainWindow.loadFile(loadingPath, { query: { v: pkgVersion, port: String(USER_PORT) } });

    // Start server in background — loading.html polls /health and navigates itself
    // Keep electron-main as fallback in case polling fails
    ensureServerStarted().then(() => {
        writeDesktopLog('ensureServerStarted resolved appUrl=' + APP_URL);
        global.__desktopManualUpdateCheck = triggerManualUpdateCheck;
        // loading.html will navigate on its own via polling — no need to loadURL here
        // but as a safety net, navigate after a short delay if window is still on loading page
        setTimeout(() => {
            if (mainWindow && mainWindow.webContents.getURL().startsWith('file://')) {
                writeDesktopLog('fallback loadURL ' + APP_URL);
                mainWindow.loadURL(APP_URL);
            }
        }, 1000);
        setupAutoUpdater();
    }).catch(err => {
        console.error('Server start failed:', err.message);
        if (err && err.code) console.error('Server start failed code:', err.code);
        if (err && err.stack) console.error('Server start failed stack:', err.stack);
        writeDesktopLog('server start failed message=' + (err && err.message ? err.message : 'unknown'));
        if (err && err.code) writeDesktopLog('server start failed code=' + err.code);
        if (err && err.stack) writeDesktopLog('server start failed stack=' + err.stack.replace(/\r?\n/g, ' | '));
        if (mainWindow) {
            const raw = (err && err.message ? err.message : 'Unknown startup error') + (err && err.code ? ' [' + err.code + ']' : '');
            const details = (err && err.code === 'EADDRINUSE')
                ? ('Port ' + USER_PORT + ' er allerede i brug. Luk andre Efterkalk-processer og prov igen.')
                : raw;
            const msg = details.replace(/</g, '&lt;').replace(/>/g, '&gt;');
            const errorHtml = '<!DOCTYPE html><html><head><meta charset="UTF-8"><style>body{font-family:Arial;background:#c0392b;color:#fff;display:flex;align-items:center;justify-content:center;height:100vh;margin:0}.box{background:rgba(0,0,0,.3);border-radius:12px;padding:40px;max-width:700px;text-align:center}</style></head><body><div class="box"><h2>Server kunne ikke starte</h2><p style="word-break:break-word">' + msg + '</p><p style="margin-top:20px;font-size:12px;opacity:.75">Luk og prov igen.</p></div></body></html>';
            mainWindow.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(errorHtml));
        }
    });
}

process.on('uncaughtException', (err) => {
    writeDesktopLog('uncaughtException message=' + (err && err.message ? err.message : 'unknown'));
    if (err && err.stack) writeDesktopLog('uncaughtException stack=' + err.stack.replace(/\r?\n/g, ' | '));
});

process.on('unhandledRejection', (reason) => {
    const message = reason && reason.message ? reason.message : String(reason);
    writeDesktopLog('unhandledRejection message=' + message);
    if (reason && reason.stack) writeDesktopLog('unhandledRejection stack=' + reason.stack.replace(/\r?\n/g, ' | '));
});

app.whenReady().then(bootDesktopApp).catch((err) => {
    console.error('Desktop startup failed:', err);
    writeDesktopLog('Desktop startup failed message=' + (err && err.message ? err.message : String(err)));
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
