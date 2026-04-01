const { app, BrowserWindow, shell, dialog } = require('electron');
const { autoUpdater } = require('electron-updater');
const { ensureServerStarted } = require('./server');

const APP_URL = process.env.EFTERKALK_URL || 'http://localhost:3000';
const APP_NAME = 'Gantech Efterkalk';
const SHOULD_AUTO_START = String(process.env.EFTERKALK_AUTO_START || '1') === '1';

let mainWindow = null;

console.info('Desktop process booting...');

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
    if (process.platform !== 'win32') return;

    app.setLoginItemSettings({
        openAtLogin: SHOULD_AUTO_START,
        openAsHidden: false,
        path: process.execPath,
        args: []
    });
}

function createMainWindow() {
    mainWindow = new BrowserWindow({
        title: APP_NAME,
        width: 1440,
        height: 960,
        minWidth: 1100,
        minHeight: 760,
        autoHideMenuBar: true,
        backgroundColor: '#f5f5f5',
        webPreferences: {
            nodeIntegration: false,
            contextIsolation: true,
            sandbox: true
        }
    });

    mainWindow.removeMenu();

    mainWindow.webContents.setWindowOpenHandler(({ url }) => {
        shell.openExternal(url).catch(() => {});
        return { action: 'deny' };
    });

    mainWindow.loadURL(APP_URL);

    mainWindow.on('closed', () => {
        mainWindow = null;
    });
}

async function bootDesktopApp() {
    configureAutoStart();
    await ensureServerStarted();
    createMainWindow();
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
