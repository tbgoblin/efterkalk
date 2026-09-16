const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('GohDesktopLogin', {
    info: () => ipcRenderer.invoke('goh-login:info'),
    restore: () => ipcRenderer.invoke('goh-login:restore'),
    remember: credentials => ipcRenderer.invoke('goh-login:remember', credentials),
    forget: () => ipcRenderer.invoke('goh-login:forget')
});