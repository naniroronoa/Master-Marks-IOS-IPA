const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    getHWID: () => ipcRenderer.invoke('get-hwid'),
    saveAppState: (state) => ipcRenderer.invoke('save-app-state', state),
    loadAppState: () => ipcRenderer.invoke('load-app-state'),
    toggleFullScreen: () => ipcRenderer.invoke('toggle-fullscreen'),
    saveFile: (options) => ipcRenderer.invoke('save-file', options),
    closeApp: () => ipcRenderer.invoke('close-app')
});
