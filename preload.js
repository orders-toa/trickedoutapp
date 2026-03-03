const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('toaAPI', {
  getTabs: () => ipcRenderer.invoke('get-tabs'),
  getStackRanker: () => ipcRenderer.invoke('get-stack-ranker'),
  getNextCode: (tabName) => ipcRenderer.invoke('get-next-code', tabName),
  markCodeUsed: (data) => ipcRenderer.invoke('mark-code-used', data),
  submitSuggestion: (data) => ipcRenderer.invoke('submit-suggestion', data),
  closeApp: () => ipcRenderer.send('close-app'),
  minimizeApp: () => ipcRenderer.send('minimize-app'),
});
