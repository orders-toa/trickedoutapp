const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('toaAPI', {
  getTabs: () => ipcRenderer.invoke('get-tabs'),
  getStoreEmployeeDirectory: () => ipcRenderer.invoke('get-store-employee-directory'),
  getEmployeeHighlights: (data) => ipcRenderer.invoke('get-employee-highlights', data),
  getCompanyLeaders: () => ipcRenderer.invoke('get-company-leaders'),
  getStoreStats: (data) => ipcRenderer.invoke('get-store-stats', data),
  getStackRanker: () => ipcRenderer.invoke('get-stack-ranker'),
  submitSupplyOrder: (data) => ipcRenderer.invoke('submit-supply-order', data),
  getRecentShoutOuts: () => ipcRenderer.invoke('get-recent-shout-outs'),
  reactToShoutOut: (data) => ipcRenderer.invoke('react-to-shout-out', data),
  getNextCode: (tabName) => ipcRenderer.invoke('get-next-code', tabName),
  markCodeUsed: (data) => ipcRenderer.invoke('mark-code-used', data),
  submitSuggestion: (data) => ipcRenderer.invoke('submit-suggestion', data),
  closeApp: () => ipcRenderer.send('close-app'),
  minimizeApp: () => ipcRenderer.send('minimize-app'),
  toggleMaximize: () => ipcRenderer.send('toggle-maximize'),
});
