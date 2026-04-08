const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('qsApi', {
  searchDrawing:   (input)   => ipcRenderer.invoke('search-drawing', input),
  openFile:        (fp)      => ipcRenderer.invoke('open-file', fp),
  closeQuickSearch: ()       => ipcRenderer.invoke('close-quick-search'),
  resizeToContent: (h)       => ipcRenderer.invoke('resize-quick-search', h),
  getConfig:       ()        => ipcRenderer.invoke('get-config'),
  getSystemDark:   ()        => ipcRenderer.invoke('get-system-dark'),
  addRecentDrawing: (q)      => ipcRenderer.invoke('add-recent-drawing', q),
  addRecentOpenedDrawing: (n, q) => ipcRenderer.invoke('add-recent-opened-drawing', n, q),
});
