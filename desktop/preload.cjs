const { contextBridge, ipcRenderer } = require('electron');
contextBridge.exposeInMainWorld('mop', {
  uploadWorkbook: (bytes, name) => ipcRenderer.invoke('upload-workbook', bytes, name),
  exportWorkbook: model => ipcRenderer.invoke('export-workbook', model),
  uploadGenericWorkbook: (bytes, name) => ipcRenderer.invoke('upload-generic-workbook', bytes, name),
  exportMerge: model => ipcRenderer.invoke('export-merge', model),
  openResults: () => ipcRenderer.invoke('open-results'),
});
