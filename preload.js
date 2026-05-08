const { contextBridge, ipcRenderer } = require('electron');

// 暴露安全的API给渲染进程
contextBridge.exposeInMainWorld('electronAPI', {
  // 平台信息
  platform: process.platform,  // 'darwin' | 'win32' | 'linux'
  
  // 文件选择
  selectImages: () => ipcRenderer.invoke('select-images'),
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  saveExcel: (defaultName) => ipcRenderer.invoke('save-excel', defaultName),
  scanFolder: (folderPath) => ipcRenderer.invoke('scan-folder', folderPath),
  
  // 图片读取
  readImage: (filePath) => ipcRenderer.invoke('read-image', filePath),
  
  // 模板管理
  getTemplates: () => ipcRenderer.invoke('get-templates'),
  saveTemplate: (template) => ipcRenderer.invoke('save-template', template),
  deleteTemplate: (templateId) => ipcRenderer.invoke('delete-template', templateId),
  exportTemplate: (template) => ipcRenderer.invoke('export-template', template),
  importTemplate: () => ipcRenderer.invoke('import-template'),
  
  // 菜单事件
  onMenuImport: (callback) => ipcRenderer.on('menu-import', callback),
  onMenuExport: (callback) => ipcRenderer.on('menu-export', callback),
  
  // 本地OCR
  ocrImage: (imagePath) => ipcRenderer.invoke('ocr-image', imagePath),
  
  // Excel模板导入
  importExcelTemplate: () => ipcRenderer.invoke('import-excel-template'),
  
  // 路径处理
  joinPath: (...parts) => ipcRenderer.invoke('join-path', ...parts),
  
  // 文件写入
  writeFile: (filePath, buffer) => ipcRenderer.invoke('write-file', filePath, buffer),
  
  // 文件读取
  readFile: (filePath) => ipcRenderer.invoke('read-file', filePath),
  
  // 临时文件（用于PDF OCR）
  saveTempImage: (base64Data) => ipcRenderer.invoke('save-temp-image', base64Data),
  deleteTempFile: (filePath) => ipcRenderer.invoke('delete-temp-file', filePath)
});