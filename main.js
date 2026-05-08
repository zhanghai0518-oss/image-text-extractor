const { app, BrowserWindow, ipcMain, dialog, Menu } = require('electron');
const path = require('path');
const fs = require('fs');

// 捕获未处理的异常
process.on('uncaughtException', (error) => {
  console.error('未捕获的异常:', error);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('未处理的Promise拒绝:', reason);
});

// 主窗口
let mainWindow;

// 模板配置文件路径
const templatesPath = path.join(app.getPath('userData'), 'templates');
const templatesFile = path.join(templatesPath, 'templates.json');

// 确保模板目录存在
function ensureTemplatesDir() {
  if (!fs.existsSync(templatesPath)) {
    fs.mkdirSync(templatesPath, { recursive: true });
  }
  if (!fs.existsSync(templatesFile)) {
    fs.writeFileSync(templatesFile, JSON.stringify([
      {
        id: 'default',
        name: '税务完税证明模板',
        fields: [
          { name: '所属公司', alias: '纳税人名称', enabled: true },
          { name: '所属公司代码', alias: '纳税人识别号', enabled: true },
          { name: '征收机关', alias: '税务机关', enabled: true },
          { name: '税种名称', alias: '税种', enabled: true },
          { name: '税目名称', alias: '品目', enabled: true },
          { name: '金额', alias: '实缴金额', enabled: true },
          { name: '税款所属期起', alias: '所属期起', enabled: true },
          { name: '税款所属期止', alias: '所属期止', enabled: true },
          { name: '缴款日期', alias: '入库日期', enabled: true },
          { name: '税票号码', alias: '凭证号码', enabled: true },
          { name: '主管税务所', alias: '税务机关', enabled: true },
          { name: '备注', alias: '备注', enabled: false }
        ]
      }
    ], null, 2));
  }
}

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 900,
    minHeight: 600,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    },
    titleBarStyle: 'hiddenInset',
    backgroundColor: '#f5f5f5'
  });

  mainWindow.loadFile('index.html');

  // 开发模式打开DevTools
  mainWindow.webContents.openDevTools();
}

// 创建菜单
function createMenu() {
  const template = [
    {
      label: '文件',
      submenu: [
        {
          label: '导入图片',
          accelerator: 'CmdOrCtrl+O',
          click: () => mainWindow.webContents.send('menu-import')
        },
        {
          label: '导出Excel',
          accelerator: 'CmdOrCtrl+E',
          click: () => mainWindow.webContents.send('menu-export')
        },
        { type: 'separator' },
        { role: 'quit', label: '退出' }
      ]
    },
    {
      label: '编辑',
      submenu: [
        { label: '撤销', role: 'undo' },
        { label: '重做', role: 'redo' },
        { type: 'separator' },
        { label: '剪切', role: 'cut' },
        { label: '复制', role: 'copy' },
        { label: '粘贴', role: 'paste' },
        { label: '全选', role: 'selectAll' }
      ]
    },
    {
      label: '视图',
      submenu: [
        { label: '重新加载', role: 'reload' },
        { label: '强制重新加载', role: 'forceReload' },
        { type: 'separator' },
        { label: '实际大小', role: 'resetZoom' },
        { label: '放大', role: 'zoomIn' },
        { label: '缩小', role: 'zoomOut' },
        { type: 'separator' },
        { label: '全屏', role: 'togglefullscreen' },
        { type: 'separator' },
        { 
          label: '开发者工具',
          role: 'toggleDevTools',
          accelerator: 'F12'
        }
      ]
    },
    {
      label: '帮助',
      submenu: [
        {
          label: '关于',
          click: () => {
            dialog.showMessageBox(mainWindow, {
              type: 'info',
              title: '关于图片文字提取工具',
              message: '图片文字提取工具 v1.0.0',
              detail: '支持批量导入图片、自定义字段模板、OCR识别、Excel导出'
            });
          }
        }
      ]
    }
  ];

  const menu = Menu.buildFromTemplate(template);
  Menu.setApplicationMenu(menu);
}

app.whenReady().then(() => {
  ensureTemplatesDir();
  createWindow();
  createMenu();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) {
      createWindow();
    }
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') {
    app.quit();
  }
});

// IPC处理程序

// 打开文件选择对话框
ipcMain.handle('select-images', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: '所有支持的文件', extensions: ['jpg', 'jpeg', 'png', 'bmp', 'gif', 'webp', 'pdf'] },
      { name: '图片文件', extensions: ['jpg', 'jpeg', 'png', 'bmp', 'gif', 'webp'] },
      { name: 'PDF文件', extensions: ['pdf'] },
      { name: '所有文件', extensions: ['*'] }
    ]
  });
  return result.filePaths;
});

// 打开文件夹选择
ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory']
  });
  return result.filePaths[0];
});

// 保存Excel文件
ipcMain.handle('save-excel', async (event, defaultName) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: '保存Excel文件',
    defaultPath: defaultName || '提取结果.xlsx',
    filters: [
      { name: 'Excel文件', extensions: ['xlsx'] }
    ]
  });
  return result.filePath;
});

// 读取模板列表
ipcMain.handle('get-templates', () => {
  try {
    const data = fs.readFileSync(templatesFile, 'utf-8');
    return JSON.parse(data);
  } catch {
    return [];
  }
});

// 保存模板
ipcMain.handle('save-template', (event, template) => {
  ensureTemplatesDir();
  let templates = [];
  try {
    templates = JSON.parse(fs.readFileSync(templatesFile, 'utf-8'));
  } catch {}
  
  const index = templates.findIndex(t => t.id === template.id);
  if (index >= 0) {
    templates[index] = template;
  } else {
    template.id = Date.now().toString();
    templates.push(template);
  }
  
  fs.writeFileSync(templatesFile, JSON.stringify(templates, null, 2));
  return template;
});

// 删除模板
ipcMain.handle('delete-template', (event, templateId) => {
  let templates = [];
  try {
    templates = JSON.parse(fs.readFileSync(templatesFile, 'utf-8'));
  } catch {}
  
  templates = templates.filter(t => t.id !== templateId);
  fs.writeFileSync(templatesFile, JSON.stringify(templates, null, 2));
  return true;
});

// 导出模板到文件
ipcMain.handle('export-template', async (event, template) => {
  const result = await dialog.showSaveDialog(mainWindow, {
    title: '导出模板',
    defaultPath: `${template.name || '模板'}.json`,
    filters: [
      { name: 'JSON文件', extensions: ['json'] }
    ]
  });
  
  if (result.filePath) {
    const exportData = {
      version: '1.0',
      exportDate: new Date().toISOString(),
      template: template
    };
    fs.writeFileSync(result.filePath, JSON.stringify(exportData, null, 2), 'utf-8');
    return result.filePath;
  }
  return null;
});

// 导入模板文件
ipcMain.handle('import-template', async (event) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '导入模板配置',
    filters: [
      { name: 'JSON文件', extensions: ['json'] }
    ],
    properties: ['openFile']
  });
  
  if (result.filePaths && result.filePaths.length > 0) {
    try {
      const data = JSON.parse(fs.readFileSync(result.filePaths[0], 'utf-8'));
      return data;
    } catch (err) {
      dialog.showErrorBox('导入失败', `无法读取模板文件: ${err.message}`);
      return null;
    }
  }
  return null;
});

// 读取图片文件（返回base64）
ipcMain.handle('read-image', (event, filePath) => {
  try {
    const buffer = fs.readFileSync(filePath);
    const ext = path.extname(filePath).toLowerCase();
    
    // PDF文件特殊处理
    if (ext === '.pdf') {
      return {
        data: `data:application/pdf;base64,${buffer.toString('base64')}`,
        name: path.basename(filePath),
        size: buffer.length,
        isPDF: true
      };
    }
    
    // 图片文件
    const mime = ext === '.png' ? 'image/png' : 'image/jpeg';
    return {
      data: `data:${mime};base64,${buffer.toString('base64')}`,
      name: path.basename(filePath),
      size: buffer.length
    };
  } catch (err) {
    return { error: err.message };
  }
});

// 扫描文件夹中的图片和PDF
ipcMain.handle('scan-folder', (event, folderPath) => {
  try {
    const exts = ['.jpg', '.jpeg', '.png', '.bmp', '.gif', '.webp', '.pdf'];
    const files = fs.readdirSync(folderPath)
      .filter(f => exts.includes(path.extname(f).toLowerCase()))
      .map(f => path.join(folderPath, f));
    return files;
  } catch (err) {
    return [];
  }
});

// 使用本地OCR工具识别图片（跨平台支持）
ipcMain.handle('ocr-image', async (event, imagePath) => {
  const { execFile, exec } = require('child_process');
  const util = require('util');
  const execFileAsync = util.promisify(execFile);
  
  const platform = process.platform;
  
  try {
    // macOS: 使用Vision OCR工具
    if (platform === 'darwin') {
      const ocrTool = path.join(app.getPath('desktop'), 'OCR工具包', 'ocr_image');
      
      if (fs.existsSync(ocrTool)) {
        const { stdout } = await execFileAsync(ocrTool, [imagePath]);
        return { success: true, text: stdout };
      } else {
        return { success: false, error: 'OCR工具未找到，请确保 ~/Desktop/OCR工具包/ocr_image 存在' };
      }
    }
    
    // Windows: 使用 PaddleOCR-json (高性能中文OCR)
    if (platform === 'win32') {
      // 查找 PaddleOCR-json.exe 路径
      const possiblePaddlePaths = [
        // 打包后的路径 (extraResources 会将 resources/ 复制到 process.resourcesPath/)
        path.join(process.resourcesPath, 'PaddleOCR-json', 'PaddleOCR-json.exe'),
        // 另一种可能的打包路径
        path.join(path.dirname(process.execPath), 'resources', 'PaddleOCR-json', 'PaddleOCR-json.exe'),
        // 开发环境路径
        path.join(__dirname, 'resources', 'PaddleOCR-json', 'PaddleOCR-json.exe'),
        // 当前目录
        path.join('.', 'resources', 'PaddleOCR-json', 'PaddleOCR-json.exe'),
      ];
      
      console.log('[Win OCR] 搜索 PaddleOCR-json...');
      console.log('[Win OCR] process.resourcesPath:', process.resourcesPath);
      console.log('[Win OCR] process.execPath:', process.execPath);
      console.log('[Win OCR] __dirname:', __dirname);
      
      let paddleOcrPath = null;
      let checkedPaths = [];
      
      for (const p of possiblePaddlePaths) {
        const exists = fs.existsSync(p);
        checkedPaths.push({ path: p, exists: exists });
        console.log('[Win OCR] 检查:', p, '存在:', exists);
        if (exists && !paddleOcrPath) {
          paddleOcrPath = p;
        }
      }
      
      if (paddleOcrPath) {
        console.log('[Win OCR] 使用 PaddleOCR-json:', paddleOcrPath);
        
        // 解决中文路径问题：先将图片复制到临时目录（纯英文路径）
        const tempDir = path.join(app.getPath('temp'), 'image-text-extractor');
        if (!fs.existsSync(tempDir)) {
          fs.mkdirSync(tempDir, { recursive: true });
        }
        
        // 生成临时文件名（纯英文）
        let tempImagePath = path.join(tempDir, `ocr_${Date.now()}${path.extname(imagePath)}`);
        let useTempPath = false;
        
        try {
          // 复制图片到临时目录
          await fs.promises.copyFile(imagePath, tempImagePath);
          console.log('[PaddleOCR] 复制图片到临时目录:', tempImagePath);
          useTempPath = true;
        } catch (copyErr) {
          // 如果复制失败，可能是因为路径已经相同（如临时文件）
          console.log('[PaddleOCR] 复制失败，使用原路径:', copyErr.message);
          tempImagePath = null;
        }
        
        try {
          // PaddleOCR-json 命令行参数
          const actualPath = useTempPath && fs.existsSync(tempImagePath) ? tempImagePath : imagePath;
          const args = [
            '-image_path=' + actualPath
          ];
          
          console.log('[PaddleOCR] 使用路径:', actualPath);
          
          console.log('[PaddleOCR] 执行:', paddleOcrPath, args.join(' '));
          
          const { stdout, stderr } = await execFileAsync(paddleOcrPath, args, {
            maxBuffer: 50 * 1024 * 1024,
            encoding: 'utf-8',
            cwd: path.dirname(paddleOcrPath),
            timeout: 120000
          });
          
          console.log('[PaddleOCR] 完整输出:', stdout);
          console.log('[PaddleOCR] stderr:', stderr);
          
          // PaddleOCR 输出包含版本信息等，需要提取最后的 JSON 行
          // 输出格式：
          // PaddleOCR-json v1.4.1
          // OCR single image mode. Path: xxx
          // OCR init completed.
          // {"code":100,"data":[...]}
          
          let jsonStr = '';
          
          // 查找最后一个完整的 JSON 对象（以 {"code" 开头的行）
          const lines = stdout.split('\n');
          console.log('[PaddleOCR] 总行数:', lines.length);
          
          for (let i = lines.length - 1; i >= 0; i--) {
            const line = lines[i].trim();
            console.log('[PaddleOCR] 行', i, ':', line.substring(0, 50));
            if (line.startsWith('{"code"') || line.startsWith('{"code":')) {
              jsonStr = line;
              console.log('[PaddleOCR] 找到JSON行，索引:', i);
              break;
            }
          }
          
          if (!jsonStr) {
            console.log('[PaddleOCR] 未找到JSON行，尝试在整个输出中搜索');
            // 尝试在整个输出中查找JSON
            const jsonMatch = stdout.match(/\{"code":\d+,"data":\[[^\]]+\]\}/);
            if (jsonMatch) {
              jsonStr = jsonMatch[0];
              console.log('[PaddleOCR] 正则匹配找到JSON');
            }
          }
          
          if (!jsonStr) {
            console.log('[PaddleOCR] 无法提取JSON，OCR失败');
            return { success: false, error: '无法解析OCR输出' };
          }
          
          console.log('[PaddleOCR] 提取的JSON长度:', jsonStr.length);
          
          // 解析 JSON 输出
          try {
            const result = JSON.parse(jsonStr);
            
            if (result.code === 100 && Array.isArray(result.data)) {
              // 合并所有文本块
              const text = result.data.map(item => item.text).join('\n');
              console.log('[PaddleOCR] 成功，文本长度:', text.length);
              return { success: true, text: text };
            } else if (result.code === 100) {
              // 单个结果
              return { success: true, text: result.data || '' };
            } else if (result.code === 101) {
              // 无文字
              console.log('[PaddleOCR] 未检测到文字');
              return { success: true, text: '' };
            } else {
              console.log('[PaddleOCR] 错误码:', result.code, '完整结果:', JSON.stringify(result));
              // 返回详细错误信息，包括原始输出
              return { 
                success: false, 
                error: `OCR错误: ${result.data || '未知错误'} (code: ${result.code})`,
                rawOutput: stdout,
                stderr: stderr
              };
            }
          } catch (parseError) {
            // JSON 解析失败，可能是 PaddleOCR 输出了错误信息
            console.log('[PaddleOCR] JSON解析失败:', parseError.message);
            console.log('[PaddleOCR] 尝试解析的内容:', jsonStr.substring(0, 300));
            
            return { success: false, error: 'OCR输出格式错误，无法解析JSON' };
          }
        } catch (execError) {
          console.error('[PaddleOCR] 执行错误:', execError.message);
          return { success: false, error: 'PaddleOCR执行失败: ' + execError.message };
        } finally {
          // 清理临时图片文件
          if (tempImagePath && fs.existsSync(tempImagePath)) {
            try {
              fs.unlinkSync(tempImagePath);
              console.log('[PaddleOCR] 清理临时文件:', tempImagePath);
            } catch (e) {
              console.log('[PaddleOCR] 清理临时文件失败:', e.message);
            }
          }
        }
      }
      
      return { success: false, error: '未找到PaddleOCR引擎' };
    }
    
    return { success: false, error: '不支持的平台: ' + platform };
  } catch (err) {
    console.error('[OCR] 错误:', err.message);
    return { success: false, error: err.message };
  }
});

// 导入Excel模板（读取列头）
ipcMain.handle('import-excel-template', async (event) => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: '导入Excel模板',
    filters: [
      { name: 'Excel文件', extensions: ['xlsx', 'xls'] }
    ],
    properties: ['openFile']
  });
  
  if (result.canceled || !result.filePaths.length) {
    return null;
  }
  
  const filePath = result.filePaths[0];
  try {
    const ExcelJS = require('exceljs');
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    
    const worksheet = workbook.worksheets[0];
    const headerRow = worksheet.getRow(1);
    
    const fields = [];
    headerRow.eachCell((cell, colNumber) => {
      const value = cell.value?.toString().trim();
      if (value) {
        fields.push({
          name: value,
          alias: '',
          enabled: true
        });
      }
    });
    
    return {
      name: path.basename(filePath, path.extname(filePath)),
      fields: fields
    };
  } catch (err) {
    dialog.showErrorBox('导入失败', `无法读取Excel文件: ${err.message}`);
    return null;
  }
});

// 构建文件路径（用于渲染进程）
ipcMain.handle('join-path', (event, ...parts) => {
  return path.join(...parts);
});

// 文件写入API（用于Excel导出）
ipcMain.handle('write-file', async (event, filePath, buffer) => {
  try {
    const bufferData = Buffer.from(buffer);
    await fs.promises.writeFile(filePath, bufferData);
    return { success: true };
  } catch (error) {
    console.error('写入文件失败:', error);
    return { success: false, error: error.message };
  }
});

// 文件读取API（用于加载字段配置）
ipcMain.handle('read-file', async (event, filePath) => {
  try {
    const data = await fs.promises.readFile(filePath, 'utf8');
    return { success: true, data };
  } catch (error) {
    console.error('读取文件失败:', error);
    return { success: false, error: error.message };
  }
});

// 保存临时图片文件（用于PDF转换后的OCR）
ipcMain.handle('save-temp-image', async (event, base64Data) => {
  try {
    const tempDir = path.join(app.getPath('temp'), 'image-text-extractor');
    if (!fs.existsSync(tempDir)) {
      fs.mkdirSync(tempDir, { recursive: true });
    }
    
    const tempFile = path.join(tempDir, `temp_${Date.now()}.png`);
    const buffer = Buffer.from(base64Data, 'base64');
    await fs.promises.writeFile(tempFile, buffer);
    
    return { success: true, path: tempFile };
  } catch (error) {
    console.error('保存临时文件失败:', error);
    return { success: false, error: error.message };
  }
});

// 删除临时文件
ipcMain.handle('delete-temp-file', async (event, filePath) => {
  try {
    if (fs.existsSync(filePath)) {
      await fs.promises.unlink(filePath);
    }
    return { success: true };
  } catch (error) {
    console.error('删除临时文件失败:', error);
    return { success: false, error: error.message };
  }
});