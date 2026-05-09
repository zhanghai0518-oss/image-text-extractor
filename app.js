/**
 * 图片文字提取工具 - 前端逻辑
 */

// 状态管理
const state = {
  images: [],           // 图片列表
  currentIndex: -1,     // 当前选中图片索引
  template: null,       // 当前模板（旧格式）
  fieldConfig: null,    // 字段配置（新格式）
  results: [],          // 提取结果
  extractedData: [],    // 汇总数据
  isProcessing: false,  // 是否正在处理
  currentGroup: '__all__' // 当前显示的分组
};

// PDF处理函数 - 提取图片和文本
async function pdfToImages(fileData) {
  try {
    // 处理base64编码的PDF数据
    let pdfData;
    if (fileData.data.startsWith('data:application/pdf;base64,')) {
      const base64String = fileData.data.split(',')[1];
      pdfData = Uint8Array.from(atob(base64String), c => c.charCodeAt(0));
    } else {
      pdfData = fileData.data;
    }
    
    const pdf = await pdfjsLib.getDocument({ data: pdfData }).promise;
    const images = [];
    const allText = []; // 收集所有页面的文本
    
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      
      // 提取文本内容（核心改进！）
      const textContent = await page.getTextContent();
      const pageText = textContent.items.map(item => item.str).join(' ');
      allText.push(pageText);
      console.log(`[PDF] 第${i}页提取文本长度: ${pageText.length}`);
      
      // 同时生成图片预览
      const scale = 2;
      const viewport = page.getViewport({ scale });
      
      const canvas = document.createElement('canvas');
      const context = canvas.getContext('2d');
      canvas.width = viewport.width;
      canvas.height = viewport.height;
      
      await page.render({
        canvasContext: context,
        viewport: viewport
      }).promise;
      
      const imageData = canvas.toDataURL('image/png');
      images.push({
        data: imageData,
        name: `${fileData.name.replace('.pdf', '')} - 第${i}页`,
        pageNum: i,
        pdfText: pageText // 关键：保存PDF提取的文本！
      });
      console.log(`[PDF] 第${i}页提取文本:`, pageText.substring(0, 100));
    }
    
    // 合并所有页面文本
    const combinedText = allText.join('\n');
    console.log(`[PDF] 总文本长度: ${combinedText.length}`);
    
    return {
      images,
      fullText: combinedText
    };
  } catch (err) {
    console.error('PDF处理错误:', err);
    throw new Error(`PDF处理失败: ${err.message}`);
  }
}

// 自动加载字段配置（启动时）- 从localStorage恢复上次保存的配置
async function loadFieldConfig() {
  console.log('[启动] 尝试从localStorage恢复字段配置');
  
  try {
    // 从localStorage读取上次保存的配置
    const savedConfig = localStorage.getItem('fieldConfig');
    if (savedConfig) {
      const data = JSON.parse(savedConfig);
      state.fieldConfig = data;
      console.log('[启动] 从localStorage恢复字段配置成功，共', data.length, '个字段');
      
      // 渲染字段列表
      const fields = data.map(item => ({
        name: item.columnHeader || item.name || '',
        alias: item.fieldName || item.alias || '',
        enabled: item.enabled !== false,
        aggregate: item.aggregate === true,
        valueType: item.valueType || 'text'
      }));
      
      renderFieldList(fields);
      dom.templateName.value = '已保存的配置';
      
      const aggCount = fields.filter(f => f.aggregate).length;
      console.log('[启动] 分表字段数:', aggCount);
      
      return true;
    } else {
      console.log('[启动] 未找到已保存的字段配置，请导入配置文件');
      return false;
    }
  } catch (err) {
    console.log('[启动] 字段配置加载失败:', err.message);
    return false;
  }
}

// DOM元素
const dom = {
  imageList: document.getElementById('imageList'),
  imageCount: document.getElementById('imageCount'),
  previewContainer: document.getElementById('previewContainer'),
  currentFileName: document.getElementById('currentFileName'),
  resultContainer: document.getElementById('resultContainer'),
  templateSelect: document.getElementById('templateSelect'),
  tableHeader: document.getElementById('tableHeader'),
  tableBody: document.getElementById('tableBody'),
  recordCount: document.getElementById('recordCount'),
  groupTabs: document.getElementById('groupTabs'),
  
  // 按钮
  importImages: document.getElementById('importImages'),
  importFolder: document.getElementById('importFolder'),
  clearAll: document.getElementById('clearAll'),
  editTemplate: document.getElementById('editTemplate'),
  extractAll: document.getElementById('extractAll'),
  exportExcel: document.getElementById('exportExcel'),
  
  // 弹窗
  templateModal: document.getElementById('templateModal'),
  progressModal: document.getElementById('progressModal'),
  fieldList: document.getElementById('fieldList'),
  templateName: document.getElementById('templateName'),
  addField: document.getElementById('addField'),
  saveTemplateBtn: document.getElementById('saveTemplateBtn'),
  cancelTemplate: document.getElementById('cancelTemplate'),
  closeModal: document.getElementById('closeModal'),
  
  // 进度
  progressText: document.getElementById('progressText'),
  progressPercent: document.getElementById('progressPercent'),
  progressFill: document.getElementById('progressFill'),
  cancelProgress: document.getElementById('cancelProgress')
};

// 初始化
async function init() {
  // 检测平台
  const platform = window.electronAPI.platform;
  console.log('[初始化] 平台:', platform);
  console.log('[初始化] OCR引擎:', platform === 'darwin' ? 'Vision OCR (macOS)' : 'PaddleOCR (Windows)');
  
  // 加载模板
  await loadTemplates();
  
  // 自动加载字段配置
  await loadFieldConfig();
  
  // 绑定事件
  bindEvents();
  
  // 监听菜单事件
  window.electronAPI.onMenuImport(() => importImages());
  window.electronAPI.onMenuExport(() => exportExcel());
}

// 加载模板列表
async function loadTemplates() {
  const templates = await window.electronAPI.getTemplates();
  
  dom.templateSelect.innerHTML = '<option value="">选择模板...</option>';
  templates.forEach(t => {
    const option = document.createElement('option');
    option.value = t.id;
    option.textContent = t.name;
    dom.templateSelect.appendChild(option);
  });
  
  // 默认选择第一个模板
  if (templates.length > 0) {
    dom.templateSelect.value = templates[0].id;
    state.template = templates[0];
    updateTableHeader();
  }
}

// 绑定事件
function bindEvents() {
  // 导入按钮
  dom.importImages.addEventListener('click', () => importImages());
  dom.importFolder.addEventListener('click', () => importFolder());
  dom.clearAll.addEventListener('click', clearAll);
  
  // 模板选择
  dom.templateSelect.addEventListener('change', onTemplateChange);
  dom.editTemplate.addEventListener('click', openTemplateEditor);
  
  // 提取和导出
  dom.extractAll.addEventListener('click', extractAll);
  dom.exportExcel.addEventListener('click', exportExcel);
  
  // 模板编辑弹窗
  dom.addField.addEventListener('click', addFieldRow);
  dom.saveTemplateBtn.addEventListener('click', saveTemplate);
  dom.cancelTemplate.addEventListener('click', closeTemplateModal);
  dom.closeModal.addEventListener('click', closeTemplateModal);
  
  // Excel模板导入
  document.getElementById('importExcelTemplate').addEventListener('click', importExcelTemplate);
  
  // 模板文件导入导出
  document.getElementById('exportTemplateBtn').addEventListener('click', exportTemplateFile);
  document.getElementById('importTemplateBtn').addEventListener('click', importTemplateFile);
  
  // 进度弹窗
  dom.cancelProgress.addEventListener('click', cancelProcessing);
  
  // 选择全部复选框
  document.getElementById('selectAll').addEventListener('change', toggleSelectAll);
  
  // 翻条按钮（新增）
  document.getElementById('prevRecord').addEventListener('click', () => navigateRecord(-1));
  document.getElementById('nextRecord').addEventListener('click', () => navigateRecord(1));
  
  // 保存修改按钮（新增）
  document.getElementById('saveChanges').addEventListener('click', saveManualChanges);
  
  // 重新提取按钮（修正功能）
  document.getElementById('retryExtract').addEventListener('click', retryCurrentExtract);
}

// 导入图片
async function importImages() {
  const filePaths = await window.electronAPI.selectImages();
  if (filePaths && filePaths.length > 0) {
    await addImages(filePaths);
  }
}

// 导入文件夹
async function importFolder() {
  const folderPath = await window.electronAPI.selectFolder();
  if (folderPath) {
    // 通过预加载脚本扫描文件夹中的图片
    const files = await window.electronAPI.scanFolder(folderPath);
    
    if (files && files.length > 0) {
      await addImages(files);
    }
  }
}

// 添加图片到列表
async function addImages(filePaths) {
  for (const filePath of filePaths) {
    const ext = filePath.split('.').pop().toLowerCase();
    
    // 处理PDF文件
    if (ext === 'pdf') {
      try {
        const imageData = await window.electronAPI.readImage(filePath);
        if (!imageData.error) {
          // PDF转换为图片 + 提取文本
          const pdfResult = await pdfToImages({
            name: imageData.name,
            data: imageData.data,
            size: imageData.size
          });
          
          // 添加每一页作为一个独立的图片
          for (const img of pdfResult.images) {
            console.log(`[添加] PDF页面 ${img.pageNum}, pdfText长度: ${img.pdfText ? img.pdfText.length : 0}`);
            state.images.push({
              path: filePath,
              name: img.name,
              data: img.data,
              size: imageData.size,
              status: 'pending',
              result: null,
              isPDF: true,
              pageNum: img.pageNum,
              pdfText: img.pdfText // PDF提取的原始文本
            });
          }
        }
      } catch (err) {
        console.error('PDF处理失败:', err);
        alert(`PDF文件处理失败: ${err.message}`);
      }
    } else {
      // 处理图片文件
      const imageData = await window.electronAPI.readImage(filePath);
      if (!imageData.error) {
        state.images.push({
          path: filePath,
          name: imageData.name,
          data: imageData.data,
          size: imageData.size,
          status: 'pending',  // pending, processing, done, error
          result: null
        });
      }
    }
  }
  
  renderImageList();
  updateButtons();
  
  // 自动选择第一张
  if (state.currentIndex === -1 && state.images.length > 0) {
    selectImage(0);
  }
}

// 渲染图片列表
function renderImageList() {
  if (state.images.length === 0) {
    dom.imageList.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">🖼️</div>
        <p>点击上方按钮导入图片</p>
        <p class="hint">支持 JPG, PNG, BMP, GIF, WebP</p>
      </div>
    `;
  } else {
    dom.imageList.innerHTML = state.images.map((img, index) => {
      // 校验状态标记
      let validationIcon = '';
      if (img.status === 'done' && img.validation) {
        if (img.validation.valid) {
          validationIcon = '<span class="validation-pass" title="校验通过">✅</span>';
        } else {
          validationIcon = '<span class="validation-fail" title="需要校验">❗️</span>';
        }
      }
      
      return `
        <div class="image-item ${index === state.currentIndex ? 'active' : ''}" data-index="${index}">
          <img src="${img.data}" alt="${img.name}">
          <span class="name" title="${img.name}">${img.name}</span>
          ${validationIcon}
          <span class="status ${img.status}"></span>
        </div>
      `;
    }).join('');
    
    // 绑定点击事件
    dom.imageList.querySelectorAll('.image-item').forEach(item => {
      item.addEventListener('click', () => selectImage(parseInt(item.dataset.index)));
    });
  }
  
  dom.imageCount.textContent = `${state.images.length} 张图片`;
}

// 选择图片
function selectImage(index) {
  if (index >= 0 && index < state.images.length) {
    state.currentIndex = index;
    const img = state.images[index];
    
    // 更新列表高亮
    renderImageList();
    
    // 显示预览
    dom.previewContainer.innerHTML = `<img src="${img.data}" alt="${img.name}">`;
    dom.currentFileName.textContent = img.name;
    
    // 显示结果
    renderResult(img);
  }
}

// 渲染提取结果（支持多条记录翻条）
function renderResult(img) {
  if (!img.result || Object.keys(img.result).length === 0) {
    dom.resultContainer.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📝</div>
        <p>点击"提取全部"开始识别</p>
      </div>
    `;
    // 隐藏翻条按钮
    document.getElementById('prevRecord').style.display = 'none';
    document.getElementById('nextRecord').style.display = 'none';
    document.getElementById('recordIndicator').style.display = 'none';
    document.getElementById('saveChanges').style.display = 'none';
    return;
  }
  
  // 初始化当前记录索引
  if (!img.currentRecordIndex && img.currentRecordIndex !== 0) {
    img.currentRecordIndex = 0;
  }
  
  // 获取当前要显示的记录
  const allRecords = img.allResults || [img.result];
  const currentRecord = allRecords[img.currentRecordIndex] || img.result;
  const totalRecords = allRecords.length;
  const currentIndex = img.currentRecordIndex + 1;
  
  // 显示/隐藏翻条按钮
  const prevBtn = document.getElementById('prevRecord');
  const nextBtn = document.getElementById('nextRecord');
  const indicator = document.getElementById('recordIndicator');
  const saveBtn = document.getElementById('saveChanges');
  
  if (totalRecords > 1) {
    prevBtn.style.display = 'inline-block';
    nextBtn.style.display = 'inline-block';
    indicator.style.display = 'inline-block';
    indicator.textContent = `${currentIndex}/${totalRecords}`;
    prevBtn.disabled = img.currentRecordIndex === 0;
    nextBtn.disabled = img.currentRecordIndex === totalRecords - 1;
  } else {
    prevBtn.style.display = 'none';
    nextBtn.style.display = 'none';
    indicator.style.display = 'none';
  }
  
  // 显示保存按钮（如果已经有结果）
  saveBtn.style.display = 'inline-block';
  
  // 渲染字段
  const fields = state.template?.fields || [];
  dom.resultContainer.innerHTML = fields.filter(f => f.enabled).map(field => `
    <div class="result-field">
      <span class="label">${field.name}</span>
      <span class="value">
        <input type="text" 
               data-field="${field.name}" 
               value="${currentRecord[field.name] || ''}"
               placeholder="${field.alias || field.name}">
      </span>
    </div>
  `).join('');
  
  // 绑定输入事件（实时更新当前记录）
  dom.resultContainer.querySelectorAll('input').forEach(input => {
    input.addEventListener('change', (e) => {
      const fieldName = e.target.dataset.field;
      currentRecord[fieldName] = e.target.value;
      // 标记为已修改（未保存）
      img.hasUnsavedChanges = true;
      saveBtn.classList.add('btn-warning');
    });
  });
}

// 翻条功能（上一条/下一条）
function navigateRecord(direction) {
  if (state.currentIndex < 0) return;
  const img = state.images[state.currentIndex];
  
  if (!img.allResults || img.allResults.length <= 1) return;
  
  // 计算新索引
  const newIndex = img.currentRecordIndex + direction;
  if (newIndex < 0 || newIndex >= img.allResults.length) return;
  
  // 更新索引并重新渲染
  img.currentRecordIndex = newIndex;
  renderResult(img);
}

// 保存手动修改
function saveManualChanges() {
  if (state.currentIndex < 0) return;
  const img = state.images[state.currentIndex];
  
  // 获取当前显示的记录
  const allRecords = img.allResults || [img.result];
  const currentRecord = allRecords[img.currentRecordIndex || 0];
  
  // 收集所有输入框的值
  const inputs = dom.resultContainer.querySelectorAll('input');
  inputs.forEach(input => {
    const fieldName = input.dataset.field;
    currentRecord[fieldName] = input.value;
  });
  
  // 更新数据汇总表格
  updateDataTable();
  
  // 标记为手动修正
  img.manuallyModified = true;
  img.hasUnsavedChanges = false;
  
  // 更新按钮状态
  const saveBtn = document.getElementById('saveChanges');
  saveBtn.classList.remove('btn-warning');
  
  // 显示成功提示
  alert('✅ 修改已保存！数据汇总表格已更新，导出Excel时会包含这些修改。');
}

// 重新提取当前图片
async function retryCurrentExtract() {
  if (state.currentIndex < 0) return;
  
  const img = state.images[state.currentIndex];
  
  // 确认操作
  if (!confirm(`确定要重新提取 "${img.name}" 吗？\n这将覆盖当前的提取结果。`)) {
    return;
  }
  
  // 清空之前的结果
  img.result = null;
  img.allResults = null;
  img.currentRecordIndex = 0;
  img.status = 'pending';
  img.validation = null;
  img.ocrText = null;
  
  // 更新UI
  renderImageList();
  dom.resultContainer.innerHTML = `
    <div class="empty-state">
      <div class="empty-icon">⏳</div>
      <p>正在重新提取...</p>
    </div>
  `;
  
  try {
    // 重新执行OCR和提取
    let text = '';
    let tempFilePath = null;
    
    // PDF文件：直接用提取的文本
    if (img.isPDF && img.pdfText) {
      text = img.pdfText;
      console.log(`[重新提取-PDF] 文本长度: ${text.length}`);
    } else {
      // 图片文件：使用OCR
      let imagePath = img.path;
      
      // 如果是PDF转换的图片或没有路径，需要先保存到临时文件
      if (img.isPDF || !img.path || img.path.endsWith('.pdf')) {
        const base64Data = img.data.split(',')[1] || img.data;
        const tempResult = await window.electronAPI.saveTempImage(base64Data);
        if (tempResult.success) {
          imagePath = tempResult.path;
          tempFilePath = tempResult.path;
        } else {
          throw new Error('保存临时文件失败: ' + tempResult.error);
        }
      }
      
      // 调用OCR
      const result = await window.electronAPI.ocrImage(imagePath);
      
      // 清理临时文件
      if (tempFilePath) {
        await window.electronAPI.deleteTempFile(tempFilePath);
      }
      
      if (result.success) {
        text = result.text;
        console.log(`[重新提取-OCR] 文本长度: ${text.length}`);
      } else {
        throw new Error(result.error || 'OCR识别失败');
      }
    }
    
    // 解析文本提取字段
    const isPDF = img.isPDF || false;
    const records = parseExtractedText(text, img.name, isPDF);
    
    if (Array.isArray(records) && records.length > 0) {
      img.result = records[0];
      img.allResults = records;
      img.currentRecordIndex = 0;
      img.status = 'done';
      img.ocrText = text;
      
      // 数据校验
      const totalAmount = extractTotalAmount(text);
      const validation = validateRecords(img.allResults, text, totalAmount);
      img.validation = validation;
    } else {
      img.result = records[0] || {};
      img.status = 'done';
    }
    
    // 更新UI
    renderImageList();
    renderResult(img);
    updateDataTable();
    
    alert('✅ 重新提取完成！');
    
  } catch (err) {
    console.error('重新提取失败:', err);
    img.status = 'error';
    img.result = { error: err.message };
    renderImageList();
    
    dom.resultContainer.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">❌</div>
        <p>重新提取失败</p>
        <p style="color: #999; font-size: 12px; margin-top: 8px;">${err.message}</p>
      </div>
    `;
  }
}

// 显示提取日志在界面上
function showExtractLogs(logs, filename) {
  if (!logs || logs.length === 0) return;
  
  const logHtml = `
    <div class="extract-logs" style="margin: 10px 0; padding: 10px; background: #f5f5f5; border-radius: 4px; font-family: monospace; font-size: 12px;">
      <div style="font-weight: bold; margin-bottom: 8px; color: #666;">📋 提取过程 - ${filename}</div>
      ${logs.map(log => `<div style="color: #333; margin: 4px 0;">${escapeHtml(log)}</div>`).join('')}
    </div>
  `;
  
  // 追加到结果区域
  const existingLogs = dom.resultContainer.querySelector('.extract-logs');
  if (existingLogs) {
    existingLogs.outerHTML = logHtml;
  } else {
    dom.resultContainer.innerHTML = logHtml + dom.resultContainer.innerHTML;
  }
}

// HTML转义函数
function escapeHtml(text) {
  const div = document.createElement('div');
  div.textContent = text;
  return div.innerHTML;
}

// 显示验证警告
function showValidationWarning(validation) {
  if (!validation) return;
  
  const warningHtml = `
    <div class="validation-warning" style="margin: 10px 0; padding: 10px; background: ${validation.valid ? '#e8f5e9' : '#fff3e0'}; border-left: 4px solid ${validation.valid ? '#4caf50' : '#ff9800'}; border-radius: 4px;">
      <div style="font-weight: bold; color: ${validation.valid ? '#2e7d32' : '#e65100'};">
        ${validation.message}
      </div>
    </div>
  `;
  
  // 追加到结果区域
  dom.resultContainer.innerHTML = warningHtml + dom.resultContainer.innerHTML;
}

// 提取合计金额（用于验证）
function extractTotalAmount(text) {
  // 匹配格式：金额合计 (大写)xxx ¥ 5,780.23 或 合计 ¥5780.23
  // 支持半角¥和全角￥符号
  // 支持逗号分隔(5,780.23)和点号分隔(5.780.13)
  const totalMatch = text.match(/(?:金额)?合计[^¥￥]*[¥￥]\s*([\d,.\s]+\.\d{2})/);
  if (totalMatch) {
    return parseFloat(totalMatch[1].replace(/[,.\s]/g, '').replace(/(\d{2})$/, '.$1'));
  }
  
  // 匹配格式：¥ 5,780.23 或 ￥20.010.13（最后一行带¥/￥符号的金额）
  const yenMatch = text.match(/[¥￥]\s*([\d,.\s]+\.\d{2})/);
  if (yenMatch) {
    // 处理点号分隔的金额（如 20.010.13 → 20010.13）
    let amountStr = yenMatch[1].replace(/,/g, '');
    // 如果有多于一个点号，说明是千位分隔符
    const dots = (amountStr.match(/\./g) || []).length;
    if (dots > 1) {
      // 移除所有点号，然后重新添加小数点
      amountStr = amountStr.replace(/\./g, '');
      amountStr = amountStr.replace(/(\d{2})$/, '.$1');
    }
    return parseFloat(amountStr);
  }
  
  return 0;
}

// 验证记录金额之和
function validateRecordAmounts(records, totalAmount) {
  if (!totalAmount || totalAmount === 0 || !records || records.length === 0) {
    return { valid: true, message: '无需验证' };
  }
  
  let sum = 0;
  console.log('=== 金额验证详情 ===');
  records.forEach((r, idx) => {
    const amountStr = r['金额'] || r.amount || '0';
    const amount = parseFloat(amountStr.toString().replace(/,/g, '')) || 0;
    console.log(`记录${idx + 1}: 金额字段="${amountStr}", 数值=${amount}`);
    sum += amount;
  });
  console.log(`记录合计: ${sum.toFixed(2)}, 凭证合计: ${totalAmount.toFixed(2)}`);
  
  const diff = Math.abs(sum - totalAmount);
  const isValid = diff < 0.01;
  
  return {
    valid: isValid,
    totalAmount: totalAmount,
    recordsSum: sum.toFixed(2),
    difference: diff.toFixed(2),
    message: isValid 
      ? '✅ 金额验证通过' 
      : `⚠️ 金额验证失败：记录合计 ${sum.toFixed(2)} 元，凭证合计 ${totalAmount.toFixed(2)} 元，差异 ${diff.toFixed(2)} 元`
  };
}

// 尝试补全遗漏的记录
function trySupplementRecords(records, text, totalAmount) {
  // 计算已提取的金额之和
  const existingSum = records.reduce((sum, r) => {
    const amount = parseFloat((r['金额'] || r.amount || '0').toString().replace(/,/g, '')) || 0;
    return sum + amount;
  }, 0);
  
  const missingAmount = totalAmount - existingSum;
  if (missingAmount < 0.01) return null; // 差异太小，不需要补全
  
  console.log('[补全] 检测到遗漏金额:', missingAmount.toFixed(2));
  
  // 从文本中提取所有可能的金额（包括OCR错误的格式）
  // 格式1: 标准格式 34,661.02
  // 格式2: 点号分隔 11.879.21
  const allAmounts = [];
  
  // 匹配标准金额格式
  const standardMatches = text.match(/\d{1,3}(?:,\d{3})*\.\d{2}/g) || [];
  standardMatches.forEach(m => {
    const clean = m.replace(/,/g, '');
    const val = parseFloat(clean);
    if (val > 0.01 && val < 10000000) {
      allAmounts.push({ original: m, value: val, standard: true });
    }
  });
  
  // 匹配点号分隔格式（如 11.879.21）
  const dotMatches = text.match(/\d{1,3}\.\d{3}\.\d{2}/g) || [];
  dotMatches.forEach(m => {
    // 转换为标准格式：11.879.21 → 11879.21
    const parts = m.split('.');
    if (parts.length === 3) {
      const clean = parts[0] + parts[1] + '.' + parts[2];
      const val = parseFloat(clean);
      if (val > 0.01 && val < 10000000) {
        allAmounts.push({ original: m, value: val, standard: false });
      }
    }
  });
  
  console.log('[补全] 找到的所有金额:', allAmounts.map(a => `${a.original}→${a.value}`));
  
  // 找出已使用的金额
  const usedAmounts = records.map(r => parseFloat((r['金额'] || r.amount || '0').toString().replace(/,/g, '')) || 0);
  console.log('[补全] 已使用的金额:', usedAmounts);
  
  // 找出未使用的金额（接近遗漏金额）
  const missingAmounts = allAmounts.filter(a => {
    // 排除合计金额
    if (Math.abs(a.value - totalAmount) < 0.01) return false;
    // 排除已使用的金额
    return !usedAmounts.some(u => Math.abs(u - a.value) < 0.01);
  });
  
  console.log('[补全] 未使用的金额:', missingAmounts.map(a => a.value));
  
  // 检查是否有遗漏金额匹配
  const matched = missingAmounts.find(a => Math.abs(a.value - missingAmount) < 0.01);
  if (matched) {
    console.log('[补全] 找到遗漏金额:', matched.original, '→', matched.value);
    
    // 创建新记录（复制第一条记录的信息，修改金额）
    if (records.length > 0) {
      const template = { ...records[0] };
      template['金额'] = matched.value.toString();
      template['备注'] = '[自动补全]';
      
      // 尝试从文本中提取更多信息（如税目）
      const taxItemMatch = text.match(/工资薪金所得/g) || [];
      if (taxItemMatch.length > records.length) {
        template['税目名称'] = '工资薪金所得';
      }
      
      return [...records, template];
    }
  }
  
  return null;
}

// 数据校验：检查字段完整性和金额
function validateRecords(records, text, totalAmount) {
  const result = {
    valid: true,
    issues: [],
    emptyFields: []
  };
  
  // 1. 检查字段完整性（关键字段不能为空）
  const keyFields = ['所属公司', '金额', '税种名称', '税款所属期起', '税款所属期止'];
  
  records.forEach((record, idx) => {
    keyFields.forEach(field => {
      const value = record[field];
      if (!value || value.trim() === '' || value === 'undefined') {
        result.valid = false;
        result.issues.push(`记录${idx + 1}: ${field}为空`);
        result.emptyFields.push({ recordIndex: idx, field });
      }
    });
  });
  
  // 2. 检查金额校验
  if (totalAmount > 0 && records.length > 0) {
    const validation = validateRecordAmounts(records, totalAmount);
    if (!validation.valid) {
      result.valid = false;
      result.issues.push(`金额校验失败: 记录合计${validation.recordsSum}元, 凭证合计${validation.totalAmount}元, 差异${validation.difference}元`);
      result.amountValidation = validation;
    }
  }
  
  return result;
}

// 模板变化
async function onTemplateChange(e) {
  if (!e.target.value) {
    state.template = null;
    return;
  }
  
  const templates = await window.electronAPI.getTemplates();
  state.template = templates.find(t => t.id === e.target.value);
  updateTableHeader();
  
  // 重新渲染当前结果
  if (state.currentIndex >= 0) {
    renderResult(state.images[state.currentIndex]);
  }
}

// 更新表格表头
function updateTableHeader() {
  let headers = ['序号', '来源文件'];
  
  if (state.fieldConfig) {
    // 使用新配置格式
    headers.push(...state.fieldConfig.filter(f => f.enabled !== false).map(f => f.columnHeader));
  } else if (state.template) {
    // 使用旧模板格式
    headers.push(...state.template.fields.filter(f => f.enabled).map(f => f.name));
  }
  
  dom.tableHeader.innerHTML = `
    <th><input type="checkbox" id="selectAll"></th>
    ${headers.map(h => `<th>${h}</th>`).join('')}
  `;
  
  document.getElementById('selectAll').addEventListener('change', toggleSelectAll);
}

// 更新数据表格（支持分组）
function updateDataTable() {
  // 展开所有记录
  const allRecords = [];
  state.images.forEach((img, imgIndex) => {
    if (img.status === 'done') {
      if (img.allResults && img.allResults.length > 0) {
        img.allResults.forEach((record, recordIndex) => {
          allRecords.push({
            imgIndex,
            recordIndex,
            img,
            record,
            fileName: img.name,
            isFirstRecord: recordIndex === 0,
            totalRecords: img.allResults.length
          });
        });
      } else if (img.result) {
        allRecords.push({
          imgIndex,
          recordIndex: 0,
          img,
          record: img.result,
          fileName: img.name,
          isFirstRecord: true,
          totalRecords: 1
        });
      }
    }
  });
  
  // 检查是否需要分组显示
  if (state.fieldConfig && window.FieldConfigManager) {
    window.FieldConfigManager.loadConfig(state.fieldConfig);
    const aggregateFields = window.FieldConfigManager.getAggregateFields();
    
    if (aggregateFields.length > 0 && allRecords.length > 0) {
      // 按aggregate字段分组（保留完整对象，包括fileName）
      const groups = window.FieldConfigManager.groupByAggregate(allRecords, 'record');
      
      // 渲染分组标签
      renderGroupTabs(groups, aggregateFields[0].columnHeader);
      
      // 只显示当前分组
      const currentGroupRecords = groups[state.currentGroup] || allRecords;
      renderTableBody(currentGroupRecords, allRecords);
      
      dom.recordCount.textContent = `(${currentGroupRecords.length}/${allRecords.length}条)`;
      return;
    }
  }
  
  // 无分组时隐藏标签
  dom.groupTabs.style.display = 'none';
  dom.recordCount.textContent = `(${allRecords.length}条)`;
  renderTableBody(allRecords, allRecords);
}

// 渲染分组标签
function renderGroupTabs(groups, groupField) {
  const groupNames = Object.keys(groups);
  
  if (groupNames.length <= 1) {
    dom.groupTabs.style.display = 'none';
    return;
  }
  
  dom.groupTabs.style.display = 'flex';
  dom.groupTabs.innerHTML = groupNames.map(name => `
    <button class="group-tab ${name === state.currentGroup ? 'active' : ''}" data-group="${name}">
      ${name} (${groups[name].length}条)
    </button>
  `).join('');
  
  // 绑定点击事件
  dom.groupTabs.querySelectorAll('.group-tab').forEach(tab => {
    tab.addEventListener('click', () => {
      state.currentGroup = tab.dataset.group;
      updateDataTable();
    });
  });
}

// 渲染表格内容
function renderTableBody(displayRecords, allRecords) {
  if (displayRecords.length === 0) {
    dom.tableBody.innerHTML = `
      <tr class="empty-row">
        <td colspan="100">
          <div class="empty-state small">
            <p>暂无数据</p>
          </div>
        </td>
      </tr>
    `;
    return;
  }
  
  // 获取列头
  let headers = [];
  if (state.fieldConfig) {
    headers = state.fieldConfig.filter(f => f.enabled !== false).map(f => f.columnHeader);
  } else if (state.template) {
    headers = state.template.fields.filter(f => f.enabled).map(f => f.name);
  }
  
  // 构建表格行
  dom.tableBody.innerHTML = displayRecords.map((item, index) => {
    const record = item.record || item;
    const fileName = item.fileName || item.img?.name || '';
    const totalRecords = item.totalRecords || 1;
    const recordIndex = item.recordIndex || 0;
    
    return `
      <tr>
        <td><input type="checkbox" class="row-check" data-index="${index}"></td>
        <td>${index + 1}</td>
        <td title="${fileName}${totalRecords > 1 ? ' (第' + (recordIndex + 1) + '条)' : ''}">${fileName}${totalRecords > 1 ? '<span style="color:#999">[' + (recordIndex + 1) + '/' + totalRecords + ']</span>' : ''}</td>
        ${headers.map(h => `<td>${record?.[h] || ''}</td>`).join('')}
      </tr>
    `;
  }).join('');
}

// 提取全部
async function extractAll() {
  if (!state.template) {
    alert('请先选择模板');
    return;
  }
  
  if (state.images.length === 0) {
    alert('请先导入图片');
    return;
  }
  
  state.isProcessing = true;
  showProgressModal();
  
  const total = state.images.length;
  
  for (let i = 0; i < total; i++) {
    if (!state.isProcessing) break;
    
    const img = state.images[i];
    img.status = 'processing';
    renderImageList();
    
    updateProgress(i, `正在处理: ${img.name}`);
    
    try {
      let text = '';
      let tempFilePath = null;
      
      console.log(`[提取] 文件: ${img.name}, isPDF: ${img.isPDF}, pdfText: ${img.pdfText ? '有(' + img.pdfText.length + ')' : '无'}`);
      
      // PDF文件：直接用 pdf.js 提取文本（不走OCR）
      if (img.isPDF && img.pdfText) {
        text = img.pdfText;
        console.log(`[PDF] 直接提取PDF文本，长度: ${text.length}`);
      } else {
        // 图片文件：使用本地OCR工具
        // Windows: PaddleOCR-json（已集成）
        // macOS: Vision OCR
        
        let imagePath = img.path;
        
        // 如果是PDF转换的图片或没有路径，需要先保存到临时文件
        if (img.isPDF || !img.path || img.path.endsWith('.pdf')) {
          const base64Data = img.data.split(',')[1] || img.data;
          const tempResult = await window.electronAPI.saveTempImage(base64Data);
          if (tempResult.success) {
            imagePath = tempResult.path;
            tempFilePath = tempResult.path;
            console.log(`[OCR] 保存临时文件: ${tempFilePath}`);
          } else {
            throw new Error('保存临时文件失败: ' + tempResult.error);
          }
        }
        
        // 调用本地OCR（main.js会根据平台自动选择引擎）
        const result = await window.electronAPI.ocrImage(imagePath);
        
        // 清理临时文件
        if (tempFilePath) {
          await window.electronAPI.deleteTempFile(tempFilePath);
        }
        
        if (result.success) {
          text = result.text;
          console.log(`[OCR] 图片识别文本长度: ${text.length}`);
          console.log(`[OCR] 前200字符: ${text.substring(0, 200)}`);
        } else {
          // 显示详细错误信息
          console.error('[OCR] 识别失败:', result.error);
          if (result.rawOutput) {
            console.error('[OCR] PaddleOCR原始输出:', result.rawOutput);
          }
          if (result.stderr) {
            console.error('[OCR] PaddleOCR stderr:', result.stderr);
          }
          throw new Error(result.error || 'OCR识别失败');
        }
      }
      
      // 解析文本，提取字段（返回数组，支持多税目）
      const isPDF = img.isPDF || false;
      const records = parseExtractedText(text, img.name, isPDF);
      if (Array.isArray(records) && records.length > 0) {
        // 多条记录的情况
        img.result = records[0]; // 主记录用于显示
        img.extraResults = records.slice(1); // 额外记录
        img.allResults = records; // 所有记录用于导出
      } else {
        img.result = records[0] || {};
      }
      img.status = 'done';
      img.ocrText = text; // 保存原始OCR文本
      
      // 数据校验
      const totalAmount = extractTotalAmount(text);
      const validation = validateRecords(img.allResults || [img.result], text, totalAmount);
      img.validation = validation;
      
      if (!validation.valid) {
        console.warn(`[校验] ${img.name}: 校验失败`, validation.issues);
      } else {
        console.log(`[校验] ${img.name}: 校验通过`);
      }
      
    } catch (err) {
      console.error('OCR错误:', err);
      img.status = 'error';
      img.result = { error: err.message };
    }
    
    renderImageList();
  }
  
  hideProgressModal();
  state.isProcessing = false;
  
  // 显示第一张完成的结果
  const firstDone = state.images.findIndex(img => img.status === 'done');
  if (firstDone >= 0) {
    selectImage(firstDone);
  }
  
  updateDataTable();
  updateButtons();
}

// 图片预处理函数（提高OCR识别率）
function preprocessImage(imageData, options = {}) {
  const { scale = 2, contrast = 1.5, brightness = 10, grayscale = true } = options;
  
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => {
      // 创建canvas
      const canvas = document.createElement('canvas');
      const ctx = canvas.getContext('2d');
      
      // 放大图片
      canvas.width = img.width * scale;
      canvas.height = img.height * scale;
      
      // 绘制放大后的图片
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);
      
      // 获取像素数据
      const imageDataObj = ctx.getImageData(0, 0, canvas.width, canvas.height);
      const data = imageDataObj.data;
      
      // 遍历每个像素进行增强
      for (let i = 0; i < data.length; i += 4) {
        let r = data[i];
        let g = data[i + 1];
        let b = data[i + 2];
        
        // 灰度化（可选）
        if (grayscale) {
          const gray = 0.299 * r + 0.587 * g + 0.114 * b;
          r = g = b = gray;
        }
        
        // 调整亮度
        r += brightness;
        g += brightness;
        b += brightness;
        
        // 调整对比度
        r = ((r - 128) * contrast) + 128;
        g = ((g - 128) * contrast) + 128;
        b = ((b - 128) * contrast) + 128;
        
        // 二值化（黑白化，提高文字清晰度）
        const threshold = 180;
        const avg = (r + g + b) / 3;
        if (avg < threshold) {
          // 深色变黑（文字）
          r = g = b = 0;
        } else {
          // 浅色变白（背景）
          r = g = b = 255;
        }
        
        // 限制范围
        data[i] = Math.min(255, Math.max(0, r));
        data[i + 1] = Math.min(255, Math.max(0, g));
        data[i + 2] = Math.min(255, Math.max(0, b));
      }
      
      // 放回canvas
      ctx.putImageData(imageDataObj, 0, 0);
      
      // 返回处理后的base64
      resolve(canvas.toDataURL('image/png'));
    };
    img.src = imageData;
  });
}

// OCR错误修正函数（Windows OCR常见错误）
function correctOCRErrors(text) {
  console.log('[OCR修正] 原始文本长度:', text.length);
  
  let corrected = text;
  
  // 1. 处理字符间多余空格（Windows OCR常见问题）
  // 检测是否每个字符间都有空格："中 税 华 收" → "中税华收"
  // 但要保留正常的单词间空格
  const spacePattern = /([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])/g;
  
  // 统计中文字符间空格的密度
  const chineseChars = corrected.match(/[\u4e00-\u9fa5]/g) || [];
  const chineseWithSpaces = corrected.match(/[\u4e00-\u9fa5]\s+[\u4e00-\u9fa5]/g) || [];
  const spaceRatio = chineseChars.length > 0 ? chineseWithSpaces.length / chineseChars.length : 0;
  
  if (spaceRatio > 0.3) {
    // 如果超过30%的中文字符间有空格，说明是OCR问题，去掉中文字符间的空格
    corrected = corrected.replace(/([\u4e00-\u9fa5])\s+([\u4e00-\u9fa5])/g, '$1$2');
    console.log('[OCR修正] 去除中文字符间空格（比例:', (spaceRatio * 100).toFixed(1) + '%）');
  }
  
  // 2. 常见OCR字符错误映射
  const corrections = {
    // 常见字符错误
    '杌关': '机关',
    '汌': '局',
    '叹': '汉',
    '發': '发',
    '岔': '区',
    '热务局': '税务局',
    '热务所': '税务所',
    
    // 数字/日期格式错误
    '在': '年',  // "202在03" → "202年03"
    '。': '.',   // "03。01" → "03.01"
    '•': '.',    // 日期分隔符
    
    // 税务相关错误
    '个人所保税': '个人所得税',
    '个人所报税': '个人所得税',
    '企业所保税': '企业所得税',
  };
  
  for (const [wrong, right] of Object.entries(corrections)) {
    corrected = corrected.split(wrong).join(right);
  }
  
  // 3. 修正日期格式：2026在03在01 → 2026年03月01
  corrected = corrected.replace(/(\d{4})在(\d{2})在(\d{2})/g, '$1年$2月$3日');
  corrected = corrected.replace(/(\d{4})在(\d{1})在(\d{2})/g, '$1年0$2月$3日');
  corrected = corrected.replace(/(\d{4})在(\d{2})在(\d{1})/g, '$1年$2月0$3日');
  
  // 4. 修正日期分隔符：03。01 → 03.01
  corrected = corrected.replace(/(\d{2})。(\d{2})/g, '$1.$2');
  corrected = corrected.replace(/(\d{1})。(\d{2})/g, '0$1.$2');
  
  // 5. 修正金额格式：5, 566.54 → 5,566.54（去掉逗号后空格）
  corrected = corrected.replace(/(\d),\s+(\d{3})/g, '$1,$2');
  
  // 6. 修正公司名称前缀：或汉 → 武汉
  corrected = corrected.replace(/^或汉/gm, '武汉');
  
  // 7. 修正No.格式
  corrected = corrected.replace(/No[.•·]/g, 'No.');
  
  if (corrected !== text) {
    console.log('[OCR修正] 已修正OCR错误');
    console.log('[OCR修正] 修正后文本（前500字符）:', corrected.substring(0, 500));
  }
  
  // 打印修正后的完整文本（用于调试）
  console.log('[OCR修正] 修正后完整文本:\n', corrected);
  
  return corrected;
}

// 解析提取的文本 - 返回数组支持多税目
// isPDF: true表示PDF文件（用FieldMatcher），false表示图片（用TaxParser）
function parseExtractedText(text, filename = '', isPDF = false) {
  console.log('[parseExtractedText] isPDF:', isPDF, 'fieldConfig:', !!state.fieldConfig, 'FieldMatcher:', !!window.FieldMatcher);
  
  // 修正OCR错误
  text = correctOCRErrors(text);
  
  // PDF文件：使用字段配置匹配器（支持动态识别税种）
  if (isPDF && window.FieldMatcher) {
    // 显示提取过程日志
    const logs = [];
    const matcher = window.FieldMatcher;
    const originalLog = matcher.log.bind(matcher);
    matcher.log = (...args) => {
      logs.push(args.join(' '));
      originalLog(...args);
    };
    
    // 传入字段配置（如果有），否则使用默认字段列表
    const result = matcher.match(text, state.fieldConfig || null);
    
    // 恢复原始日志函数
    matcher.log = originalLog;
    
    // 在结果区域显示提取日志
    showExtractLogs(logs, filename);
    
    // 处理新的返回格式（包含验证信息）
    let records = result.records || result;
    const validation = result.validation;
    
    // 显示验证结果
    if (validation && !validation.valid) {
      console.warn('[验证]', validation.message);
      // 在界面上显示验证警告
      showValidationWarning(validation);
    } else if (validation && validation.valid) {
      console.log('[验证]', validation.message);
    }
    
    // 检查记录是否有效 - 关键字段不能全为空
    if (records && records.length > 0) {
      const hasValidData = records.some(r => {
        // 检查是否有至少一个关键字段有值
        return (r['所属公司'] && r['所属公司'].length > 2) ||
               (r['金额'] && r['金额'].length > 0) ||
               (r['税种名称'] && r['税种名称'].length > 0);
      });
      
      if (hasValidData) {
        // 备注字段只保留原始提取的内容，不再追加来源信息
        console.log('[parseExtractedText] FieldMatcher返回有效记录:', records.length, '条');
        return records;
      } else {
        console.log('[parseExtractedText] FieldMatcher返回的记录关键字段为空，回退到TaxParser');
      }
    }
  }
  
  // 图片文件：使用增强版TaxParser（针对Vision OCR文本优化）
  if (window.TaxParser) {
    const records = window.TaxParser.parse(text, filename);
    if (records && records.length > 0) {
      // 为图片格式也添加验证
      const totalAmount = extractTotalAmount(text);
      if (totalAmount > 0) {
        const validation = validateRecordAmounts(records, totalAmount);
        if (!validation.valid) {
          console.warn('[验证-图片]', validation.message);
          
          // 尝试补全遗漏的记录
          const supplemented = trySupplementRecords(records, text, totalAmount);
          if (supplemented && supplemented.length > records.length) {
            console.log('[补全] 成功补充', supplemented.length - records.length, '条遗漏记录');
            showValidationWarning({
              valid: validateRecordAmounts(supplemented, totalAmount).valid,
              message: `⚠️ 原始提取${records.length}条记录金额不匹配，已自动补全${supplemented.length}条`
            });
            return supplemented;
          }
          
          showValidationWarning(validation);
        } else {
          console.log('[验证-图片]', validation.message);
        }
      }
      return records; // 返回数组
    }
  }
  
  // 回退到原有逻辑
  const result = {};
  
  if (!state.template) return [result];
  
  const fields = state.template.fields.filter(f => f.enabled);
  
  // 通用正则匹配规则
  const patterns = {
    '所属公司': [/纳税人名称[：:]\s*([^\n]+)/, /公司名称[：:]\s*([^\n]+)/],
    '所属公司代码': [/纳税人识别号[：:]\s*([A-Za-z0-9]+)/, /统一社会信用代码[：:]\s*([A-Za-z0-9]+)/],
    '金额': [/实缴金额[（(]元[)）][：:]\s*([\d,.]+)|金额[：:]\s*([\d,.]+)\s*元|¥\s*([\d,.]+)/],
    '税票号码': [/税票号码[：:]\s*(\d+)/, /No[.•]?\s*(\d+)/],
    '税款所属期起': [/税款所属期[起]*[：:]\s*(\d{4}[-年]\d{1,2}[-月]\d{1,2}[日]?)/],
    '税款所属期止': [/至\s*(\d{4}[-年]\d{1,2}[-月]\d{1,2}[日]?)/],
    '缴款日期': [/缴款日期[：:]\s*(\d{4}[-年]\d{1,2}[-月]\d{1,2}[日]?)/, /入.*库日期[：:]\s*(\d{4}[-年]\d{1,2}[-月]\d{1,2}[日]?)/],
    '征收机关': [/税务机关[：:]\s*([^\n]+)/, /征收机关[：:]\s*([^\n]+)/],
    '主管税务所': [/主管税务所[（(]科、分局[)）][：:]\s*([^\n]+)/]
  };
  
  fields.forEach(field => {
    const fieldName = field.name;
    const aliases = [fieldName, field.alias].filter(Boolean);
    
    // 先尝试用别名在文本中直接查找
    for (const alias of aliases) {
      const regex = new RegExp(alias + '[：:]*\\s*([^\\n]+)', 'i');
      const match = text.match(regex);
      if (match) {
        result[fieldName] = match[1].trim();
        break;
      }
    }
    
    // 如果没找到，尝试用预定义模式
    if (!result[fieldName] && patterns[fieldName]) {
      for (const pattern of patterns[fieldName]) {
        const match = text.match(pattern);
        if (match) {
          result[fieldName] = match[1] || match[2] || match[3];
          break;
        }
      }
    }
  });
  
  return result;
}

// 导出Excel（支持按分表字段分组导出到不同文件或不同sheet）
async function exportExcel() {
  // 获取所有记录
  const allRecords = [];
  state.images.forEach((img) => {
    if (img.status === 'done') {
      if (img.allResults && img.allResults.length > 0) {
        img.allResults.forEach((record) => {
          allRecords.push({ fileName: img.name, record });
        });
      } else if (img.result) {
        allRecords.push({ fileName: img.name, record: img.result });
      }
    }
  });
  
  if (allRecords.length === 0) {
    alert('没有可导出的数据');
    return;
  }
  
  // 获取列头
  let headers = ['序号', '来源文件'];
  if (state.fieldConfig) {
    headers.push(...state.fieldConfig.filter(f => f.enabled !== false).map(f => f.columnHeader));
  } else if (state.template) {
    headers.push(...state.template.fields.filter(f => f.enabled).map(f => f.name));
  }
  
  // 检查是否需要分表导出
  console.log('[导出] state.fieldConfig:', state.fieldConfig ? '已设置' : '未设置');
  console.log('[导出] FieldConfigManager:', window.FieldConfigManager ? '已加载' : '未加载');
  
  // 如果字段配置未设置，尝试加载
  if (!state.fieldConfig) {
    console.log('[导出] 字段配置未设置，尝试加载...');
    await loadFieldConfig();
  }
  
  if (state.fieldConfig && window.FieldConfigManager) {
    window.FieldConfigManager.loadConfig(state.fieldConfig);
    const aggregateFields = window.FieldConfigManager.getAggregateFields();
    
    console.log('[导出] 分表字段:', aggregateFields);
    console.log('[导出] 分表字段数量:', aggregateFields.length);
    
    if (aggregateFields.length > 0) {
      // 有分表字段，按字段分组
      const groupField = aggregateFields[0].columnHeader;
      // 分组时保留 fileName 信息
      const groups = {};
      allRecords.forEach(item => {
        const key = item.record[groupField] || '未分类';
        if (!groups[key]) {
          groups[key] = [];
        }
        // 将 fileName 添加到 record 中
        const recordWithFileName = { ...item.record, '来源文件': item.fileName };
        groups[key].push(recordWithFileName);
      });
      const groupCount = Object.keys(groups).length;
      
      console.log('[导出] 分组字段:', groupField, '分组数:', groupCount, '分组:', Object.keys(groups));
      
      // 无论有几组，都弹出选择框
      const choice = await showExportChoiceDialog(groupField, groupCount);
      console.log('[导出] 用户选择:', choice);
      if (!choice) return; // 用户取消
      
      if (choice === 'files') {
        // 导出到多个Excel文件
        await exportToMultipleFiles(headers, groups, groupField);
        return;
      } else {
        // 导出到单个Excel的不同sheet
        await exportToMultipleSheets(headers, groups, groupField);
        return;
      }
    }
  }
  
  // 无分组时导出单个Sheet
  console.log('[导出] 无分组或无分表字段，导出单个Sheet');
  const filePath = await window.electronAPI.saveExcel(`提取结果_${new Date().toISOString().slice(0,10)}.xlsx`);
  if (!filePath) return;
  
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('提取结果');
  worksheet.addRow(headers);
  worksheet.getRow(1).font = { bold: true };
  worksheet.getRow(1).fill = {
    type: 'pattern', pattern: 'solid',
    fgColor: { argb: 'FFE5E5EA' }
  };
  
  allRecords.forEach((item, index) => {
    const row = [index + 1, item.fileName];
    headers.slice(2).forEach(h => {
      row.push(item.record?.[h] || '');
    });
    worksheet.addRow(row);
  });
  
  worksheet.columns.forEach((col, i) => {
    let maxLen = (headers[i] || '').length;
    worksheet.eachRow((row, rowNum) => {
      if (rowNum > 1) {
        const cell = row.getCell(i + 1);
        const len = (cell.value || '').toString().length;
        maxLen = Math.max(maxLen, len);
      }
    });
    col.width = Math.min(maxLen + 2, 50);
  });
  
  // 使用Buffer方式写入
  const buffer = await workbook.xlsx.writeBuffer();
  await window.electronAPI.writeFile(filePath, buffer);
  alert(`已导出到: ${filePath}`);
}

// 显示导出选择对话框
function showExportChoiceDialog(groupField, groupCount) {
  return new Promise((resolve) => {
    // 创建遮罩和对话框
    const overlay = document.createElement('div');
    overlay.style.cssText = `
      position: fixed;
      top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(0,0,0,0.5);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 10000;
    `;
    
    const dialog = document.createElement('div');
    dialog.style.cssText = `
      background: white;
      padding: 24px;
      border-radius: 12px;
      box-shadow: 0 4px 24px rgba(0,0,0,0.2);
      max-width: 400px;
      text-align: center;
    `;
    
    dialog.innerHTML = `
      <h3 style="margin: 0 0 16px; color: #333;">选择导出方式</h3>
      <p style="color: #666; margin-bottom: 20px;">
        数据将按 <strong>"${groupField}"</strong> 字段分组<br>
        共 ${groupCount} 个分组
      </p>
      <div style="display: flex; gap: 12px; justify-content: center;">
        <button id="exportSheets" style="
          padding: 12px 24px;
          border: 1px solid #007AFF;
          background: white;
          color: #007AFF;
          border-radius: 8px;
          cursor: pointer;
          font-size: 14px;
        ">📁 单文件多Sheet</button>
        <button id="exportFiles" style="
          padding: 12px 24px;
          border: none;
          background: #007AFF;
          color: white;
          border-radius: 8px;
          cursor: pointer;
          font-size: 14px;
        ">📂 多个Excel文件</button>
      </div>
      <button id="exportCancel" style="
        margin-top: 16px;
        padding: 8px 16px;
        border: none;
        background: transparent;
        color: #999;
        cursor: pointer;
        font-size: 13px;
      ">取消</button>
    `;
    
    overlay.appendChild(dialog);
    document.body.appendChild(overlay);
    
    // 绑定事件
    document.getElementById('exportSheets').onclick = () => {
      document.body.removeChild(overlay);
      resolve('sheets');
    };
    document.getElementById('exportFiles').onclick = () => {
      document.body.removeChild(overlay);
      resolve('files');
    };
    document.getElementById('exportCancel').onclick = () => {
      document.body.removeChild(overlay);
      resolve(null);
    };
  });
}

// 导出到多个Excel文件
async function exportToMultipleFiles(headers, groups, groupField) {
  // 选择保存目录
  const folderPath = await window.electronAPI.selectFolder();
  if (!folderPath) return;
  
  let exportedCount = 0;
  const safeDate = new Date().toISOString().slice(0, 10);
  
  for (const [groupName, records] of Object.entries(groups)) {
    // 清理文件名中的非法字符
    const safeName = groupName.replace(/[\\/:*?"<>|]/g, '_').substring(0, 50);
    const fileName = `${safeName}_${safeDate}.xlsx`;
    const filePath = await window.electronAPI.joinPath(folderPath, fileName);
    
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('数据');
    
    // 添加表头
    worksheet.addRow(headers);
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: 'pattern', pattern: 'solid',
      fgColor: { argb: 'FFE5E5EA' }
    };
    
    // 添加数据（包含来源文件列）
    records.forEach((record, idx) => {
      const row = [idx + 1, record['来源文件'] || record.fileName || ''];
      headers.slice(2).forEach(h => {
        row.push(record[h] || '');
      });
      worksheet.addRow(row);
    });
    
    // 调整列宽
    worksheet.columns.forEach((col, i) => {
      let maxLen = (headers[i] || '').length;
      worksheet.eachRow((row, rowNum) => {
        if (rowNum > 1) {
          const cell = row.getCell(i + 1);
          const len = (cell.value || '').toString().length;
          maxLen = Math.max(maxLen, len);
        }
      });
      col.width = Math.min(maxLen + 2, 50);
    });
    
    // 使用Buffer方式写入
    const buffer = await workbook.xlsx.writeBuffer();
    await window.electronAPI.writeFile(filePath, buffer);
    exportedCount++;
    console.log(`导出: ${fileName} (${records.length}条)`);
  }
  
  alert(`已导出 ${exportedCount} 个Excel文件到:\n${folderPath}`);
}

// 导出到单个Excel的多个sheet
async function exportToMultipleSheets(headers, groups, groupField) {
  console.log('[导出] 开始多Sheet导出，分组数:', Object.keys(groups).length);
  const filePath = await window.electronAPI.saveExcel(`提取结果_${new Date().toISOString().slice(0,10)}.xlsx`);
  if (!filePath) return;
  
  const workbook = new ExcelJS.Workbook();
  
  // 首先创建汇总sheet
  const summarySheet = workbook.addWorksheet('汇总');
  summarySheet.addRow([...headers.slice(0, 2), groupField, '记录数']);
  summarySheet.getRow(1).font = { bold: true };
  summarySheet.getRow(1).fill = {
    type: 'pattern', pattern: 'solid',
    fgColor: { argb: 'FFE5E5EA' }
  };
  
  let totalRecords = 0;
  Object.entries(groups).forEach(([groupName, records], idx) => {
    summarySheet.addRow([idx + 1, '', groupName, records.length]);
    totalRecords += records.length;
  });
  summarySheet.addRow(['', '', '合计', totalRecords]);
  summarySheet.columns.forEach(col => { col.width = 15; });
  
  // 为每个分组创建sheet
  for (const [groupName, records] of Object.entries(groups)) {
    const safeName = groupName.replace(/[\\/:*?"<>|]/g, '_').substring(0, 25);
    const worksheet = workbook.addWorksheet(safeName);
    
    worksheet.addRow(headers);
    worksheet.getRow(1).font = { bold: true };
    worksheet.getRow(1).fill = {
      type: 'pattern', pattern: 'solid',
      fgColor: { argb: 'FFE5E5EA' }
    };
    
    records.forEach((record, idx) => {
      const row = [idx + 1, record['来源文件'] || ''];
      headers.slice(2).forEach(h => {
        row.push(record[h] || '');
      });
      worksheet.addRow(row);
    });
    
    worksheet.columns.forEach((col, i) => {
      let maxLen = (headers[i] || '').length;
      worksheet.eachRow((row, rowNum) => {
        if (rowNum > 1) {
          const cell = row.getCell(i + 1);
          const len = (cell.value || '').toString().length;
          maxLen = Math.max(maxLen, len);
        }
      });
      col.width = Math.min(maxLen + 2, 50);
    });
  }
  
  // 使用Buffer方式写入
  const buffer = await workbook.xlsx.writeBuffer();
  await window.electronAPI.writeFile(filePath, buffer);
  alert(`已导出到: ${filePath}\n共 ${Object.keys(groups).length + 1} 个工作表（含汇总）`);
}

// 模板编辑器
function openTemplateEditor() {
  const template = state.template || {
    id: Date.now().toString(),
    name: '新模板',
    fields: [
      { name: '所属公司', alias: '纳税人名称', enabled: true },
      { name: '所属公司代码', alias: '纳税人识别号', enabled: true },
      { name: '金额', alias: '实缴金额', enabled: true }
    ]
  };
  
  dom.templateName.value = template.name;
  renderFieldList(template.fields);
  dom.templateModal.classList.add('show');
}

function renderFieldList(fields) {
  // 转换为统一的编辑格式
  const editFields = fields.map(f => {
    if (f.columnHeader) {
      // 新配置格式
      return {
        name: f.columnHeader,
        alias: f.fieldName,
        enabled: f.enabled !== false,
        aggregate: f.aggregate === true,
        valueType: f.valueType || 'text'
      };
    } else {
      // 旧模板格式
      return f;
    }
  });
  
  dom.fieldList.innerHTML = editFields.map((f, i) => `
    <div class="field-row" data-index="${i}">
      <div class="col-check">
        <input type="checkbox" class="field-enabled" ${f.enabled ? 'checked' : ''} title="启用">
      </div>
      <div class="col-name">
        <input type="text" class="field-name" value="${f.name || ''}" placeholder="Excel列头">
      </div>
      <div class="col-alias">
        <input type="text" class="field-alias" value="${f.alias || ''}" placeholder="图片字段名">
      </div>
      <div class="col-type">
        <select class="field-type">
          <option value="text" ${f.valueType === 'text' ? 'selected' : ''}>文本</option>
          <option value="number" ${f.valueType === 'number' ? 'selected' : ''}>数字</option>
          <option value="date" ${f.valueType === 'date' ? 'selected' : ''}>日期</option>
        </select>
      </div>
      <div class="col-agg">
        <input type="checkbox" class="field-agg" ${f.aggregate ? 'checked' : ''} title="按此字段分表">
      </div>
      <div class="col-action">
        <button class="btn-remove" data-action="remove" title="删除">×</button>
      </div>
    </div>
  `).join('');
  
  // 事件委托
  dom.fieldList.querySelectorAll('.btn-remove').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const index = parseInt(e.target.closest('.field-row').dataset.index);
      removeField(index);
    });
  });
}

function addFieldRow() {
  const currentFields = getFieldsFromList();
  currentFields.push({ name: '', alias: '', enabled: true, aggregate: false, valueType: 'text' });
  renderFieldList(currentFields);
}

function removeField(index) {
  const currentFields = getFieldsFromList();
  currentFields.splice(index, 1);
  renderFieldList(currentFields);
}

function getFieldsFromList() {
  return Array.from(dom.fieldList.querySelectorAll('.field-row')).map(item => ({
    name: item.querySelector('.field-name').value,
    alias: item.querySelector('.field-alias').value,
    enabled: item.querySelector('.field-enabled').checked,
    aggregate: item.querySelector('.field-agg').checked,
    valueType: item.querySelector('.field-type').value
  }));
}

async function saveTemplate() {
  const editFields = getFieldsFromList();
  
  // 转换为新配置格式
  const fieldConfig = editFields.map(f => ({
    fieldName: f.alias || f.name,
    columnHeader: f.name,
    description: '',
    positionHint: '根据图片内容自动识别',
    valueType: f.valueType || 'text',
    enabled: f.enabled,
    aggregate: f.aggregate
  }));
  
  // 同时保存两种格式
  const template = {
    id: state.template?.id || Date.now().toString(),
    name: dom.templateName.value || '自定义配置',
    fields: editFields  // 旧格式
  };
  
  // 保存到状态
  state.template = template;
  state.fieldConfig = fieldConfig;
  
  // 保存到localStorage（持久化）
  localStorage.setItem('fieldConfig', JSON.stringify(fieldConfig));
  console.log('[保存] 字段配置已保存到localStorage，共', fieldConfig.length, '个字段');
  
  await window.electronAPI.saveTemplate(template);
  
  await loadTemplates();
  dom.templateSelect.value = template.id;
  updateTableHeader();
  closeTemplateModal();
  
  alert(`已保存 ${fieldConfig.length} 个字段配置`);
}

function closeTemplateModal() {
  dom.templateModal.classList.remove('show');
}

// 导出模板文件
async function exportTemplateFile() {
  const currentFields = getFieldsFromList();
  const template = {
    name: dom.templateName.value || '未命名模板',
    fields: currentFields
  };
  
  if (template.fields.length === 0) {
    alert('请先添加字段');
    return;
  }
  
  const filePath = await window.electronAPI.exportTemplate(template);
  if (filePath) {
    alert(`模板已导出到: ${filePath}`);
  }
}

// 导入模板文件
async function importTemplateFile() {
  const data = await window.electronAPI.importTemplate();
  if (!data) return;
  
  try {
    // 先打开模态框
    dom.templateModal.classList.add('show');
    
    // 检测配置格式
    if (Array.isArray(data)) {
      // 检查是否是字段配置格式（有fieldName或columnHeader属性）
      const isFieldConfig = data[0] && (
        data[0].hasOwnProperty('fieldName') ||
        data[0].hasOwnProperty('columnHeader') ||
        data[0].hasOwnProperty('alias')
      );
      
      // 检查是否是数据记录格式（有金额、公司等业务字段的值）
      const isDataRecord = data[0] && !isFieldConfig && (
        (data[0].hasOwnProperty('金额') && typeof data[0]['金额'] === 'string') ||
        (data[0].hasOwnProperty('所属公司') && typeof data[0]['所属公司'] === 'string') ||
        (data[0].hasOwnProperty('税种名称') && typeof data[0]['税种名称'] === 'string')
      );
      
      console.log('[导入] 数组类型，isFieldConfig:', isFieldConfig, 'isDataRecord:', isDataRecord);
      
      if (isDataRecord) {
        // 这是一个数据记录JSON，自动生成字段配置
        const sampleRecord = data[0];
        const fields = Object.keys(sampleRecord).map(key => ({
          name: key,
          alias: '',
          enabled: true,
          aggregate: key === '所属公司', // 默认按公司分表
          valueType: typeof sampleRecord[key] === 'number' ? 'number' : 
                     (sampleRecord[key] && String(sampleRecord[key]).match(/^\d{4}-\d{2}-\d{2}/) ? 'date' : 'text')
        }));
        
        dom.templateName.value = '数据字段配置';
        renderFieldList(fields);
        
        // 将数据直接导入到state中
        importDataFromJson(data);
        alert(`已导入 ${fields.length} 个字段，共 ${data.length} 条数据记录`);
      } else if (isFieldConfig) {
        // 这是字段配置数组
        console.log('[导入] 检测到字段配置数组，共', data.length, '个字段');
        const fields = data.map(item => ({
          name: item.columnHeader || item.name || '',
          alias: item.fieldName || item.alias || '',
          enabled: item.enabled !== false,
          aggregate: item.aggregate === true,
          valueType: item.valueType || 'text'
        }));
        
        dom.templateName.value = '导入模板';
        state.fieldConfig = data; // 保存原始配置
        
        // 同时保存到localStorage
        localStorage.setItem('fieldConfig', JSON.stringify(data));
        console.log('[导入] 字段配置已保存到localStorage');
        
        renderFieldList(fields);
        
        const aggCount = fields.filter(f => f.aggregate).length;
        console.log('[导入] 渲染完成，第一个字段:', fields[0], '分表字段:', aggCount);
        alert(`已导入 ${fields.length} 个字段配置\n其中 ${aggCount} 个设置了分表`);
      } else {
        // 无法识别的格式
        alert('无法识别的JSON数组格式\n\n第一个元素应该是字段配置或数据记录');
      }
    } else if (data.fields) {
      // 旧格式：{ name, fields: [...] } 或新格式：{ name, fields: [...], ... }
      dom.templateName.value = data.name || '';
      
      // 检查字段格式
      const fields = data.fields.map(f => {
        // 兼容新旧格式
        if (f.columnHeader) {
          return {
            name: f.columnHeader,
            alias: f.fieldName || '',
            enabled: f.enabled !== false,
            aggregate: f.aggregate === true,
            valueType: f.valueType || 'text'
          };
        } else {
          return f;
        }
      });
      
      renderFieldList(fields);
      alert(`已导入模板: ${data.name || '未命名模板'}，共 ${fields.length} 个字段`);
    } else if (data.template) {
      // 导出格式：{ version, exportDate, template }
      dom.templateName.value = data.template.name || '';
      renderFieldList(data.template.fields || []);
      alert(`已导入模板: ${data.template.name || '未命名模板'}`);
    } else {
      alert('无法识别的JSON格式，请检查文件内容');
    }
  } catch (err) {
    console.error('导入模板失败:', err);
    alert(`导入失败: ${err.message}`);
  }
}

// 从JSON数据导入到state
function importDataFromJson(records) {
  if (!Array.isArray(records) || records.length === 0) return;
  
  // 创建虚拟图片对象来存储数据
  const mockImage = {
    name: '导入数据',
    status: 'done',
    result: records[0],
    allResults: records,
    data: null,
    path: null
  };
  
  // 清空现有图片，添加导入的数据
  state.images = [mockImage];
  state.currentIndex = 0;
  
  // 更新UI
  renderImageList();
  updateDataTable();
  updateButtons();
  
  // 显示导入的数据
  dom.previewContainer.innerHTML = `
    <div class="empty-state">
      <div class="empty-icon">📊</div>
      <p>已导入 ${records.length} 条数据记录</p>
      <p class="hint">可以在下方表格查看数据，或导出Excel</p>
    </div>
  `;
  dom.currentFileName.textContent = `导入数据 (${records.length}条)`;
}

// 导入Excel模板
async function importExcelTemplate() {
  const result = await window.electronAPI.importExcelTemplate();
  
  if (result && result.fields && result.fields.length > 0) {
    dom.templateName.value = result.name;
    renderFieldList(result.fields);
    alert(`已导入 ${result.fields.length} 个字段`);
  }
}

// 进度弹窗
function showProgressModal() {
  dom.progressModal.classList.add('show');
}

function hideProgressModal() {
  dom.progressModal.classList.remove('show');
}

function updateProgress(index, text) {
  const total = state.images.length;
  const percent = Math.round(((index + 1) / total) * 100);
  
  dom.progressText.textContent = text;
  dom.progressPercent.textContent = `${percent}%`;
  dom.progressFill.style.width = `${percent}%`;
}

function cancelProcessing() {
  state.isProcessing = false;
  hideProgressModal();
}

function toggleSelectAll(e) {
  const checked = e.target.checked;
  document.querySelectorAll('.row-check').forEach(cb => cb.checked = checked);
}

// 工具函数
function clearAll() {
  if (confirm('确定要清空所有图片和数据吗?')) {
    state.images = [];
    state.currentIndex = -1;
    state.results = [];
    
    renderImageList();
    dom.previewContainer.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📄</div>
        <p>选择左侧图片查看预览</p>
      </div>
    `;
    dom.resultContainer.innerHTML = `
      <div class="empty-state">
        <div class="empty-icon">📝</div>
        <p>提取结果将显示在这里</p>
      </div>
    `;
    dom.currentFileName.textContent = '';
    
    updateDataTable();
    updateButtons();
  }
}

function updateButtons() {
  const hasImages = state.images.length > 0;
  const hasTemplate = state.template !== null;
  const hasData = state.images.some(img => img.status === 'done');
  
  dom.extractAll.disabled = !hasImages || !hasTemplate;
  dom.exportExcel.disabled = !hasData;
  dom.clearAll.disabled = !hasImages;
}

// 启动应用
// 添加全局错误处理
window.onerror = function(message, source, lineno, colno, error) {
  console.error('[全局错误]', message, 'at', source, ':', lineno, ':', colno);
  console.error('[错误堆栈]', error?.stack);
  alert('应用错误: ' + message + '\n\n请查看控制台获取详细信息');
  return false;
};

window.addEventListener('unhandledrejection', function(event) {
  console.error('[未处理的Promise错误]', event.reason);
  alert('异步错误: ' + event.reason);
});

// 可拖动分割条逻辑
(function initResizer() {
  const resizer = document.getElementById('resizer');
  const contentArea = document.querySelector('.content-area');
  const dataSection = document.querySelector('.data-section');
  const mainContent = document.querySelector('.main-content');
  
  let isResizing = false;
  let startY = 0;
  let startHeight = 0;
  
  resizer.addEventListener('mousedown', (e) => {
    isResizing = true;
    startY = e.clientY;
    startHeight = contentArea.offsetHeight;
    
    document.body.style.cursor = 'ns-resize';
    document.body.style.userSelect = 'none';
    
    e.preventDefault();
  });
  
  document.addEventListener('mousemove', (e) => {
    if (!isResizing) return;
    
    const deltaY = e.clientY - startY;
    const newHeight = startHeight + deltaY;
    
    // 限制最小和最大高度
    const minHeight = 200;
    const maxHeight = mainContent.offsetHeight - 300;
    
    if (newHeight >= minHeight && newHeight <= maxHeight) {
      contentArea.style.height = newHeight + 'px';
      contentArea.style.flex = 'none';
    }
  });
  
  document.addEventListener('mouseup', () => {
    if (isResizing) {
      isResizing = false;
      document.body.style.cursor = '';
      document.body.style.userSelect = '';
    }
  });
  
  // 双击重置
  resizer.addEventListener('dblclick', () => {
    contentArea.style.height = '';
    contentArea.style.flex = '1';
  });
})();

init();