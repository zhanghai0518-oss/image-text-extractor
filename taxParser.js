/**
 * 增强版税务文本解析器
 */

class TaxParser {
  constructor() {
    this.debug = true;
  }

  log(...args) {
    if (this.debug) console.log('[TaxParser]', ...args);
  }

  /**
   * 主解析入口
   */
  parse(text, filename = '') {
    this.log('开始解析:', filename);
    this.log('文本长度:', text.length);
    this.log('文本内容（前500字符）:', text.substring(0, 500));
    this.log('文本内容（全文）:\n', text);
    
    // 智能判断文档类型
    if (this.isPersonalIncomeTax(text)) {
      return this.parsePersonalIncomeTax(text, filename);
    } else if (this.isEnterpriseTax(text)) {
      return this.parseEnterpriseTax(text, filename);
    }
    
    // 默认使用个人所得税解析
    return this.parsePersonalIncomeTax(text, filename);
  }

  /**
   * 判断是否为个人所得税完税证明
   */
  isPersonalIncomeTax(text) {
    return /个人所[得保税]|综合所得|预扣预缴/i.test(text);
  }

  /**
   * 判断是否为企业税收完税证明
   */
  isEnterpriseTax(text) {
    return /税收完税证明|企业所得税|增值税|印花税/i.test(text) && !this.isPersonalIncomeTax(text);
  }

  /**
   * 解析个人所得税完税证明（支持多税目）
   */
  parsePersonalIncomeTax(text, filename) {
    const records = [];
    
    // 基本信息
    const taxpayerId = this.extractValue(text, ['纳税人识别号', '纳税人识别码']);
    
    // 修复问题1：纳税人名称提取右边单元格内容
    const taxpayerName = this.extractCompanyName(text);
    
    const taxOffice = this.extractTaxOffice(text);
    const voucherNo = this.extractVoucherNo(text);
    const fillDate = this.extractDate(text, '填发日期');
    this.log('提取填发日期:', fillDate);
    
    // 修复问题2：提取品目名称（按OCR实际识别的内容）
    const taxItems = this.extractTaxItems(text);
    this.log('提取的品目名称:', taxItems);
    
    // 修复问题3：提取税目金额（正确处理逗号和空格）
    const taxAmounts = this.extractTaxAmounts(text);
    this.log('提取的税目金额:', taxAmounts);
    
    // 修复问题2：提取税款所属期（支持多个）
    const periods = this.extractAllPeriods(text);
    this.log('提取的税款所属期:', periods);
    
    // 提取公共信息
    const taxType = '个人所得税';
    
    // 为每个税目金额创建记录
    const recordCount = taxAmounts.length || taxItems.length || 1;
    
    // 提取备注
    const remarks = this.extractRemarks(text);
    
    for (let i = 0; i < recordCount; i++) {
      const period = periods[i] || periods[0] || { start: '', end: '' };
      const itemName = taxItems[i] || taxItems[0] || '';
      
      records.push({
        所属公司: taxpayerName,
        所属公司代码: taxpayerId,
        征收机关: taxOffice,
        税种名称: taxType,
        税目名称: itemName,
        金额: taxAmounts[i] || '',
        税款所属期起: period.start,
        税款所属期止: period.end,
        税款类型名称: '正税',
        缴款日期: fillDate,
        申报日期: fillDate,
        税票号码: voucherNo,
        主管税务所: taxOffice,
        实缴年份: fillDate ? fillDate.substring(0, 4) : '',
        实缴月份: fillDate ? fillDate.substring(5, 7) : '',
        备注: remarks || ''  // 只保留原始备注内容
      });
    }
    
    this.log('解析结果:', records.length, '条记录');
    return records;
  }

  /**
   * 新增：提取品目名称（按OCR实际识别的内容，不去重）
   */
  extractTaxItems(text) {
    const items = [];
    
    // 已知的个人所得税品目
    const knownItems = [
      '工资薪金所得', '劳务报酬所得', '稿酬所得', '特许权使用费所得',
      '经营所得', '利息股息红利所得', '财产租赁所得', '财产转让所得',
      '偶然所得', '其他所得'
    ];
    
    // 按行分割，查找品目名称（保留所有匹配，不去重）
    const lines = text.split('\n');
    
    for (const line of lines) {
      // 跳过表头行
      if (line.includes('品目名称') || line.includes('品日名称')) continue;
      
      // 匹配已知的品目名称
      for (const item of knownItems) {
        if (line.includes(item)) {
          items.push(item);
          this.log('找到品目:', item);
          break; // 每行只匹配一个品目
        }
      }
    }
    
    return items;
  }

  /**
   * 解析企业税收完税证明
   */
  parseEnterpriseTax(text, filename) {
    const records = [];
    
    // 基本信息
    const taxpayerId = this.extractValue(text, ['纳税人识别号']);
    const taxpayerName = this.extractCompanyName(text);
    const taxOffice = this.extractTaxOffice(text);
    const fillDate = this.extractDate(text, '填发日期');
    
    // 提取表格行
    const tablePattern = /(\d{15,})\s+([\u4e00-\u9fa5]+)\s+([\u4e00-\u9fa5]+)\s+(\d{4}-\d{2}-\d{2})\s+至\s+(\d{4}-\d{2}-\d{2})\s+(\d{4}-\d{2}-\d{2})\s+([\d,.]+)/g;
    
    let match;
    while ((match = tablePattern.exec(text)) !== null) {
      records.push({
        所属公司: taxpayerName,
        所属公司代码: taxpayerId,
        征收机关: taxOffice,
        税种名称: match[2],
        税目名称: match[3],
        金额: match[7].replace(/,/g, ''),
        税款所属期起: match[4],
        税款所属期止: match[5],
        税款类型名称: '正税',
        缴款日期: match[6],
        税票号码: match[1],
        主管税务所: taxOffice,
        实缴年份: fillDate ? fillDate.substring(0, 4) : '',
        实缴月份: fillDate ? fillDate.substring(5, 7) : '',
        备注: `来源: ${filename}`
      });
    }
    
    return records;
  }

  /**
   * 修复问题1：提取纳税人名称（右边单元格内容）
   */
  extractCompanyName(text) {
    const lines = text.split('\n').map(l => l.trim()).filter(l => l);
    
    // 方法1: 直接匹配"纳税人名称"后面的内容（同一行）
    const sameLineMatch = text.match(/纳税人名称\s*[\n\r]+\s*([^\n\r]+)/);
    if (sameLineMatch && sameLineMatch[1].includes('公司')) {
      const name = sameLineMatch[1].trim();
      if (!name.includes('税务局') && !name.includes('税务所')) {
        this.log('提取公司名称（方法1-下一行）:', name);
        return name;
      }
    }
    
    // 方法2: 找到"纳税人名称"行，向下查找包含"公司"的行
    let nameIndex = -1;
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].includes('纳税人名称')) {
        nameIndex = i;
        this.log('找到纳税人名称行，索引:', i);
        break;
      }
    }
    
    if (nameIndex === -1) {
      this.log('未找到纳税人名称行');
      return '';
    }
    
    // 向下查找，找到第一个包含"公司"的行（排除税务机关）
    for (let i = nameIndex + 1; i < lines.length && i < nameIndex + 10; i++) {
      const line = lines[i];
      
      // 跳过其他字段名
      if (line.includes('税款所属') || line.includes('No.') || line.includes('No，') || 
          line.includes('税务机关') || line.includes('填发日期') ||
          line.includes('税种') || line.includes('品目') || line.includes('实缴')) {
        continue;
      }
      
      // 找到包含"公司"的行
      if (line.includes('公司')) {
        // 排除税务机关
        if (line.includes('税务局') || line.includes('税务所')) {
          continue;
        }
        
        // 修正OCR错误：或汉 -> 武汉
        const name = line.replace(/^或汉/, '武汉').trim();
        this.log('找到公司名称:', name);
        return name;
      }
    }
    
    // 方法3: 直接从文本中提取包含"公司"的名称
    const companyPattern = /([^\n]*?[^\s]{2,}公司[^\s]*)/g;
    const companies = text.match(companyPattern) || [];
    for (const c of companies) {
      const name = c.trim();
      // 排除税务机关
      if (name.includes('税务局') || name.includes('税务所') || name.length < 4) {
        continue;
      }
      this.log('提取公司名称（方法3-全文搜索）:', name);
      return name;
    }
    
    return '';
  }

  /**
   * 提取备注（图片右下角的文本）
   */
  extractRemarks(text) {
    const lines = text.split('\n').map(l => l.trim()).filter(l => l);
    
    this.log('开始提取备注，总行数:', lines.length);
    
    // 方法1：找"备注"关键字（不匹配"原凭证号"）
    for (let i = 0; i < lines.length; i++) {
      const line = lines[i];
      
      // 只匹配"备注"关键字，不匹配"凭证"
      if (line.includes('备注') || line.includes('说明')) {
        this.log('找到备注关键字行:', i, line);
        
        // 获取关键字后面的内容
        let afterKey = line.replace(/^.*备注[：:]\s*/, '').trim() ||
                       line.replace(/^.*说明[：:]\s*/, '').trim();
        
        if (afterKey && afterKey.length > 2) {
          this.log('提取备注（同一行）:', afterKey);
          
          // 截取到税务局/税务所（包含税号等完整信息）
          // 格式: 一般申报 正税 主管税务所...税务局左岭税务所
          const taxOfficeMatch = afterKey.match(/^(.*?)税务[局所][^\s,，]*/);
          if (taxOfficeMatch) {
            afterKey = taxOfficeMatch[0].trim();
          }
          
          // 如果包含"土地编号"等信息，也要保留
          if (afterKey.includes('土地编号')) {
            // 截取到土地编号结束
            const landMatch = afterKey.match(/^(.*?土地编号\s*:\s*\S+)/);
            if (landMatch) {
              afterKey = landMatch[1].trim();
            }
          }
          
          return afterKey;
        }
        
        // 如果当前行只有关键字，获取下一行
        if (i + 1 < lines.length) {
          const nextLine = lines[i + 1].trim();
          if (nextLine && !nextLine.includes('：') && !nextLine.includes(':') && 
              nextLine.length > 2 && /[\u4e00-\u9fa5]/.test(nextLine)) {
            this.log('提取备注（下一行）:', nextLine);
            // 截取到税务局
            const taxOfficeEnd = nextLine.match(/.*?税务[局所]/);
            return taxOfficeEnd ? taxOfficeEnd[0] : nextLine;
          }
        }
      }
    }
    
    // 方法2：提取图片右下角的文本（通常是最后几行）
    // 跳过金额、日期、税票号码等格式的内容
    if (lines.length > 5) {
      for (let i = lines.length - 1; i >= lines.length - 5; i--) {
        const line = lines[i];
        
        this.log('检查右下角行:', i, line);
        
        // 跳过金额、日期、税票号码等格式
        if (/^\d+[\d,\s]*\.\s*\d{2}$/.test(line)) {
          this.log('跳过金额:', line);
          continue;
        }
        if (/^\d{4}[.\s-]\d{2}/.test(line)) {
          this.log('跳过日期:', line);
          continue;
        }
        if (/^No\./i.test(line)) {
          this.log('跳过税票号码:', line);
          continue;
        }
        if (/^\d{15,}$/.test(line)) {
          this.log('跳过纯数字:', line);
          continue;
        }
        if (line.length < 3) {
          this.log('跳过短行:', line);
          continue;
        }
        
        // 检查是否是有意义的文本（包含中文）
        if (/[\u4e00-\u9fa5]/.test(line)) {
          this.log('提取备注（右下角中文）:', line);
          return line;
        }
      }
    }
    
    this.log('未找到备注');
    return '';
  }

  /**
   * 提取税务机关（修正OCR错误+合并多行）
   */
  extractTaxOffice(text) {
    const lines = text.split('\n').map(l => l.trim());
    
    // 找"税务机关"行
    let officeIndex = -1;
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].includes('税务机关')) {
        officeIndex = i;
        this.log('找到税务机关行，索引:', i, '内容:', lines[i]);
        break;
      }
    }
    
    if (officeIndex === -1) {
      return '';
    }
    
    // 提取税务机关内容
    let office = '';
    
    // 格式1：税务机关：国家税务总局xxx（内容在同一行）
    const line = lines[officeIndex];
    const afterKey = line.replace(/^税务机关[：:\s]*/, '').trim();
    
    if (afterKey && afterKey.length > 2) {
      office = afterKey;
      this.log('税务机关在同一行:', office);
    } else if (officeIndex + 1 < lines.length) {
      // 格式2：税务机关在下一行
      office = lines[officeIndex + 1];
      this.log('税务机关在下一行:', office);
    }
    
    // 检查后续行是否有补充内容（如"一税务所（办税服务厅）"）
    const startIdx = office.length > 0 ? officeIndex + 2 : officeIndex + 1;
    
    for (let i = startIdx; i < Math.min(startIdx + 3, lines.length); i++) {
      const nextLine = lines[i];
      
      // 跳过字段名和空行
      if (!nextLine || nextLine.includes('纳税人') || nextLine.includes('税款') || 
          nextLine.includes('No') || nextLine.includes('填发') || nextLine.includes('税种') ||
          nextLine.includes('品目') || nextLine.includes('金额') || nextLine.includes('日期')) {
        break;
      }
      
      // 如果是税务机关的后续内容（以"一"、"第"、或包含"税务所"/"办税"）
      if (nextLine.startsWith('一') || nextLine.startsWith('第') || 
          nextLine.includes('税务所') || nextLine.includes('办税')) {
        office += nextLine;
        this.log('添加后续行:', nextLine);
      } else {
        break;
      }
    }
    
    // 修正OCR错误：热→税
    office = office.replace(/热务局/g, '税务局')
                   .replace(/热务所/g, '税务所')
                   .replace(/监税/g, '完税');
    
    this.log('提取税务机关:', office);
    return office;
  }

  /**
   * 修复问题2：提取所有税款所属期（支持多条，处理OCR连在一起的日期）
   */
  extractAllPeriods(text) {
    const periods = [];
    
    this.log('开始提取税款所属期');
    
    // 特殊处理：OCR可能把两个日期连在一起，中间可能有空格
    // 格式：2026. 03. 012026. 03. 31 或 2026.03.012026.03.31
    // 需要处理：2026. 03. 01 和 2026. 03. 31
    const joinedPattern = /(\d{4})[.\s]+(\d{1,2})[.\s]+(\d{1,2})(\d{4})[.\s]+(\d{1,2})[.\s]+(\d{1,2})/g;
    
    let match;
    while ((match = joinedPattern.exec(text)) !== null) {
      // 第一个日期
      const start = `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
      // 第二个日期（紧接在第一个日期后面）
      const end = `${match[4]}-${match[5].padStart(2, '0')}-${match[6].padStart(2, '0')}`;
      
      this.log('匹配到连在一起的日期: 起=' + start + ' 止=' + end);
      periods.push({ start, end });
    }
    
    // 如果没有匹配到连在一起的日期，尝试标准格式
    if (periods.length === 0) {
      // 匹配格式：2026.03.01-2026.03.31 或 2026. 03. 01-2026. 03. 31
      const pattern = /(\d{4})[.\s-]*(\d{1,2})[.\s-]*(\d{1,2})\s*[-至到~]\s*(\d{4})[.\s-]*(\d{1,2})[.\s-]*(\d{1,2})/g;
      
      while ((match = pattern.exec(text)) !== null) {
        const start = `${match[1]}-${match[2].padStart(2, '0')}-${match[3].padStart(2, '0')}`;
        const end = `${match[4]}-${match[5].padStart(2, '0')}-${match[6].padStart(2, '0')}`;
        
        this.log('匹配到日期范围: 起=' + start + ' 止=' + end);
        periods.push({ start, end });
      }
    }
    
    return periods;
  }

  /**
   * 修复问题3：提取税目金额（正确处理OCR的金额格式）
   * OCR格式可能���：5, 566.54 或 5,566.54 或 5566.54
   */
  extractTaxAmounts(text) {
    const amounts = [];
    
    this.log('开始提取税目金额');
    
    // 找合计金额（带*号或¥/￥符号的数字）
    const totalMatch = text.match(/[¥￥*]\s*([\d,\s.]+\.\d{2})/);
    
    let totalAmount = 0;
    if (totalMatch && totalMatch[1]) {
      // 处理点号分隔的金额（如 46.540.23）
      let totalStr = totalMatch[1].replace(/[\s,]/g, '');
      const dots = (totalStr.match(/\./g) || []).length;
      if (dots > 1) {
        totalStr = totalStr.replace(/\./g, '').replace(/(\d{2})$/, '.$1');
      }
      totalAmount = parseFloat(totalStr);
      this.log('找到合计:', totalAmount);
    }
    
    // 策略：按行分割，逐行匹配金额（避免跨行拼接）
    const lines = text.split('\n');
    const allAmounts = [];
    
    for (const line of lines) {
      // 跳过带*号或合计的行
      if (line.includes('*') || line.includes('合计')) {
        this.log('跳过合计行:', line);
        continue;
      }
      
      // 匹配点号分隔格式（如 11.879.21 → 11879.21）
      const dotPattern = /(\d{1,3})\.(\d{3})\.(\d{2})/g;
      let dotMatch;
      while ((dotMatch = dotPattern.exec(line)) !== null) {
        const combined = dotMatch[1] + dotMatch[2] + '.' + dotMatch[3];
        const val = parseFloat(combined);
        if (val >= 100 && val <= 100000000) {
          this.log('点号分隔格式匹配:', dotMatch[0], '→', combined);
          const exists = allAmounts.some(a => a.raw === combined);
          if (!exists) {
            allAmounts.push({ raw: combined, value: val });
          }
        }
      }
      
      // 匹配金额格式
      // 特殊处理：OCR可能把"5,566.54"识别成"5, 566.54"（逗号后加空格）
      const specialPattern = /(\d{1,3})\s*,\s*(\d{1,3}(?:,\d{3})*\.\d{2})/g;
      let specialMatch;
      const specialAmounts = []; // 记录特殊格式匹配的金额
      
      while ((specialMatch = specialPattern.exec(line)) !== null) {
        const combined = specialMatch[1] + specialMatch[2].replace(/,/g, '');
        const val = parseFloat(combined);
        if (val >= 100 && val <= 100000000) {
          this.log('特殊格式匹配:', specialMatch[0], '→ 组合:', combined);
          const exists = allAmounts.some(a => a.raw === combined);
          if (!exists) {
            allAmounts.push({ raw: combined, value: val });
            // 记录这个金额，用于后面排除标准格式的重复匹配
            specialAmounts.push(combined);
          }
        }
      }
      
      // 匹配标准金额格式（排除已被特殊格式匹配的）
      const pattern = /(\d{1,3}(?:,\d{3})*\.\d{2})/g;
      let match;
      
      while ((match = pattern.exec(line)) !== null) {
        const raw = match[1];
        const cleaned = raw.replace(/,/g, '');
        
        // 检查是否已被特殊格式匹配过
        const isSpecialMatched = specialAmounts.some(s => {
          // 检查cleaned是否是特殊格式金额的一部分
          return s.includes(cleaned) || cleaned.length < 5;
        });
        if (isSpecialMatched) continue;
        
        // 跳过日期格式
        if (/^\d{4}[-\/.]\d{1,2}[-\/.]\d{1,2}$/.test(cleaned)) continue;
        if (/^0\d{1}\.\d{2}$/.test(cleaned)) continue;
        
        const val = parseFloat(cleaned);
        if (val < 100 || val > 100000000) continue;
        
        // 去重
        const exists = allAmounts.some(a => a.raw === cleaned);
        if (!exists) {
          allAmounts.push({ raw: cleaned, value: val });
          this.log('标准格式匹配:', raw, '→', cleaned);
        }
      }
    }
    
    // 去重
    const uniqueAmounts = [];
    const seen = new Set();
    for (const a of allAmounts) {
      if (!seen.has(a.raw)) {
        seen.add(a.raw);
        uniqueAmounts.push(a);
      }
    }
    
    this.log('去重后的候选金额:', uniqueAmounts.map(a => a.raw));
    
    // 找出相加等于合计的金额组合
    if (totalAmount > 0 && uniqueAmounts.length >= 1) {
      // 过滤掉等于合计的候选
      const candidates = uniqueAmounts.filter(a => Math.abs(a.value - totalAmount) > 0.01);
      
      this.log('过滤后的候选:', candidates.map(a => a.raw + '=' + a.value));
      this.log('合计:', totalAmount);
      
      // 如果候选为空，只有一条税目（金额=合计）
      if (candidates.length === 0) {
        // 单条记录，金额就是合计
        amounts.push(totalAmount.toFixed(2));
        this.log('只有一条税目金额:', amounts);
        return amounts;
      }
      
      // 检查所有候选相加是否等于合计
      const sum = candidates.reduce((s, a) => s + a.value, 0);
      this.log('候选相加:', sum, '合计:', totalAmount, '差值:', Math.abs(sum - totalAmount));
      
      if (Math.abs(sum - totalAmount) < 0.01) {
        candidates.forEach(a => amounts.push(a.raw));
        this.log('找到税目金额（相加等于合计）:', amounts);
        return amounts;
      }
      
      // 尝试找部分组合
      for (let i = 0; i < candidates.length; i++) {
        let partialSum = 0;
        const selected = [];
        
        this.log('尝试组合，起始索引:', i);
        
        for (let j = i; j < candidates.length; j++) {
          this.log('  检查:', candidates[j].raw, '=', candidates[j].value, '当前累计:', partialSum);
          if (partialSum + candidates[j].value <= totalAmount + 0.01) {
            partialSum += candidates[j].value;
            selected.push(candidates[j].raw);
            this.log('    加入，累计:', partialSum);
          }
        }
        
        this.log('  组合结果:', selected, '累计:', partialSum, '差值:', Math.abs(partialSum - totalAmount));
        
        if (Math.abs(partialSum - totalAmount) < 0.01) {
          this.log('找到税目金额（部分组合）:', selected);
          return selected;
        }
      }
      
      // 如果所有组合都不匹配，直接返回所有候选
      this.log('所有组合都不匹配合计，直接返回候选');
      candidates.forEach(a => amounts.push(a.raw));
      return amounts;
    }
    
    // 如果没有合计，取所有候选（排除等于合计的）
    for (const item of uniqueAmounts) {
      if (totalAmount === 0 || Math.abs(item.value - totalAmount) > 0.01) {
        amounts.push(item.raw);
      }
    }
    
    this.log('最终税目金额:', amounts);
    return amounts;
  }

  /**
   * 提取字段值
   */
  extractValue(text, keywords) {
    for (const keyword of keywords) {
      const regex = new RegExp(keyword + '[：:]*\\s*([^\\n]+)', 'i');
      const match = text.match(regex);
      if (match) {
        return match[1].trim();
      }
    }
    return '';
  }

  /**
   * 提取税票号码（支持OCR全角逗号和多种格式）
   */
  extractVoucherNo(text) {
    // 支持多种格式：No. / No， / No． / No•
    const match = text.match(/No[.,，.•·]\s*(\d{15,})/i) ||
                  text.match(/税票号码[：:]*\s*(\d{15,})/i) ||
                  text.match(/(\d{15,})/);  // 兜底：提取15位以上数字
    
    return match ? match[1] : '';
  }

  /**
   * 提取日期
   */
  extractDate(text, keyword) {
    // 1. 先尝试带关键字的格式
    const patterns = [
      // yyyy年mm月dd日 格式
      new RegExp(keyword + '[：:]*\\s*(\\d{4})\\s*年\\s*(\\d{1,2})\\s*月\\s*(\\d{1,2})\\s*日?', 'i'),
      // yyyy-mm-dd 或 yyyy/mm/dd 格式
      new RegExp(keyword + '[：:]*\\s*(\\d{4})[-/](\\d{1,2})[-/](\\d{1,2})', 'i'),
      // yyyy.mm.dd 格式
      new RegExp(keyword + '[：:]*\\s*(\\d{4})\\.(\\d{1,2})\\.(\\d{1,2})', 'i')
    ];
    
    for (const regex of patterns) {
      const match = text.match(regex);
      if (match) {
        const year = match[1];
        const month = match[2].padStart(2, '0');
        const day = match[3].padStart(2, '0');
        this.log('提取日期:', keyword, '→', `${year}-${month}-${day}`);
        return `${year}-${month}-${day}`;
      }
    }
    
    // 2. 如果关键字提取失败，尝试提取任意位置的日期
    const anyDatePatterns = [
      /(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/g,
      /(\d{4})[-/](\d{1,2})[-/](\d{1,2})/g,
      /(\d{4})\.(\d{1,2})\.(\d{1,2})/g
    ];
    
    // 收集所有日期
    const dates = [];
    for (const regex of anyDatePatterns) {
      let match;
      while ((match = regex.exec(text)) !== null) {
        const year = parseInt(match[1]);
        const month = parseInt(match[2]);
        const day = parseInt(match[3]);
        // 验证日期有效性
        if (year >= 2020 && year <= 2030 && month >= 1 && month <= 12 && day >= 1 && day <= 31) {
          dates.push(`${match[1]}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`);
        }
      }
    }
    
    // 去重后返回第一个有效日期
    if (dates.length > 0) {
      const uniqueDates = [...new Set(dates)];
      this.log('提取日期(无关键字):', keyword, '→', uniqueDates[0], '候选:', uniqueDates);
      return uniqueDates[0];
    }
    
    return '';
  }

  /**
   * 标准化税种名称
   */
  normalizeTaxType(type) {
    if (!type) return '';
    // 修正OCR错误
    if (/个人所[得保]税|个税/i.test(type)) return '个人所得税';
    if (/企业所[得保]税/i.test(type)) return '企业所得税';
    return type.trim();
  }
}

// 导出到全局
window.TaxParser = new TaxParser();
