/**
 * 智能字段匹配器
 * 针对税务完税证明文本优化
 * 处理PDF提取的单行文本格式
 */

class FieldMatcher {
  constructor() {
    this.debug = true;
  }

  log(...args) {
    if (this.debug) console.log('[FieldMatcher]', ...args);
  }

  /**
   * 根据字段配置提取数据
   */
  match(text, fieldConfig) {
    this.log('========== 开始匹配 ==========');
    this.log('文本长度:', text.length);
    this.log('');
    this.log('========== 文本内容 ==========');
    this.log(text);
    this.log('========== 文本结束 ==========');
    this.log('');

    // 如果没有字段配置，使用默认字段列表
    const defaultFields = [
      { columnHeader: '所属公司', fieldName: 'companyName', enabled: true },
      { columnHeader: '所属公司代码', fieldName: 'companyCode', enabled: true },
      { columnHeader: '征收机关', fieldName: 'taxOffice', enabled: true },
      { columnHeader: '税票号码', fieldName: 'voucherNo', enabled: true },
      { columnHeader: '缴款日期', fieldName: 'date', enabled: true },
      { columnHeader: '税款所属期起', fieldName: 'periodStart', enabled: true },
      { columnHeader: '税款所属期止', fieldName: 'periodEnd', enabled: true },
      { columnHeader: '税种名称', fieldName: 'taxType', enabled: true },
      { columnHeader: '税目名称', fieldName: 'taxItem', enabled: true },
      { columnHeader: '金额', fieldName: 'amount', enabled: true },
      { columnHeader: '备注', fieldName: 'remarks', enabled: true }
    ];
    
    const enabledFields = (fieldConfig || defaultFields).filter(f => f.enabled !== false);
    
    // 判断是个人所得税还是企业税
    const isPersonalIncomeTax = this.isPersonalIncomeTax(text);
    this.log('文档类型:', isPersonalIncomeTax ? '个人所得税完税证明' : '企业税收完税证明');
    
    // 1. 提取头部信息
    const headerInfo = this.extractHeaderInfo(text);
    this.log('头部信息:', headerInfo);
    
    // 2. 根据类型提取记录
    let records;
    if (isPersonalIncomeTax) {
      records = this.extractPersonalIncomeTaxRecords(text);
    } else {
      records = this.extractEnterpriseTaxRecords(text);
    }
    this.log('提取到', records.length, '条记录');
    
    // 提取合计金额用于验证
    const totalAmount = this.extractTotalAmount(text);
    this.log('提取合计金额:', totalAmount);
    
    // 3. 组装最终结果
    const finalRecords = records.map((record, idx) => {
      const result = {};
      
      enabledFields.forEach(field => {
        const col = field.columnHeader;
        const fieldName = field.fieldName || field.alias || col;
        
        let value = '';
        
        if (col === '所属公司' || fieldName === 'companyName') {
          value = headerInfo.companyName;
        } else if (col === '所属公司代码' || fieldName === 'companyCode') {
          value = headerInfo.companyCode;
        } else if (col === '征收机关' || col === '主管税务所') {
          value = headerInfo.taxOffice;
        } else if (col === '税票号码') {
          value = headerInfo.voucherNo;
        } else if (col === '缴款日期' || col === '申报日期') {
          value = record.date || headerInfo.fillDate;
        } else if (col === '税款所属期起') {
          value = record.periodStart || headerInfo.periodStart;
        } else if (col === '税款所属期止') {
          value = record.periodEnd || headerInfo.periodEnd;
        } else if (col === '税种名称') {
          value = record.taxType || '';
        } else if (col === '税目名称') {
          value = record.taxItem || '';
        } else if (col === '金额') {
          value = record.amount || '';
        } else if (col === '税款类型名称') {
          value = '正税';
        } else if (col === '实缴年份') {
          const date = record.date || headerInfo.fillDate;
          value = date ? date.substring(0, 4) : '';
        } else if (col === '实缴月份') {
          const date = record.date || headerInfo.fillDate;
          value = date ? date.substring(5, 7) : '';
        } else if (col === '备注') {
          value = headerInfo.remarks;
        }
        
        result[col] = value;
      });
      
      return result;
    });
    
    // 返回结果和验证信息
    return {
      records: finalRecords,
      totalAmount: totalAmount,
      validation: this.validateAmounts(finalRecords, totalAmount)
    };
  }

  /**
   * 判断是否为个人所得税完税证明
   */
  isPersonalIncomeTax(text) {
    return /个人所[得保税]|综合所得|预扣预缴|工资薪金所得|偶然所得/i.test(text);
  }

  /**
   * 提取个人所得税记录
   */
  extractPersonalIncomeTaxRecords(text) {
    const records = [];
    
    // 已知的个人所得税品目
    const knownItems = [
      '工资薪金所得', '劳务报酬所得', '稿酬所得', '特许权使用费所得',
      '经营所得', '利息股息红利所得', '财产租赁所得', '财产转让所得',
      '偶然所得', '其他所得'
    ];
    
    // 提取品目名称（按出现顺序）
    const items = [];
    for (const item of knownItems) {
      const regex = new RegExp(item, 'g');
      let match;
      while ((match = regex.exec(text)) !== null) {
        items.push({ name: item, position: match.index });
      }
    }
    
    // 按位置排序并去重
    items.sort((a, b) => a.position - b.position);
    this.log('个人所得税品目:', items.map(i => i.name));
    
    // 提取金额
    const amounts = [];
    const amountRegex = /(\d{1,3}(?:,\d{3})*\.\d{2})/g;
    let match;
    while ((match = amountRegex.exec(text)) !== null) {
      const amount = match[1].replace(/,/g, '');
      amounts.push({ value: amount, position: match.index });
    }
    
    // 过滤掉最大的金额（合计）
    if (amounts.length > 0) {
      const maxAmount = Math.max(...amounts.map(a => parseFloat(a.value)));
      const filteredAmounts = amounts.filter(a => parseFloat(a.value) < maxAmount);
      this.log('个人所得税金额:', filteredAmounts.map(a => a.value));
      
      // 匹配品目和金额
      for (let i = 0; i < items.length && i < filteredAmounts.length; i++) {
        records.push({
          taxType: '个人所得税',
          taxItem: items[i].name,
          amount: filteredAmounts[i].value,
          date: '',
          periodStart: '',
          periodEnd: ''
        });
      }
    }
    
    // 提取税款所属期
    const periodMatch = text.match(/(\d{4})[\.-](\d{2})[\.-](\d{2})\s*[-—]\s*(\d{4})[\.-](\d{2})[\.-](\d{2})/);
    if (periodMatch) {
      const periodStart = `${periodMatch[1]}-${periodMatch[2]}-${periodMatch[3]}`;
      const periodEnd = `${periodMatch[4]}-${periodMatch[5]}-${periodMatch[6]}`;
      records.forEach(r => {
        r.periodStart = periodStart;
        r.periodEnd = periodEnd;
      });
    }
    
    return records;
  }

  /**
   * 提取企业税记录（动态识别税种，不依赖预设列表）
   */
  extractEnterpriseTaxRecords(text) {
    const allRecords = [];
    const matchedPositions = [];
    
    this.log('=== 动态识别税种模式 ===');
    
    // 方法1: 基于税票号码定位（18位数字）
    // 格式: 税票号码 税种 品目 日期范围 入库日期 金额
    const voucherPattern = /(\d{18})\s+(\S+)\s+(\S+)\s+(\d{4}-\d{2}-\d{2})\s+至\s+(\d{4}-\d{2}-\d{2})\s+(\d{4}-\d{2}-\d{2})\s+([\d,]+\.\d{2})/g;
    
    let match;
    while ((match = voucherPattern.exec(text)) !== null) {
      const [fullMatch, voucherNo, taxType, taxItem, periodStart, periodEnd, date, amount] = match;
      
      // 过滤掉明显不是税种的内容（如"原凭证号"、"税种"等表头）
      if (taxType === '原凭证号' || taxType === '税种' || taxType === '品目名称') {
        continue;
      }
      
      this.log('找到记录 (税票号码定位):', taxType, amount);
      allRecords.push({
        taxType,
        taxItem,
        periodStart,
        periodEnd,
        date,
        amount: amount.replace(/,/g, ''),
        position: match.index
      });
    }
    
    // 如果找到了记录，直接返回
    if (allRecords.length > 0) {
      this.log('动态识别成功，共', allRecords.length, '条记录');
      return allRecords;
    }
    
    // 方法2: 基于"税种"表头定位
    const taxTypeHeader = text.indexOf('税   种') !== -1 ? text.indexOf('税   种') : text.indexOf('税种');
    if (taxTypeHeader !== -1) {
      this.log('找到税种表头位置:', taxTypeHeader);
      
      // 从表头后开始查找
      const afterHeader = text.substring(taxTypeHeader + 10);
      
      // 匹配格式: 税种 品目 日期范围 入库日期 金额
      // 税种通常是2-8个中文字符
      const recordPattern = /([^\s\d]{2,8})\s+(\S+)\s+(\d{4}-\d{2}-\d{2})\s+至\s+(\d{4}-\d{2}-\d{2})\s+(\d{4}-\d{2}-\d{2})\s+([\d,]+\.\d{2})/g;
      
      while ((match = recordPattern.exec(afterHeader)) !== null) {
        const [fullMatch, taxType, taxItem, periodStart, periodEnd, date, amount] = match;
        
        // 过滤掉表头和无效内容
        if (taxType === '品目名称' || taxType === '税款所属时期' || taxType.includes('金额')) {
          continue;
        }
        
        this.log('找到记录 (表头定位):', taxType, amount);
        allRecords.push({
          taxType,
          taxItem,
          periodStart,
          periodEnd,
          date,
          amount: amount.replace(/,/g, ''),
          position: taxTypeHeader + match.index
        });
      }
    }
    
    // 如果找到了记录，直接返回
    if (allRecords.length > 0) {
      this.log('动态识别成功，共', allRecords.length, '条记录');
      return allRecords;
    }
    
    // 方法3: 回退到已知税种列表匹配（兼容旧格式）
    this.log('动态识别失败，回退到税种列表匹配');
    return this.extractByKnownTaxTypes(text);
  }
  
  /**
   * 回退方法：基于已知税种列表匹配
   */
  extractByKnownTaxTypes(text) {
    const allRecords = [];
    const matchedPositions = [];
    
    // 已知的税种列表（按长度降序排列，优先匹配长的税种名）
    const knownTaxTypes = [
      '城市维护建设税',
      '地方教育附加',
      '教育费附加',
      '土地增值税',
      '环境保护税',
      '个人所得税',
      '企业所得税',
      '增值税',
      '印花税',
      '房产税',
      '车船税',
      '契税',
      '资源税',
      '城镇土地使用税',
      '土地使用税',
      '耕地占用税'
    ];
    
    for (const taxType of knownTaxTypes) {
      let searchPos = 0;
      
      while (true) {
        const taxTypeIndex = text.indexOf(taxType, searchPos);
        if (taxTypeIndex === -1) break;
        
        // 检查这个位置是否已经被匹配过
        const isAlreadyMatched = matchedPositions.some(pos => 
          taxTypeIndex >= pos.start && taxTypeIndex < pos.end
        );
        
        if (isAlreadyMatched) {
          searchPos = taxTypeIndex + taxType.length;
          continue;
        }
        
        // 检查是否是完整的税种名称（前后是空格或边界）
        const beforeChar = taxTypeIndex > 0 ? text[taxTypeIndex - 1] : ' ';
        const afterChar = taxTypeIndex + taxType.length < text.length ? text[taxTypeIndex + taxType.length] : ' ';
        
        const isValidMatch = /\s|\d/.test(beforeChar) && /\s/.test(afterChar);
        
        if (!isValidMatch) {
          searchPos = taxTypeIndex + 1;
          continue;
        }
        
        // 从税种位置开始，向后找150字符内的数据
        const afterTaxType = text.substring(taxTypeIndex + taxType.length, taxTypeIndex + taxType.length + 150);
        
        // 提取金额（第一个金额是实缴金额，不是最后一个！）
        const amounts = afterTaxType.match(/(\d{1,3}(?:,\d{3})*\.\d{2})/g) || [];
        // 取第一个金额，但要排除可能的税票号码（超过10位数字的金额）
        let amount = '';
        for (const a of amounts) {
          const cleanAmount = a.replace(/,/g, '');
          // 金额应该是小于100万的数值（排除税票号码）
          if (parseFloat(cleanAmount) < 1000000) {
            amount = cleanAmount;
            break;
          }
        }
        
        // 提取入库日期（最后一个日期）
        const dates = afterTaxType.match(/(\d{4}-\d{2}-\d{2})/g) || [];
        const date = dates.length > 0 ? dates[dates.length - 1] : '';
        
        // 提取税款所属期（起始 至 结束）
        const periodMatch = afterTaxType.match(/(\d{4}-\d{2}-\d{2})\s+至\s+(\d{4}-\d{2}-\d{2})/);
        const periodStart = periodMatch ? periodMatch[1] : '';
        const periodEnd = periodMatch ? periodMatch[2] : '';
        
        // 提取税目（税种后到第一个日期之前的内容）
        const beforeFirstDate = afterTaxType.split(/\d{4}-\d{2}/)[0] || '';
        const taxItemMatch = beforeFirstDate.match(/^\s*(\S+)/);
        const taxItem = taxItemMatch ? taxItemMatch[1].trim() : '';
        
        if (amount && date) {
          const record = {
            taxType: taxType,
            taxItem,
            periodStart,
            periodEnd,
            date,
            amount,
            position: taxTypeIndex
          };
          
          matchedPositions.push({
            start: taxTypeIndex,
            end: taxTypeIndex + taxType.length
          });
          
          this.log('找到记录[' + taxType + '] 位置:' + taxTypeIndex + ':', record);
          allRecords.push(record);
        }
        
        searchPos = taxTypeIndex + taxType.length;
      }
    }
    
    // 按PDF文本中的位置排序
    allRecords.sort((a, b) => a.position - b.position);
    this.log('按位置排序后的记录顺序:', allRecords.map(r => r.taxType));
    
    // 移除 position 字段
    return allRecords.map(function(record) {
      const result = {
        taxType: record.taxType,
        taxItem: record.taxItem,
        periodStart: record.periodStart,
        periodEnd: record.periodEnd,
        date: record.date,
        amount: record.amount
      };
      return result;
    });
  }

  /**
   * 提取头部信息
   */  /**
   * 提取头部信息
   */
  extractHeaderInfo(text) {
    const info = {
      companyName: '',
      companyCode: '',
      taxOffice: '',
      voucherNo: '',
      fillDate: '',
      remarks: ''
    };
    
    // 提取税票号码
    const voucherMatch = text.match(/No\s*[●.•]?\s*(\d{18})/i);
    if (voucherMatch) {
      info.voucherNo = voucherMatch[1];
      this.log('提取税票号码:', info.voucherNo);
    }
    
    // 提取纳税人名称
    const companyMatch = text.match(/纳税人名称\s+([^\s]+公司[^\s]*)/);
    if (companyMatch) {
      info.companyName = companyMatch[1].trim();
      this.log('提取纳税人名称:', info.companyName);
    }
    
    // 提取纳税人识别号
    const codeMatch = text.match(/纳税人识别号\s+([A-Z0-9]{18})/i);
    if (codeMatch) {
      info.companyCode = codeMatch[1];
      this.log('提取纳税人识别号:', info.companyCode);
    } else {
      const codes = text.match(/([A-Z0-9]{18})/g) || [];
      for (const c of codes) {
        if (c.match(/[A-Z]/)) {
          info.companyCode = c;
          break;
        }
      }
    }
    
    // 提取填发日期
    const fillDateMatch = text.match(/填发日期[：:]*\s*(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日/i);
    if (fillDateMatch) {
      info.fillDate = fillDateMatch[1] + '-' + fillDateMatch[2].padStart(2, '0') + '-' + fillDateMatch[3].padStart(2, '0');
      this.log('提取填发日期:', info.fillDate);
    }
    
    // 提取税务机关
    const taxOfficeMatch = text.match(/税务机关[：:]\s*([^\s]+税务局)\s+([^\s]+科)/i);
    if (taxOfficeMatch) {
      info.taxOffice = taxOfficeMatch[1].trim() + ' ' + taxOfficeMatch[2].trim();
      this.log('提取税务机关:', info.taxOffice);
    } else {
      const taxOfficeMatch2 = text.match(/税务机关[：:]\s*([^\s]+税务局)/i);
      if (taxOfficeMatch2) {
        info.taxOffice = taxOfficeMatch2[1].trim();
      }
    }
    
    // 提取主管税务所
    const taxOfficeMatch3 = text.match(/主管税务所[（(][^)）]+[)）][：:]\s*(国家税务\s*总局[^\s]+税务所)/i);
    if (taxOfficeMatch3) {
      info.taxOffice = taxOfficeMatch3[1].replace(/\s+/g, '');
      this.log('提取主管税务所:', info.taxOffice);
    } else {
      const taxOfficeMatch4 = text.match(/国家税务\s*总局[^\s]*税务局[^\s]*税务所/i);
      if (taxOfficeMatch4) {
        info.taxOffice = taxOfficeMatch4[0].replace(/\s+/g, '');
        this.log('提取主管税务所(备选):', info.taxOffice);
      }
    }
    
    // 提取备注（改进版：截取到税务局）
    // 格式1: 备注: xxx 或 备注 : xxx
    const remarkMatch = text.match(/备注\s*[：:]\s*([^\n]+)/i);
    if (remarkMatch) {
      let remarks = remarkMatch[1].trim();
      // 截取到税务局（税务所）关键字
      const taxOfficeEnd = remarks.match(/.*?税务[局所]/);
      if (taxOfficeEnd) {
        remarks = taxOfficeEnd[0];
      }
      info.remarks = remarks;
    }
    
    // 格式2: 如果备注后没有内容，尝试匹配下一行
    if (!info.remarks && text.includes('备注')) {
      const lines = text.split('\n');
      for (let i = 0; i < lines.length; i++) {
        const line = lines[i].trim();
        if (line.startsWith('备注') || line.includes('备注')) {
          // 检查同行是否有内容
          const afterRemark = line.replace(/^.*?备注\s*[：:]\s*/, '').trim();
          if (afterRemark && afterRemark.length > 0) {
            info.remarks = afterRemark;
            break;
          }
          // 检查下一行
          if (i + 1 < lines.length) {
            const nextLine = lines[i + 1].trim();
            // 跳过空行
            let j = i + 1;
            while (j < lines.length && lines[j].trim() === '') j++;
            if (j < lines.length) {
              const candidate = lines[j].trim();
              // 备注通常是中文内容，不包含数字金额等
              if (/[\u4e00-\u9fa5]/.test(candidate) && !/\d{4}[-.\s]\d{2}/.test(candidate)) {
                info.remarks = candidate;
                break;
              }
            }
          }
        }
      }
    }
    
    // 格式3: 从文本末尾提取可能的备注（税票右下角）
    if (!info.remarks) {
      const lines = text.split('\n').map(l => l.trim()).filter(l => l);
      // 从后往前找包含中文的行
      for (let i = lines.length - 1; i >= 0; i--) {
        const line = lines[i];
        // 跳过金额、日期、税票号等
        if (/^\d+[\d,\s]*\.\s*\d{2}$/.test(line)) continue;
        if (/^\d{4}[.\s-]\d{2}/.test(line)) continue;
        if (/^No\./i.test(line)) continue;
        if (line.length < 3) continue;
        
        // 找到包含中文且可能是备注的内容
        if (/[\u4e00-\u9fa5]/.test(line) && !line.includes('税务机关') && !line.includes('纳税人')) {
          info.remarks = line;
          break;
        }
      }
    }
    
    this.log('提取备注结果:', info.remarks);
    
    this.log('头部信息结果:', info);
    return info;
  }
  
  /**
   * 提取合计金额
   */
  extractTotalAmount(text) {
    // 匹配格式：金额合计 (大写)xxx ¥ 5,780.23 或 合计 ¥5780.23
    // 支持半角¥和全角￥符号
    // 支持逗号分隔(5,780.23)和点号分隔(5.780.13)
    const totalMatch = text.match(/(?:金额)?合计[^¥￥]*[¥￥]\s*([\d,.\s]+\.\d{2})/);
    if (totalMatch) {
      let amountStr = totalMatch[1].replace(/,/g, '').replace(/\s/g, '');
      // 如果有多于一个点号，说明是千位分隔符
      const dots = (amountStr.match(/\./g) || []).length;
      if (dots > 1) {
        amountStr = amountStr.replace(/\./g, '').replace(/(\d{2})$/, '.$1');
      }
      return parseFloat(amountStr);
    }
    
    // 匹配格式：¥ 5,780.23 或 ￥20.010.13（带¥/￥符号的金额）
    const yenMatch = text.match(/[¥￥]\s*([\d,.\s]+\.\d{2})/);
    if (yenMatch) {
      let amountStr = yenMatch[1].replace(/,/g, '').replace(/\s/g, '');
      // 如果有多于一个点号，说明是千位分隔符
      const dots = (amountStr.match(/\./g) || []).length;
      if (dots > 1) {
        amountStr = amountStr.replace(/\./g, '').replace(/(\d{2})$/, '.$1');
      }
      return parseFloat(amountStr);
    }
    
    return 0;
  }
  
  /**
   * 验证金额之和是否等于合计
   */
  validateAmounts(records, totalAmount) {
    if (!totalAmount || totalAmount === 0 || !records || records.length === 0) {
      return { valid: true, message: '无需验证' };
    }
    
    // 计算记录金额之和（详细日志）
    let sum = 0;
    this.log('=== 金额验证详情 ===');
    records.forEach((r, idx) => {
      const amountStr = r['金额'] || r.amount || '0';
      const amount = parseFloat(amountStr.toString().replace(/,/g, '')) || 0;
      this.log(`记录${idx + 1}: 金额字段="${amountStr}", 数值=${amount}`);
      sum += amount;
    });
    this.log(`记录合计: ${sum.toFixed(2)}, 凭证合计: ${totalAmount.toFixed(2)}`);
    
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
}

window.FieldMatcher = new FieldMatcher();