/**
 * 字段配置管理器
 * 支持完整的字段配置格式
 */

class FieldConfigManager {
  constructor() {
    this.config = [];
    this.debug = true;
  }

  log(...args) {
    if (this.debug) console.log('[FieldConfig]', ...args);
  }

  /**
   * 加载字段配置（支持多种格式）
   */
  loadConfig(data) {
    if (Array.isArray(data)) {
      // 直接是配置数组
      this.config = data;
    } else if (data.fields) {
      // 包含fields属性
      this.config = data.fields;
    } else if (data.template && data.template.fields) {
      // 导出格式
      this.config = data.template.fields;
    }
    
    this.log('加载配置:', this.config.length, '个字段');
    return this.config;
  }

  /**
   * 获取启用的字段
   */
  getEnabledFields() {
    return this.config.filter(f => f.enabled !== false);
  }

  /**
   * 获取aggregate字段（需要分表的字段）
   */
  getAggregateFields() {
    return this.config.filter(f => f.aggregate === true && f.enabled !== false);
  }

  /**
   * 获取字段映射 (fieldName -> columnHeader)
   */
  getFieldMapping() {
    const mapping = {};
    this.getEnabledFields().forEach(f => {
      mapping[f.fieldName] = f.columnHeader;
    });
    return mapping;
  }

  /**
   * 获取列头列表
   */
  getHeaders() {
    return this.getEnabledFields().map(f => f.columnHeader);
  }

  /**
   * 按aggregate字段分组数据
   */
  groupByAggregate(records) {
    const aggregateFields = this.getAggregateFields();
    
    if (aggregateFields.length === 0) {
      return { '__all__': records };
    }

    const groups = {};
    
    // 使用第一个aggregate字段分组
    const groupField = aggregateFields[0].columnHeader;
    
    records.forEach(record => {
      const key = record[groupField] || '未分类';
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(record);
    });
    
    this.log('分组结果:', Object.keys(groups), '组');
    return groups;
  }

  /**
   * 转换为UI编辑格式
   */
  toEditFormat() {
    return this.config.map(f => ({
      name: f.columnHeader,
      alias: f.fieldName,
      enabled: f.enabled !== false,
      aggregate: f.aggregate === true,
      valueType: f.valueType || 'text',
      description: f.description || ''
    }));
  }

  /**
   * 从UI编辑格式转换回配置格式
   */
  fromEditFormat(editData, name = '自定义模板') {
    return editData.map(item => ({
      fieldName: item.alias || item.name,
      columnHeader: item.name,
      description: item.description || '',
      positionHint: '根据图片内容自动识别',
      valueType: item.valueType || 'text',
      enabled: item.enabled !== false,
      aggregate: item.aggregate === true
    }));
  }

  /**
   * 验证配置
   */
  validate() {
    const errors = [];
    const headers = new Set();
    
    this.config.forEach((f, i) => {
      if (!f.fieldName) {
        errors.push(`第${i + 1}个字段缺少fieldName`);
      }
      if (!f.columnHeader) {
        errors.push(`第${i + 1}个字段缺少columnHeader`);
      }
      if (headers.has(f.columnHeader)) {
        errors.push(`列头"${f.columnHeader}"重复`);
      }
      headers.add(f.columnHeader);
    });
    
    return { valid: errors.length === 0, errors };
  }
}

window.FieldConfigManager = new FieldConfigManager();