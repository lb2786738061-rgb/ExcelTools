/**
 * 专供小程序的纯逻辑 Excel 处理引擎
 * 剥离了所有 DOM 渲染和 Browser API，仅保留核心计算算法
 */
import * as XLSX from 'xlsx';

export function getColNumericStats(cleanedData, colIndex) {
  const headerStr = String(cleanedData[0][colIndex] || '');
  if (/(IP|电话|手机|联系|账号|密码|编号|代码|流水号|号|ID|邮编|日期|时间|年份|月份)/i.test(headerStr)) {
    return { sum: 0, count: 0, max: 0, min: 0 };
  }

  let sum = 0;
  let count = 0;
  let max = -Infinity;
  let min = Infinity;

  for (let rowIndex = 1; rowIndex < cleanedData.length; rowIndex++) {
    const val = cleanedData[rowIndex][colIndex];
    if (val === null || val === undefined || String(val).trim() === '') continue;

    const strVal = String(val).trim();
    if (/^1[3-9]\d{9}$/.test(strVal) || /^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(strVal) || /^\d{1,2}[-/.]\d{1,2}$/.test(strVal)) {
      continue;
    }

    const numVal = Number(strVal);
    if (typeof val === 'number' || (!isNaN(numVal) && isFinite(numVal) && !strVal.startsWith('0') && strVal !== '')) {
      sum += numVal;
      count++;
      if (numVal > max) max = numVal;
      if (numVal < min) min = numVal;
    }
  }

  return { sum, count, max: max === -Infinity ? 0 : max, min: min === Infinity ? 0 : min };
}

/**
 * 纯逻辑：从二位数组数据中进行高阶分析
 */
export function processExcelDataLogic(originalAoa, options) {
  let cleanedData = [...originalAoa.map(row => [...row])];
  if (cleanedData.length === 0) return { error: '空数据' };

  const { trimText, maskSensitive, autoSum, calcFormula, advancedFormula } = options;

  if (trimText) {
    cleanedData = cleanedData.map(row => row.map(cell => typeof cell === 'string' ? cell.trim().replace(/\s+/g, ' ') : cell));
  }

  if (maskSensitive) {
    cleanedData = cleanedData.map(row => row.map(cell => {
      if (typeof cell === 'string') {
        if (/^1[3-9]\d{9}$/.test(cell.trim())) return cell.trim().replace(/(\d{3})\d{4}(\d{4})/, '$1****$2');
        if (/^\d{17}[\dXx]$/.test(cell.trim())) return cell.trim().replace(/(\d{6})\d{8}(\d{4})/, '$1********$2');
      }
      return cell;
    }));
  }

  let hasSummaryRow = false;
  if (autoSum && cleanedData.length > 1) {
    const colCount = Math.max(...cleanedData.map(r => r.length));
    
    if (calcFormula === 'SUM') {
      const summaryRow = new Array(colCount).fill('');
      summaryRow[0] = '自动汇总合计 (SUM)';
      for (let colIndex = 1; colIndex < colCount; colIndex++) {
        const stats = getColNumericStats(cleanedData, colIndex);
        if (stats.count > 0) {
          summaryRow[colIndex] = Number(stats.sum.toFixed(2));
          hasSummaryRow = true;
        }
      }
      if (hasSummaryRow) cleanedData.push(summaryRow);
    } 
    // 其他公式逻辑为了精简先略过，如果是正式生产可以完整 copy 原逻辑
  }

  // SUMIF VIP 功能
  if (advancedFormula === 'sumif' && cleanedData.length > 1) {
    const colCount = Math.max(...cleanedData.map(r => r.length));
    const categoryMap = new Map();
    const originLength = cleanedData.length;
    
    for (let r = 1; r < originLength; r++) {
      const rowTitle = String(cleanedData[r][0] || '');
      if (rowTitle.includes('合计')) continue;
      
      let cat = rowTitle.trim() || '未分类';
      if (!categoryMap.has(cat)) categoryMap.set(cat, new Array(colCount).fill(0));
      
      const sums = categoryMap.get(cat);
      for (let c = 1; c < colCount; c++) {
        const headerStr = String(cleanedData[0][c] || '');
        if (/(IP|电话|手机|账号|密码|编号|代码|流水号|ID|邮编|日期|时间)/i.test(headerStr)) continue;

        const num = Number(cleanedData[r][c]);
        if (!isNaN(num) && isFinite(num) && !/^1[3-9]\d{9}$/.test(String(cleanedData[r][c]))) {
          sums[c] += num;
        }
      }
    }
    
    cleanedData.push(new Array(colCount).fill(''));
    const titleRow = new Array(colCount).fill('');
    titleRow[0] = '📊 SUMIF 智能透视表';
    cleanedData.push(titleRow);
    
    for (const [cat, sums] of categoryMap.entries()) {
      const row = new Array(colCount).fill('');
      row[0] = cat;
      let hasData = false;
      for (let c = 1; c < colCount; c++) {
        if (sums[c] !== 0) {
          row[c] = Number(sums[c].toFixed(2));
          hasData = true;
        }
      }
      if (hasData) cleanedData.push(row);
    }
  }

  // 返回的是一个纯二维数组，方便 Vue 在 template 里做 v-for 渲染
  return {
    success: true,
    data: cleanedData
  };
}
