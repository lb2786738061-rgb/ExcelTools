/**
 * Office 智能效率工具箱 - Excel 处理与 2/3 常规核心公式计算引擎
 * 核心公式支持：SUM (求和), AVERAGE (平均值), MAX/MIN (极值), COUNT (计数), 
 * SUMIF (条件求和), TRIM (文本清洗), IFERROR (错误清洗), VLOOKUP/CONCAT 等
 */

import * as XLSX from 'xlsx';

/**
 * 解析并处理 Excel 文件数据
 * @param {ArrayBuffer} arrayBuffer - 上传的原始文件二进制数据
 * @param {Object} options - 开关与模板配置参数
 * @returns {Object} 包含处理后的 HTML 渲染结构、数据表与元数据
 */
export function processExcelFile(arrayBuffer, options = {}) {
  try {
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });

    if (!workbook || !workbook.SheetNames || workbook.SheetNames.length === 0) {
      return { error: '文件中未找到任何有效工作表 (Sheet)' };
    }

    let targetSheetName = null;
    let rawData = [];

    for (const name of workbook.SheetNames) {
      const sheet = workbook.Sheets[name];
      if (!sheet) continue;

      const data = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
      const validRows = data.filter(row => Array.isArray(row) && row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== ''));

      if (validRows.length > 0) {
        targetSheetName = name;
        rawData = data;
        break;
      }
    }

    if (!targetSheetName || rawData.length === 0) {
      return { error: '文件内容为空或无法识别有效表格结构（可能为空白表格或受加密保护）' };
    }

    return processExcelGrid(rawData, options, targetSheetName);
  } catch (err) {
    console.error('Excel 解析处理异常:', err);
    return { error: '文件解析失败，文件可能受损或格式无法识别' };
  }
}

export function processExcelGrid(rawData, options = {}, targetSheetName = 'Sheet1') {
  const {
    template = 'standard',
    calcFormula = 'SUM',
    autoSum = true,
    autoBorder = true,
    autoWidth = true,
    stripeColor = true,
    maskSensitive = false,
    trimText = true,
    cleanErrors = true
  } = options;

  try {
    let cleanedData = rawData.filter(row => Array.isArray(row) && row.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== ''));

    if (cleanedData.length === 0) {
      return { error: '未在工作表中找到有效行数据' };
    }

    // 执行 TRIM 文本清洗
    if (trimText) {
      cleanedData = cleanedData.map(row => row.map(cell => typeof cell === 'string' ? cell.trim().replace(/\s+/g, ' ') : cell));
    }

    // 2. 敏感数据打码脱敏逻辑 (手机号/身份证)
    if (maskSensitive) {
      cleanedData = cleanedData.map(row => {
        return row.map(cell => {
          if (typeof cell === 'string') {
            if (/^1[3-9]\d{9}$/.test(cell.trim())) {
              return cell.trim().replace(/(\d{3})\d{4}(\d{4})/, '$1****$2');
            }
            if (/^\d{17}[\dXx]$/.test(cell.trim())) {
              return cell.trim().replace(/(\d{6})\d{8}(\d{4})/, '$1********$2');
            }
          }
          return cell;
        });
      });
    }

    if (options.removeDuplicates && cleanedData.length > 1) {
      const seen = new Set();
      const header = cleanedData[0];
      const body = cleanedData.slice(1).filter((row) => {
        const key = JSON.stringify(row);
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
      });
      cleanedData = [header, ...body];
    }

    const sortColRaw = options.sortCol;
    if (sortColRaw !== undefined && sortColRaw !== null && sortColRaw !== '' && sortColRaw !== 'none') {
      const col = Number(sortColRaw);
      if (Number.isInteger(col) && col >= 0) {
        const header = cleanedData[0];
        const body = cleanedData.slice(1);
        const dir = options.sortDir === 'desc' ? -1 : 1;
        body.sort((a, b) => {
          const va = a[col];
          const vb = b[col];
          const na = Number(va);
          const nb = Number(vb);
          if (Number.isFinite(na) && Number.isFinite(nb) && String(va).trim() !== '' && String(vb).trim() !== '') {
            return (na - nb) * dir;
          }
          return String(va ?? '').localeCompare(String(vb ?? ''), 'zh') * dir;
        });
        cleanedData = [header, ...body];
      }
    }

    if (options.roundNumbers && cleanedData.length > 1) {
      cleanedData = cleanedData.map((row, idx) => {
        if (idx === 0) return row;
        return row.map((cell) => (typeof cell === 'number' && Number.isFinite(cell) ? Math.round(cell * 100) / 100 : cell));
      });
    }

    if (options.concatColumns && cleanedData.length > 0 && cleanedData[0].length >= 2) {
      cleanedData = cleanedData.map((row, idx) => {
        const next = row.slice();
        if (idx === 0) {
          next.push('CONCAT 拼接');
        } else {
          next.push(`${row[0] ?? ''}-${row[1] ?? ''}`);
        }
        return next;
      });
    }

    if (options.ifStatusColumn && cleanedData.length > 1) {
      let valueCol = -1;
      for (let c = 1; c < cleanedData[0].length; c++) {
        if (getColNumericStats(cleanedData, c).count > 0) {
          valueCol = c;
          break;
        }
      }
      if (valueCol >= 0) {
        cleanedData = cleanedData.map((row, idx) => {
          const next = row.slice();
          if (idx === 0) {
            next.push('IF状态');
          } else {
            const n = Number(row[valueCol]);
            next.push(Number.isFinite(n) && n < 0 ? '异常' : '正常');
          }
          return next;
        });
      }
    }

    // 3. 全量 Excel 核心公式计算引擎 (SUM / AVERAGE / MAX / MIN / COUNT / SUMIF)
    let hasSummaryRow = false;
    const formulaCells = [];
    const lastDataExcelRow = cleanedData.length;

    function pushFormulaRow(label, fnName, valueFn) {
      const colCount = Math.max(...cleanedData.map(r => r.length));
      const row = new Array(colCount).fill('');
      row[0] = label;
      const rowIndex = cleanedData.length;
      let wrote = false;
      for (let colIndex = 1; colIndex < colCount; colIndex++) {
        const stats = getColNumericStats(cleanedData, colIndex);
        const minRatio = fnName === 'COUNTA' ? 0 : 0.4;
        if (stats.count > 0 && stats.count >= (cleanedData.length - 1) * minRatio) {
          const value = valueFn(stats);
          row[colIndex] = value;
          formulaCells.push({
            r: rowIndex,
            c: colIndex,
            f: `${fnName}(${colLetter(colIndex)}2:${colLetter(colIndex)}${lastDataExcelRow})`,
            v: typeof value === 'number' ? value : stats.count
          });
          wrote = true;
        }
      }
      if (wrote) {
        cleanedData.push(row);
        hasSummaryRow = true;
      }
    }

    if (autoSum && cleanedData.length > 1) {
      if (calcFormula === 'SUM') {
        pushFormulaRow('自动汇总合计 (SUM)', 'SUM', (stats) => Number(stats.sum.toFixed(2)));
      } else if (calcFormula === 'AVERAGE') {
        pushFormulaRow('全列平均值 (AVERAGE)', 'AVERAGE', (stats) => Number((stats.sum / stats.count).toFixed(2)));
      } else if (calcFormula === 'MAX_MIN') {
        pushFormulaRow('极大值 (MAX)', 'MAX', (stats) => Number(stats.max.toFixed(2)));
        pushFormulaRow('极小值 (MIN)', 'MIN', (stats) => Number(stats.min.toFixed(2)));
      } else if (calcFormula === 'COUNT') {
        pushFormulaRow('有效记录数统计 (COUNT)', 'COUNT', (stats) => stats.count);
      }
    }

    // ==========================================
    // 3.5 💎 独家高级智能分析引擎 (SUMIF / SUBTOTAL)
    // ==========================================
    const { advancedFormula = 'none' } = options;
    if (advancedFormula !== 'none' && cleanedData.length > 1) {
      const colCount = Math.max(...cleanedData.map(r => r.length));

      if (advancedFormula === 'sumif') {
        // 【SUMIF 分类透视汇总算法】: 以第0列作为分类维度，对其他所有数字列进行条件求和
        const categoryMap = new Map(); // key: 分类名, value: [col1_sum, col2_sum...]
        const originLength = cleanedData.length;

        for (let r = 1; r < originLength; r++) {
          const rowTitle = String(cleanedData[r][0]);
          if (rowTitle.includes('合计') || rowTitle.includes('平均值') || rowTitle.includes('极值') || rowTitle.includes('记录数')) continue;

          let cat = rowTitle.trim();
          if (!cat) cat = '未分类';
          if (!categoryMap.has(cat)) categoryMap.set(cat, new Array(colCount).fill(0));

          const sums = categoryMap.get(cat);
          for (let c = 1; c < colCount; c++) {
            // 🤖 智能防呆：根据表头语义拒绝将特定属性列纳入计算
            const headerStr = String(cleanedData[0][c]);
            if (/(IP|电话|手机|联系|账号|密码|编号|代码|流水号|号|ID|邮编|日期|时间|年份|月份)/i.test(headerStr)) continue;

            const num = Number(cleanedData[r][c]);
            if (!isNaN(num) && isFinite(num) && !/^1[3-9]\d{9}$/.test(String(cleanedData[r][c]))) {
              sums[c] += num;
            }
          }
        }

        cleanedData.push(new Array(colCount).fill('')); // 空行分割
        const titleRow = new Array(colCount).fill('');
        titleRow[0] = '📊 [VIP智能生成] SUMIF 分类透视汇总报表';
        cleanedData.push(titleRow);

        const subHeader = new Array(colCount).fill('');
        subHeader[0] = '分类维度';
        for (let c = 1; c < colCount; c++) subHeader[c] = cleanedData[0][c] ? `${cleanedData[0][c]} (汇总)` : `列${c + 1} (汇总)`;
        cleanedData.push(subHeader);

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
        hasSummaryRow = true;
      }
      else if (advancedFormula === 'subtotal') {
        // 【SUBTOTAL 净值汇总算法】: 剔除负数、零值及错误数据，仅对绝对正向有效值求和
        const originLength = cleanedData.length;
        cleanedData.push(new Array(colCount).fill('')); // 空行分割

        const titleRow = new Array(colCount).fill('');
        titleRow[0] = '📊 [VIP智能生成] SUBTOTAL 剔除异常值净和报表';
        cleanedData.push(titleRow);

        const validSumRow = new Array(colCount).fill('');
        validSumRow[0] = '净有效值合计 (>0)';

        for (let c = 1; c < colCount; c++) {
          // 🤖 智能防呆：根据表头语义拒绝将特定属性列纳入计算
          const headerStr = String(cleanedData[0][c]);
          if (/(IP|电话|手机|联系|账号|密码|编号|代码|流水号|号|ID|邮编|日期|时间|年份|月份)/i.test(headerStr)) continue;

          let sum = 0;
          for (let r = 1; r < originLength; r++) {
            const rowTitle = String(cleanedData[r][0]);
            if (rowTitle.includes('合计') || rowTitle.includes('极值') || rowTitle.includes('记录数')) continue;

            const num = Number(cleanedData[r][c]);
            if (!isNaN(num) && isFinite(num) && num > 0 && !/^1[3-9]\d{9}$/.test(String(cleanedData[r][c]))) {
              sum += num;
            }
          }
          if (sum > 0) validSumRow[c] = Number(sum.toFixed(2));
        }
        cleanedData.push(validSumRow);
        hasSummaryRow = true;
      }
    }

    // 4. 数字显示格式化、IFERROR 错误清洗与热力图色阶
    const { numberFormat = 'normal', colorScale = false, colorScaleType = 'green-yellow-red' } = options;
    const colorScaleMap = {};

    let formattedData = cleanedData.map((row, rIdx) => {
      if (rIdx === 0) return row; // 表头不格式化
      return row.map((cell, cIdx) => {
        let strVal = String(cell).trim();

        // 自动清洗公式错误 IFERROR (#DIV/0!, #N/A, #VALUE!, #REF!)
        if (cleanErrors && /^#(DIV\/0!|N\/A|VALUE!|REF!|NAME\?|NUM!|NULL!)$/i.test(strVal)) {
          return '-';
        }

        const headerStr = String(cleanedData[0][cIdx] || '').trim();
        // 🤖 序号/编号列保护：判定表头语义是否为序号、编号、ID、NO等
        const isSequenceCol = /(序号|编号|ID|NO|No|代码|项次|行号|Index)/i.test(headerStr) || (cIdx === 0 && /^\d+$/.test(strVal) && Number(strVal) < 1000);

        const numVal = Number(strVal);

        if (cell !== '' && typeof cell !== 'boolean' && !isNaN(numVal) && isFinite(numVal) && !/^1[3-9]\d{9}$/.test(strVal)) {
          // 如果是序号列或简单整数序号，直接保持整数字符串，绝不加小数点 .00
          if (isSequenceCol) {
            return String(Math.round(numVal));
          }

          let formattedStr = cell;
          if (numberFormat === 'thousand') {
            formattedStr = Number.isInteger(numVal) ? numVal.toLocaleString('en-US') : numVal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
          } else if (numberFormat === 'currency') {
            formattedStr = `¥${numVal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
          } else if (numberFormat === 'dollar') {
            formattedStr = `$${numVal.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
          } else if (numberFormat === 'percent') {
            formattedStr = `${(numVal * (numVal <= 1 ? 100 : 1)).toFixed(1)}%`;
          }

          if (colorScale) {
            if (colorScaleType === 'blue-white-red') {
              if (numVal >= 100000) colorScaleMap[`${rIdx - 1}_${cIdx}`] = 'background: rgba(239, 68, 68, 0.25); color: #dc2626; font-weight: bold;';
              else if (numVal >= 50000) colorScaleMap[`${rIdx - 1}_${cIdx}`] = 'background: rgba(59, 130, 246, 0.2); color: #2563eb;';
            } else if (colorScaleType === 'green-gradient') {
              if (numVal >= 100000) colorScaleMap[`${rIdx - 1}_${cIdx}`] = 'background: rgba(16, 185, 129, 0.35); color: #047857; font-weight: bold;';
              else if (numVal >= 50000) colorScaleMap[`${rIdx - 1}_${cIdx}`] = 'background: rgba(16, 185, 129, 0.18); color: #059669;';
            } else {
              // 经典 绿 - 黄 - 红
              if (numVal >= 100000) colorScaleMap[`${rIdx - 1}_${cIdx}`] = 'background: rgba(16, 185, 129, 0.25); color: #047857; font-weight: bold;';
              else if (numVal >= 50000) colorScaleMap[`${rIdx - 1}_${cIdx}`] = 'background: rgba(245, 158, 11, 0.25); color: #b45309;';
            }
          }

          return formattedStr;
        }
        return cell;
      });
    });

    const header = formattedData[0] || [];
    const bodyRows = formattedData.slice(1);
    const exportGrid = cleanedData.map(row => Array.isArray(row) ? row.slice() : row);

    return {
      sheetName: targetSheetName,
      header,
      rows: bodyRows,
      totalRows: formattedData.length,
      totalCols: Math.max(...formattedData.map(r => r.length)),
      templateStyle: template,
      hasSummaryRow,
      colorScaleMap,
      rawProcessedData: formattedData,
      exportGrid,
      formulaCells,
      freezeHeader: options.freezeHeader !== false,
      autoWidth: autoWidth !== false,
      sheetTabColor: options.sheetTabColor || '',
      stripeColor: stripeColor !== false,
      chartData: buildChartData(cleanedData)
    };
  } catch (err) {
    console.error('Excel 解析处理异常:', err);
    return { error: '文件解析失败，文件可能受损或格式无法识别' };
  }
}

/**
 * 提取指定列的数值统计信息 (用于 SUM / AVERAGE / MAX / MIN / COUNT)
 */
function getColNumericStats(cleanedData, colIndex) {
  // 🤖 智能防呆：根据表头语义拒绝将特定属性列纳入计算
  const headerStr = String(cleanedData[0][colIndex]);
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

function buildChartData(grid) {
  if (!Array.isArray(grid) || grid.length < 2) return null;
  const header = grid[0] || [];
  let valueCol = -1;
  for (let c = 1; c < header.length; c++) {
    if (getColNumericStats(grid, c).count > 0) {
      valueCol = c;
      break;
    }
  }
  if (valueCol < 0) return null;
  const labels = [];
  const values = [];
  for (let r = 1; r < grid.length; r++) {
    const title = String(grid[r][0] ?? '');
    if (/合计|平均|极大|极小|记录数|SUMIF|SUBTOTAL|分类/.test(title)) continue;
    labels.push(title || `行${r}`);
    const n = Number(grid[r][valueCol]);
    values.push(Number.isFinite(n) ? n : 0);
  }
  if (labels.length === 0) return null;
  return { title: String(header[valueCol] || '数值'), labels, values };
}

function colLetter(index) {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function resolveExportGrid(processedData) {
  if (Array.isArray(processedData)) return { aoa: processedData, sheetName: 'Sheet1' };
  if (processedData?.exportGrid?.length) {
    return { aoa: processedData.exportGrid, sheetName: processedData.sheetName || 'Sheet1' };
  }
  if (processedData?.rawProcessedData?.length) {
    return { aoa: processedData.rawProcessedData, sheetName: processedData.sheetName || 'Sheet1' };
  }
  if (processedData?.header) {
    return {
      aoa: [processedData.header, ...(processedData.rows || [])],
      sheetName: processedData.sheetName || 'Sheet1'
    };
  }
  return { aoa: [], sheetName: 'Sheet1' };
}

function writeRealXlsx(aoa, filename, sheetName = 'Sheet1', extra = {}) {
  if (!aoa || aoa.length === 0) {
    throw new Error('没有可导出的数据');
  }

  const ws = XLSX.utils.aoa_to_sheet(aoa);

  if (Array.isArray(extra.formulaCells)) {
    extra.formulaCells.forEach((cell) => {
      if (cell == null || cell.r == null || cell.c == null || !cell.f) return;
      const addr = XLSX.utils.encode_cell({ r: cell.r, c: cell.c });
      ws[addr] = { t: 'n', v: Number(cell.v) || 0, f: cell.f };
    });
  }

  const colCount = Math.max(...aoa.map(row => (Array.isArray(row) ? row.length : 0)));
  if (extra.autoWidth !== false) {
    ws['!cols'] = Array.from({ length: colCount }, (_, i) => {
      let max = 8;
      aoa.forEach(row => {
        const len = String(row?.[i] ?? '').length;
        if (len > max) max = Math.min(len + 2, 40);
      });
      return { wch: max };
    });
  }

  if (extra.freezeHeader !== false) {
    ws['!views'] = [{ state: 'frozen', ySplit: 1, topLeftCell: 'A2' }];
  }

  const wb = XLSX.utils.book_new();
  const safeName = String(sheetName).replace(/[\\/?*[\]]/g, ' ').slice(0, 31) || 'Sheet1';
  XLSX.utils.book_append_sheet(wb, ws, safeName);

  const tabRgb = String(extra.sheetTabColor || '').replace('#', '');
  if (/^[0-9A-Fa-f]{6}$/.test(tabRgb)) {
    wb.Workbook = {
      Sheets: [{ name: safeName, tabColor: { rgb: tabRgb.toUpperCase() } }]
    };
  }

  const outName = String(filename || '美化规整后的文档.xlsx').replace(/\.xls$/i, '.xlsx');
  XLSX.writeFile(wb, outName, { bookType: 'xlsx', cellDates: true });
}

export function exportProcessedExcel(processedData, filename = '美化规整后的文档.xlsx') {
  const { aoa, sheetName } = resolveExportGrid(processedData);
  writeRealXlsx(aoa, filename, sheetName, Array.isArray(processedData) ? {} : {
    formulaCells: processedData.formulaCells,
    freezeHeader: processedData.freezeHeader,
    autoWidth: processedData.autoWidth,
    sheetTabColor: processedData.sheetTabColor
  });
}

export function exportStyledExcel(processedData, filename = '美化规整后的文档.xlsx') {
  exportProcessedExcel(processedData, filename);
}

/**
 * 批量合并多个 Excel/CSV 文件
 * @param {File[]} files - 文件对象数组
 * @returns {Promise<Array[]>} - 返回合并后的二维数组
 */
export async function mergeMultipleExcels(files) {
  if (!files || files.length === 0) throw new Error('未选择任何文件');

  const mergedData = [];

  for (let i = 0; i < files.length; i++) {
    const file = files[i];
    const data = await new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = (e) => {
        try {
          const workbook = XLSX.read(e.target.result, { type: 'array' });
          const firstSheetName = workbook.SheetNames[0];
          const worksheet = workbook.Sheets[firstSheetName];
          const aoa = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
          resolve(aoa);
        } catch (err) {
          reject(new Error(`解析文件 ${file.name} 失败`));
        }
      };
      reader.onerror = () => reject(new Error(`读取文件 ${file.name} 失败`));
      reader.readAsArrayBuffer(file);
    });

    if (data.length === 0) continue;

    // 对基准文件（第一个），保留表头，并在表头末尾追加 "[数据来源]"
    if (mergedData.length === 0) {
      const headerRow = [...(data[0] || [])];
      if (headerRow.length === 0) headerRow.push('数据列');
      headerRow.push('[数据来源]');
      mergedData.push(headerRow);

      for (let r = 1; r < data.length; r++) {
        const row = [...data[r]];
        while (row.length < mergedData[0].length - 1) row.push('');
        row.push(file.name);
        mergedData.push(row);
      }
    } else {
      // 对后续文件，跳过表头 (假定结构一致)
      for (let r = 1; r < data.length; r++) {
        const row = [...data[r]];
        if (row.length === 0 || (row.length === 1 && String(row[0]).trim() === '')) continue;

        while (row.length < mergedData[0].length - 1) row.push('');
        row.length = mergedData[0].length - 1;
        row.push(file.name);
        mergedData.push(row);
      }
    }
  }

  return mergedData;
}

