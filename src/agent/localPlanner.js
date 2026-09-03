/**
 * 不会 Office 的用户用自然语言；无大模型时也能落地到编辑器命令。
 */

import { composeWatermark, parseWatermarkFromSpeech } from '../engine/watermarkEngine.js';

function colByHeader(headers, keys) {
  const list = headers || [];
  const want = keys.map((k) => String(k).trim().toLowerCase()).filter(Boolean);
  for (let i = 0; i < list.length; i++) {
    const h = String(list[i] || '').trim().toLowerCase();
    if (want.some((k) => h === k)) return i;
  }
  for (let i = 0; i < list.length; i++) {
    const h = String(list[i] || '').trim().toLowerCase();
    if (want.some((k) => h.includes(k) || k.includes(h))) return i;
  }
  return -1;
}

function namedColFromText(text, headers) {
  const m = text.match(/把\s*[「"']?([A-Za-z0-9_\u4e00-\u9fa5]+)[」"']?\s*(这|那)?列/)
    || text.match(/[「"']?([A-Za-z0-9_\u4e00-\u9fa5]+)[」"']?\s*(这|那)?一?列/)
    || text.match(/按\s*[「"']?([A-Za-z0-9_\u4e00-\u9fa5]+)[」"']?/)
    || text.match(/给\s*[「"']?([A-Za-z0-9_\u4e00-\u9fa5]+)[」"']?\s*列/);
  if (!m) return -1;
  return colByHeader(headers, [m[1]]);
}

function colFromText(text, headers) {
  const named = namedColFromText(text, headers);
  if (named >= 0) return named;
  const letter = text.match(/\b([A-Za-z]{1,3})\s*列/);
  if (letter) {
    const i = letterToIndex(letter[1]);
    if (i >= 0) return i;
  }
  const num = text.match(/第\s*(\d+)\s*列/);
  if (num) return Math.max(0, Number(num[1]) - 1);
  return -1;
}

function colContainingValue(snap, value) {
  const preview = snap.preview || [];
  const headers = snap.headers || [];
  const want = String(value).trim();
  const hi = colByHeader(headers, [want]);
  if (hi >= 0) return hi;
  for (let c = 0; c < headers.length; c++) {
    for (let r = 1; r < preview.length; r++) {
      const cell = String(preview[r]?.[c] ?? '');
      if (cell && (cell.includes(want) || want.includes(cell))) return c;
    }
  }
  return 0;
}

function letterToIndex(letters) {
  let c = 0;
  for (const ch of String(letters).toUpperCase()) {
    if (ch < 'A' || ch > 'Z') return -1;
    c = c * 26 + (ch.charCodeAt(0) - 64);
  }
  return c - 1;
}

function parseTwoColumns(text, headers) {
  const letters = text.match(/([A-Za-z]{1,3})\s*列?\s*[、,，和与及\/]\s*([A-Za-z]{1,3})\s*两?列?/);
  if (letters) {
    const a = letterToIndex(letters[1]);
    const b = letterToIndex(letters[2]);
    if (a >= 0 && b >= 0) return [Math.min(a, b), Math.max(a, b)];
  }
  const nums = text.match(/第?\s*(\d+)\s*列.{0,6}第?\s*(\d+)\s*列/);
  if (nums) {
    const a = Number(nums[1]) - 1;
    const b = Number(nums[2]) - 1;
    if (a >= 0 && b >= 0) return [Math.min(a, b), Math.max(a, b)];
  }
  const names = text.match(/把\s*[「"']?([A-Za-z0-9_\u4e00-\u9fa5]+)[」"']?.{0,4}[「"']?([A-Za-z0-9_\u4e00-\u9fa5]+)[」"']?\s*(两)?列/);
  if (names && headers) {
    const a = colByHeader(headers, [names[1]]);
    const b = colByHeader(headers, [names[2]]);
    if (a >= 0 && b >= 0) return [Math.min(a, b), Math.max(a, b)];
  }
  return null;
}

function parseFillColor(text) {
  if (/浅\s*绿|淡\s*绿|浅绿/.test(text)) return '#d9ead3';
  if (/深绿/.test(text)) return '#38761d';
  if (/绿/.test(text)) return '#93c47d';
  if (/浅红|粉/.test(text)) return '#f4cccc';
  if (/红/.test(text)) return '#ea9999';
  if (/浅黄|米黄/.test(text)) return '#fff2cc';
  if (/黄/.test(text)) return '#ffe599';
  if (/浅蓝|淡蓝/.test(text)) return '#cfe2f3';
  if (/蓝/.test(text)) return '#9fc5e8';
  if (/橙/.test(text)) return '#fce5cd';
  if (/褐|棕/.test(text)) return '#b45f06';
  if (/灰/.test(text)) return '#efefef';
  if (/白/.test(text)) return '#ffffff';
  return '#d9ead3';
}

const COLOR_WORD = /绿|红|黄|蓝|橙|灰|白|粉|褐|棕/;

function parseRowFromText(text) {
  if (/表头|首行|标题行/.test(text)) return 0;
  const cn = text.match(/第\s*([一二三四五六七八九十]+)\s*行/);
  if (cn) {
    const map = { 一: 1, 二: 2, 三: 3, 四: 4, 五: 5, 六: 6, 七: 7, 八: 8, 九: 9, 十: 10 };
    const n = map[cn[1]];
    if (n) return n - 1;
  }
  const num = text.match(/第\s*(\d+)\s*行/);
  if (num) return Math.max(0, Number(num[1]) - 1);
  return 0;
}

export function parseSlashHeader(text) {
  const src = String(text || '');
  let r = 0;
  let c = 0;
  if (/A1|左上角|第一格/.test(src)) {
    r = 0;
    c = 0;
  } else {
    if (/第\s*[1一]\s*行|表头|首行/.test(src)) r = 0;
    else r = parseRowFromText(src);
    const colNum = src.match(/第\s*[1一]\s*列/) || /A列/.test(src);
    c = colNum ? 0 : 0;
  }
  const topM = src.match(/上(?:部分|边|面|半)?(?:写|填|是为?)?\s*[「"'“]?([^\s，,。；;下「"'”]+)[」"'”]?/);
  const botM = src.match(/下(?:部分|边|面|半)?(?:写|填|是为?)?\s*[「"'“]?([^\s，,。；;上「"'”]+)[」"'”]?/);
  return {
    r,
    c,
    top: topM ? String(topM[1]).trim() : '部分',
    bottom: botM ? String(botM[1]).trim() : '序号'
  };
}

/** 模型把「行」做成「列」、或把斜线表头做成涂色时，按用户原话改回来。 */
export function fixRowColActions(message, actions) {
  const text = String(message || '');
  if (/斜线/.test(text)) {
    return [{ cmd: 'slashHeader', payload: parseSlashHeader(text) }];
  }
  const talksRow = /行|表头/.test(text);
  const talksCol = /列/.test(text);
  if (!Array.isArray(actions) || !actions.length) return actions;
  if (talksRow && !talksCol) {
    const row = parseRowFromText(text);
    const color = parseFillColor(text);
    return actions.map((a) => {
      if (a?.cmd === 'fillColumn') return { cmd: 'fillRow', payload: { row, color: a.payload?.color || color } };
      return a;
    });
  }
  if (talksCol && !talksRow) {
    return actions.map((a) => {
      if (a?.cmd === 'fillRow') {
        const c = Number(a.payload?.col ?? a.payload?.c ?? a.payload?.row ?? 0);
        return { cmd: 'fillColumn', payload: { col: c, color: a.payload?.color || parseFillColor(text) } };
      }
      return a;
    });
  }
  return actions;
}

function lastDataRow(snap) {
  return Math.max(1, Number(snap.maxR) || 1);
}

function numericCol(snap) {
  const headers = snap.headers || [];
  const named = colByHeader(headers, ['工资', '金额', '奖金', '合计', '数量', '业绩', '收入', '费用']);
  if (named >= 0) return named;
  const preview = snap.preview || [];
  for (let c = 1; c < headers.length; c++) {
    let n = 0;
    for (let r = 1; r < preview.length; r++) {
      if (Number.isFinite(Number(preview[r][c])) && preview[r][c] !== '') n++;
    }
    if (n >= 2) return c;
  }
  return 1;
}

export function planLocally(message, mode, snap) {
  const text = String(message || '').trim();
  const actions = [];
  let reply = '';

  if (mode === 'word') {
    if (/红头|公文|通知/.test(text)) {
      actions.push({ cmd: 'redhead' }, { cmd: 'beautify' });
      reply = '已套公文红头，并统一了正文宋体和首行缩进。你可以直接改标题文字。';
    } else if (/水印/.test(text)) {
      const payload = parseWatermarkFromSpeech(text);
      const line = composeWatermark(payload);
      if (!line) {
        reply = '请说清公司和姓名，例如：水印写上华润集团，仅供张三使用。';
      } else {
        actions.push({ cmd: 'watermark', payload: { ...payload, text: line } });
        reply = `已按你的意思盖上「${line}」。`;
      }
    } else if (/目录/.test(text)) {
      actions.push({ cmd: 'toc' });
      reply = '已按文中的标题插入目录。';
    } else if (/表格/.test(text)) {
      actions.push({ cmd: 'table' });
      reply = '已插入表格，点格子就能打字。';
    } else if (/勾选|打勾|复选/.test(text)) {
      actions.push({ cmd: 'checkbox' });
      reply = '已插入勾选框，点一下就能打勾。';
    } else if (/是\s*否|是否/.test(text)) {
      actions.push({ cmd: 'yesno' });
      reply = '已插入「是 / 否」，点其中一个即可。';
    } else if (/多选/.test(text)) {
      actions.push({ cmd: 'multiselect' });
      reply = '请在弹出框里填写选项，然后可以同时勾多个。';
    } else if (/单选/.test(text)) {
      actions.push({ cmd: 'radio' });
      reply = '请填写选项，一组里只能选一个。';
    } else if (/下拉/.test(text)) {
      actions.push({ cmd: 'dropdown' });
      reply = '请填写下拉选项。';
    } else if (/加粗|黑体/.test(text)) {
      actions.push({ cmd: 'style', payload: 'H1' });
      reply = '已把当前段落设成标题。也可先选中文字再让我加粗。';
    } else if (/标题|写成|改成|插入文字/.test(text)) {
      const m = text.match(/(?:写成|改成|插入)\s*[「"']?(.+?)[」"']?$/);
      actions.push({ cmd: 'insertText', payload: { text: m ? m[1] : text } });
      reply = '已把文字写进文档。';
    } else {
      actions.push({ cmd: 'beautify' });
      reply = '已按公文习惯整理字体和段落。你也可以说「加红头」「插目录」「插是非勾选」。';
      return { reply, actions, matched: false };
    }
    return { reply, actions, matched: true };
  }

  const maxR = lastDataRow(snap);
  const col = numericCol(snap);
  const headerName = (snap.headers && snap.headers[col]) || `第${col + 1}列`;

  if (/水印/.test(text)) {
    const payload = parseWatermarkFromSpeech(text);
    const line = composeWatermark(payload);
    if (!line) {
      reply = '请说清公司和姓名，例如：水印写上华润集团，仅供张三使用。光写「内部机密」看不出是谁的文件。';
    } else {
      actions.push({ cmd: 'watermark', payload: { ...payload, text: line } });
      reply = `已按你的意思盖上「${line}」。`;
    }
  } else if (/斜线/.test(text)) {
    const payload = parseSlashHeader(text);
    actions.push({ cmd: 'slashHeader', payload });
    reply = `已在第 ${payload.r + 1} 行第 ${payload.c + 1} 列做成斜线表头：右上写「${payload.top}」，左下写「${payload.bottom}」。`;
  } else if (/把\s*([A-Za-z]{1,3}\d+)\s*(改成|写成|填成|设为|变成)\s*[「"']?(.+?)[」"']?\s*$/.test(text)
    || /([A-Za-z]{1,3}\d+)\s*(单元格)?\s*(改成|写成|填)\s*[「"']?(.+?)[」"']?\s*$/.test(text)) {
    const m = text.match(/([A-Za-z]{1,3}\d+)\s*.*(?:改成|写成|填成|设为|变成|填)\s*[「"']?(.+?)[」"']?\s*$/);
    actions.push({ cmd: 'setCell', payload: { addr: m[1].toUpperCase(), v: m[2] } });
    reply = `已把 ${m[1].toUpperCase()} 写成「${m[2]}」。`;
  } else if (/序号/.test(text) && /列|插入|加|填/.test(text)) {
    const nextCol = (Number(snap.maxC) || 0) + 1;
    actions.push(
      { cmd: 'addColumn', payload: { header: '序号' } },
      { cmd: 'fillSeries', payload: { col: nextCol, r1: 1, r2: maxR, start: 1 } }
    );
    reply = '已在表右侧新增「序号」列，并填入 1、2、3…';
  } else if (/插入一?列|加一列|新增一?列/.test(text)) {
    const m = text.match(/叫\s*[「"']?([^\s「"']+)[」"']?/) || text.match(/列名\s*[「"']?([^\s「"']+)/);
    actions.push({ cmd: 'addColumn', payload: { header: m ? m[1] : '新列' } });
    reply = `已在表右侧加一列「${m ? m[1] : '新列'}」。`;
  } else if (/插入一?行|加一?行|新增一?行/.test(text)) {
    const at = /表头下|第2行|数据前/.test(text) ? 1 : maxR + 1;
    actions.push({ cmd: 'insertRow', payload: { at } });
    reply = '已插入一行。';
  } else if (/空行/.test(text) && /删|去|清/.test(text)) {
    actions.push({ cmd: 'deleteBlankRows' });
    reply = '已删掉中间的空行，表头还在。';
  } else if (/分列|拆分/.test(text)) {
    const c = colFromText(text, snap.headers);
    const sep = /空格/.test(text) ? '空格' : /顿号|、/.test(text) ? '、' : /；|;/.test(text) ? ';' : ',';
    actions.push({ cmd: 'splitColumn', payload: { col: c >= 0 ? c : 0, sep } });
    reply = '已按分隔符把该列拆成多列。';
  } else if (/空格/.test(text) && /去|删|清|trim/i.test(text)) {
    actions.push({ cmd: 'trimAll' });
    reply = '已去掉格子里前后和多余的空格。';
  } else if (/只要|筛选|过滤/.test(text) && !/打开筛选|自动筛选/.test(text)) {
    const m = text.match(/只要\s*[「"']?(.+?)[」"']?\s*$/)
      || text.match(/(?:筛选|过滤)\s*(?:出)?\s*[「"']?([^\s，,。]+)/);
    const value = m ? m[1].replace(/的?行$/, '').trim() : '';
    if (!value) {
      actions.push({ cmd: 'filter' });
      reply = '已打开自动筛选。点表头右边的 ▾，勾你要看的项即可。';
    } else {
      const c = colContainingValue(snap, value);
      actions.push({ cmd: 'filterBy', payload: { col: c, value } });
      reply = `已只留下「${value}」相关的行，其它行藏起来了。`;
    }
  } else if (/表头/.test(text) && /加粗|黑体/.test(text)) {
    actions.push({ cmd: 'bold', payload: { r1: 0, c1: 0, r2: 0, c2: snap.maxC || 0 } });
    reply = '已把表头整行加粗。';
  } else if (/表头/.test(text) && /居中/.test(text)) {
    actions.push({ cmd: 'align', payload: { r1: 0, c1: 0, r2: 0, c2: snap.maxC || 0, align: 'center' } });
    reply = '已把表头居中。';
  } else if (/边框|框线/.test(text)) {
    actions.push({ cmd: 'border', payload: { r1: 0, c1: 0, r2: maxR, c2: snap.maxC || 0, border: 'all' } });
    reply = '已给当前表加上框线。';
  } else if (/(大于|小于|超过|低于)/.test(text) && /(标红|红色|高亮)/.test(text)) {
    const n = text.match(/(-?\d+(?:\.\d+)?)/);
    const op = /小于|低于/.test(text) ? 'lt' : 'gt';
    const c = colFromText(text, snap.headers);
    actions.push({ cmd: 'highlightIf', payload: { col: c >= 0 ? c : undefined, op, value: n ? Number(n[1]) : 0, color: '#dc2626' } });
    reply = `已把${op === 'lt' ? '小于' : '大于'} ${n ? n[1] : 0} 的数字标红。`;
  } else if (/汇总|按.+统计|分组/.test(text)) {
    const by = colByHeader(snap.headers, ['部门', '组别', '分类', '类别', '城市', 'city', '地区']) >= 0
      ? colByHeader(snap.headers, ['部门', '组别', '分类', '类别', '城市', 'city', '地区'])
      : 0;
    const val = numericCol(snap);
    actions.push({ cmd: 'groupSum', payload: { byCol: by, valueCol: val } });
    reply = `已按「${(snap.headers || [])[by] || '第1列'}」汇总「${(snap.headers || [])[val] || '数值列'}」，写在表下面。`;
  } else if (/下拉/.test(text)) {
    const opts = text.match(/选项[：:是]?\s*(.+)$/);
    const options = opts ? opts[1].split(/[,，、]/).map((s) => s.trim()).filter(Boolean) : ['是', '否', '待定'];
    const c = colFromText(text, snap.headers);
    actions.push({
      cmd: 'dropdown',
      payload: { r1: 1, c1: c >= 0 ? c : 0, r2: maxR, c2: c >= 0 ? c : 0, options }
    });
    reply = '已在该列加上下拉选项，点格子就能选，不用手打。';
  } else if (!COLOR_WORD.test(text) && (/把\s*[「"']?(.+?)[」"']?\s*(改成|换成|替换成)\s*[「"']?(.+?)[」"']?/.test(text) || /替换/.test(text))) {
    const m = text.match(/把\s*[「"']?([^「"']+?)[」"']?\s*(?:改成|换成|替换成)\s*[「"']?([^「"']+?)[」"']?\s*$/)
      || text.match(/替换\s*[「"']?([^「"']+?)[」"']?\s*(?:为|成)\s*[「"']?([^「"']+?)[」"']?\s*$/);
    if (m) {
      actions.push({ cmd: 'replace', payload: { find: m[1], replace: m[2] } });
      reply = `已把「${m[1]}」全部换成「${m[2]}」。`;
    } else {
      actions.push({ cmd: 'replace' });
      reply = '已打开替换窗口。';
    }
  } else if (/表头/.test(text) && /改名|改成|叫/.test(text) && !COLOR_WORD.test(text)) {
    const m = text.match(/[「"']?([A-Za-z0-9_\u4e00-\u9fa5]+)[」"']?\s*(?:改成|改名|叫)\s*[「"']?([A-Za-z0-9_\u4e00-\u9fa5]+)[」"']?/);
    if (m) {
      const c = colByHeader(snap.headers, [m[1]]);
      actions.push({ cmd: 'renameHeader', payload: { col: c >= 0 ? c : 0, to: m[2] } });
      reply = `已把表头改成「${m[2]}」。`;
    }
  } else if (/合并/.test(text) && /[A-Za-z]+\d+\s*[:：~到至-]\s*[A-Za-z]+\d+/.test(text)) {
    const m = text.match(/([A-Za-z]{1,3}\d+)\s*[:：~到至-]\s*([A-Za-z]{1,3}\d+)/);
    actions.push({ cmd: 'merge', payload: { range: `${m[1].toUpperCase()}:${m[2].toUpperCase()}` } });
    reply = `已合并 ${m[1].toUpperCase()} 到 ${m[2].toUpperCase()}。`;
  } else if ((/表头|首行|标题行/.test(text) || /行/.test(text)) && /标|涂|填|底色|背景|颜色|变成|改成/.test(text) && COLOR_WORD.test(text) && !/列/.test(text)) {
    const r = parseRowFromText(text);
    const color = parseFillColor(text);
    actions.push({ cmd: 'fillRow', payload: { row: r, color } });
    reply = r === 0
      ? '已把表头那一整行底色改好，不是改第一列。'
      : `已把第 ${r + 1} 行整行底色改好。`;
  } else if (/合并/.test(text) && /列/.test(text)) {
    const pair = parseTwoColumns(text, snap.headers);
    if (!pair) {
      reply = '请写清要合并哪两列，例如：把 D、E 两列合并。';
    } else {
      actions.push({ cmd: 'mergeColumns', payload: { c1: pair[0], c2: pair[1] } });
      const a = (snap.headers && snap.headers[pair[0]]) || `第${pair[0] + 1}列`;
      const b = (snap.headers && snap.headers[pair[1]]) || `第${pair[1] + 1}列`;
      reply = `已把 ${a} 和 ${b} 两列逐行合并成一格，内容拼在一起。`;
    }
  } else if (/列/.test(text) && /标|涂|填|底色|背景|颜色|变成|改成/.test(text) && COLOR_WORD.test(text)) {
    const c = namedColFromText(text, snap.headers);
    if (c < 0) {
      reply = '没找到你说的那一列。请用表头上的原名，例如：把 city 这列标成浅绿色。';
    } else {
      const color = parseFillColor(text);
      actions.push({ cmd: 'fillColumn', payload: { col: c, color } });
      reply = `已把「${snap.headers[c]}」整列（含表头和下面所有数据行）底色改成你要的颜色。`;
    }
  } else if (/求和|加总|合计|总共|一共/.test(text) && !/汇总/.test(text)) {
    const c = colFromText(text, snap.headers);
    const use = c >= 0 ? c : col;
    actions.push({ cmd: 'autosum', payload: { r1: 1, c1: use, r2: maxR, c2: use } });
    reply = `已对「${(snap.headers && snap.headers[use]) || headerName}」从第 2 行加总，合计写在下面。`;
  } else if (/平均/.test(text)) {
    const c = colFromText(text, snap.headers);
    const use = c >= 0 ? c : col;
    actions.push({ cmd: 'average', payload: { r1: 1, c1: use, r2: maxR, c2: use } });
    reply = `已计算「${(snap.headers && snap.headers[use]) || headerName}」的平均值。`;
  } else if (/最大/.test(text)) {
    const c = colFromText(text, snap.headers);
    const use = c >= 0 ? c : col;
    actions.push({ cmd: 'max', payload: { r1: 1, c1: use, r2: maxR, c2: use } });
    reply = `已标出「${(snap.headers && snap.headers[use]) || headerName}」的最大值。`;
  } else if (/最小/.test(text)) {
    const c = colFromText(text, snap.headers);
    const use = c >= 0 ? c : col;
    actions.push({ cmd: 'min', payload: { r1: 1, c1: use, r2: maxR, c2: use } });
    reply = `已标出「${(snap.headers && snap.headers[use]) || headerName}」的最小值。`;
  } else if (/降序|从高到低|从大到小/.test(text)) {
    const sortCol = colFromText(text, snap.headers);
    const use = sortCol >= 0 ? sortCol : col;
    actions.push({ cmd: 'sort', payload: { col: use, dir: 'desc' } });
    reply = `已按「${(snap.headers && snap.headers[use]) || headerName}」从大到小排序，整行一起动。`;
  } else if (/升序|从低到高|从小到大|排序/.test(text)) {
    const sortCol = colFromText(text, snap.headers);
    const use = sortCol >= 0 ? sortCol : 0;
    actions.push({ cmd: 'sort', payload: { col: use, dir: 'asc' } });
    reply = `已按「${(snap.headers && snap.headers[use]) || '第1列'}」排序，整行一起动。`;
  } else if (/去重|重复/.test(text)) {
    actions.push(
      { cmd: 'select', payload: { r1: 1, c1: 0, r2: maxR, c2: snap.maxC || 0 } },
      { cmd: 'unique' }
    );
    reply = '已删除完全重复的行。';
  } else if (/冻结|表头不要动|滚/.test(text)) {
    actions.push({ cmd: 'freezeTop' });
    reply = '已冻结首行，往下滚动时表头还在。';
  } else if (/负数|异常|标红/.test(text)) {
    actions.push({ cmd: 'highlightNegatives' });
    reply = '已把表里的负数标成红色。';
  } else if (/千分位|货币|人民币|百分比/.test(text)) {
    const fmt = /百分/.test(text) ? 'percent' : /千分/.test(text) ? 'thousand' : 'currency';
    actions.push(
      { cmd: 'select', payload: { r1: 1, c1: col, r2: maxR, c2: col } },
      { cmd: 'numFmt', payload: fmt }
    );
    reply = `已把「${headerName}」改成${fmt === 'percent' ? '百分比' : fmt === 'thousand' ? '千分位' : '人民币'}格式。`;
  } else if (/勾选|打勾/.test(text)) {
    actions.push({ cmd: 'checkbox' });
    reply = '已在当前选区插入勾选框，点一下就能打勾。';
  } else if (/是\s*否|是否/.test(text)) {
    actions.push({ cmd: 'yesno' });
    reply = '已插入「是 / 否」，点其中一个即可。';
  } else if (/多选/.test(text)) {
    actions.push({ cmd: 'multiselect' });
    reply = '请填写多选项，可以同时勾几个。';
  } else if (/单选/.test(text)) {
    actions.push({ cmd: 'radio' });
    reply = '请选择一组格子再插入单选圈，一组只能选一个。';
  } else if (/查找|搜/.test(text)) {
    const m = text.match(/查找\s*[「"]?(.+?)[」"]?$/) || text.match(/搜(索)?\s*(.+)$/);
    const q = m ? m[m.length - 1] : '';
    if (q) {
      actions.push({ cmd: 'find', payload: { find: q.replace(/[「」]/g, '') } });
      reply = `正在查找「${q}」。`;
    } else {
      actions.push({ cmd: 'find' });
      reply = '已打开查找，输入要找的字即可。';
    }
  } else if (/图|柱状|折线|饼图/.test(text)) {
    actions.push(
      { cmd: 'select', payload: { r1: 0, c1: Math.max(0, col - 1), r2: maxR, c2: col } },
      { cmd: 'chart', payload: /折/.test(text) ? 'line' : 'bar' }
    );
    reply = '已按当前数据画图。';
  } else if (/美化|排版|好看|规整|整理|专业/.test(text)) {
    actions.push({ cmd: 'beautify' });
    reply = '已整理表头颜色、隔行底纹，并加上合计行。';
  } else {
    actions.push({ cmd: 'beautify' });
    reply = '我先按「整理报表」做了一步。你可以直接说：只要销售部、按部门汇总工资、插入序号、去掉空格、大于100标红、把A1写成合计。';
    return { reply, actions, matched: false };
  }

  return { reply, actions, matched: true };
}
