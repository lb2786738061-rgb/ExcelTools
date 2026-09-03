import * as XLSX from 'xlsx';
import { a1, colName, evaluateCellFormula, letterToIndex, parseA1 } from './formulas.js';
import { applyWatermarkOverlay, composeWatermark, maskSensitiveText, resolveWatermarkText } from '../engine/watermarkEngine.js';

function emptyStyle() {
  return {
    bold: false,
    italic: false,
    underline: false,
    strike: false,
    fontFamily: '',
    fontSize: '',
    color: '',
    fill: '',
    align: '',
    vAlign: 'middle',
    wrap: false,
    border: '',
    numFmt: 'general'
  };
}

function formatValue(cell) {
  if (!cell) return '';
  if (cell.f) {
    const v = cell.computed;
    if (v == null) return '';
    return formatNumber(v, cell.s?.numFmt);
  }
  return formatNumber(cell.v, cell.s?.numFmt);
}

function formatNumber(v, fmt) {
  if (v == null || v === '') return '';
  if (typeof v !== 'number' || !Number.isFinite(v) || !fmt || fmt === 'general') return String(v);
  if (fmt === 'thousand') return v.toLocaleString('en-US', { maximumFractionDigits: 2 });
  if (fmt === 'currency') return `¥${v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  if (fmt === 'dollar') return `$${v.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  if (fmt === 'percent') return `${(v * (Math.abs(v) <= 1 ? 100 : 1)).toFixed(1)}%`;
  if (fmt === 'int') return String(Math.round(v));
  return String(v);
}

export function createExcelSheet() {
  const root = document.createElement('div');
  root.className = 'xl-root';
  root.innerHTML = `
    <div class="xl-formula-bar">
      <span class="xl-name-box" id="xlName">A1</span>
      <span class="xl-fx">fx</span>
      <input class="xl-formula-input" id="xlFormula" placeholder="输入数值或公式，例如 =SUM(B2:B7)" />
    </div>
    <div class="xl-scroll" id="xlScroll">
      <table class="xl-grid" id="xlGrid"></table>
    </div>
    <div class="xl-status" id="xlStatus">就绪 · 单击选中，双击编辑，公式以 = 开头</div>
    <div class="xl-dialog" id="xlFindDlg" hidden>
      <div class="xl-dlg-head"><strong id="xlFindTitle">查找</strong><button type="button" class="xl-dlg-x" id="xlFindClose">×</button></div>
      <label class="xl-dlg-row">查找内容<input id="xlFindWhat" /></label>
      <label class="xl-dlg-row" id="xlReplaceRow">替换为<input id="xlReplaceWith" /></label>
      <label class="xl-dlg-check"><input type="checkbox" id="xlFindCase" /> 区分大小写</label>
      <label class="xl-dlg-check"><input type="checkbox" id="xlFindWild" /> 通配符 * ?</label>
      <div class="xl-dlg-actions">
        <button type="button" id="xlFindNext">查找下一个</button>
        <button type="button" id="xlReplaceOne">替换</button>
        <button type="button" id="xlReplaceAll">全部替换</button>
      </div>
    </div>
    <div class="xl-filter-pop" id="xlFilterPop" hidden></div>
    <div class="watermark-overlay xl-wm" id="xlWm" hidden></div>
  `;

  const gridEl = root.querySelector('#xlGrid');
  const nameBox = root.querySelector('#xlName');
  const formulaInput = root.querySelector('#xlFormula');
  const statusEl = root.querySelector('#xlStatus');

  let rows = 40;
  let cols = 16;
  const cells = new Map();
  const merges = new Map();
  let sel = { r1: 0, c1: 0, r2: 0, c2: 0 };
  let active = { r: 0, c: 0 };
  let editing = false;
  let undo = [];
  let chartType = 'none';
  let freezeRows = 0;
  let freezeCols = 0;
  let filterEnabled = false;
  const filters = {};
  let findCursor = -1;

  const findDlg = root.querySelector('#xlFindDlg');
  const findWhat = root.querySelector('#xlFindWhat');
  const replaceWith = root.querySelector('#xlReplaceWith');
  const replaceRow = root.querySelector('#xlReplaceRow');
  const findCase = root.querySelector('#xlFindCase');
  const findWild = root.querySelector('#xlFindWild');
  const filterPop = root.querySelector('#xlFilterPop');
  const scrollEl = root.querySelector('#xlScroll');
  const wmEl = root.querySelector('#xlWm');
  let watermarkKey = '';
  let watermarkOpacity = 0.22;

  function key(r, c) { return a1(r, c); }

  function getCell(r, c) {
    return cells.get(key(r, c)) || { v: '', f: '', s: emptyStyle() };
  }

  function setCell(r, c, patch) {
    const cur = { ...getCell(r, c), s: { ...emptyStyle(), ...getCell(r, c).s } };
    const next = { ...cur, ...patch };
    if (patch.s) next.s = { ...cur.s, ...patch.s };
    if (patch.ctrl === null) delete next.ctrl;
    if (patch.slash === null) delete next.slash;
    const hasCtrl = !!next.ctrl;
    const hasSlash = !!next.slash;
    const hasVal = next.f || (next.v !== '' && next.v != null && next.v !== false);
    const hasStyle = Object.entries(next.s).some(([k, v]) => v && v !== 'middle' && v !== 'general');
    if (!hasCtrl && !hasSlash && !hasVal && !hasStyle) {
      cells.delete(key(r, c));
    } else {
      cells.set(key(r, c), next);
    }
  }

  function snapshot() {
    undo.push(JSON.stringify({
      cells: [...cells.entries()],
      merges: [...merges.entries()]
    }));
    if (undo.length > 40) undo.shift();
  }

  function restoreSnapshot(prev) {
    const data = JSON.parse(prev);
    cells.clear();
    merges.clear();
    if (Array.isArray(data)) {
      data.forEach(([k, v]) => cells.set(k, v));
      return;
    }
    (data.cells || []).forEach(([k, v]) => cells.set(k, v));
    (data.merges || []).forEach(([k, v]) => merges.set(k, v));
  }

  function mergeKey(r, c) {
    return `${r},${c}`;
  }

  function isCovered(r, c) {
    for (const [k, m] of merges) {
      const [sr, sc] = k.split(',').map(Number);
      if (r >= sr && r < sr + m.rs && c >= sc && c < sc + m.cs && (r !== sr || c !== sc)) return true;
    }
    return false;
  }

  function eachSel(fn) {
    const r1 = Math.min(sel.r1, sel.r2);
    const r2 = Math.max(sel.r1, sel.r2);
    const c1 = Math.min(sel.c1, sel.c2);
    const c2 = Math.max(sel.c1, sel.c2);
    for (let r = r1; r <= r2; r++) {
      for (let c = c1; c <= c2; c++) fn(r, c);
    }
    return { r1, r2, c1, c2 };
  }

  function getRaw(addr) {
    const p = parseA1(addr);
    if (!p) return '';
    return getCell(p.r, p.c);
  }

  function recalc() {
    for (const [k, cell] of cells) {
      if (cell.f) {
        cell.computed = evaluateCellFormula(cell.f, getRaw);
      }
    }
  }

  function loadAoA(aoa, usedName = 'Sheet1') {
    cells.clear();
    merges.clear();
    undo = [];
    let maxR = 1;
    let maxC = 1;
    (aoa || []).forEach((row, r) => {
      (row || []).forEach((v, c) => {
        if (v === '' || v == null) return;
        const str = String(v);
        if (str.startsWith('=')) setCell(r, c, { f: str, v: '', s: emptyStyle() });
        else setCell(r, c, { v, f: '', s: emptyStyle() });
        maxR = Math.max(maxR, r + 1);
        maxC = Math.max(maxC, c + 1);
      });
    });
    rows = Math.max(30, maxR + 12);
    cols = Math.max(12, maxC + 4);
    sel = { r1: 0, c1: 0, r2: 0, c2: 0 };
    active = { r: 0, c: 0 };
    freezeRows = 0;
    freezeCols = 0;
    filterEnabled = false;
    Object.keys(filters).forEach((k) => delete filters[k]);
    statusEl.textContent = `${usedName} · ${maxR} 行 × ${maxC} 列 · 可直接编辑`;
    recalc();
    render();
  }

  function toAoA() {
    recalc();
    const aoa = [];
    let maxR = 0;
    let maxC = 0;
    for (const [k, cell] of cells) {
      const p = parseA1(k);
      if (!p) continue;
      maxR = Math.max(maxR, p.r);
      maxC = Math.max(maxC, p.c);
    }
    for (let r = 0; r <= maxR; r++) {
      const row = [];
      for (let c = 0; c <= maxC; c++) {
        const cell = getCell(r, c);
        row.push(cell.f ? (cell.computed ?? '') : (cell.v ?? ''));
      }
      aoa.push(row);
    }
    return aoa;
  }

  function exportXlsx(filename = '工作簿.xlsx') {
    recalc();
    const aoa = [];
    const formulas = [];
    let maxR = 0;
    let maxC = 0;
    for (const [k] of cells) {
      const p = parseA1(k);
      if (!p) continue;
      maxR = Math.max(maxR, p.r);
      maxC = Math.max(maxC, p.c);
    }
    for (let r = 0; r <= maxR; r++) {
      const row = [];
      for (let c = 0; c <= maxC; c++) {
        const cell = getCell(r, c);
        if (cell.f) {
          row.push(cell.computed ?? 0);
          formulas.push({ r, c, f: cell.f.replace(/^=/, ''), v: cell.computed });
        } else row.push(cell.v ?? '');
      }
      aoa.push(row);
    }
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    formulas.forEach((item) => {
      ws[XLSX.utils.encode_cell({ r: item.r, c: item.c })] = { t: 'n', v: Number(item.v) || 0, f: item.f };
    });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    XLSX.writeFile(wb, filename, { bookType: 'xlsx' });
  }

  function dataMax() {
    let maxR = 0;
    let maxC = 0;
    for (const [k] of cells) {
      const p = parseA1(k);
      if (!p) continue;
      maxR = Math.max(maxR, p.r);
      maxC = Math.max(maxC, p.c);
    }
    return { maxR, maxC };
  }

  function cellText(r, c) {
    const cell = getCell(r, c);
    if (cell.slash) return `${cell.slash.bottom || ''}/${cell.slash.top || ''}`;
    if (cell.f) return String(cell.computed ?? '');
    if (cell.v == null) return '';
    return String(cell.v);
  }

  function slashEditValue(cell) {
    if (!cell?.slash) return cell?.f || (cell?.v ?? '');
    return `${cell.slash.top || ''}/${cell.slash.bottom || ''}`;
  }

  function rowPassesFilter(r) {
    if (!filterEnabled || r === 0) return true;
    for (const [cStr, allowed] of Object.entries(filters)) {
      if (!allowed) continue;
      const token = cellText(r, Number(cStr)) || '(空白)';
      if (allowed.has(token)) continue;
      let hit = false;
      for (const a of allowed) {
        if (a && a !== '(空白)' && (token.includes(a) || a.includes(token))) { hit = true; break; }
      }
      if (!hit) return false;
    }
    return true;
  }

  function styleAttr(s) {
    const parts = [];
    if (s.bold) parts.push('font-weight:700');
    if (s.italic) parts.push('font-style:italic');
    const dec = [];
    if (s.underline) dec.push('underline');
    if (s.strike) dec.push('line-through');
    if (dec.length) parts.push(`text-decoration:${dec.join(' ')}`);
    if (s.fontFamily) parts.push(`font-family:${s.fontFamily}`);
    if (s.fontSize) parts.push(`font-size:${s.fontSize}px`);
    if (s.color) parts.push(`color:${s.color}`);
    if (s.fill) parts.push(`background:${s.fill}`);
    if (s.align) parts.push(`text-align:${s.align}`);
    parts.push(`vertical-align:${s.vAlign || 'middle'}`);
    if (s.wrap) parts.push('white-space:normal');
    if (s.border === 'none') parts.push('border:none');
    if (s.border === 'all') parts.push('border:1px solid #94a3b8');
    return parts.join(';');
  }

  function render() {
    recalc();
    const r1 = Math.min(sel.r1, sel.r2);
    const r2 = Math.max(sel.r1, sel.r2);
    const c1 = Math.min(sel.c1, sel.c2);
    const c2 = Math.max(sel.c1, sel.c2);

    const HEAD_H = 26;
    const ROW_W = 40;
    const COL_W = 88;
    let html = '<thead><tr><th class="xl-corner"></th>';
    for (let c = 0; c < cols; c++) {
      const freezeLeft = c < freezeCols ? `left:${ROW_W + c * COL_W}px;` : '';
      html += `<th class="xl-colh ${c >= c1 && c <= c2 ? 'hl' : ''} ${c < freezeCols ? 'xl-frozen-col' : ''}" style="${freezeLeft}">${colName(c)}</th>`;
    }
    html += '</tr></thead><tbody>';
    for (let r = 0; r < rows; r++) {
      const hidden = !rowPassesFilter(r);
      let rowHasSlash = false;
      for (let c = 0; c < cols; c++) {
        if (getCell(r, c).slash) { rowHasSlash = true; break; }
      }
      const top = r < freezeRows ? `top:${HEAD_H + r * 24}px;` : '';
      html += `<tr class="${r < freezeRows ? 'xl-frozen-row' : ''} ${hidden ? 'xl-filtered' : ''} ${rowHasSlash ? 'xl-slash-row' : ''}" style="${top}"><th class="xl-rowh ${r >= r1 && r <= r2 ? 'hl' : ''}">${r + 1}</th>`;
      for (let c = 0; c < cols; c++) {
        if (isCovered(r, c)) continue;
        const cell = getCell(r, c);
        const selected = r >= r1 && r <= r2 && c >= c1 && c <= c2;
        const act = r === active.r && c === active.c;
        const left = c < freezeCols ? `left:${ROW_W + c * COL_W}px;` : '';
        const frozen = (r < freezeRows || c < freezeCols) ? ' xl-frozen-cell' : '';
        const filterBtn = (filterEnabled && r === 0 && !cell.slash)
          ? `<button type="button" class="xl-filter-btn ${filters[c] ? 'on' : ''}" data-filter-col="${c}">▾</button>`
          : '';
        const span = merges.get(mergeKey(r, c));
        const spanAttr = span ? ` colspan="${span.cs}" rowspan="${span.rs}"` : '';
        html += `<td${spanAttr} class="xl-cell${selected ? ' sel' : ''}${act ? ' act' : ''}${cell.ctrl ? ' has-ctrl' : ''}${cell.slash ? ' xl-slash-cell' : ''}${frozen}${c < freezeCols ? ' xl-frozen-col' : ''}${span ? ' xl-merged' : ''}" data-r="${r}" data-c="${c}" style="${styleAttr(cell.s)};${left}">${cellInner(cell)}${filterBtn}</td>`;
      }
      html += '</tr>';
    }
    html += '</tbody>';
    gridEl.innerHTML = html;
    const oldChart = root.querySelector('.xl-chart');
    if (oldChart) oldChart.remove();
    if (chartType !== 'none') {
      const holder = document.createElement('div');
      holder.innerHTML = buildChart();
      root.querySelector('#xlScroll').appendChild(holder.firstChild);
    }

    nameBox.textContent = a1(active.r, active.c);
    const ac = getCell(active.r, active.c);
    formulaInput.value = slashEditValue(ac);
    bindCells();
    bindFilterButtons();
  }

  function buildChart() {
    const { r1, r2, c1, c2 } = eachSel(() => {});
    const labels = [];
    const values = [];
    for (let r = r1; r <= r2; r++) {
      labels.push(String(getCell(r, c1).v ?? a1(r, c1)));
      const cell = getCell(r, Math.max(c1, c2));
      const n = Number(cell.f ? cell.computed : cell.v);
      values.push(Number.isFinite(n) ? n : 0);
    }
    const max = Math.max(...values, 1);
    const w = 480;
    const h = 160;
    const bars = values.map((v, i) => {
      const bh = (v / max) * 120;
      const x = 30 + i * ((w - 40) / values.length);
      const bw = Math.max(8, (w - 40) / values.length * 0.7);
      return `<rect x="${x}" y="${140 - bh}" width="${bw}" height="${bh}" fill="#217346"/>`;
    }).join('');
    return `<div class="xl-chart"><svg viewBox="0 0 ${w} ${h}" width="100%" height="${h}">${bars}</svg></div>`;
  }

  function escapeHtml(s) {
    return String(s)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  function isOn(v) {
    return v === true || v === 'TRUE' || v === '是' || v === '✓' || v === 1 || v === '1';
  }

  function cellInner(cell) {
    if (cell.slash) {
      return `<div class="xl-slash"><span class="xl-slash-top">${escapeHtml(cell.slash.top || '')}</span><span class="xl-slash-bot">${escapeHtml(cell.slash.bottom || '')}</span></div>`;
    }
    const ctrl = cell.ctrl;
    if (!ctrl) return escapeHtml(formatValue(cell));
    if (ctrl.type === 'checkbox' || ctrl.type === 'radio') {
      const on = isOn(cell.v);
      const mark = ctrl.type === 'radio' ? (on ? '●' : '○') : (on ? '☑' : '☐');
      return `<span class="xl-check ${on ? 'on' : ''}" data-ctrl="${ctrl.type}">${mark}</span>`;
    }
    if (ctrl.type === 'yesno') {
      const v = String(cell.v || '');
      return `<span class="xl-yesno">
        <button type="button" class="xl-opt ${v === '是' ? 'on' : ''}" data-ctrl="yesno" data-val="是">是</button>
        <button type="button" class="xl-opt ${v === '否' ? 'on' : ''}" data-ctrl="yesno" data-val="否">否</button>
      </span>`;
    }
    if (ctrl.type === 'dropdown') {
      const opts = ctrl.options || ['是', '否'];
      const cur = String(cell.v ?? '');
      return `<select class="xl-select" data-ctrl="dropdown">${opts.map((o) =>
        `<option value="${escapeHtml(o)}" ${cur === o ? 'selected' : ''}>${escapeHtml(o)}</option>`).join('')}</select>`;
    }
    if (ctrl.type === 'multiselect') {
      const selected = String(cell.v || '').split(/[,，]/).map((s) => s.trim()).filter(Boolean);
      const opts = ctrl.options || [];
      return `<div class="xl-multi">${opts.map((o) => {
        const on = selected.includes(o);
        return `<label class="xl-multi-item"><input type="checkbox" data-ctrl="multi" data-val="${escapeHtml(o)}" ${on ? 'checked' : ''}/> ${escapeHtml(o)}</label>`;
      }).join('')}</div>`;
    }
    return escapeHtml(formatValue(cell));
  }

  function paintSelection() {
    const r1 = Math.min(sel.r1, sel.r2);
    const r2 = Math.max(sel.r1, sel.r2);
    const c1 = Math.min(sel.c1, sel.c2);
    const c2 = Math.max(sel.c1, sel.c2);
    gridEl.querySelectorAll('.xl-cell').forEach((td) => {
      const r = Number(td.dataset.r);
      const c = Number(td.dataset.c);
      td.classList.toggle('sel', r >= r1 && r <= r2 && c >= c1 && c <= c2);
      td.classList.toggle('act', r === active.r && c === active.c);
    });
    nameBox.textContent = a1(active.r, active.c);
    const ac = getCell(active.r, active.c);
    formulaInput.value = slashEditValue(ac);
  }

  function bindCells() {
    gridEl.querySelectorAll('.xl-cell').forEach((td) => {
      td.addEventListener('mousedown', (e) => {
        if (editing) commitEdit();
        const r = Number(td.dataset.r);
        const c = Number(td.dataset.c);
        active = { r, c };
        sel = e.shiftKey
          ? { ...sel, r2: r, c2: c }
          : { r1: r, c1: c, r2: r, c2: c };
        paintSelection();
        const move = (ev) => {
          const t = ev.target.closest?.('.xl-cell') || document.elementFromPoint(ev.clientX, ev.clientY)?.closest('.xl-cell');
          if (!t) return;
          sel.r2 = Number(t.dataset.r);
          sel.c2 = Number(t.dataset.c);
          paintSelection();
        };
        const up = () => {
          window.removeEventListener('mousemove', move);
          window.removeEventListener('mouseup', up);
        };
        window.addEventListener('mousemove', move);
        window.addEventListener('mouseup', up);
      });
      td.addEventListener('dblclick', (e) => {
        if (getCell(Number(td.dataset.r), Number(td.dataset.c)).ctrl) {
          e.preventDefault();
          return;
        }
        startEdit();
      });
    });

    gridEl.querySelectorAll('[data-ctrl]').forEach((el) => {
      el.addEventListener('mousedown', (e) => e.stopPropagation());
      const apply = (e) => {
        e.stopPropagation();
        const td = el.closest('.xl-cell');
        if (!td) return;
        const r = Number(td.dataset.r);
        const c = Number(td.dataset.c);
        const type = el.dataset.ctrl;
        snapshot();
        if (type === 'checkbox') {
          const cell = getCell(r, c);
          setCell(r, c, { v: !isOn(cell.v), ctrl: cell.ctrl });
        } else if (type === 'radio') {
          const cell = getCell(r, c);
          const gid = cell.ctrl?.groupId;
          if (gid) {
            for (const [k, other] of cells) {
              if (other.ctrl?.type === 'radio' && other.ctrl.groupId === gid) {
                const p = parseA1(k);
                if (p) setCell(p.r, p.c, { v: false, ctrl: other.ctrl });
              }
            }
          }
          setCell(r, c, { v: true, ctrl: cell.ctrl });
        } else if (type === 'yesno') {
          const cell = getCell(r, c);
          setCell(r, c, { v: el.dataset.val, ctrl: cell.ctrl });
        } else if (type === 'dropdown') {
          const cell = getCell(r, c);
          setCell(r, c, { v: el.value, ctrl: cell.ctrl });
        } else if (type === 'multi') {
          const cell = getCell(r, c);
          const picked = String(cell.v || '').split(/[,，]/).map((s) => s.trim()).filter(Boolean);
          const val = el.dataset.val;
          const next = el.checked ? [...new Set([...picked, val])] : picked.filter((x) => x !== val);
          setCell(r, c, { v: next.join('，'), ctrl: cell.ctrl });
        }
        render();
      };
      if (el.tagName === 'SELECT') el.addEventListener('change', apply);
      else if (el.type === 'checkbox') el.addEventListener('change', apply);
      else el.addEventListener('click', apply);
    });
  }

  function bindFilterButtons() {
    gridEl.querySelectorAll('.xl-filter-btn').forEach((btn) => {
      btn.addEventListener('mousedown', (e) => e.stopPropagation());
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        openFilterPopup(Number(btn.dataset.filterCol), btn);
      });
    });
  }

  function openFilterPopup(col, anchor) {
    const { maxR } = dataMax();
    const counts = new Map();
    for (let r = 1; r <= maxR; r++) {
      let ok = true;
      for (const [cStr, allowed] of Object.entries(filters)) {
        if (Number(cStr) === col || !allowed) continue;
        const token = cellText(r, Number(cStr)) || '(空白)';
        if (!allowed.has(token)) { ok = false; break; }
      }
      if (!ok) continue;
      const token = cellText(r, col) || '(空白)';
      counts.set(token, (counts.get(token) || 0) + 1);
    }
    const values = [...counts.keys()].sort((a, b) => a.localeCompare(b, 'zh'));
    const selected = filters[col] ? new Set(filters[col]) : new Set(values);
    filterPop.innerHTML = `
      <div class="xl-filter-title">${colName(col)} 列筛选</div>
      <label class="xl-filter-all"><input type="checkbox" id="xlFilterAll" ${selected.size === values.length ? 'checked' : ''}/> 全选</label>
      <div class="xl-filter-list">${values.map((v) => `
        <label><input type="checkbox" data-fv="${escapeHtml(v)}" ${selected.has(v) ? 'checked' : ''}/> ${escapeHtml(v)} <span class="muted">(${counts.get(v)})</span></label>
      `).join('')}</div>
      <div class="xl-dlg-actions">
        <button type="button" id="xlFilterOk">确定</button>
        <button type="button" id="xlFilterClear">清除此列</button>
      </div>
    `;
    filterPop.hidden = false;
    const rect = anchor.getBoundingClientRect();
    const host = root.getBoundingClientRect();
    filterPop.style.left = `${Math.min(Math.max(8, rect.left - host.left), host.width - 230)}px`;
    filterPop.style.top = `${rect.bottom - host.top + 4}px`;

    const allBox = filterPop.querySelector('#xlFilterAll');
    allBox.addEventListener('change', () => {
      filterPop.querySelectorAll('input[data-fv]').forEach((box) => { box.checked = allBox.checked; });
    });
    filterPop.querySelector('#xlFilterOk').addEventListener('click', () => {
      const picked = new Set();
      filterPop.querySelectorAll('input[data-fv]:checked').forEach((box) => picked.add(box.dataset.fv));
      if (picked.size === 0 || picked.size === values.length) delete filters[col];
      else filters[col] = picked;
      filterPop.hidden = true;
      render();
    });
    filterPop.querySelector('#xlFilterClear').addEventListener('click', () => {
      delete filters[col];
      filterPop.hidden = true;
      render();
    });
  }

  function wildcardToRegExp(pattern, caseSensitive) {
    const src = caseSensitive ? pattern : pattern.toLowerCase();
    const escaped = src.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*').replace(/\?/g, '.');
    return new RegExp(escaped);
  }

  function textMatches(text, pattern) {
    if (!pattern) return false;
    const cs = findCase.checked;
    const t = cs ? text : text.toLowerCase();
    const p = cs ? pattern : pattern.toLowerCase();
    if (findWild.checked && /[*?]/.test(pattern)) return wildcardToRegExp(pattern, cs).test(t);
    return t.includes(p);
  }

  function collectHits(pattern) {
    const { maxR, maxC } = dataMax();
    const hits = [];
    for (let r = 0; r <= maxR; r++) {
      if (!rowPassesFilter(r)) continue;
      for (let c = 0; c <= maxC; c++) {
        if (textMatches(cellText(r, c), pattern)) hits.push({ r, c });
      }
    }
    return hits;
  }

  function openFind(replaceMode) {
    findDlg.hidden = false;
    root.querySelector('#xlFindTitle').textContent = replaceMode ? '查找和替换' : '查找';
    replaceRow.style.display = replaceMode ? 'flex' : 'none';
    root.querySelector('#xlReplaceOne').hidden = !replaceMode;
    root.querySelector('#xlReplaceAll').hidden = !replaceMode;
    findWhat.focus();
    findWhat.select();
  }

  function findNext() {
    const pattern = findWhat.value;
    if (!pattern) return;
    const hits = collectHits(pattern);
    if (!hits.length) {
      statusEl.textContent = '找不到指定内容';
      return;
    }
    const cur = hits.findIndex((h) => h.r > active.r || (h.r === active.r && h.c > active.c));
    findCursor = cur === -1 ? 0 : cur;
    const hit = hits[findCursor];
    active = { r: hit.r, c: hit.c };
    sel = { r1: hit.r, c1: hit.c, r2: hit.r, c2: hit.c };
    render();
    statusEl.textContent = `找到 ${hits.length} 处 · 当前第 ${findCursor + 1} 处`;
    const td = gridEl.querySelector(`.xl-cell[data-r="${hit.r}"][data-c="${hit.c}"]`);
    td?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
  }

  function replaceOne() {
    const pattern = findWhat.value;
    if (!pattern) return;
    if (textMatches(cellText(active.r, active.c), pattern)) {
      snapshot();
      const cell = getCell(active.r, active.c);
      const next = replaceInText(cellText(active.r, active.c), pattern, replaceWith.value);
      setCell(active.r, active.c, { f: '', v: next, ctrl: cell.ctrl || null });
    }
    findNext();
  }

  function replaceAll() {
    const pattern = findWhat.value;
    if (!pattern) return;
    snapshot();
    const hits = collectHits(pattern);
    hits.forEach(({ r, c }) => {
      const cell = getCell(r, c);
      const next = replaceInText(cellText(r, c), pattern, replaceWith.value);
      setCell(r, c, { f: '', v: next, ctrl: cell.ctrl || null });
    });
    render();
    statusEl.textContent = `已替换 ${hits.length} 处`;
  }

  function replaceInText(text, pattern, replacement) {
    const cs = findCase.checked;
    if (findWild.checked && /[*?]/.test(pattern)) {
      return text.replace(wildcardToRegExp(pattern, cs), replacement);
    }
    if (cs) return text.split(pattern).join(replacement);
    const re = new RegExp(pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&'), 'gi');
    return text.replace(re, replacement);
  }

  function startEdit() {
    editing = true;
    const td = gridEl.querySelector(`.xl-cell[data-r="${active.r}"][data-c="${active.c}"]`);
    if (!td) return;
    const cell = getCell(active.r, active.c);
    td.innerHTML = '';
    const input = document.createElement('input');
    input.className = 'xl-edit';
    input.value = slashEditValue(cell);
    td.appendChild(input);
    input.focus();
    input.select();
    input.addEventListener('keydown', (e) => {
      if (e.key === 'Enter') { e.preventDefault(); commitEdit(input.value); move(1, 0); }
      if (e.key === 'Escape') { editing = false; render(); }
      if (e.key === 'Tab') { e.preventDefault(); commitEdit(input.value); move(0, 1); }
    });
    input.addEventListener('blur', () => { if (editing) commitEdit(input.value); });
  }

  function commitEdit(raw) {
    if (!editing && raw == null) return;
    editing = false;
    const value = raw ?? formulaInput.value;
    snapshot();
    if (String(value).startsWith('=')) setCell(active.r, active.c, { f: String(value), v: '', slash: null });
    else {
      const prev = getCell(active.r, active.c);
      const slashParts = String(value).split('/');
      if (prev.slash && slashParts.length >= 2) {
        const top = slashParts[0].trim();
        const bottom = slashParts.slice(1).join('/').trim();
        setCell(active.r, active.c, {
          f: '',
          v: `${bottom}/${top}`,
          slash: { top, bottom }
        });
      } else {
        const n = Number(value);
        setCell(active.r, active.c, {
          f: '',
          slash: null,
          v: value !== '' && Number.isFinite(n) && String(value).trim() !== '' ? n : value
        });
      }
    }
    render();
  }

  function move(dr, dc) {
    active = { r: Math.max(0, Math.min(rows - 1, active.r + dr)), c: Math.max(0, Math.min(cols - 1, active.c + dc)) };
    sel = { r1: active.r, c1: active.c, r2: active.r, c2: active.c };
    render();
  }

  formulaInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') commitEdit(formulaInput.value);
  });

  root.tabIndex = 0;
  root.addEventListener('keydown', (e) => {
    if (editing) return;
    if (e.target === formulaInput) return;
    if (e.target.closest('.xl-dialog, .xl-filter-pop')) return;
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'f') { e.preventDefault(); command('find'); return; }
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'h') { e.preventDefault(); command('replace'); return; }
    if (e.ctrlKey && e.key.toLowerCase() === 'z') { e.preventDefault(); command('undo'); return; }
    if (e.ctrlKey && e.key.toLowerCase() === 'b') { e.preventDefault(); command('bold'); return; }
    if (e.ctrlKey && e.key.toLowerCase() === 'c') { e.preventDefault(); command('copy'); return; }
    if (e.ctrlKey && e.key.toLowerCase() === 'v') { e.preventDefault(); command('paste'); return; }
    if (e.key === 'Delete' || e.key === 'Backspace') { e.preventDefault(); command('clear'); return; }
    if (e.key === 'F2') { e.preventDefault(); startEdit(); return; }
    if (e.key === 'Enter') { move(1, 0); return; }
    if (e.key === 'Tab') { e.preventDefault(); move(0, e.shiftKey ? -1 : 1); return; }
    if (e.key === 'ArrowUp') { e.preventDefault(); move(-1, 0); }
    if (e.key === 'ArrowDown') { e.preventDefault(); move(1, 0); }
    if (e.key === 'ArrowLeft') { e.preventDefault(); move(0, -1); }
    if (e.key === 'ArrowRight') { e.preventDefault(); move(0, 1); }
    if (e.key.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) {
      startEdit();
      const input = gridEl.querySelector('.xl-edit');
      if (input) input.value = e.key;
    }
  });

  function toggleStyle(prop) {
    snapshot();
    const first = getCell(Math.min(sel.r1, sel.r2), Math.min(sel.c1, sel.c2)).s[prop];
    eachSel((r, c) => setCell(r, c, { s: { [prop]: !first } }));
    render();
  }

  function applyStyle(patch) {
    snapshot();
    eachSel((r, c) => setCell(r, c, { s: patch }));
    render();
  }

  function clearMergesIn(r1, c1, r2, c2) {
    for (const [k, m] of [...merges]) {
      const [sr, sc] = k.split(',').map(Number);
      const er = sr + m.rs - 1;
      const ec = sc + m.cs - 1;
      if (sr <= r2 && er >= r1 && sc <= c2 && ec >= c1) merges.delete(k);
    }
  }

  function mergeRowSpan(r, c1, c2) {
    if (c2 <= c1) return;
    const parts = [];
    for (let c = c1; c <= c2; c++) {
      const t = cellText(r, c).trim();
      if (t) parts.push(t);
    }
    const left = getCell(r, c1);
    setCell(r, c1, {
      v: parts.join(' '),
      f: '',
      s: { ...left.s, align: left.s.align || 'center' }
    });
    for (let c = c1 + 1; c <= c2; c++) cells.delete(key(r, c));
    clearMergesIn(r, c1, r, c2);
    merges.set(mergeKey(r, c1), { rs: 1, cs: c2 - c1 + 1 });
  }

  function mergeBlock(r1, c1, r2, c2) {
    if (r1 === r2 && c2 > c1) {
      mergeRowSpan(r1, c1, c2);
      return;
    }
    if (c1 === c2 && r2 > r1) {
      const parts = [];
      for (let r = r1; r <= r2; r++) {
        const t = cellText(r, c1).trim();
        if (t) parts.push(t);
      }
      const top = getCell(r1, c1);
      setCell(r1, c1, { v: parts.join(' '), f: '', s: { ...top.s, align: 'center' } });
      for (let r = r1 + 1; r <= r2; r++) cells.delete(key(r, c1));
      clearMergesIn(r1, c1, r2, c2);
      merges.set(mergeKey(r1, c1), { rs: r2 - r1 + 1, cs: 1 });
      return;
    }
    if (r2 > r1 && c2 > c1) {
      for (let r = r1; r <= r2; r++) mergeRowSpan(r, c1, c2);
    }
  }

  function insertAgg(fn) {
    snapshot();
    const box = eachSel(() => {});
    const targetR = box.r2 + 1;
    const targetC = box.c1;
    const start = a1(box.r1, box.c1);
    const end = a1(box.r2, box.c2);
    setCell(targetR, targetC, { f: `=${fn}(${start}:${end})`, v: '' });
    rows = Math.max(rows, targetR + 6);
    active = { r: targetR, c: targetC };
    sel = { r1: targetR, c1: targetC, r2: targetR, c2: targetC };
    render();
  }

  function rangeFromPayload(payload) {
    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) return null;
    if (payload.addr) {
      const p = parseA1(payload.addr);
      if (p) return { r1: p.r, c1: p.c, r2: p.r, c2: p.c };
    }
    if (payload.range) {
      const parts = String(payload.range).split(':');
      const a = parseA1(parts[0]);
      const b = parseA1(parts[1] || parts[0]);
      if (a && b) {
        return {
          r1: Math.min(a.r, b.r),
          c1: Math.min(a.c, b.c),
          r2: Math.max(a.r, b.r),
          c2: Math.max(a.c, b.c)
        };
      }
    }
    if (payload.r != null && payload.c != null && payload.r1 == null) {
      const r = Number(payload.r);
      const c = Number(payload.c);
      if (Number.isInteger(r) && Number.isInteger(c) && r >= 0 && c >= 0) return { r1: r, c1: c, r2: r, c2: c };
    }
    if (payload.r1 != null || payload.c1 != null) {
      const r1 = Math.max(0, Number(payload.r1 ?? 0));
      const c1 = Math.max(0, Number(payload.c1 ?? 0));
      const r2 = Math.max(r1, Number(payload.r2 ?? r1));
      const c2 = Math.max(c1, Number(payload.c2 ?? c1));
      return { r1, c1, r2, c2 };
    }
    return null;
  }

  function applyPayloadRange(payload) {
    const box = rangeFromPayload(payload);
    if (!box) return null;
    sel = box;
    active = { r: box.r1, c: box.c1 };
    return box;
  }

  function payloadStyleValue(payload, keys) {
    if (payload == null) return undefined;
    if (typeof payload !== 'object') return payload;
    for (const k of keys) {
      if (payload[k] != null && payload[k] !== '') return payload[k];
    }
    return undefined;
  }

  function resolveCol(payload) {
    if (payload == null) return -1;
    if (typeof payload === 'number' && Number.isInteger(payload)) return payload;
    if (typeof payload === 'string') {
      const asLetter = letterToIndex(payload);
      if (asLetter >= 0 && /^[A-Za-z]{1,3}$/.test(payload)) return asLetter;
      const { maxC } = dataMax();
      const want = payload.trim().toLowerCase();
      for (let c = 0; c <= maxC; c++) {
        if (String(cellText(0, c)).trim().toLowerCase() === want) return c;
      }
      return -1;
    }
    if (payload.col != null && payload.col !== '') {
      const n = Number(payload.col);
      if (Number.isInteger(n)) return n;
      return resolveCol(String(payload.col));
    }
    if (payload.letter) return letterToIndex(payload.letter);
    if (payload.header) return resolveCol(String(payload.header));
    return -1;
  }

  function writeValue(r, c, v) {
    rows = Math.max(rows, r + 8);
    cols = Math.max(cols, c + 4);
    if (v == null) {
      setCell(r, c, { f: '', v: '' });
      return;
    }
    if (typeof v === 'number' && Number.isFinite(v)) {
      setCell(r, c, { f: '', v });
      return;
    }
    const s = String(v);
    if (s.startsWith('=')) {
      setCell(r, c, { f: s, v: '' });
      return;
    }
    const n = Number(s);
    const asNum = s.trim() !== '' && Number.isFinite(n) && !/^0\d+/.test(s.trim());
    setCell(r, c, { f: '', v: asNum ? n : v });
  }

  function classifyCol(c, maxR) {
    const h = String(cellText(0, c) || '');
    if (/(手机|电话|联系|身份证|账号|卡号|邮编|工号)/.test(h)) return 'id';
    if (/(工资|奖金|金额|收入|费用|价格|合计|绩效|津贴|补贴|成本|利润|收款|付款)/.test(h)) return 'money';
    if (/(百分|比率|占比|折扣)/.test(h)) return 'percent';
    if (/序号/.test(h)) return 'index';
    if (/(日期|时间|年月)/.test(h)) return 'date';
    let num = 0;
    let nonempty = 0;
    for (let r = 1; r <= maxR; r++) {
      const v = getCell(r, c).f ? getCell(r, c).computed : getCell(r, c).v;
      if (v === '' || v == null) continue;
      nonempty++;
      if (typeof v === 'number' || /^-?\d+(\.\d+)?$/.test(String(v).replace(/,/g, ''))) num++;
    }
    if (nonempty >= 2 && num / nonempty >= 0.8) return 'number';
    return 'text';
  }

  function pickTheme(kinds) {
    const money = kinds.some((k) => k === 'money');
    if (money) {
      return { name: 'finance', header: '#217346', fg: '#ffffff', stripe: '#eef7f0', total: '#fff2cc' };
    }
    return { name: 'report', header: '#1f4e79', fg: '#ffffff', stripe: '#eef3f8', total: '#fff2cc' };
  }

  function beautifyLocal(payload) {
    snapshot();
    const { maxR, maxC } = dataMax();
    if (maxR === 0 && maxC === 0 && !cellText(0, 0)) {
      statusEl.textContent = '表是空的，先放入数据再排版';
      return;
    }
    const kinds = [];
    for (let c = 0; c <= maxC; c++) kinds.push(classifyCol(c, maxR));
    const theme = {
      header: payload?.headerFill || payload?.header,
      fg: payload?.headerColor,
      stripe: payload?.stripe,
      total: payload?.totalFill
    };
    const fallback = pickTheme(kinds);
    const headerFill = theme.header || fallback.header;
    const headerFg = theme.fg || fallback.fg;
    const stripe = theme.stripe || fallback.stripe;
    const totalFill = theme.total || fallback.total;

    for (let c = 0; c <= maxC; c++) {
      const h = getCell(0, c);
      if (h.slash) {
        setCell(0, c, { ...h, slash: h.slash, s: { ...h.s, bold: true, align: 'center', wrap: true } });
        continue;
      }
      setCell(0, c, {
        ...h,
        s: {
          ...h.s,
          bold: true,
          fill: headerFill,
          color: headerFg,
          align: 'center',
          vAlign: 'middle',
          wrap: true
        }
      });
    }

    for (let r = 1; r <= maxR; r++) {
      const lastLabel = String(cellText(r, 0) || '');
      if (lastLabel === '合计' || lastLabel === '总计') continue;
      for (let c = 0; c <= maxC; c++) {
        const kind = kinds[c];
        const patch = { wrap: false };
        patch.fill = r % 2 === 0 ? stripe : '#ffffff';
        if (kind === 'money' || kind === 'number' || kind === 'percent') patch.align = 'right';
        else if (kind === 'index') patch.align = 'center';
        else patch.align = 'left';
        if (kind === 'money') patch.numFmt = 'currency';
        else if (kind === 'percent') patch.numFmt = 'percent';
        else if (kind === 'index') patch.numFmt = 'int';
        else if (kind === 'number') patch.numFmt = 'thousand';
        setCell(r, c, { s: patch });
      }
    }

    for (let r = 0; r <= maxR; r++) {
      for (let c = 0; c <= maxC; c++) {
        const cell = getCell(r, c);
        if (cell.slash) continue;
        setCell(r, c, { s: { ...cell.s, border: 'all' } });
      }
    }

    for (let r = 1; r <= maxR; r++) {
      for (let c = 0; c <= maxC; c++) {
        if (kinds[c] !== 'money' && kinds[c] !== 'number') continue;
        const n = Number(getCell(r, c).f ? getCell(r, c).computed : getCell(r, c).v);
        if (Number.isFinite(n) && n < 0) setCell(r, c, { s: { color: '#dc2626', bold: true } });
      }
    }

    const lastLabel = String(cellText(maxR, 0) || '');
    const moneyCols = kinds.map((k, c) => (k === 'money' ? c : -1)).filter((c) => c >= 0);
    if (moneyCols.length && lastLabel !== '合计' && lastLabel !== '总计' && maxR >= 1) {
      const tr = maxR + 1;
      setCell(tr, 0, { v: '合计', f: '', s: { bold: true, fill: totalFill, align: 'center', border: 'all' } });
      for (let c = 1; c <= maxC; c++) {
        if (kinds[c] === 'money') {
          setCell(tr, c, {
            f: `=SUM(${a1(1, c)}:${a1(maxR, c)})`,
            v: '',
            s: { bold: true, fill: totalFill, numFmt: 'currency', align: 'right', border: 'all' }
          });
        } else {
          setCell(tr, c, { v: '', s: { fill: totalFill, border: 'all' } });
        }
      }
    }

    freezeRows = 1;
    freezeCols = 0;
    statusEl.textContent = '已按报表规范排版：表头、对齐、数字格式、框线、冻结首行、合计行';
    render();
  }

  function listOptions(payload, fallback) {
    if (Array.isArray(payload?.options)) return payload.options.map((s) => String(s).trim()).filter(Boolean);
    if (typeof payload === 'string' && payload.includes(',')) {
      return payload.split(/[,，]/).map((s) => s.trim()).filter(Boolean);
    }
    if (typeof payload?.options === 'string') {
      return payload.options.split(/[,，]/).map((s) => s.trim()).filter(Boolean);
    }
    if (fallback) {
      const raw = window.prompt(fallback.title, fallback.sample);
      if (raw == null) return null;
      return raw.split(/[,，]/).map((s) => s.trim()).filter(Boolean);
    }
    return [];
  }

  function command(name, payload) {
    if (name === 'undo') {
      const prev = undo.pop();
      if (!prev) return;
      restoreSnapshot(prev);
      render();
      return;
    }
    if (name === 'bold' || name === 'italic' || name === 'underline' || name === 'strike' || name === 'wrap') {
      applyPayloadRange(payload);
      if (payload && typeof payload === 'object') {
        if (payload.on === true || payload.set === true || payload[name] === true) return applyStyle({ [name]: true });
        if (payload.on === false || payload[name] === false) return applyStyle({ [name]: false });
      }
      return toggleStyle(name);
    }
    if (name === 'styleRange' || name === 'applyStyle') {
      applyPayloadRange(payload);
      const patch = {};
      if (!payload || typeof payload !== 'object') return;
      if (payload.bold != null) patch.bold = !!payload.bold;
      if (payload.italic != null) patch.italic = !!payload.italic;
      if (payload.wrap != null) patch.wrap = !!payload.wrap;
      if (payload.fill) patch.fill = payload.fill;
      if (payload.color || payload.fontColor) patch.color = payload.color || payload.fontColor;
      if (payload.align) patch.align = payload.align;
      if (payload.vAlign) patch.vAlign = payload.vAlign;
      if (payload.numFmt) patch.numFmt = payload.numFmt;
      if (payload.border) patch.border = payload.border;
      if (payload.fontSize) patch.fontSize = payload.fontSize;
      if (payload.fontFamily) patch.fontFamily = payload.fontFamily;
      if (!Object.keys(patch).length) return;
      return applyStyle(patch);
    }
    if (name === 'fontFamily') {
      applyPayloadRange(payload);
      return applyStyle({ fontFamily: payloadStyleValue(payload, ['fontFamily', 'family']) ?? payload });
    }
    if (name === 'fontSize') {
      applyPayloadRange(payload);
      return applyStyle({ fontSize: payloadStyleValue(payload, ['fontSize', 'size']) ?? payload });
    }
    if (name === 'fontColor') {
      applyPayloadRange(payload);
      return applyStyle({ color: payloadStyleValue(payload, ['color', 'fontColor']) ?? payload });
    }
    if (name === 'fill') {
      applyPayloadRange(payload);
      return applyStyle({ fill: payloadStyleValue(payload, ['color', 'fill']) ?? payload });
    }
    if (name === 'fillColumn') {
      const c = Number(payload?.col ?? payload?.c);
      const color = payload?.color || '#d9ead3';
      if (!Number.isInteger(c) || c < 0) return;
      snapshot();
      const { maxR } = dataMax();
      for (let r = 0; r <= maxR; r++) {
        setCell(r, c, { s: { fill: color } });
      }
      statusEl.textContent = `已填充第 ${c + 1} 列（共 ${maxR + 1} 行）`;
      sel = { r1: 0, c1: c, r2: maxR, c2: c };
      active = { r: 0, c };
      render();
      return;
    }
    if (name === 'fillRow') {
      const r = Number(payload?.row ?? payload?.r);
      const color = payload?.color || '#b45f06';
      if (!Number.isInteger(r) || r < 0) return;
      snapshot();
      const { maxC } = dataMax();
      for (let c = 0; c <= maxC; c++) {
        setCell(r, c, { s: { fill: color } });
      }
      statusEl.textContent = `已填充第 ${r + 1} 行（共 ${maxC + 1} 列）`;
      sel = { r1: r, c1: 0, r2: r, c2: maxC };
      active = { r, c: 0 };
      render();
      return;
    }
    if (name === 'slashHeader') {
      const r = Number(payload?.r ?? payload?.row ?? Math.min(sel.r1, sel.r2));
      const c = Number(payload?.c ?? payload?.col ?? Math.min(sel.c1, sel.c2));
      const top = String(payload?.top ?? payload?.upper ?? '部分').trim() || '部分';
      const bottom = String(payload?.bottom ?? payload?.lower ?? '序号').trim() || '序号';
      if (!Number.isInteger(r) || r < 0 || !Number.isInteger(c) || c < 0) return;
      snapshot();
      setCell(r, c, {
        f: '',
        v: `${bottom}/${top}`,
        slash: { top, bottom },
        s: { wrap: true, bold: true, align: 'center' }
      });
      sel = { r1: r, c1: c, r2: r, c2: c };
      active = { r, c };
      statusEl.textContent = `已做成斜线表头：右上「${top}」，左下「${bottom}」`;
      render();
      return;
    }
    if (name === 'align') {
      applyPayloadRange(payload);
      return applyStyle({ align: payloadStyleValue(payload, ['align']) ?? payload });
    }
    if (name === 'vAlign') {
      applyPayloadRange(payload);
      return applyStyle({ vAlign: payloadStyleValue(payload, ['vAlign']) ?? payload });
    }
    if (name === 'border') {
      applyPayloadRange(payload);
      return applyStyle({ border: payloadStyleValue(payload, ['border']) ?? payload });
    }
    if (name === 'numFmt') {
      applyPayloadRange(payload);
      return applyStyle({ numFmt: payloadStyleValue(payload, ['numFmt', 'fmt']) ?? payload });
    }
    if (name === 'clear') {
      applyPayloadRange(payload);
      snapshot();
      eachSel((r, c) => cells.delete(key(r, c)));
      render();
      return;
    }
    if (name === 'clearFormat') {
      applyPayloadRange(payload);
      snapshot();
      eachSel((r, c) => setCell(r, c, { s: emptyStyle(), slash: null }));
      render();
      return;
    }
    if (name === 'autosum') { applyPayloadRange(payload); return insertAgg('SUM'); }
    if (name === 'average') { applyPayloadRange(payload); return insertAgg('AVERAGE'); }
    if (name === 'max') { applyPayloadRange(payload); return insertAgg('MAX'); }
    if (name === 'min') { applyPayloadRange(payload); return insertAgg('MIN'); }
    if (name === 'count') { applyPayloadRange(payload); return insertAgg('COUNT'); }
    if (name === 'insertRow') {
      snapshot();
      const at = payload?.at != null ? Number(payload.at) : Math.min(sel.r1, sel.r2);
      const next = new Map();
      for (const [k, cell] of cells) {
        const p = parseA1(k);
        if (!p) continue;
        const nr = p.r >= at ? p.r + 1 : p.r;
        next.set(key(nr, p.c), cell);
      }
      cells.clear();
      next.forEach((v, k) => cells.set(k, v));
      rows++;
      render();
      return;
    }
    if (name === 'insertCol') {
      snapshot();
      const at = payload?.at != null ? Number(payload.at) : Math.min(sel.c1, sel.c2);
      const next = new Map();
      for (const [k, cell] of cells) {
        const p = parseA1(k);
        if (!p) continue;
        const nc = p.c >= at ? p.c + 1 : p.c;
        next.set(key(p.r, nc), cell);
      }
      cells.clear();
      next.forEach((v, k) => cells.set(k, v));
      cols++;
      if (payload?.header) {
        setCell(0, at, { v: payload.header, f: '', s: { bold: true, align: 'center' } });
      }
      render();
      return;
    }
    if (name === 'deleteRow') {
      snapshot();
      const at = payload?.at != null ? Number(payload.at) : Math.min(sel.r1, sel.r2);
      const next = new Map();
      for (const [k, cell] of cells) {
        const p = parseA1(k);
        if (!p || p.r === at) continue;
        const nr = p.r > at ? p.r - 1 : p.r;
        next.set(key(nr, p.c), cell);
      }
      cells.clear();
      next.forEach((v, k) => cells.set(k, v));
      render();
      return;
    }
    if (name === 'deleteCol') {
      snapshot();
      const at = payload?.at != null ? Number(payload.at) : Math.min(sel.c1, sel.c2);
      const next = new Map();
      for (const [k, cell] of cells) {
        const p = parseA1(k);
        if (!p || p.c === at) continue;
        const nc = p.c > at ? p.c - 1 : p.c;
        next.set(key(p.r, nc), cell);
      }
      cells.clear();
      next.forEach((v, k) => cells.set(k, v));
      render();
      return;
    }
    if (name === 'sortAsc' || name === 'sortDesc' || name === 'sort') {
      snapshot();
      const { maxR, maxC } = dataMax();
      const dir = (name === 'sortDesc' || payload?.dir === 'desc' || payload?.dir === '降序') ? -1 : 1;
      let keyCol = resolveCol(payload);
      if (keyCol < 0) keyCol = Math.min(sel.c1, sel.c2);
      const r1 = payload?.r1 != null ? Number(payload.r1) : 1;
      const r2 = payload?.r2 != null ? Number(payload.r2) : maxR;
      const block = [];
      for (let r = r1; r <= r2; r++) {
        const row = [];
        for (let c = 0; c <= maxC; c++) row.push(getCell(r, c));
        block.push(row);
      }
      block.sort((a, b) => {
        const va = a[keyCol]?.v;
        const vb = b[keyCol]?.v;
        const na = Number(va);
        const nb = Number(vb);
        if (Number.isFinite(na) && Number.isFinite(nb) && va !== '' && vb !== '') return (na - nb) * dir;
        return String(va ?? '').localeCompare(String(vb ?? ''), 'zh') * dir;
      });
      block.forEach((row, i) => {
        row.forEach((cell, j) => cells.set(key(r1 + i, j), cell));
      });
      statusEl.textContent = `已按第 ${keyCol + 1} 列${dir < 0 ? '降序' : '升序'}排序（整行一起动）`;
      render();
      return;
    }
    if (name === 'select') {
      const r1 = Math.max(0, Number(payload?.r1 ?? 0));
      const c1 = Math.max(0, Number(payload?.c1 ?? 0));
      const r2 = Math.max(r1, Number(payload?.r2 ?? r1));
      const c2 = Math.max(c1, Number(payload?.c2 ?? c1));
      sel = { r1, c1, r2, c2 };
      active = { r: r1, c: c1 };
      render();
      return;
    }
    if (name === 'find') {
      if (payload?.find) {
        findWhat.value = String(payload.find);
        openFind(false);
        findNext();
        return;
      }
      openFind(false);
      return;
    }
    if (name === 'replace') {
      if (payload?.find) {
        findWhat.value = String(payload.find);
        replaceWith.value = String(payload.replace ?? '');
        replaceAll();
        return;
      }
      openFind(true);
      return;
    }
    if (name === 'filter') {
      filterEnabled = !filterEnabled;
      if (!filterEnabled) Object.keys(filters).forEach((k) => delete filters[k]);
      filterPop.hidden = true;
      statusEl.textContent = filterEnabled ? '已开启自动筛选 · 点表头 ▾ 选条件' : '已关闭筛选';
      render();
      return;
    }
    if (name === 'freezeTop') {
      freezeRows = 1;
      freezeCols = 0;
      statusEl.textContent = '已冻结首行';
      render();
      return;
    }
    if (name === 'freezeLeft') {
      freezeRows = 0;
      freezeCols = 1;
      statusEl.textContent = '已冻结首列';
      render();
      return;
    }
    if (name === 'freeze') {
      freezeRows = active.r;
      freezeCols = active.c;
      statusEl.textContent = freezeRows || freezeCols
        ? `已冻结活动单元格上方 ${freezeRows} 行、左侧 ${freezeCols} 列`
        : '活动格在 A1，未冻结（请先点要冻结边界的格子）';
      render();
      return;
    }
    if (name === 'unfreeze') {
      freezeRows = 0;
      freezeCols = 0;
      statusEl.textContent = '已取消冻结';
      render();
      return;
    }
    if (name === 'unique') {
      snapshot();
      applyPayloadRange(payload);
      let box = eachSel(() => {});
      if (box.r1 === box.r2 && box.c1 === box.c2) {
        const { maxR, maxC } = dataMax();
        box = { r1: 1, c1: 0, r2: maxR, c2: maxC };
      }
      const seen = new Set();
      const kept = [];
      for (let r = box.r1; r <= box.r2; r++) {
        const row = [];
        for (let c = box.c1; c <= box.c2; c++) row.push(getCell(r, c));
        const sig = JSON.stringify(row.map((x) => x.v));
        if (seen.has(sig)) continue;
        seen.add(sig);
        kept.push(row);
      }
      for (let r = box.r1; r <= box.r2; r++) {
        for (let c = box.c1; c <= box.c2; c++) cells.delete(key(r, c));
      }
      kept.forEach((row, i) => row.forEach((cell, j) => cells.set(key(box.r1 + i, box.c1 + j), cell)));
      render();
      return;
    }
    if (name === 'copy') {
      const box = eachSel(() => {});
      const lines = [];
      for (let r = box.r1; r <= box.r2; r++) {
        const line = [];
        for (let c = box.c1; c <= box.c2; c++) {
          const cell = getCell(r, c);
          line.push(cell.f || cell.v);
        }
        lines.push(line.join('\t'));
      }
      navigator.clipboard?.writeText(lines.join('\n'));
      statusEl.textContent = '已复制选区';
      return;
    }
    if (name === 'paste') {
      navigator.clipboard?.readText().then((text) => {
        snapshot();
        const lines = text.split(/\r?\n/);
        lines.forEach((line, i) => {
          line.split('\t').forEach((val, j) => {
            const r = active.r + i;
            const c = active.c + j;
            if (String(val).startsWith('=')) setCell(r, c, { f: val, v: '' });
            else setCell(r, c, { f: '', v: val });
          });
        });
        render();
      });
      return;
    }
    if (name === 'merge') {
      applyPayloadRange(payload);
      snapshot();
      const box = eachSel(() => {});
      mergeBlock(box.r1, box.c1, box.r2, box.c2);
      statusEl.textContent = '已合并选区';
      render();
      return;
    }
    if (name === 'mergeColumns') {
      let c1 = Number(payload?.c1);
      let c2 = Number(payload?.c2);
      if (payload?.a && payload?.b) {
        c1 = letterToIndex(payload.a);
        c2 = letterToIndex(payload.b);
      }
      if (!Number.isInteger(c1) || !Number.isInteger(c2) || c1 < 0 || c2 < 0) return;
      if (c1 > c2) [c1, c2] = [c2, c1];
      snapshot();
      const { maxR } = dataMax();
      for (let r = 0; r <= maxR; r++) mergeRowSpan(r, c1, c2);
      sel = { r1: 0, c1, r2: maxR, c2 };
      active = { r: 0, c: c1 };
      statusEl.textContent = `已逐行合并 ${colName(c1)} 到 ${colName(c2)} 列`;
      render();
      return;
    }
    if (name === 'chart') {
      applyPayloadRange(payload);
      chartType = (typeof payload === 'string' ? payload : (payload?.type || payload?.chart)) || 'bar';
      render();
      return;
    }
    if (name === 'checkbox') {
      applyPayloadRange(payload);
      snapshot();
      eachSel((r, c) => setCell(r, c, { v: false, f: '', ctrl: { type: 'checkbox' }, s: { align: 'center' } }));
      statusEl.textContent = '已插入勾选框 · 单击格子打勾/取消';
      render();
      return;
    }
    if (name === 'yesno') {
      applyPayloadRange(payload);
      snapshot();
      eachSel((r, c) => setCell(r, c, { v: '', f: '', ctrl: { type: 'yesno' }, s: { align: 'center' } }));
      statusEl.textContent = '已插入是/否单选 · 点「是」或「否」';
      render();
      return;
    }
    if (name === 'radio') {
      applyPayloadRange(payload);
      snapshot();
      const groupId = 'g' + Date.now();
      eachSel((r, c) => setCell(r, c, { v: false, f: '', ctrl: { type: 'radio', groupId }, s: { align: 'center' } }));
      statusEl.textContent = '已插入单选圈 · 选区里只能勾中一个';
      render();
      return;
    }
    if (name === 'dropdown') {
      applyPayloadRange(payload);
      const options = listOptions(payload, { title: '下拉选项（逗号分隔）', sample: '是,否,待定' });
      if (!options || !options.length) return;
      snapshot();
      eachSel((r, c) => setCell(r, c, { v: options[0], f: '', ctrl: { type: 'dropdown', options }, s: { align: 'center' } }));
      statusEl.textContent = '已插入下拉单选';
      render();
      return;
    }
    if (name === 'multiselect') {
      applyPayloadRange(payload);
      const options = listOptions(payload, { title: '多选项（逗号分隔）', sample: '选项A,选项B,选项C' });
      if (!options || !options.length) return;
      snapshot();
      eachSel((r, c) => setCell(r, c, { v: '', f: '', ctrl: { type: 'multiselect', options }, s: { wrap: true, vAlign: 'top' } }));
      rows = Math.max(rows, Math.max(sel.r1, sel.r2) + 4);
      statusEl.textContent = '已插入多选 · 可同时勾多个';
      render();
      return;
    }
    if (name === 'setCell' || name === 'setCells') {
      snapshot();
      const items = payload?.cells || (Array.isArray(payload) ? payload : [payload]);
      let last = { r: 0, c: 0 };
      items.forEach((item) => {
        if (!item) return;
        let r = Number(item.r);
        let c = Number(item.c);
        if (item.addr) {
          const p = parseA1(item.addr);
          if (p) { r = p.r; c = p.c; }
        }
        if (!Number.isInteger(r) || !Number.isInteger(c) || r < 0 || c < 0) return;
        if (item.f) writeValue(r, c, String(item.f).startsWith('=') ? item.f : `=${item.f}`);
        else writeValue(r, c, item.v ?? item.value);
        if (item.s) setCell(r, c, { s: item.s });
        last = { r, c };
      });
      sel = { r1: last.r, c1: last.c, r2: last.r, c2: last.c };
      active = last;
      statusEl.textContent = '已写入单元格';
      render();
      return;
    }
    if (name === 'renameHeader') {
      const c = resolveCol(payload);
      const title = payload?.to || payload?.name || payload?.value || payload?.v;
      if (c < 0 || title == null) return;
      snapshot();
      const cur = getCell(0, c);
      setCell(0, c, { ...cur, v: title, f: '' });
      statusEl.textContent = `表头已改为「${title}」`;
      render();
      return;
    }
    if (name === 'filterBy') {
      if (payload?.clear) {
        Object.keys(filters).forEach((k) => delete filters[k]);
        filterEnabled = false;
        statusEl.textContent = '已关闭筛选';
        render();
        return;
      }
      let c = resolveCol(payload);
      if (c < 0) c = 0;
      const values = Array.isArray(payload?.values)
        ? payload.values.map(String)
        : [String(payload?.value ?? payload?.q ?? '')].filter((s) => s !== '');
      if (!values.length) {
        filterEnabled = true;
        render();
        return;
      }
      filterEnabled = true;
      filters[c] = new Set(values);
      statusEl.textContent = `已筛选第 ${c + 1} 列：${values.join('、')}`;
      render();
      return;
    }
    if (name === 'fillSeries') {
      const c = resolveCol(payload);
      const col = c >= 0 ? c : 0;
      const { maxR } = dataMax();
      const r1 = payload?.r1 != null ? Number(payload.r1) : 1;
      const r2 = payload?.r2 != null ? Number(payload.r2) : Math.max(maxR, r1);
      const start = payload?.start != null ? Number(payload.start) : 1;
      const step = payload?.step != null ? Number(payload.step) : 1;
      snapshot();
      if (payload?.header) setCell(0, col, { v: payload.header, f: '', s: { bold: true, align: 'center' } });
      for (let r = r1, n = start; r <= r2; r++, n += step) writeValue(r, col, n);
      statusEl.textContent = `已在第 ${col + 1} 列填入序号`;
      render();
      return;
    }
    if (name === 'trimAll') {
      snapshot();
      const { maxR, maxC } = dataMax();
      for (let r = 0; r <= maxR; r++) {
        for (let c = 0; c <= maxC; c++) {
          const cell = getCell(r, c);
          if (cell.f || cell.slash || cell.ctrl) continue;
          if (typeof cell.v === 'string') setCell(r, c, { v: cell.v.replace(/\s+/g, ' ').trim() });
        }
      }
      statusEl.textContent = '已去掉多余空格';
      render();
      return;
    }
    if (name === 'splitColumn') {
      let c = resolveCol(payload);
      if (c < 0) c = Math.min(sel.c1, sel.c2);
      const sepRaw = payload?.sep ?? payload?.separator ?? ',';
      const sep = sepRaw === '空格' ? /\s+/ : sepRaw === 'tab' ? '\t' : sepRaw;
      const { maxR } = dataMax();
      let width = 1;
      const parts = [];
      for (let r = 0; r <= maxR; r++) {
        const t = cellText(r, c);
        const bits = t === '' ? [''] : String(t).split(sep).map((s) => s.trim());
        parts[r] = bits;
        width = Math.max(width, bits.length);
      }
      snapshot();
      for (let i = 1; i < width; i++) {
        const next = new Map();
        const at = c + i;
        for (const [k, cell] of cells) {
          const p = parseA1(k);
          if (!p) continue;
          const nc = p.c >= at ? p.c + 1 : p.c;
          next.set(key(p.r, nc), cell);
        }
        cells.clear();
        next.forEach((v, k) => cells.set(k, v));
        cols++;
      }
      for (let r = 0; r <= maxR; r++) {
        (parts[r] || []).forEach((bit, i) => writeValue(r, c + i, bit));
      }
      statusEl.textContent = `已把第 ${c + 1} 列拆成 ${width} 列`;
      render();
      return;
    }
    if (name === 'deleteBlankRows') {
      snapshot();
      const { maxR, maxC } = dataMax();
      const kept = [];
      if (maxR >= 0) {
        const header = [];
        for (let c = 0; c <= maxC; c++) header.push(getCell(0, c));
        kept.push(header);
      }
      for (let r = 1; r <= maxR; r++) {
        let empty = true;
        const row = [];
        for (let c = 0; c <= maxC; c++) {
          const cell = getCell(r, c);
          row.push(cell);
          if (cell.f || (cell.v !== '' && cell.v != null) || cell.ctrl || cell.slash) empty = false;
        }
        if (!empty) kept.push(row);
      }
      const next = new Map();
      kept.forEach((row, r) => row.forEach((cell, c) => next.set(key(r, c), cell)));
      cells.clear();
      next.forEach((v, k) => cells.set(k, v));
      statusEl.textContent = `已删除空行，剩余 ${kept.length} 行（含表头）`;
      render();
      return;
    }
    if (name === 'highlightIf') {
      snapshot();
      const { maxR, maxC } = dataMax();
      let c = resolveCol(payload);
      const op = String(payload?.op || payload?.operator || 'gt');
      const threshold = payload?.value ?? payload?.n;
      const color = payload?.color || '#dc2626';
      const startC = c >= 0 ? c : 0;
      const endC = c >= 0 ? c : maxC;
      for (let r = 1; r <= maxR; r++) {
        for (let cc = startC; cc <= endC; cc++) {
          const raw = getCell(r, cc).f ? getCell(r, cc).computed : getCell(r, cc).v;
          const n = Number(raw);
          let hit = false;
          if (op === 'contains') hit = String(raw).includes(String(threshold));
          else if (op === 'eq') hit = String(raw) === String(threshold) || n === Number(threshold);
          else if (Number.isFinite(n)) {
            const t = Number(threshold);
            if (op === 'gt' || op === '>') hit = n > t;
            else if (op === 'gte' || op === '>=') hit = n >= t;
            else if (op === 'lt' || op === '<') hit = n < t;
            else if (op === 'lte' || op === '<=') hit = n <= t;
          }
          if (hit) setCell(r, cc, { s: { color, bold: true } });
        }
      }
      statusEl.textContent = '已按条件标出格子';
      render();
      return;
    }
    if (name === 'groupSum') {
      const by = payload?.byCol != null ? Number(payload.byCol) : resolveCol(payload?.by ?? payload);
      const val = payload?.valueCol != null ? Number(payload.valueCol) : resolveCol(payload?.value ?? payload?.val ?? payload?.sum ?? payload?.valHeader);
      if (by < 0 || val < 0) return;
      snapshot();
      const { maxR } = dataMax();
      const keys = [];
      const seen = new Set();
      for (let r = 1; r <= maxR; r++) {
        const k = cellText(r, by);
        if (!k || seen.has(k)) continue;
        seen.add(k);
        keys.push(k);
      }
      const outR = maxR + 2;
      setCell(outR, 0, { v: cellText(0, by) || '分组', s: { bold: true, fill: '#fff2cc' } });
      setCell(outR, 1, { v: (cellText(0, val) || '合计') + '合计', s: { bold: true, fill: '#fff2cc' } });
      keys.forEach((k, i) => {
        setCell(outR + 1 + i, 0, { v: k });
        const crit = `"${String(k).replace(/"/g, '')}"`;
        setCell(outR + 1 + i, 1, { f: `=SUMIF(${a1(1, by)}:${a1(maxR, by)},${crit},${a1(1, val)}:${a1(maxR, val)})`, v: '' });
      });
      rows = Math.max(rows, outR + keys.length + 6);
      statusEl.textContent = `已按「${cellText(0, by)}」汇总「${cellText(0, val)}」`;
      render();
      return;
    }
    if (name === 'addColumn') {
      const { maxC } = dataMax();
      const at = payload?.at != null ? Number(payload.at) : maxC + 1;
      command('insertCol', { at, header: payload?.header || payload?.name || '新列' });
      return;
    }
    if (name === 'watermark') {
      if (payload?.opacity != null) watermarkOpacity = Number(payload.opacity);
      const shown = composeWatermark(payload);
      watermarkKey = shown;
      applyWatermarkOverlay(wmEl, shown, watermarkOpacity);
      statusEl.textContent = shown ? `已盖水印「${shown}」` : '已去掉水印';
      return;
    }
    if (name === 'maskSensitive') {
      snapshot();
      const { maxR, maxC } = dataMax();
      let n = 0;
      for (let r = 0; r <= maxR; r++) {
        for (let c = 0; c <= maxC; c++) {
          const cell = getCell(r, c);
          if (cell.f || cell.slash || cell.ctrl) continue;
          const next = maskSensitiveText(cell.v);
          if (next !== cell.v && next !== String(cell.v ?? '')) {
            setCell(r, c, { v: next, f: '' });
            n++;
          } else if (typeof cell.v === 'number') {
            const asText = maskSensitiveText(String(Math.trunc(cell.v)));
            if (asText !== String(Math.trunc(cell.v))) {
              setCell(r, c, { v: asText, f: '' });
              n++;
            }
          }
        }
      }
      statusEl.textContent = n ? `已打码 ${n} 处手机号/身份证` : '没找到可打码的手机号或身份证';
      render();
      return;
    }
    if (name === 'highlightNegatives') {
      snapshot();
      const { maxR, maxC } = dataMax();
      for (let r = 1; r <= maxR; r++) {
        for (let c = 0; c <= maxC; c++) {
          const n = Number(getCell(r, c).f ? getCell(r, c).computed : getCell(r, c).v);
          if (Number.isFinite(n) && n < 0) setCell(r, c, { s: { color: '#dc2626', bold: true } });
        }
      }
      statusEl.textContent = '已把负数标红';
      render();
      return;
    }
    if (name === 'beautify') {
      beautifyLocal(payload);
      return;
    }
  }

  root.querySelector('#xlFindNext').addEventListener('click', findNext);
  root.querySelector('#xlReplaceOne').addEventListener('click', replaceOne);
  root.querySelector('#xlReplaceAll').addEventListener('click', replaceAll);
  root.querySelector('#xlFindClose').addEventListener('click', () => { findDlg.hidden = true; });
  findWhat.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') { e.preventDefault(); findNext(); }
  });
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      findDlg.hidden = true;
      filterPop.hidden = true;
    }
  });
  document.addEventListener('mousedown', (e) => {
    if (!filterPop.hidden && !filterPop.contains(e.target) && !e.target.closest('.xl-filter-btn')) {
      filterPop.hidden = true;
    }
  });

  loadAoA([[]]);

  function getSnapshot() {
    recalc();
    const { maxR, maxC } = dataMax();
    const headers = [];
    for (let c = 0; c <= maxC; c++) headers.push(cellText(0, c));
    const preview = [];
    for (let r = 0; r <= Math.min(maxR, 20); r++) {
      const row = [];
      for (let c = 0; c <= maxC; c++) row.push(cellText(r, c));
      preview.push(row);
    }
    return {
      headers,
      preview,
      maxR,
      maxC,
      selection: { ...sel },
      active: { ...active },
      note: '第0行是表头，第1行起是数据。列A的索引是0。写格子用 setCell 的 addr 如 A2。'
    };
  }

  function buildPrintElement() {
    recalc();
    const { maxR, maxC } = dataMax();
    const wrap = document.createElement('div');
    wrap.className = 'xl-print-page';
    wrap.style.cssText = 'background:#fff;padding:28px 32px;color:#0f172a;font-family:Microsoft YaHei,"微软雅黑",sans-serif;';
    const table = document.createElement('table');
    table.style.cssText = 'border-collapse:collapse;font-size:12px;line-height:1.45;';
    let rowsOut = 0;
    for (let r = 0; r <= maxR; r++) {
      if (!rowPassesFilter(r)) continue;
      let empty = true;
      for (let c = 0; c <= maxC; c++) {
        const cell = getCell(r, c);
        if (cellText(r, c) || cell.slash || cell.s.fill || cell.ctrl) { empty = false; break; }
      }
      if (empty) continue;
      const tr = document.createElement('tr');
      for (let c = 0; c <= maxC; c++) {
        if (isCovered(r, c)) continue;
        const cell = getCell(r, c);
        const td = document.createElement('td');
        const span = merges.get(mergeKey(r, c));
        if (span?.cs > 1) td.colSpan = span.cs;
        if (span?.rs > 1) td.rowSpan = span.rs;
        const st = styleAttr(cell.s);
        td.style.cssText = `${st};padding:7px 10px;border:1px solid #94a3b8;min-width:72px;white-space:pre-wrap;`;
        if (!cell.s.border || cell.s.border === 'none') td.style.border = '1px solid #94a3b8';
        if (cell.slash) td.textContent = `${cell.slash.bottom || ''} / ${cell.slash.top || ''}`;
        else td.textContent = formatValue(cell);
        tr.appendChild(td);
      }
      table.appendChild(tr);
      rowsOut += 1;
    }
    if (!rowsOut) {
      const p = document.createElement('p');
      p.textContent = '表里没有可导出的内容。';
      wrap.appendChild(p);
    } else {
      wrap.appendChild(table);
    }
    const wmText = resolveWatermarkText(watermarkKey) || (watermarkKey && composeWatermark({ text: watermarkKey }));
    if (wmText) {
      wrap.style.position = 'relative';
      const overlay = document.createElement('div');
      overlay.className = 'watermark-overlay';
      overlay.style.position = 'absolute';
      overlay.style.inset = '0';
      applyWatermarkOverlay(overlay, wmText, watermarkOpacity);
      wrap.appendChild(overlay);
    }
    return wrap;
  }

  return {
    element: root,
    loadAoA,
    command,
    exportXlsx,
    toAoA,
    getSnapshot,
    buildPrintElement,
    focus: () => root.focus({ preventScroll: true })
  };
}
