/**
 * 安全水印：弹出窗体填写单位和使用人，避免藏在功能区左侧。
 */

import { composeWatermark } from '../engine/watermarkEngine.js';

const WM_STORE = 'exceltools.watermark.identity';

export function createWatermarkModal({ onStamp, onMask } = {}) {
  const el = document.createElement('div');
  el.className = 'modal-backdrop';
  el.style.display = 'none';
  el.innerHTML = `
    <div class="modal-card" style="max-width: 460px; width: 92%;">
      <div class="modal-header">
        <h3 class="modal-title">填写水印信息</h3>
        <button type="button" class="modal-close" id="wmClose">&times;</button>
      </div>
      <p style="font-size: 0.85rem; color: var(--text-muted); margin: 0 0 1rem;">
        水印要能看出是哪家单位、哪个人的。填在这个窗口里，点「盖上水印」就会打到当前表或文档上。
      </p>
      <form id="wmForm">
        <div class="form-group">
          <label class="form-label">单位 / 公司</label>
          <input class="form-select" id="wmCompany" placeholder="例如：华润集团" autocomplete="organization" />
        </div>
        <div class="form-group">
          <label class="form-label">使用人 / 接收人</label>
          <input class="form-select" id="wmPerson" placeholder="例如：张三" autocomplete="name" />
        </div>
        <div class="form-group">
          <label class="form-label">补充（可选）</label>
          <input class="form-select" id="wmNote" placeholder="部门、日期、用途，例如：销售部 2026-09-03" />
        </div>
        <div class="form-group">
          <label class="form-label">水印说法</label>
          <div class="wm-tone-row">
            <label class="wm-tone"><input type="radio" name="wmTone" value="for" checked /> 仅供此人使用</label>
            <label class="wm-tone"><input type="radio" name="wmTone" value="confidential" /> 本单位内部机密</label>
            <label class="wm-tone"><input type="radio" name="wmTone" value="draft" /> 草案仅供参考</label>
            <label class="wm-tone"><input type="radio" name="wmTone" value="custom" /> 按补充原句</label>
          </div>
        </div>
        <p class="wm-dialog-preview" id="wmDialogPreview">预览：请填写单位或使用人</p>
        <div style="display:flex; gap: 0.6rem; margin-top: 1.1rem; flex-wrap: wrap;">
          <button type="submit" class="btn-primary" style="flex:1; min-width: 140px;">盖上水印</button>
          <button type="button" class="btn-secondary" id="wmMask">手机号/身份证打码</button>
          <button type="button" class="btn-secondary" id="wmClear">去掉水印</button>
          <button type="button" class="btn-secondary" id="wmCancel">取消</button>
        </div>
      </form>
    </div>
  `;

  const company = el.querySelector('#wmCompany');
  const person = el.querySelector('#wmPerson');
  const note = el.querySelector('#wmNote');
  const preview = el.querySelector('#wmDialogPreview');

  function tone() {
    return el.querySelector('input[name="wmTone"]:checked')?.value || 'for';
  }

  function fields() {
    return {
      company: company.value.trim(),
      person: person.value.trim(),
      note: note.value.trim(),
      tone: tone()
    };
  }

  function updatePreview() {
    const f = fields();
    const text = composeWatermark(f) || '请填写单位或使用人';
    preview.textContent = `预览：${text}`;
  }

  function loadSaved() {
    try {
      const saved = JSON.parse(localStorage.getItem(WM_STORE) || 'null');
      if (!saved) return;
      company.value = saved.company || '';
      person.value = saved.person || '';
      note.value = saved.note || '';
      const t = saved.tone || 'for';
      const radio = el.querySelector(`input[name="wmTone"][value="${t}"]`);
      if (radio) radio.checked = true;
    } catch { /* ignore */ }
    updatePreview();
  }

  function persist() {
    try { localStorage.setItem(WM_STORE, JSON.stringify(fields())); } catch { /* ignore */ }
  }

  function hide() {
    el.style.display = 'none';
  }

  function show() {
    loadSaved();
    el.style.display = 'flex';
    setTimeout(() => company.focus(), 30);
  }

  el.querySelector('#wmForm').addEventListener('submit', (e) => {
    e.preventDefault();
    const f = fields();
    if (!f.company && !f.person && !f.note) {
      company.focus();
      preview.textContent = '预览：请先填写单位或使用人，才能看出是谁的文件。';
      return;
    }
    persist();
    const text = composeWatermark(f);
    onStamp?.({ ...f, text });
    hide();
  });

  el.querySelector('#wmMask').addEventListener('click', () => {
    onMask?.();
  });

  el.querySelector('#wmClear').addEventListener('click', () => {
    onStamp?.({ tone: 'none' });
    hide();
  });

  el.querySelector('#wmClose').addEventListener('click', hide);
  el.querySelector('#wmCancel').addEventListener('click', hide);
  el.addEventListener('click', (e) => { if (e.target === el) hide(); });
  el.querySelectorAll('#wmCompany, #wmPerson, #wmNote, input[name="wmTone"]').forEach((node) => {
    node.addEventListener('input', () => { updatePreview(); persist(); });
    node.addEventListener('change', () => { updatePreview(); persist(); });
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && el.style.display !== 'none') hide();
  });

  return { element: el, show, hide };
}
