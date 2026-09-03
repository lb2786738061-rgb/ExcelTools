/**
 * Office 智能效率工具箱 - 防泄密水印与数据脱敏引擎
 * 核心功能：全屏倾斜文本水印防伪生成、敏感字段（手机号/身份证/金额）打码脱敏
 */

/**
 * 预置常用防伪水印模版选项
 */
export const WATERMARK_PRESETS = [
  { id: 'none', label: '无水印' },
  { id: 'confidential', label: '内部机密 严禁外传' },
  { id: 'bank', label: '仅供办理业务使用' },
  { id: 'draft', label: '草案样本 仅供参考' }
];

/**
 * 敏感文本快速脱敏遮挡
 * @param {string} text - 原始文本字符串
 * @returns {string} 替换掩码后的字符串
 */
export function maskSensitiveText(text) {
  if (text == null) return text;
  const src = String(text);
  let result = src.replace(/(1[3-9]\d)\d{4}(\d{4})/g, '$1****$2');
  result = result.replace(/(\d{6})\d{8}(\d{3}[\dXx])/g, '$1********$2');
  return result;
}

/**
 * 生成动态平铺水印元素的样式对象列表
 * @param {string} watermarkText - 水印文案
 * @param {number} count - 生成密度数量
 * @returns {Array} 水印标签列表
 */
export function generateWatermarkGrid(watermarkText, count = 12) {
  if (!watermarkText || watermarkText === '无水印' || watermarkText === 'none') {
    return [];
  }

  const list = [];
  for (let i = 0; i < count; i++) {
    list.push({
      id: `wm-${i}`,
      text: watermarkText
    });
  }
  return list;
}

export function resolveWatermarkText(idOrText) {
  const raw = String(idOrText || '').trim();
  if (!raw || raw === 'none' || raw === '无水印') return '';
  const hit = WATERMARK_PRESETS.find((p) => p.id === raw);
  return hit && hit.id !== 'none' ? hit.label : raw;
}

/**
 * 按客户填写的公司、使用人、用途拼出水印。空身份不盖通用套话。
 */
export function composeWatermark(opts = {}) {
  if (typeof opts === 'string') {
    if (opts === 'none' || opts === '无水印') return '';
    if (opts === 'confidential' || opts === 'bank' || opts === 'draft' || opts === 'for' || opts === 'custom') return '';
    return resolveWatermarkText(opts);
  }
  const company = String(opts.company || opts.org || '').trim();
  const person = String(opts.person || opts.name || opts.user || '').trim();
  const note = String(opts.note || opts.extra || opts.purpose || '').trim();
  const tone = String(opts.tone || opts.id || 'confidential');
  if (opts.text) return String(opts.text).trim();
  if (tone === 'none') return '';
  const who = [company, person].filter(Boolean).join(' ');
  if (!who && !note) return '';
  if (tone === 'custom') return note && who ? `${who} ${note}` : (note || who);
  if (tone === 'for' || tone === 'bank') {
    const core = who || note;
    const tail = who && note ? ` ${note}` : '';
    return `仅供 ${core} 使用${tail} 严禁外传`;
  }
  if (tone === 'draft') {
    return who ? `${who} 草案 仅供参考${note ? ` ${note}` : ''}` : `${note} 草案 仅供参考`;
  }
  if (who && note) return `${who} ${note} 严禁外传`;
  if (who) return `${who} 内部机密 严禁外传`;
  return `${note} 严禁外传`;
}

function pickOrgName(t) {
  const suffixes = ['事务所', '公司', '集团', '医院', '银行'];
  for (const s of suffixes) {
    let from = 0;
    while (from < t.length) {
      const idx = t.indexOf(s, from);
      if (idx < 0) break;
      let prefix = t.slice(Math.max(0, idx - 6), idx);
      prefix = prefix.replace(/.*(写上|水印|给|的)/, '');
      const name = prefix.slice(-4);
      if (name.length >= 2) return name + s;
      from = idx + s.length;
    }
  }
  return '';
}

export function parseWatermarkFromSpeech(message) {
  const t = String(message || '').trim();
  const quoted = t.match(/[「"']([^」"']{2,40})[」"']/);
  if (quoted) return { text: quoted[1], tone: 'custom' };
  const company = (t.match(/公司[是为:：]?\s*([^\s，,。的]+)/)?.[1]) || pickOrgName(t) || '';
  let person = (t.match(/(?:使用人|接收人|姓名)[是为:：]?\s*([^\s，,。的]+)/)
    || t.match(/(?:发给|给)\s*([^\s，,。盖的]{2,8})/)
    || t.match(/仅供\s*([^\s使用]{1,12})\s*使用/))?.[1] || '';
  if (person && /公司|集团|水印|内部/.test(person)) person = '';
  const note = (t.match(/补充[是为:：]?\s*(.+)$/) || t.match(/用途[是为:：]?\s*([^\s，,。]+)/))?.[1] || '';
  let tone = 'confidential';
  if (/仅供|给.+用|发给/.test(t)) tone = 'for';
  else if (/草案|草稿|参考/.test(t)) tone = 'draft';
  else if (/机密|外传/.test(t)) tone = 'confidential';
  return { company, person, note, tone };
}

export function applyWatermarkOverlay(host, text, opacity = 0.28) {
  if (!host) return;
  host.innerHTML = '';
  const label = resolveWatermarkText(text);
  if (!label) {
    host.hidden = true;
    host.setAttribute('hidden', '');
    return;
  }
  host.hidden = false;
  host.removeAttribute('hidden');
  generateWatermarkGrid(label, 20).forEach((item) => {
    const el = document.createElement('div');
    el.className = 'watermark-item';
    el.style.opacity = String(opacity);
    el.textContent = item.text;
    host.appendChild(el);
  });
}
