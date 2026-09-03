import { renderAsync } from 'docx-preview';
import { applyWatermarkOverlay, composeWatermark, maskSensitiveText } from '../engine/watermarkEngine.js';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  HeadingLevel,
  AlignmentType,
  WidthType,
  BorderStyle,
  Header,
  Footer,
  PageNumber
} from 'docx';

function exec(cmd, value) {
  document.execCommand(cmd, false, value);
}

export function createWordDoc() {
  const root = document.createElement('div');
  root.className = 'wd-root';
  root.innerHTML = `
    <div class="wd-page-host">
      <div class="wd-paper" id="wdPaper" contenteditable="true" spellcheck="false"></div>
      <div class="watermark-overlay" id="wdWm" hidden></div>
    </div>
  `;
  const paper = root.querySelector('#wdPaper');
  const wmEl = root.querySelector('#wdWm');
  let watermarkKey = '';

  function focus() {
    paper.focus({ preventScroll: true });
  }

  function beautifyLocal(payload) {
    const theme = payload?.headerFill || '#1f4e79';
    paper.querySelectorAll('.wd-toc, .wd-redhead').forEach((el) => el.remove());
    const keepTables = paper.querySelector('table');
    const lines = (paper.innerText || '')
      .replace(/\r/g, '')
      .split('\n')
      .map((s) => cleanVisibleLine(s))
      .filter((s) => s && s !== '目录' && !/^目录$/.test(s));
    const outline = looksLikeOutline(lines);
    if (outline && !keepTables) {
      paper.innerHTML = buildWordLayout(lines, theme);
    }
    applyWordTheme(paper, theme);
  }

  function looksLikeOutline(lines) {
    if (lines.length < 3) return false;
    const sectionHits = lines.filter((l) => isSectionLine(l)).length;
    const listHits = lines.filter((l) => isListLine(l)).length;
    return sectionHits >= 1 || listHits >= 2 || lines[0].length <= 24;
  }

  function isSectionLine(line) {
    const t = line.replace(/[：:]\s*$/, '');
    if (/^(项目名称|项目|部署|注意|备注|说明|概述|背景|目标|环境|账号|清单|交接|目录)$/.test(t)) return true;
    if (/^.{1,10}[：:]$/.test(line) && !isListLine(line)) return true;
    return false;
  }

  function isListLine(line) {
    return /^(\d+、|\d+\.\s*|（\d+）|\(\d+\)|[•·●]|[-—])/.test(line);
  }

  function stripListMark(line) {
    return line.replace(/^(\d+、|\d+\.\s*|（\d+）|\(\d+\)|[•·●]\s*|[-—]\s*)/, '').trim();
  }

  function cleanVisibleLine(s) {
    return String(s || '')
      .replace(/<\/?[a-zA-Z][^>]*>/g, '')
      .replace(/&nbsp;/g, ' ')
      .replace(/class="wd-[^"]*"/g, '')
      .replace(/\s+/g, ' ')
      .trim();
  }

  function formatInline(text) {
    const clean = cleanVisibleLine(text);
    return escapeHtml(clean).replace(/(账号|密码|IP|手机号?|接口)[：:]/g, '<strong>$1：</strong>');
  }

  function buildWordLayout(lines, theme) {
    let title = (lines[0] || '文档').replace(/^目录\s*/, '');
    const kicker = /交接|备注|部署|项目/.test(title) ? '项目交接文档' : '内部文档';
    const rest = lines.slice(1);
    const parts = [];
    let i = 0;
    while (i < rest.length) {
      const line = rest[i];
      if (isSectionLine(line)) {
        const name = line.replace(/[：:]\s*$/, '');
        i += 1;
        const items = [];
        const paras = [];
        while (i < rest.length && !isSectionLine(rest[i])) {
          if (isListLine(rest[i])) items.push(stripListMark(rest[i]));
          else paras.push(rest[i]);
          i += 1;
        }
        parts.push({ name, items, paras, warn: /注意|警告|风险/.test(name) });
        continue;
      }
      if (isListLine(line)) {
        const items = [];
        while (i < rest.length && isListLine(rest[i]) && !isSectionLine(rest[i])) {
          items.push(stripListMark(rest[i]));
          i += 1;
        }
        parts.push({ name: '', items, paras: [], warn: false });
        continue;
      }
      parts.push({ name: '', items: [], paras: [line], warn: false });
      i += 1;
    }

    let html = `<header class="wd-doc-head" data-theme="${escapeHtml(theme)}">
      <div class="wd-doc-kicker">${escapeHtml(kicker)}</div>
      <h1>${escapeHtml(title)}</h1>
      <div class="wd-doc-rule"></div>
    </header>`;
    parts.forEach((sec) => {
      if (sec.warn) {
        html += `<aside class="wd-callout"><div class="wd-callout-t">${escapeHtml(sec.name || '注意')}</div>`;
        sec.paras.forEach((p) => { html += `<p>${formatInline(p)}</p>`; });
        if (sec.items.length) {
          html += `<ol class="wd-ol">${sec.items.map((it) => `<li>${formatInline(it)}</li>`).join('')}</ol>`;
        }
        html += '</aside>';
        return;
      }
      if (sec.name) html += `<h2>${escapeHtml(sec.name)}</h2>`;
      if (sec.items.length) {
        html += `<ol class="wd-ol">${sec.items.map((it) => `<li>${formatInline(it)}</li>`).join('')}</ol>`;
      }
      sec.paras.forEach((p) => { html += `<p class="wd-body">${formatInline(p)}</p>`; });
    });
    return html;
  }

  function applyWordTheme(rootEl, theme) {
    rootEl.style.fontFamily = 'SimSun, 宋体, Songti SC, serif';
    rootEl.style.fontSize = '14pt';
    rootEl.style.lineHeight = '1.75';
    rootEl.style.color = '#1e293b';
    rootEl.style.background = '#fff';
    rootEl.querySelectorAll('.wd-doc-head, .wd-doc-rule').forEach((el) => {
      el.style.setProperty('--wd-theme', theme);
    });
    rootEl.querySelectorAll('h1').forEach((h) => {
      h.style.fontFamily = 'Microsoft YaHei, 微软雅黑, SimHei, sans-serif';
      h.style.fontSize = '22pt';
      h.style.textAlign = 'center';
      h.style.textIndent = '0';
      h.style.color = theme;
      h.style.fontWeight = '800';
      h.style.margin = '0.15em 0 0.35em';
      h.style.letterSpacing = '3px';
      h.style.lineHeight = '1.35';
    });
    rootEl.querySelectorAll('h2').forEach((h) => {
      h.style.fontFamily = 'Microsoft YaHei, 微软雅黑, SimHei, sans-serif';
      h.style.fontSize = '13.5pt';
      h.style.textIndent = '0';
      h.style.color = '#fff';
      h.style.background = theme;
      h.style.margin = '1.15em 0 0.55em';
      h.style.padding = '0.28em 0.7em';
      h.style.fontWeight = '700';
      h.style.letterSpacing = '1px';
      h.style.borderRadius = '2px';
    });
    rootEl.querySelectorAll('h3').forEach((h) => {
      h.style.fontFamily = 'Microsoft YaHei, sans-serif';
      h.style.fontSize = '12pt';
      h.style.textIndent = '0';
      h.style.color = theme;
    });
    rootEl.querySelectorAll('p.wd-body, .wd-callout p').forEach((p) => {
      p.style.textIndent = '0';
      p.style.margin = '0.35em 0';
      p.style.textAlign = 'left';
      p.style.fontSize = '12pt';
      p.style.lineHeight = '1.7';
    });
    rootEl.querySelectorAll('p:not(.wd-body):not(.wd-doc-kicker)').forEach((p) => {
      if (p.closest('li,aside,header,table')) return;
      p.style.fontSize = '12pt';
      p.style.lineHeight = '1.75';
      p.style.textIndent = '2em';
      p.style.margin = '0.45em 0';
      p.style.textAlign = 'justify';
    });
    rootEl.querySelectorAll('table').forEach((table) => {
      table.style.width = '100%';
      table.style.borderCollapse = 'collapse';
      table.style.margin = '1em 0';
      table.style.fontSize = '12pt';
      [...table.rows].forEach((tr, i) => {
        [...tr.cells].forEach((td) => {
          td.style.border = '1px solid #94a3b8';
          td.style.padding = '8px 10px';
          td.style.textIndent = '0';
          if (i === 0) {
            td.style.background = theme;
            td.style.color = '#fff';
            td.style.fontWeight = '700';
            td.style.textAlign = 'center';
            td.style.fontFamily = 'Microsoft YaHei, sans-serif';
          } else if (i % 2 === 0) td.style.background = '#f8fafc';
          else td.style.background = '#fff';
        });
      });
    });
  }

  function askOptions(title, sample) {
    const raw = window.prompt(title, sample);
    if (raw == null) return null;
    const options = raw.split(/[,，]/).map((s) => s.trim()).filter(Boolean);
    return options.length ? options : null;
  }

  function escapeHtml(s) {
    return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }

  paper.addEventListener('click', (e) => {
    const check = e.target.closest('.wd-check');
    if (check) {
      e.preventDefault();
      const on = check.dataset.checked === 'true';
      check.dataset.checked = on ? 'false' : 'true';
      check.textContent = on ? '☐' : '☑';
      return;
    }
    const yn = e.target.closest('.wd-yesno button');
    if (yn) {
      e.preventDefault();
      yn.parentElement.querySelectorAll('button').forEach((b) => b.classList.remove('on'));
      yn.classList.add('on');
      return;
    }
    const radio = e.target.closest('.wd-radio');
    if (radio) {
      e.preventDefault();
      const group = radio.closest('.wd-radio-group');
      group.querySelectorAll('.wd-radio').forEach((item) => {
        item.dataset.on = 'false';
        const mk = item.querySelector('.mk');
        if (mk) mk.textContent = '○';
      });
      radio.dataset.on = 'true';
      const mk = radio.querySelector('.mk');
      if (mk) mk.textContent = '●';
    }
  });

  async function loadDocx(buffer) {
    paper.innerHTML = '';
    try {
      await renderAsync(buffer, paper, null, {
        className: 'wd-docx',
        inWrapper: true,
        ignoreWidth: false,
        experimental: true
      });
    } catch {
      paper.innerHTML = '<p>无法解析该 Word 文件，已打开空白文档。</p>';
    }
    paper.contentEditable = 'true';
    focus();
  }

  function loadHtml(html) {
    paper.innerHTML = html;
    paper.contentEditable = 'true';
    focus();
  }

  function command(name, payload) {
    focus();
    const map = {
      bold: () => exec('bold'),
      italic: () => exec('italic'),
      underline: () => exec('underline'),
      strike: () => exec('strikeThrough'),
      subscript: () => exec('subscript'),
      superscript: () => exec('superscript'),
      fontFamily: () => exec('fontName', payload),
      fontSize: () => exec('fontSize', payload),
      fontColor: () => exec('foreColor', payload),
      highlight: () => exec('hiliteColor', payload),
      align: () => {
        const m = { left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull' };
        exec(m[payload] || 'justifyLeft');
      },
      bullet: () => exec('insertUnorderedList'),
      number: () => exec('insertOrderedList'),
      indent: () => exec('indent'),
      outdent: () => exec('outdent'),
      style: () => exec('formatBlock', payload),
      undo: () => exec('undo'),
      redo: () => exec('redo'),
      removeFormat: () => exec('removeFormat'),
      hr: () => exec('insertHorizontalRule'),
      link: () => {
        const url = payload || window.prompt('输入链接地址', 'https://');
        if (url) exec('createLink', url);
      },
      unlink: () => exec('unlink'),
      pageBreak: () => exec('insertHTML', '<p style="page-break-before:always;"></p>'),
      insertText: () => {
        const t = payload?.text || payload || '';
        exec('insertHTML', `<p>${escapeHtml(String(t))}</p>`);
      },
      table: () => {
        const rows = Number(payload?.rows) || 3;
        const cols = Number(payload?.cols) || 3;
        let t = '<table class="wd-table"><tbody>';
        for (let r = 0; r < rows; r++) {
          t += '<tr>';
          for (let c = 0; c < cols; c++) t += '<td>&nbsp;</td>';
          t += '</tr>';
        }
        t += '</tbody></table><p></p>';
        exec('insertHTML', t);
      },
      image: () => {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = 'image/*';
        input.onchange = () => {
          const file = input.files?.[0];
          if (!file) return;
          const reader = new FileReader();
          reader.onload = () => exec('insertImage', reader.result);
          reader.readAsDataURL(file);
        };
        input.click();
      },
      pageColor: () => { paper.style.background = payload || '#fff'; },
      pageBorder: () => {
        paper.style.border = paper.style.border ? '' : '3px double #1d4ed8';
        paper.style.padding = '24mm 22mm';
      },
      columns: () => {
        paper.style.columnCount = payload || '1';
        paper.style.columnGap = '1.5rem';
      },
      orientation: () => {
        paper.classList.toggle('landscape', payload === 'landscape');
      },
      toc: () => {
        const heads = [...paper.querySelectorAll('h1,h2,h3')];
        const html = `<div class="wd-toc"><strong>目录</strong>${heads.map((h, i) =>
          `<div class="lv${h.tagName[1]}">${i + 1}. ${h.textContent}</div>`).join('')}</div>`;
        exec('insertHTML', html);
      },
      redhead: () => {
        exec('insertHTML', `<div class="wd-redhead"><div class="t">文件标题</div><div class="s">★ 公文版头 ★</div></div>`);
      },
      checkbox: () => exec('insertHTML', '<span class="wd-check" contenteditable="false" data-checked="false">☐</span>&nbsp;'),
      yesno: () => exec('insertHTML', '<span class="wd-yesno" contenteditable="false"><button type="button">是</button><button type="button">否</button></span>&nbsp;'),
      radio: () => {
        const options = Array.isArray(payload?.options) ? payload.options : askOptions('单选项（逗号分隔，只能选一个）', '同意,不同意,弃权');
        if (!options) return;
        exec('insertHTML', `<div class="wd-radio-group" contenteditable="false">${options.map((o) =>
          `<div class="wd-radio" data-on="false"><span class="mk">○</span> ${escapeHtml(o)}</div>`).join('')}</div><p></p>`);
      },
      dropdown: () => {
        const options = Array.isArray(payload?.options) ? payload.options : askOptions('下拉选项（逗号分隔）', '是,否,待定');
        if (!options) return;
        exec('insertHTML', `<select class="wd-select" contenteditable="false">${options.map((o) =>
          `<option>${escapeHtml(o)}</option>`).join('')}</select>&nbsp;`);
      },
      multiselect: () => {
        const options = Array.isArray(payload?.options) ? payload.options : askOptions('多选项（逗号分隔，可同时勾多个）', '选项A,选项B,选项C');
        if (!options) return;
        exec('insertHTML', `<div class="wd-multi" contenteditable="false">${options.map((o) =>
          `<label><input type="checkbox"/> ${escapeHtml(o)}</label>`).join('')}</div><p></p>`);
      },
      dropcap: () => exec('insertHTML', '<span class="wd-drop"> </span>'),
      beautify: () => beautifyLocal(payload),
      watermark: () => {
        const shown = composeWatermark(payload);
        watermarkKey = shown;
        applyWatermarkOverlay(wmEl, shown, payload?.opacity ?? 0.22);
      },
      maskSensitive: () => {
        const walk = (node) => {
          if (node.nodeType === Node.TEXT_NODE) {
            node.textContent = maskSensitiveText(node.textContent);
            return;
          }
          node.childNodes.forEach(walk);
        };
        walk(paper);
      }
    };
    if (map[name]) map[name]();
  }

  function walkRuns(node, runs) {
    if (node.nodeType === Node.TEXT_NODE) {
      const t = node.textContent || '';
      if (!t) return;
      const parent = node.parentElement;
      runs.push(new TextRun({
        text: t,
        bold: !!parent?.closest('b,strong'),
        italics: !!parent?.closest('i,em'),
        underline: parent?.closest('u') ? {} : undefined,
        strike: !!parent?.closest('s,strike,del'),
        color: cssColor(parent),
        size: 28
      }));
      return;
    }
    if (node.nodeType !== Node.ELEMENT_NODE) return;
    if (node.tagName === 'BR') {
      runs.push(new TextRun({ break: 1 }));
      return;
    }
    node.childNodes.forEach((ch) => walkRuns(ch, runs));
  }

  function cssColor(el) {
    if (!el) return undefined;
    const c = getComputedStyle(el).color;
    const m = c.match(/rgba?\((\d+),\s*(\d+),\s*(\d+)/);
    if (!m) return undefined;
    const hex = [m[1], m[2], m[3]].map((n) => Number(n).toString(16).padStart(2, '0')).join('');
    if (hex === '000000' || hex === '1e293b') return undefined;
    return hex;
  }

  function paraAlign(el) {
    const a = getComputedStyle(el).textAlign;
    if (a === 'center') return AlignmentType.CENTER;
    if (a === 'right' || a === 'end') return AlignmentType.RIGHT;
    if (a === 'justify') return AlignmentType.BOTH;
    return AlignmentType.LEFT;
  }

  async function exportDocx(filename = '文档.docx') {
    const children = [];
    const blocks = [...paper.querySelectorAll('h1,h2,h3,p,li,table,aside,header,div.wd-toc,div.wd-redhead')];
    const seen = new Set();
    const nodes = blocks.length ? blocks : [...paper.children];
    nodes.forEach((el) => {
      if (seen.has(el)) return;
      seen.add(el);
      if (el.tagName === 'TABLE') {
        const rows = [...el.rows].map((tr) => new TableRow({
          children: [...tr.cells].map((td) => new TableCell({
            width: { size: 2000, type: WidthType.DXA },
            children: [new Paragraph({ children: [new TextRun(td.textContent || '')] })]
          }))
        }));
        children.push(new Table({
          rows,
          width: { size: 9000, type: WidthType.DXA },
          borders: {
            top: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
            bottom: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
            left: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
            right: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
            insideH: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' },
            insideV: { style: BorderStyle.SINGLE, size: 4, color: '94A3B8' }
          }
        }));
        return;
      }
      const runs = [];
      walkRuns(el, runs);
      if (!runs.length) runs.push(new TextRun(''));
      const heading = el.tagName === 'H1' ? HeadingLevel.HEADING_1
        : el.tagName === 'H2' ? HeadingLevel.HEADING_2
          : el.tagName === 'H3' ? HeadingLevel.HEADING_3
            : undefined;
      children.push(new Paragraph({
        heading,
        alignment: paraAlign(el),
        children: runs
      }));
    });

    const doc = new Document({
      sections: [{
        properties: {},
        headers: { default: new Header({ children: [new Paragraph('')] }) },
        footers: {
          default: new Footer({
            children: [new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [new TextRun('第 '), new TextRun({ children: [PageNumber.CURRENT] }), new TextRun(' 页')]
            })]
          })
        },
        children: children.length ? children : [new Paragraph('')]
      }]
    });
    const blob = await Packer.toBlob(doc);
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = filename;
    a.click();
  }

  loadHtml(`
    <h1>文档标题</h1>
    <p>在这里像 Word 一样直接打字。选中文字后，用上方功能区：加粗、对齐、标题、列表、插表格。</p>
    <h2>第一节</h2>
    <p>拖入 .docx 可打开原文件继续改。导出得到标准 Word 文档。</p>
  `);

  function getSnapshot() {
    return {
      text: (paper.innerText || '').slice(0, 2500),
      html: (paper.innerHTML || '').slice(0, 4000)
    };
  }

  return {
    element: root,
    loadDocx,
    loadHtml,
    command,
    exportDocx,
    getSnapshot,
    focus
  };
}
