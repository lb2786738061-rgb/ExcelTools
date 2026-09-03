/**
 * Office 智能效率工具箱 - Word 文档重构排版与生成引擎
 * 支持基于 ControlPanel 配置参数对 Word 进行格式重构、前端富格式渲染与 .docx 二进制导出
 */

import { renderAsync } from 'docx-preview';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  Table,
  TableRow,
  TableCell,
  AlignmentType,
  WidthType,
  HeadingLevel,
  BorderStyle
} from 'docx';
import JSZip from 'jszip';
import { generateWatermarkGrid } from './watermarkEngine.js';

/**
 * 在前端浏览器中渲染预览 Word (.docx) 文件或基于控制选项应用样式
 * @param {ArrayBuffer} arrayBuffer - 原始 docx 文件数据
 * @param {HTMLElement} container - 渲染的目标容器节点
 * @param {Object} options - 美化排版控制参数
 */
export async function renderWordDocument(arrayBuffer, container, options = {}) {
  if (!container) return;
  container.innerHTML = '';

  try {
    // 渲染原始 docx
    await renderAsync(arrayBuffer, container, null, {
      className: 'rendered-word-document',
      inWrapper: true,
      ignoreWidth: false,
      ignoreHeight: false,
      experimental: true
    });

    // 动态把 ControlPanel 中的美化样式应用到 DOM 节点上
    applyStylesToWordDom(container, options);
  } catch (err) {
    console.warn('Word 自动解析预处理，回退到动态富文本样式排版渲染模式', err);
    renderFormattedWordPreview(container, options);
  }
}

/**
 * 将美化排版控制参数动态应用于渲染后的 Word DOM 视图
 */
function applyStylesToWordDom(container, options) {
  const {
    fontFamily = 'system',
    fontColor = 'default',
    fontSize = 14,
    textAlign = 'left',
    lineSpacing = '1.15',
    fontBold = false,
    fontItalic = false,
    fontUnderline = false,
    fontStrike = false,
    fontStyleItalic = false,
    firstLineIndent = false,
    indentStyle = firstLineIndent ? 'first' : 'none',
    listStyle = 'none',
    bulletList = false,
    headerText = '',
    footerText = '',
    showPageNumber = false,
    pageBorder = false,
    watermark = 'none',
    watermarkOpacity = 0.15,
    pageBgColor = 'default',
    paperSize = 'A4',
    orientation = 'portrait',
    margins = 'normal',
    columns = '1',
    spaceBefore = '6',
    spaceAfter = '6',
    highlightColor = 'none',
    h1FontSize = '24',
    headingColor = 'default',
    autoToc = false,
    autoPageBreak = false,
    officialRedHead = false,
    dropCap = false
  } = options;

  const fontMap = {
    system: '-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
    SimSun: 'SimSun, "宋体", serif',
    SimHei: 'SimHei, "黑体", sans-serif',
    KaiTi: 'KaiTi, "楷体", serif',
    Arial: 'Arial, sans-serif'
  };

  const targetFont = fontMap[fontFamily] || fontMap.system;
  const wrapper = container.querySelector('.rendered-word-document') || container;
  wrapper.style.position = 'relative';
  wrapper.style.fontFamily = targetFont;
  wrapper.style.fontSize = `${fontSize}px`;
  wrapper.style.lineHeight = String(lineSpacing);
  wrapper.style.margin = '0 auto';
  wrapper.style.backgroundColor = pageBgColor !== 'default' ? pageBgColor : '#fff';
  wrapper.style.color = '#1e293b';

  const portraitWidths = { A4: '210mm', A3: '297mm', Letter: '216mm' };
  const landscapeWidths = { A4: '297mm', A3: '420mm', Letter: '279mm' };
  wrapper.style.maxWidth = orientation === 'landscape'
    ? (landscapeWidths[paperSize] || '297mm')
    : (portraitWidths[paperSize] || '210mm');
  wrapper.style.padding = margins === 'narrow' ? '12mm' : margins === 'wide' ? '32mm' : '22mm';
  wrapper.style.columnCount = columns && columns !== '1' ? String(columns) : '1';
  wrapper.style.columnGap = columns && columns !== '1' ? '1.5rem' : 'normal';

  if (pageBorder) {
    wrapper.style.border = officialRedHead ? '3px double #dc2626' : '3px double #2563eb';
    wrapper.style.borderRadius = '4px';
  } else {
    wrapper.style.border = '1px solid #e2e8f0';
  }

  wrapper.querySelectorAll('[data-office-beautify]').forEach((el) => el.remove());
  wrapper.querySelectorAll('.watermark-overlay').forEach((el) => el.remove());

  const headingEls = [...wrapper.querySelectorAll('h1, h2, h3')];
  if (officialRedHead) {
    const banner = document.createElement('div');
    banner.dataset.officeBeautify = 'redhead';
    banner.className = 'office-redhead';
    banner.innerHTML = `
      <div class="office-redhead-title">中华人民共和国行政公文</div>
      <div class="office-redhead-star">★ 国家标准党政机关公文格式 ★</div>
      <div class="office-redhead-meta"><span>发文字号：自动生成</span><span>签发人：——</span></div>
    `;
    wrapper.insertBefore(banner, wrapper.firstChild);
  }

  if (autoToc && headingEls.length > 0) {
    const toc = document.createElement('div');
    toc.dataset.officeBeautify = 'toc';
    toc.className = 'office-toc';
    toc.innerHTML = `<div class="office-toc-title">自动提取目录</div>` + headingEls.map((el, i) => {
      const pad = el.tagName === 'H1' ? '0' : el.tagName === 'H2' ? '1.2rem' : '2.2rem';
      return `<div class="office-toc-item" style="padding-left:${pad}">${i + 1}. ${el.textContent}</div>`;
    }).join('');
    const insertBefore = wrapper.querySelector('[data-office-beautify="redhead"]')?.nextSibling || wrapper.firstChild;
    wrapper.insertBefore(toc, insertBefore);
  }

  const hColor = headingColor !== 'default' ? headingColor : (officialRedHead ? '#dc2626' : '#2563eb');
  headingEls.forEach((el) => {
    el.style.color = hColor;
    if (el.tagName === 'H1') {
      el.style.fontSize = `${h1FontSize}px`;
      if (autoPageBreak) el.style.pageBreakBefore = 'always';
    }
  });

  wrapper.style.fontWeight = fontBold ? '700' : '';
  const italicOn = fontItalic || fontStyleItalic;
  wrapper.style.fontStyle = italicOn ? 'italic' : '';
  const decorations = [];
  if (fontUnderline) decorations.push('underline');
  if (fontStrike) decorations.push('line-through');
  wrapper.style.textDecoration = decorations.join(' ');

  const effectiveList = listStyle !== 'none' ? listStyle : (bulletList ? 'bullet' : 'none');
  wrapper.classList.toggle('office-number-list', effectiveList === 'number');

  const paragraphs = wrapper.querySelectorAll('p, li');
  paragraphs.forEach((p, idx) => {
    p.style.textAlign = textAlign || 'left';
    p.style.marginTop = `${spaceBefore}pt`;
    p.style.marginBottom = `${spaceAfter}pt`;
    if (fontColor !== 'default') p.style.color = fontColor;
    p.classList.remove('office-bullet', 'office-number-item', 'office-hanging', 'office-first-indent');
    if (indentStyle === 'hanging') p.classList.add('office-hanging');
    else if (indentStyle === 'first' || firstLineIndent) p.classList.add('office-first-indent');
    if (effectiveList === 'bullet' && p.tagName === 'P') p.classList.add('office-bullet');
    if (effectiveList === 'number' && p.tagName === 'P') p.classList.add('office-number-item');
    p.classList.toggle('office-highlight', highlightColor && highlightColor !== 'none');
    if (highlightColor && highlightColor !== 'none') p.style.backgroundColor = highlightColor;
    else p.style.backgroundColor = '';
    p.classList.toggle('office-dropcap', !!dropCap && idx === 0 && p.tagName === 'P');
  });

  if (headerText || showPageNumber || footerText) {
    if (headerText) {
      const hdr = document.createElement('div');
      hdr.dataset.officeBeautify = 'header';
      hdr.className = 'office-page-header';
      hdr.textContent = headerText;
      wrapper.insertBefore(hdr, wrapper.firstChild);
    }
    if (footerText || showPageNumber) {
      const ftr = document.createElement('div');
      ftr.dataset.officeBeautify = 'footer';
      ftr.className = 'office-page-footer';
      ftr.textContent = `${footerText}${footerText && showPageNumber ? '  ·  ' : ''}${showPageNumber ? '第 1 页' : ''}`;
      wrapper.appendChild(ftr);
    }
  }

  if (watermark && watermark !== 'none' && watermark !== '无水印') {
    const labelMap = {
      confidential: '内部机密 严禁外传',
      bank: '仅供办理业务使用',
      draft: '草案样本 仅供参考'
    };
    const textLabel = labelMap[watermark] || watermark;
    const overlay = document.createElement('div');
    overlay.className = 'watermark-overlay';
    overlay.dataset.officeBeautify = 'watermark';
    generateWatermarkGrid(textLabel, 12).forEach((item) => {
      const itemEl = document.createElement('div');
      itemEl.className = 'watermark-item';
      itemEl.style.opacity = watermarkOpacity;
      itemEl.textContent = item.text;
      overlay.appendChild(itemEl);
    });
    wrapper.appendChild(overlay);
  }
}

/**
 * 离线生成与渲染 Word 文档预览
 */
export function renderFormattedWordPreview(container, options = {}) {
  container.innerHTML = '';
  const wrapper = document.createElement('div');
  wrapper.className = 'rendered-word-document';

  const sampleTitle = options.officialRedHead ? "关于印发《Office 智能美化标准规范》的通知" : "公文/报告标准化处理示例";

  const sampleContent = [
    { title: "第一章：项目背景与目标概述", type: "h1" },
    { text: "本工具致力于打造全能型 Office 移动端与桌面端智能效率神器。通过一键格式化规整、自动化样式匹配、公式小计运算与防泄密水印处理，极大地简化了用户的文档编辑流程。", type: "p" },
    { title: "1.1 移动办公排版核心痛点分析", type: "h2" },
    { text: "移动端屏幕尺寸受限，手动拖拽文本框或调整表格列宽极易导致版式打乱。本自动化规则引擎可根据文档内容智能计算最优列宽与字号。", type: "p" },
    { title: "第二章：排版规范与样式说明", type: "h1" },
    { text: "规范字符与行距，提升小屏阅读体验。", type: "p" },
    { text: "自动加盖防伪水印与敏感数据遮挡打码。", type: "p" },
    { text: "全面支持通用 Microsoft Word (.docx) 格式规范。", type: "p" }
  ];

  let html = `<h1>${sampleTitle}</h1>`;
  sampleContent.forEach((item) => {
    if (item.type === 'h1') html += `<h1>${item.title}</h1>`;
    else if (item.type === 'h2') html += `<h2>${item.title}</h2>`;
    else html += `<p>${item.text}</p>`;
  });

  wrapper.innerHTML = html;
  container.appendChild(wrapper);
  applyStylesToWordDom(container, options);
}

function xmlEscape(value) {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/"/g, '&quot;');
}

function mapWordFont(fontFamily) {
  if (fontFamily === 'SimSun') return '宋体';
  if (fontFamily === 'SimHei') return '黑体';
  if (fontFamily === 'KaiTi') return '楷体';
  if (fontFamily === 'Arial') return 'Arial';
  return '微软雅黑';
}

function wordAlign(textAlign) {
  if (textAlign === 'center') return 'center';
  if (textAlign === 'right') return 'right';
  if (textAlign === 'justify') return 'both';
  return 'left';
}

function hexColor(value, fallback) {
  if (value && String(value).startsWith('#')) return String(value).slice(1);
  return fallback;
}

function paperTwips(options) {
  const sizes = {
    A4: { w: 11906, h: 16838 },
    A3: { w: 16838, h: 23811 },
    Letter: { w: 12240, h: 15840 }
  };
  const size = sizes[options.paperSize] || sizes.A4;
  if (options.orientation === 'landscape') return { w: size.h, h: size.w, orient: 'landscape' };
  return { w: size.w, h: size.h, orient: 'portrait' };
}

function marginTwips(margins) {
  if (margins === 'narrow') return 720;
  if (margins === 'wide') return 1800;
  return 1440;
}

function patchWordStylesXml(xml, options) {
  const font = xmlEscape(mapWordFont(options.fontFamily));
  const sz = String((options.fontSize || 14) * 2);
  const h1sz = String((Number(options.h1FontSize) || 24) * 2);
  const line = String(Math.round(parseFloat(options.lineSpacing || '1.15') * 240));
  const indent = options.firstLineIndent ? '<w:ind w:firstLineChars="200"/>' : '';
  const color = (options.fontColor && options.fontColor !== 'default' && String(options.fontColor).startsWith('#'))
    ? `<w:color w:val="${String(options.fontColor).slice(1)}"/>`
    : '';
  const italic = options.fontStyleItalic ? '<w:i/><w:iCs/>' : '';
  const headingHex = hexColor(options.headingColor, options.officialRedHead ? 'DC2626' : '2563EB');
  const rPr = `<w:rPr><w:rFonts w:ascii="${font}" w:hAnsi="${font}" w:eastAsia="${font}"/><w:sz w:val="${sz}"/><w:szCs w:val="${sz}"/>${color}${italic}</w:rPr>`;
  const pPr = `<w:pPr><w:spacing w:line="${line}" w:lineRule="auto" w:before="${Number(options.spaceBefore || 6) * 20}" w:after="${Number(options.spaceAfter || 6) * 20}"/>${indent}<w:jc w:val="${wordAlign(options.textAlign)}"/></w:pPr>`;
  const defaults = `<w:docDefaults><w:rPrDefault>${rPr}</w:rPrDefault><w:pPrDefault>${pPr}</w:pPrDefault></w:docDefaults>`;

  let next = xml;
  if (/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.test(next)) {
    next = next.replace(/<w:docDefaults>[\s\S]*?<\/w:docDefaults>/, defaults);
  } else {
    next = next.replace(/<w:styles\b[^>]*>/, (match) => `${match}${defaults}`);
  }

  const headingStyle = (id, name, size) =>
    `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="240" w:after="120"/></w:pPr><w:rPr><w:rFonts w:ascii="${font}" w:eastAsia="${font}" w:hAnsi="${font}"/><w:b/><w:color w:val="${headingHex}"/><w:sz w:val="${size}"/><w:szCs w:val="${size}"/></w:rPr></w:style>`;

  if (!/w:styleId="Heading1"/.test(next)) {
    next = next.replace(/<\/w:styles>/, `${headingStyle('Heading1', 'heading 1', h1sz)}${headingStyle('Heading2', 'heading 2', String(Math.max(24, Number(h1sz) - 8)))}</w:styles>`);
  }

  return next;
}

function patchDocumentSectPr(xml, options, headerRelId) {
  const page = paperTwips(options);
  const mar = marginTwips(options.margins);
  const cols = options.columns && options.columns !== '1'
    ? `<w:cols w:num="${options.columns}" w:space="425"/>`
    : `<w:cols w:num="1"/>`;
  const headerRef = headerRelId
    ? `<w:headerReference w:type="default" r:id="${headerRelId}"/>`
    : '';
  const footerRef = (options.footerText || options.showPageNumber)
    ? `<w:footerReference w:type="default" r:id="rIdOfficeFt"/>`
    : '';
  const pgSz = `<w:pgSz w:w="${page.w}" w:h="${page.h}" w:orient="${page.orient}"/>`;
  const pgMar = `<w:pgMar w:top="${mar}" w:right="${mar}" w:bottom="${mar}" w:left="${mar}" w:header="720" w:footer="720" w:gutter="0"/>`;

  let next = xml;
  if (!/xmlns:r=/.test(next) && headerRelId) {
    next = next.replace(/<w:document\b([^>]*)>/, (m, attrs) =>
      `<w:document${attrs} xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">`);
  }

  if (/<w:pgSz\b[^/]*\/>/.test(next)) next = next.replace(/<w:pgSz\b[^/]*\/>/, pgSz);
  else if (/<w:sectPr\b[^>]*>/.test(next)) next = next.replace(/<w:sectPr\b[^>]*>/, (m) => `${m}${pgSz}`);

  if (/<w:pgMar\b[^/]*\/>/.test(next)) next = next.replace(/<w:pgMar\b[^/]*\/>/, pgMar);
  else if (/<w:sectPr\b[^>]*>/.test(next)) next = next.replace(/<w:sectPr\b[^>]*>/, (m) => `${m}${pgMar}`);

  if (/<w:cols\b[^/]*\/>/.test(next)) next = next.replace(/<w:cols\b[^/]*\/>/, cols);
  else if (/<w:sectPr\b[^>]*>/.test(next)) next = next.replace(/<w:sectPr\b[^>]*>/, (m) => `${m}${cols}`);

  if (headerRef) {
    if (/<w:headerReference\b[^/]*\/>/.test(next)) {
      next = next.replace(/<w:headerReference\b[^/]*\/>/, headerRef);
    } else if (/<w:sectPr\b[^>]*>/.test(next)) {
      next = next.replace(/<w:sectPr\b[^>]*>/, (m) => `${m}${headerRef}`);
    }
  }
  if (footerRef) {
    if (/<w:footerReference\b[^/]*\/>/.test(next)) {
      next = next.replace(/<w:footerReference\b[^/]*\/>/, footerRef);
    } else if (/<w:sectPr\b[^>]*>/.test(next)) {
      next = next.replace(/<w:sectPr\b[^>]*>/, (m) => `${m}${footerRef}`);
    }
  }

  return next;
}

function watermarkLabel(watermark) {
  const labelMap = {
    confidential: '内部机密 严禁外传',
    bank: '仅供办理业务使用',
    draft: '草案样本 仅供参考'
  };
  return labelMap[watermark] || watermark;
}

function pageFieldXml() {
  return `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
}

async function ensureWatermarkHeader(zip, options) {
  const wm = options.watermark && options.watermark !== 'none' && options.watermark !== '无水印'
    ? xmlEscape(watermarkLabel(options.watermark))
    : '';
  const headerText = xmlEscape(options.headerText || '');
  const needHeader = !!(wm || headerText);
  const needFooter = !!(options.footerText || options.showPageNumber);
  if (!needHeader && !needFooter) {
    return null;
  }

  const relsPath = 'word/_rels/document.xml.rels';
  let rels = await zip.file(relsPath)?.async('string')
    || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  let relId = null;

  if (needHeader) {
    const text = headerText || wm;
    const headerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:pPr><w:jc w:val="center"/></w:pPr>
    <w:r>
      <w:rPr><w:color w:val="${wm && !headerText ? 'C0C0C0' : '334155'}"/><w:sz w:val="22"/><w:b/></w:rPr>
      <w:t xml:space="preserve">${text}</w:t>
    </w:r>
  </w:p>
</w:hdr>`;
    zip.file('word/header_office_wm.xml', headerXml);

    let contentTypes = await zip.file('[Content_Types].xml')?.async('string');
    if (contentTypes && !contentTypes.includes('header_office_wm.xml')) {
      contentTypes = contentTypes.replace(
        '</Types>',
        '<Override PartName="/word/header_office_wm.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>'
      );
      zip.file('[Content_Types].xml', contentTypes);
    }

    relId = 'rIdOfficeWm';
    if (!rels.includes('header_office_wm.xml')) {
      rels = rels.replace('</Relationships>', `<Relationship Id="${relId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header_office_wm.xml"/></Relationships>`);
    }
  }

  zip.file(relsPath, rels);

  if (options.footerText || options.showPageNumber) {
    const footerBits = [];
    if (options.footerText) {
      footerBits.push(`<w:r><w:t xml:space="preserve">${xmlEscape(options.footerText)}</w:t></w:r>`);
    }
    if (options.showPageNumber) {
      if (footerBits.length) footerBits.push(`<w:r><w:t xml:space="preserve">  </w:t></w:r>`);
      footerBits.push(pageFieldXml());
    }
    const footerXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:pPr><w:jc w:val="center"/></w:pPr>${footerBits.join('')}</w:p>
</w:ftr>`;
    zip.file('word/footer_office.xml', footerXml);
    let types = await zip.file('[Content_Types].xml')?.async('string');
    if (types && !types.includes('footer_office.xml')) {
      types = types.replace('</Types>', '<Override PartName="/word/footer_office.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>');
      zip.file('[Content_Types].xml', types);
    }
    let rels2 = await zip.file(relsPath)?.async('string');
    if (rels2 && !rels2.includes('footer_office.xml')) {
      rels2 = rels2.replace('</Relationships>', '<Relationship Id="rIdOfficeFt" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer_office.xml"/></Relationships>');
      zip.file(relsPath, rels2);
    }
  }

  return relId;
}

function downloadBlob(blob, filename) {
  const link = document.createElement('a');
  link.href = URL.createObjectURL(blob);
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(link.href);
}

async function exportPatchedOriginalWord(arrayBuffer, options, filename) {
  const zip = await JSZip.loadAsync(arrayBuffer);
  const stylesFile = zip.file('word/styles.xml');
  if (stylesFile) {
    const xml = await stylesFile.async('string');
    zip.file('word/styles.xml', patchWordStylesXml(xml, options));
  }

  const headerRelId = await ensureWatermarkHeader(zip, options);
  const docFile = zip.file('word/document.xml');
  if (docFile) {
    const xml = await docFile.async('string');
    zip.file('word/document.xml', patchDocumentSectPr(xml, options, headerRelId));
  }

  const blob = await zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
  downloadBlob(blob, filename);
}

async function exportGeneratedSampleWord(options = {}, filename = '美化规整文档.docx') {
  const {
    fontFamily = 'SimSun',
    fontSize = 14,
    lineSpacing = 1.15,
    firstLineIndent = true,
    bulletList = false
  } = options;

  const fontName = mapWordFont(fontFamily);

  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            text: '文档智能美化标准化报告',
            heading: HeadingLevel.HEADING_1,
            alignment: AlignmentType.CENTER
          }),
          new Paragraph({ text: '' }),
          new Paragraph({
            children: [
              new TextRun({
                text: `${firstLineIndent ? '    ' : ''}${bulletList ? '• ' : ''}本文档已通过 Office 智能效率工具箱完成自动化格式规范、字号与段落重排。`,
                font: fontName,
                size: fontSize * 2
              })
            ]
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `${firstLineIndent ? '    ' : ''}${bulletList ? '• ' : ''}全局字体设置为 ${fontName}，字号为 ${fontSize}pt，段落行距为 ${lineSpacing} 倍。`,
                font: fontName,
                size: fontSize * 2
              })
            ]
          })
        ]
      }
    ]
  });

  downloadBlob(await Packer.toBlob(doc), filename);
}

/**
 * 有原始 docx 时在原文件 styles.xml 上写入默认字体/字号/行距后打包导出；
 * 无原文件时回退为范文生成。
 */
export async function exportProcessedWord(options = {}, filename = '美化规整文档.docx', originalArrayBuffer = null) {
  if (originalArrayBuffer) {
    try {
      await exportPatchedOriginalWord(originalArrayBuffer, options, filename);
      return;
    } catch (err) {
      console.warn('原文档样式补丁导出失败，回退为范文生成', err);
    }
  }
  await exportGeneratedSampleWord(options, filename);
}
