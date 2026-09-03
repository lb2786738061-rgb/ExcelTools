/**
 * Office 智能效率工具箱 - 高清 PDF 与微信长图导出工具
 * 基于 html2canvas 与 pdf-lib
 */

import html2canvas from 'html2canvas';
import { PDFDocument } from 'pdf-lib';

/**
 * 将预览容器转换为微信分享长图并自动下载 (.png)
 * @param {HTMLElement} element - 目标 DOM 容器
 * @param {string} filename - 保存的文件名
 */
export async function exportToLongImage(element, filename = '文档排版美化视图长图.png') {
  if (!element) return;

  try {
    const canvas = await html2canvas(element, {
      scale: 2, // 2倍清晰度
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false
    });

    const dataUrl = canvas.toDataURL('image/png');
    const link = document.createElement('a');
    link.href = dataUrl;
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  } catch (err) {
    console.error('生成长图失败:', err);
    alert('生成微信长图失败，请重试！');
  }
}

/**
 * 将预览容器转换为高质量 PDF 并自动下载 (.pdf)
 * @param {HTMLElement} element - 目标 DOM 容器
 * @param {string} filename - 保存的文件名
 */
export async function exportToPdf(element, filename = '文档标准化处理结果.pdf') {
  if (!element) return;

  let stash = null;
  if (!element.isConnected) {
    stash = document.createElement('div');
    stash.style.cssText = 'position:fixed;left:-12000px;top:0;background:#fff;';
    stash.appendChild(element);
    document.body.appendChild(stash);
  }

  try {
    const canvas = await html2canvas(element, {
      scale: 2,
      useCORS: true,
      backgroundColor: '#ffffff',
      logging: false
    });

    const imgDataUrl = canvas.toDataURL('image/png');
    const imgBytes = await fetch(imgDataUrl).then((res) => res.arrayBuffer());

    const pdfDoc = await PDFDocument.create();
    const pngImage = await pdfDoc.embedPng(imgBytes);
    const page = pdfDoc.addPage([pngImage.width, pngImage.height]);
    page.drawImage(pngImage, {
      x: 0,
      y: 0,
      width: pngImage.width,
      height: pngImage.height
    });

    const pdfBytes = await pdfDoc.save();
    const blob = new Blob([pdfBytes], { type: 'application/pdf' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(link.href);
  } catch (err) {
    console.error('生成 PDF 失败:', err);
    alert('生成 PDF 失败，请重试！');
  } finally {
    stash?.remove();
  }
}
