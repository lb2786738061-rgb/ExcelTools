/**
 * 文档工作区：承载可编辑的 Excel 表格 / Word 纸张，并提供导出。
 */

export function createDocumentPreview(onExportOffice, onExportPdf, onShare) {
  const container = document.createElement('div');
  container.className = 'preview-container';
  container.innerHTML = `
    <div class="preview-header" style="align-items:center; justify-content:space-between;">
      <div class="preview-title" id="previewTitle"><span>工作区</span></div>
      <div class="action-bar" id="actionBar">
        <button class="btn-secondary" id="exportExcelBtn">导出文件</button>
        <button class="btn-secondary" id="exportPdfBtn">导出 PDF</button>
        <button class="btn-primary" id="shareBtn" style="width:auto;padding:0.5rem 1rem;"><span>分享</span></button>
      </div>
    </div>
    <div class="preview-viewport editor-viewport" id="previewViewport"></div>
  `;

  const viewport = container.querySelector('#previewViewport');
  const previewTitle = container.querySelector('#previewTitle');
  container.querySelector('#exportExcelBtn').addEventListener('click', () => onExportOffice());
  container.querySelector('#exportPdfBtn').addEventListener('click', () => onExportPdf());
  container.querySelector('#shareBtn').addEventListener('click', () => onShare());

  function mount(el, title) {
    viewport.innerHTML = '';
    if (el) viewport.appendChild(el);
    if (title) previewTitle.innerHTML = `<span>${title}</span>`;
  }

  return {
    element: container,
    mount,
    getViewport: () => viewport,
    setTitle: (t) => { previewTitle.innerHTML = `<span>${t}</span>`; },
    updateQuickBarState() {}
  };
}
