/**
 * FileUploader 文件拖拽与选择上传组件
 * 支持 .xlsx 与 .docx 文件拖拽、点击浏览、快捷演示文件加载与重置
 */

export function createFileUploader(onFileSelected, onLoadDemo) {
  const container = document.createElement('div');
  container.className = 'panel-card';

  container.innerHTML = `
    <div class="panel-title" style="display: flex; justify-content: space-between; align-items: center;">
      <span>第一步：选择或拖拽 Office 文件</span>
      <button id="resetFileBtn" style="background: none; border: none; color: var(--text-muted); cursor: pointer; font-size: 0.75rem; display: none;">🧹 清空</button>
    </div>

    <!-- 新增：极简操作指引 -->
    <div style="background: rgba(16, 185, 129, 0.1); border-left: 3px solid #10b981; padding: 0.6rem; margin-bottom: 0.8rem; border-radius: 4px; font-size: 0.75rem; color: var(--text-main); line-height: 1.5;">
      <strong style="color: #10b981; display: block; margin-bottom: 4px;">💡 快速上手指南（仅需 3 步）：</strong>
      <div style="color: var(--text-muted);">
        1️⃣ <strong>传文件：</strong>在此上传您的 Word/Excel，或点击下方【试用演示报表】体验。<br/>
        2️⃣ <strong>跟助手说：</strong>右边对话框用平常话，例如「把工资加总」，不用找功能区。<br/>
        3️⃣ <strong>导出：</strong>改完点右上角导出文件。
      </div>
    </div>
    <div class="upload-zone" id="uploadDropZone">
      <svg class="upload-icon" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16a4 4 0 01-.88-7.903A5 5 0 0115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"/>
      </svg>
      <div class="upload-hint" id="uploadStatusText">点击选择或拖拽本地文件至此处</div>
      <div class="upload-subhint">支持 .xlsx / .xls（老版表格） / .csv 和 .docx（Word文档）</div>
      <input type="file" id="fileInput" accept=".xlsx,.xls,.docx,.csv" style="display: none;" />
    </div>

    <!-- 边角功能 1：一键快捷演示文件按钮组 -->
    <div style="margin-top: 0.75rem; display: flex; gap: 0.5rem; justify-content: space-between;">
      <button id="demoExcelBtn" style="flex: 1; background: rgba(59,130,246,0.1); border: 1px solid rgba(59,130,246,0.3); color: #3b82f6; border-radius: 6px; padding: 0.35rem 0.5rem; font-size: 0.75rem; cursor: pointer; transition: all 0.2s;">
        📊 试用 Excel 报表
      </button>
      <button id="demoWordBtn" style="flex: 1; background: rgba(139,92,246,0.1); border: 1px solid rgba(139,92,246,0.3); color: #8b5cf6; border-radius: 6px; padding: 0.35rem 0.5rem; font-size: 0.75rem; cursor: pointer; transition: all 0.2s;">
        📄 试用 Word 公文
      </button>
    </div>
  `;

  const dropZone = container.querySelector('#uploadDropZone');
  const fileInput = container.querySelector('#fileInput');
  const statusText = container.querySelector('#uploadStatusText');
  const resetFileBtn = container.querySelector('#resetFileBtn');

  const demoExcelBtn = container.querySelector('#demoExcelBtn');
  const demoWordBtn = container.querySelector('#demoWordBtn');

  dropZone.addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      handleFile(e.target.files[0]);
    }
  });

  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.classList.remove('dragover');
    });
  });

  dropZone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt.files && dt.files[0]) {
      handleFile(dt.files[0]);
    }
  });

  demoExcelBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    statusText.innerHTML = `使用演示文件: <span style="color: #3b82f6; font-weight: bold;">季度销售统计报表.xlsx</span>`;
    resetFileBtn.style.display = 'inline-block';
    if (onLoadDemo) onLoadDemo('excel');
  });

  demoWordBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    statusText.innerHTML = `使用演示文件: <span style="color: #8b5cf6; font-weight: bold;">标准化公文报告.docx</span>`;
    resetFileBtn.style.display = 'inline-block';
    if (onLoadDemo) onLoadDemo('word');
  });

  resetFileBtn.addEventListener('click', (e) => {
    e.stopPropagation();
    fileInput.value = '';
    statusText.innerHTML = '点击选择或拖拽本地文件至此处';
    resetFileBtn.style.display = 'none';
    if (onLoadDemo) onLoadDemo('reset');
  });

  function handleFile(file) {
    const fileName = file.name;
    const isExcel = fileName.endsWith('.xlsx') || fileName.endsWith('.xls') || fileName.endsWith('.csv');
    const isWord = fileName.endsWith('.docx');

    if (!isExcel && !isWord) {
      alert('仅支持上传 .xlsx、.xls、.csv 或 .docx 格式的 Office 文档！');
      return;
    }

    statusText.innerHTML = `已选择: <span style="color: var(--primary-500); font-weight: bold;">${fileName}</span>`;
    resetFileBtn.style.display = 'inline-block';
    onFileSelected(file, isExcel ? 'excel' : 'word');
  }

  return container;
}
