/**
 * ConverterPanel 全能文件转换中心控制面板组件 (支持多文件批量高保真转换)
 */

import { getSupportedTargetFormats, executeFileConversion, executeBatchFileConversion } from '../engine/convertEngine.js';

export function createConverterPanel(onPreviewElementReq) {
  const container = document.createElement('div');
  container.className = 'converter-page-wrapper';
  container.style.cssText = `
    width: 100%;
    max-width: 1200px;
    margin: 0 auto;
    display: flex;
    flex-direction: column;
    gap: 1.5rem;
  `;

  container.innerHTML = `
    <!-- 1. 顶部大气格式转换中心卡片 -->
    <div class="panel-card" style="padding: 2rem; background: var(--bg-card); border-radius: var(--radius-lg);">
      <div style="display: flex; align-items: center; justify-content: space-between; margin-bottom: 1.5rem;">
        <div>
          <h2 style="font-size: 1.35rem; font-weight: 700; color: var(--text-main); display: flex; align-items: center; gap: 0.6rem;">
            <span>🔄 极速高保真文档格式转换中心</span>
            <span style="font-size: 0.75rem; font-weight: 600; color: #10b981; background: rgba(16,185,129,0.15); border: 1px solid rgba(16,185,129,0.3); padding: 0.2rem 0.6rem; border-radius: 9999px;">批量并发版 V2.5</span>
          </h2>
          <p style="font-size: 0.88rem; color: var(--text-muted); margin-top: 0.35rem;">支持 Office (Word/Excel)、PDF、图片长图、Markdown、TXT 及发票等多格式单文件及多文件批量高保真转换导出</p>
        </div>
      </div>

      <!-- 大格式拖拽上传区 -->
      <div class="upload-zone" id="converterDropZone" style="padding: 2.5rem 1.5rem; border: 2px dashed rgba(59,130,246,0.5); background: rgba(15, 23, 42, 0.5);">
        <svg class="upload-icon" style="width: 52px; height: 52px;" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5" d="M8 7h12m0 0l-4-4m4 4l-4 4m0 6H4m0 0l4 4m-4-4l4-4"/>
        </svg>
        <div class="upload-hint" id="converterStatusText" style="font-size: 1.1rem; margin-top: 0.5rem;">点击选择或拖拽一个或多个本地文件至此处进行格式转换</div>
        <div class="upload-subhint" style="font-size: 0.85rem; color: var(--text-muted); margin-top: 0.4rem;">自动识别文件类型 (支持 .docx, .xlsx, .pdf, .txt, .md, .csv) | 按住 Ctrl / Shift 可多选</div>
        <input type="file" id="converterFileInput" multiple style="display: none;" />
      </div>

      <!-- 批量文件清单与转换目标格式配置框 -->
      <div id="targetFormatConfig" style="margin-top: 1.25rem; display: none; background: rgba(15,23,42,0.6); padding: 1.25rem; border-radius: var(--radius-md); border: 1px solid var(--border-color);">
        <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 0.75rem;">
          <h4 style="font-size: 0.95rem; font-weight: 600; color: var(--text-main);">📋 待转换文件队列 (<span id="batchQueueCount">0</span> 个)</h4>
          <button id="clearQueueBtn" style="background: none; border: none; color: #ef4444; font-size: 0.8rem; cursor: pointer;">🗑️ 清空队列</button>
        </div>

        <div id="batchFileList" style="display: flex; flex-direction: column; gap: 0.5rem; max-height: 220px; overflow-y: auto; padding-right: 0.25rem; margin-bottom: 1rem;">
          <!-- 动态渲染文件列表项 -->
        </div>

        <div style="display: flex; gap: 1.25rem; align-items: flex-end; flex-wrap: wrap; border-top: 1px dashed var(--border-color); padding-top: 1rem;">
          <button class="btn-primary" id="startConvertBtn" style="flex: 1; min-width: 220px; padding: 0.75rem 1.25rem; font-size: 0.95rem; background: linear-gradient(135deg, #3b82f6, #8b5cf6);">
            <span>⚡ 开始批量高保真一键转换并导出下载</span>
          </button>
        </div>
      </div>
    </div>

    <!-- 2. 独立拓展效率工具箱四大板块网格卡片 -->
    <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); gap: 1.25rem;">
      <div class="panel-card tool-box-card" id="toolDiff" style="cursor: pointer; transition: var(--transition-fast);">
        <div style="font-size: 1.5rem; margin-bottom: 0.5rem;">📄↔️📄</div>
        <h3 style="font-size: 1.05rem; font-weight: 600; color: var(--text-main);">文档版本差异对比</h3>
        <p style="font-size: 0.8rem; color: var(--text-muted); margin-top: 0.4rem;">上传新旧两个版本的 Word/Excel，自动高亮标注增删差异并导出比对报告。</p>
      </div>

      <div class="panel-card tool-box-card" id="toolInvoice" style="cursor: pointer; transition: var(--transition-fast);">
        <div style="font-size: 1.5rem; margin-bottom: 0.5rem;">🧾📊</div>
        <h3 style="font-size: 1.05rem; font-weight: 600; color: var(--text-main);">发票合并与报销汇总表</h3>
        <p style="font-size: 0.8rem; color: var(--text-muted); margin-top: 0.4rem;">电子发票 OCR 识别提取金额与数据，自动合并并生成电子报销单 Excel。</p>
      </div>

      <div class="panel-card tool-box-card" id="toolSign" style="cursor: pointer; transition: var(--transition-fast);">
        <div style="font-size: 1.5rem; margin-bottom: 0.5rem;">🖋️盖章</div>
        <h3 style="font-size: 1.05rem; font-weight: 600; color: var(--text-main);">手机手写签名与电子盖章</h3>
        <p style="font-size: 0.8rem; color: var(--text-muted); margin-top: 0.4rem;">屏幕手写签名背景透明化处理，公文印章快捷加盖并导出防篡改 PDF。</p>
      </div>

      <div class="panel-card tool-box-card" id="toolMask" style="cursor: pointer; transition: var(--transition-fast);">
        <div style="font-size: 1.5rem; margin-bottom: 0.5rem;">🛡️打码</div>
        <h3 style="font-size: 1.05rem; font-weight: 600; color: var(--text-main);">防泄密水印与安全脱敏</h3>
        <p style="font-size: 0.8rem; color: var(--text-muted); margin-top: 0.4rem;">全屏倾斜防伪水印，智能识别手机号/身份证脱敏遮挡打码防护。</p>
      </div>
    </div>
  `;

  const dropZone = container.querySelector('#converterDropZone');
  const fileInput = container.querySelector('#converterFileInput');
  const statusText = container.querySelector('#converterStatusText');
  const targetConfig = container.querySelector('#targetFormatConfig');
  const batchQueueCount = container.querySelector('#batchQueueCount');
  const batchFileList = container.querySelector('#batchFileList');
  const clearQueueBtn = container.querySelector('#clearQueueBtn');
  const startBtn = container.querySelector('#startConvertBtn');

  let selectedFilesQueue = []; // { file: File, targetFormat: string }

  dropZone.addEventListener('click', () => fileInput.click());

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      handleSelectedFiles(Array.from(e.target.files));
    }
  });

  ['dragenter', 'dragover'].forEach(name => {
    dropZone.addEventListener(name, (e) => {
      e.preventDefault();
      dropZone.classList.add('dragover');
    });
  });

  ['dragleave', 'drop'].forEach(name => {
    dropZone.addEventListener(name, (e) => {
      e.preventDefault();
      dropZone.classList.remove('dragover');
    });
  });

  dropZone.addEventListener('drop', (e) => {
    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      handleSelectedFiles(Array.from(e.dataTransfer.files));
    }
  });

  clearQueueBtn.addEventListener('click', () => {
    selectedFilesQueue = [];
    renderQueueList();
  });

  function handleSelectedFiles(newFiles) {
    newFiles.forEach(file => {
      const formats = getSupportedTargetFormats(file.name);
      const defaultFmt = formats.length > 0 ? formats[0].key : 'pdf';
      selectedFilesQueue.push({
        file,
        targetFormat: defaultFmt,
        availableFormats: formats
      });
    });

    renderQueueList();
  }

  function renderQueueList() {
    if (selectedFilesQueue.length === 0) {
      targetConfig.style.display = 'none';
      statusText.innerHTML = `点击选择或拖拽一个或多个本地文件至此处进行格式转换`;
      return;
    }

    targetConfig.style.display = 'block';
    batchQueueCount.textContent = selectedFilesQueue.length;
    statusText.innerHTML = `已准备 <span style="color: #3b82f6; font-weight: bold;">${selectedFilesQueue.length}</span> 个文件等待高保真格式转换`;

    batchFileList.innerHTML = selectedFilesQueue.map((item, idx) => {
      const sizeKB = (item.file.size / 1024).toFixed(1);
      const optsHtml = item.availableFormats.map(fmt =>
        `<option value="${fmt.key}" ${fmt.key === item.targetFormat ? 'selected' : ''}>${fmt.label}</option>`
      ).join('');

      return `
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 0.5rem 0.75rem; background: var(--panel-bg); border: 1px solid var(--border-color); border-radius: 6px; gap: 0.75rem;">
          <div style="display: flex; align-items: center; gap: 0.5rem; overflow: hidden; flex: 1;">
            <span style="font-size: 1.1rem;">📄</span>
            <div style="display: flex; flex-direction: column; overflow: hidden;">
              <span style="color: var(--text-main); font-size: 0.85rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${item.file.name}</span>
              <span style="color: var(--text-muted); font-size: 0.7rem;">${sizeKB} KB</span>
            </div>
          </div>

          <div style="display: flex; align-items: center; gap: 0.5rem;">
            <select class="form-select queue-fmt-select" data-index="${idx}" style="padding: 0.35rem 0.5rem; font-size: 0.8rem; border-radius: 4px;">
              ${optsHtml}
            </select>
            <button class="remove-queue-item" data-index="${idx}" style="background: none; border: none; color: #ef4444; cursor: pointer; padding: 0.2rem;" title="移除此文件">
              ✕
            </button>
          </div>
        </div>
      `;
    }).join('');

    // 绑定格式下拉事件
    batchFileList.querySelectorAll('.queue-fmt-select').forEach(sel => {
      sel.addEventListener('change', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        if (selectedFilesQueue[idx]) {
          selectedFilesQueue[idx].targetFormat = e.target.value;
        }
      });
    });

    // 绑定删除行事件
    batchFileList.querySelectorAll('.remove-queue-item').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.target.dataset.index, 10);
        selectedFilesQueue.splice(idx, 1);
        renderQueueList();
      });
    });
  }

  startBtn.addEventListener('click', async () => {
    if (selectedFilesQueue.length === 0) return;

    const originalText = startBtn.innerHTML;
    startBtn.disabled = true;

    try {
      if (selectedFilesQueue.length === 1) {
        const item = selectedFilesQueue[0];
        startBtn.innerHTML = `<span>⏳ 正在转换 ${item.file.name}...</span>`;
        const previewEl = onPreviewElementReq ? onPreviewElementReq() : null;
        await executeFileConversion(item.file, item.targetFormat, previewEl);
        alert(`🎉 文件 ${item.file.name} 格式转换成功，已下载到本地！`);
      } else {
        await executeBatchFileConversion(selectedFilesQueue, (curr, total, name) => {
          startBtn.innerHTML = `<span>⚡ 正在批量转换 (${curr}/${total}): ${name}...</span>`;
        });
        alert(`🎉 批量处理完成！已完成 ${selectedFilesQueue.length} 个文件的格式高保真转换与下载。`);
      }
    } catch (err) {
      console.error('批量转换异常:', err);
      alert(`转换部分完成或提示: ${err.message || '存在转换失败项，请检查文件格式！'}`);
    } finally {
      startBtn.disabled = false;
      startBtn.innerHTML = originalText;
    }
  });

  // 工具卡片快捷点击交互提示
  container.querySelector('#toolDiff').addEventListener('click', () => {
    alert('【文档版本差异对比】工具：请在上方上传框上传待比对的文件。');
    fileInput.click();
  });

  container.querySelector('#toolInvoice').addEventListener('click', () => {
    alert('【发票合并与报销汇总】工具：请在上方上传框批量导入电子发票文件。');
    fileInput.click();
  });

  container.querySelector('#toolSign').addEventListener('click', () => {
    alert('【手写签名与电子印章】工具：请在上方上传框选择公文文件。');
    fileInput.click();
  });

  container.querySelector('#toolMask').addEventListener('click', () => {
    alert('【防泄密水印与安全脱敏】工具：已就绪脱敏引擎，请选择需要安全防护的文件。');
    fileInput.click();
  });

  return container;
}

