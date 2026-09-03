/**
 * MergePanel.js - 批量合并中心工作台
 * 支持多文件拖拽上传与合并队列管理
 */

import { checkVipStatus } from '../utils/authClient.js';

export function createMergePanel(onStartMerge, onRequestVip) {
  const container = document.createElement('div');
  container.className = 'panel-card';
  container.style.marginTop = '1rem';

  container.innerHTML = `
    <div class="panel-title" style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 1rem;">
      <span style="font-size: 1.1rem;">🗂️ 批量合并工作台 (支持 Excel/CSV)</span>
      <span style="font-size: 0.8rem; color: var(--text-muted);">自动以首个文件为基准对齐表头</span>
    </div>

    <!-- VIP 拦截横幅 -->
    <div style="background: rgba(245, 158, 11, 0.1); border-left: 3px solid #f59e0b; padding: 0.75rem; border-radius: 4px; margin-bottom: 1.5rem; display: flex; justify-content: space-between; align-items: center;">
      <div>
        <strong style="color: #d97706; display: block; margin-bottom: 0.2rem; font-size: 0.85rem;">💎 提示：非 VIP 用户最多支持合并 2 个文件</strong>
        <span style="color: var(--text-muted); font-size: 0.75rem;">升级 VIP 即可享受无限制海量文件极速并发合并，并支持导出结果。</span>
      </div>
      <button class="btn-primary" id="upgradeVipBtn" style="background: linear-gradient(135deg, #f59e0b, #d97706); padding: 0.4rem 0.8rem; font-size: 0.8rem;">
        升级特权
      </button>
    </div>

    <!-- 多文件拖拽区 -->
    <div class="upload-zone" id="mergeDropZone" style="border: 2px dashed var(--primary-500); background: rgba(59, 130, 246, 0.05); padding: 2rem; border-radius: 8px; text-align: center; cursor: pointer; transition: all 0.3s ease;">
      <div style="font-size: 2.5rem; margin-bottom: 0.5rem;">📂➕</div>
      <h3 style="color: var(--text-main); margin: 0 0 0.5rem 0; font-size: 1.1rem;">点击选择或拖拽多个表格文件到此处</h3>
      <p style="color: var(--text-muted); font-size: 0.85rem; margin: 0;">支持 .xlsx, .xls, .csv 格式批量导入</p>
      <input type="file" id="mergeFileInput" accept=".xlsx,.xls,.csv" multiple style="display: none;" />
    </div>

    <!-- 文件列队区 -->
    <div id="fileQueueContainer" style="margin-top: 1.5rem; display: none;">
      <h4 style="font-size: 0.95rem; color: var(--text-main); margin-bottom: 0.8rem; border-bottom: 1px solid var(--border-color); padding-bottom: 0.5rem;">
        待合并文件队列 (<span id="queueCount">0</span>)
      </h4>
      <div id="fileList" style="max-height: 250px; overflow-y: auto; display: flex; flex-direction: column; gap: 0.5rem; margin-bottom: 1.5rem;">
        <!-- 文件项动态注入 -->
      </div>
      
      <button class="btn-primary" id="executeMergeBtn" style="width: 100%; padding: 0.85rem; font-size: 1rem; background: linear-gradient(135deg, #3b82f6, #2563eb); border-radius: 8px; font-weight: bold; box-shadow: 0 4px 15px rgba(37, 99, 235, 0.3);">
        ⚡ 立即开始合并
      </button>
    </div>
  `;

  let selectedFiles = [];

  const dropZone = container.querySelector('#mergeDropZone');
  const fileInput = container.querySelector('#mergeFileInput');
  const upgradeVipBtn = container.querySelector('#upgradeVipBtn');
  const executeMergeBtn = container.querySelector('#executeMergeBtn');
  const fileQueueContainer = container.querySelector('#fileQueueContainer');
  const fileListEl = container.querySelector('#fileList');
  const queueCountEl = container.querySelector('#queueCount');

  upgradeVipBtn.addEventListener('click', () => {
    if (onRequestVip) onRequestVip();
  });

  dropZone.addEventListener('click', () => fileInput.click());

  ['dragenter', 'dragover'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.style.background = 'rgba(59, 130, 246, 0.15)';
    });
  });

  ['dragleave', 'drop'].forEach(eventName => {
    dropZone.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropZone.style.background = 'rgba(59, 130, 246, 0.05)';
    });
  });

  dropZone.addEventListener('drop', (e) => {
    const dt = e.dataTransfer;
    if (dt.files && dt.files.length > 0) {
      addFilesToQueue(Array.from(dt.files));
    }
  });

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      addFilesToQueue(Array.from(e.target.files));
      // 清空 input value，以便重复选择同一个文件
      fileInput.value = '';
    }
  });

  function addFilesToQueue(newFiles) {
    const validFiles = newFiles.filter(f => {
      const ext = f.name.toLowerCase();
      return ext.endsWith('.xlsx') || ext.endsWith('.xls') || ext.endsWith('.csv');
    });

    if (validFiles.length < newFiles.length) {
      window.alert('部分文件格式不被支持，已被忽略。请上传 Excel 或 CSV 文件。', 'warning');
    }

    selectedFiles = [...selectedFiles, ...validFiles];
    renderFileQueue();
  }

  function removeFile(index) {
    selectedFiles.splice(index, 1);
    renderFileQueue();
  }

  function renderFileQueue() {
    if (selectedFiles.length === 0) {
      fileQueueContainer.style.display = 'none';
      return;
    }

    fileQueueContainer.style.display = 'block';
    queueCountEl.textContent = selectedFiles.length;

    fileListEl.innerHTML = selectedFiles.map((file, idx) => {
      const sizeKB = (file.size / 1024).toFixed(1);
      return `
        <div style="display: flex; justify-content: space-between; align-items: center; padding: 0.6rem 0.8rem; background: var(--panel-bg); border: 1px solid var(--border-color); border-radius: 6px;">
          <div style="display: flex; align-items: center; gap: 0.6rem; overflow: hidden;">
            <span style="font-size: 1.2rem;">${idx === 0 ? '🌟' : '📄'}</span>
            <div style="display: flex; flex-direction: column; overflow: hidden;">
              <span style="color: var(--text-main); font-size: 0.85rem; white-space: nowrap; overflow: hidden; text-overflow: ellipsis;">${file.name} ${idx === 0 ? '<span style="color: #10b981; font-size: 0.75rem;">(合并基准文件)</span>' : ''}</span>
              <span style="color: var(--text-muted); font-size: 0.7rem;">${sizeKB} KB</span>
            </div>
          </div>
          <button class="remove-file-btn" data-index="${idx}" style="background: none; border: none; color: #ef4444; cursor: pointer; padding: 0.2rem; display: flex; align-items: center; justify-content: center; border-radius: 4px;">
            <svg width="16" height="16" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
          </button>
        </div>
      `;
    }).join('');

    // 绑定删除事件
    const removeBtns = container.querySelectorAll('.remove-file-btn');
    removeBtns.forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(e.currentTarget.getAttribute('data-index'), 10);
        removeFile(idx);
      });
    });
  }

  executeMergeBtn.addEventListener('click', async () => {
    if (selectedFiles.length > 2) {
      const vip = await checkVipStatus();
      if (!vip.serverOnline) {
        window.alert('合并超过 2 个文件需要在线校验 VIP，请先启动后端服务。', 'warning');
        return;
      }
      if (!vip.isVip) {
        window.alert("💎 免费版本最多仅支持合并 2 个文件进行体验！\n升级 VIP 即可解锁无限制批量并发合并特权。", 'warning');
        if (onRequestVip) onRequestVip();
        return;
      }
    }

    if (onStartMerge) {
      onStartMerge(selectedFiles);
    }
  });

  return {
    element: container,
    getFiles: () => selectedFiles,
    clear: () => {
      selectedFiles = [];
      renderFileQueue();
    }
  };
}