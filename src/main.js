/**
 * Office 智能效率工具箱 - 商业化全栈主入口
 * 集成后端商业鉴权 API、用户登录认证、VIP 充值扣费与管理员后台
 */

import './styles/index.css';

// 注册全局精美通知中心 (会隐式劫持原生 window.alert)
import './components/Toast.js';

import { createNavbar } from './components/Navbar.js';
import { createFileUploader } from './components/FileUploader.js';
import { createControlPanel } from './components/ControlPanel.js';
import { createDocumentPreview } from './components/DocumentPreview.js';
import { createShareModal } from './components/ShareModal.js';
import { createAuthModal } from './components/AuthModal.js';
import { createVipModal } from './components/VipModal.js';
import { createAdminModal } from './components/AdminModal.js';
import { createConverterPanel } from './components/ConverterPanel.js';
import { createKnowledgePanel } from './components/KnowledgePanel.js';
import { createMergePanel } from './components/MergePanel.js';

import { mergeMultipleExcels, exportProcessedExcel } from './engine/excelEngine.js';
import { exportToPdf } from './utils/pdfExport.js';
import { consumeProcessQuota, ensureSession, getCachedUser, getApiBaseUrl, authHeaders } from './utils/authClient.js';
import { createExcelSheet } from './editor/excelSheet.js';
import { createWordDoc } from './editor/wordDoc.js';
import { createAgentPanel } from './components/AgentPanel.js';
import { planLocally, fixRowColActions } from './agent/localPlanner.js';
import * as XLSX from 'xlsx';

let currentFile = null;
let currentFileType = 'excel';
let activeTab = 'beautify';

const DEFAULT_TEST_EXCEL = [
  ['序号', '部门名称', '基本工资', '绩效奖金', '联系手机', '负责人'],
  [1, '销售一部', 50000, 15000, '13812345678', '张三'],
  [2, '销售一部', 45000, 12000, '13987654321', '李四'],
  [3, '销售二部', 60000, 18000, '13566668888', '王五'],
  [4, '销售二部', 55000, 16000, '13799990000', '赵六'],
  [5, '研发部', 80000, 5000, '18655556666', '钱七'],
  [6, '研发部', 85000, -2000, '18944445555', '孙八']
];

document.addEventListener('DOMContentLoaded', () => {
  const appContainer = document.querySelector('#app');
  if (!appContainer) return;

  let navbar = null;

  // 1. 初始化模态框组件
  const authModal = createAuthModal((user) => {
    if (navbar) navbar.updateUserState(user);
  });
  document.body.appendChild(authModal.element);

  const vipModal = createVipModal((user) => {
    if (navbar) navbar.updateUserState(user);
  });
  document.body.appendChild(vipModal.element);

  const adminModal = createAdminModal();
  document.body.appendChild(adminModal.element);

  const shareModal = createShareModal();
  document.body.appendChild(shareModal.element);

  const viewBeautify = document.createElement('main');
  viewBeautify.className = 'office-workspace';

  const sidebar = document.createElement('aside');
  sidebar.className = 'sidebar-panel';

  const excelSheet = createExcelSheet();
  const wordDoc = createWordDoc();

  const preview = createDocumentPreview(
    async () => {
      if (currentFileType === 'word') await wordDoc.exportDocx('文档.docx');
      else excelSheet.exportXlsx('工作簿.xlsx');
    },
    async () => {
      if (currentFileType === 'word') {
        const paper = wordDoc.element.querySelector('.wd-paper') || wordDoc.element;
        await exportToPdf(paper, '文档.pdf');
        return;
      }
      await exportToPdf(excelSheet.buildPrintElement(), '表格.pdf');
    },
    () => shareModal.show()
  );

  function showEditor(type) {
    const next = type === 'word' ? 'word' : 'excel';
    if (currentFileType === next) {
      if (agentPanel) agentPanel.setMode(next);
      return;
    }
    currentFileType = next;
    if (next === 'word') {
      preview.mount(wordDoc.element, 'Word 文档');
      wordDoc.focus();
    } else {
      preview.mount(excelSheet.element, 'Excel 工作表');
      excelSheet.focus();
    }
    if (agentPanel) agentPanel.setMode(next);
  }

  function applyAgentActions(mode, actions) {
    for (const act of actions || []) {
      const cmd = act.cmd;
      if (!cmd) continue;
      if (mode === 'word') wordDoc.command(cmd, act.payload);
      else excelSheet.command(cmd, act.payload !== undefined ? act.payload : (cmd === 'select' ? act : undefined));
    }
  }

  async function runBeautify() {
    const mode = currentFileType === 'word' ? 'word' : 'excel';
    const snapshot = mode === 'word' ? wordDoc.getSnapshot() : excelSheet.getSnapshot();
    let theme = null;
    let extras = [];
    let reply = '';
    try {
      const res = await fetch(`${getApiBaseUrl()}/api/agent/beautify`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json', ...authHeaders() },
        body: JSON.stringify({ mode, snapshot })
      });
      if (res.ok) {
        const data = await res.json();
        if (data.source === 'zhipu' && Array.isArray(data.actions)) {
          reply = data.reply || '';
          extras = data.actions.filter((a) => a && a.cmd && a.cmd !== 'beautify' && a.cmd !== 'toc' && a.cmd !== 'redhead');
          const b = data.actions.find((a) => a?.cmd === 'beautify');
          if (b?.payload && typeof b.payload === 'object') theme = b.payload;
        }
      }
    } catch {
      /* 无后端时用本地专业排版 */
    }
    if (mode === 'word') wordDoc.command('beautify', theme || {});
    else excelSheet.command('beautify', theme || {});
    applyAgentActions(mode, extras);
    return reply || (mode === 'word'
      ? '已按公文习惯整理标题、正文缩进和表格表头。'
      : '已按报表规范排版：表头、数字格式、对齐、框线、冻结首行，金额列已加合计。');
  }

  async function runAgent(text) {
    if (/美化|排版|规整|好看/.test(String(text || ''))) {
      return runBeautify();
    }
    const mode = currentFileType === 'word' ? 'word' : 'excel';
    const snapshot = mode === 'word' ? wordDoc.getSnapshot() : excelSheet.getSnapshot();
    const local = planLocally(text, mode, snapshot);
    let plan = null;
    if (local.matched) {
      plan = local;
    } else {
      try {
        const res = await fetch(`${getApiBaseUrl()}/api/agent/run`, {
          method: 'POST',
          headers: { 'Content-Type': 'application/json', ...authHeaders() },
          body: JSON.stringify({ message: text, mode, snapshot })
        });
        if (res.ok) {
          const data = await res.json();
          if (data.source === 'zhipu' && Array.isArray(data.actions) && data.actions.length) {
            plan = { reply: data.reply, actions: data.actions };
          }
        }
      } catch {
        /* 后端未启动时用本地规则 */
      }
      if (!plan) plan = local;
    }
    plan.actions = fixRowColActions(text, plan.actions || []);
    applyAgentActions(mode, plan.actions);
    return plan.reply || '做完了。';
  }

  const agentPanel = createAgentPanel(runAgent);

  const controlPanel = createControlPanel(async (cmd, payload) => {
    if (cmd === 'appMode') {
      showEditor(payload);
      return;
    }
    if (cmd === 'beautify') {
      controlPanel.setRunBusy(true);
      try {
        const msg = await runBeautify();
        window.alert(msg, 'success');
      } catch (err) {
        window.alert(err.message || '排版失败', 'error');
      } finally {
        controlPanel.setRunBusy(false);
      }
      return;
    }
    if (currentFileType === 'word') wordDoc.command(cmd, payload);
    else excelSheet.command(cmd, payload);
  });

  async function openFile(file, fileType) {
    const quota = await consumeProcessQuota(file.name);
    if (!quota.offline && !quota.allowed) {
      if (quota.code === 'NEED_RECHARGE') {
        alert(quota.error || '试用次数已用完');
        vipModal.show();
        return;
      }
      if (quota.code === 'NEED_LOGIN') {
        authModal.show('login');
        return;
      }
    } else if (quota.user && navbar) navbar.updateUserState(quota.user);

    const buffer = await file.arrayBuffer();
    if (fileType === 'word') {
      controlPanel.setAppMode('word');
      await wordDoc.loadDocx(buffer);
      showEditor('word');
    } else {
      controlPanel.setAppMode('excel');
      const wb = XLSX.read(buffer, { type: 'array' });
      const name = wb.SheetNames[0];
      const aoa = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
      excelSheet.loadAoA(aoa, name);
      showEditor('excel');
    }
  }

  const fileUploader = createFileUploader(
    (file, fileType) => {
      currentFile = file;
      openFile(file, fileType);
    },
    (type) => {
      if (type === 'excel') {
        currentFile = null;
        controlPanel.setAppMode('excel');
        excelSheet.loadAoA(JSON.parse(JSON.stringify(DEFAULT_TEST_EXCEL)), '演示');
        showEditor('excel');
      } else if (type === 'word') {
        currentFile = null;
        controlPanel.setAppMode('word');
        showEditor('word');
      } else if (type === 'reset') {
        currentFile = null;
        excelSheet.loadAoA([[]]);
        controlPanel.setAppMode('excel');
        showEditor('excel');
      }
    }
  );

  sidebar.appendChild(fileUploader);

  const officeBody = document.createElement('div');
  officeBody.className = 'office-body';
  officeBody.appendChild(sidebar);
  officeBody.appendChild(preview.element);
  officeBody.appendChild(agentPanel.element);

  viewBeautify.appendChild(controlPanel.element);
  viewBeautify.appendChild(officeBody);

  // 5. 格式转换专属主页面容器
  const viewConvert = document.createElement('main');
  viewConvert.style.cssText = 'flex: 1; padding: 1.5rem 2rem; display: none;';
  const converterPanel = createConverterPanel(() => preview.getViewport());
  viewConvert.appendChild(converterPanel);

  // 5.5 极速公式字典知识库主页面容器
  const viewKnowledge = document.createElement('main');
  viewKnowledge.style.cssText = 'flex: 1; display: none; background: #f8fafc; overflow-y: auto; padding: 1.5rem 2rem;';
  const knowledgePanel = createKnowledgePanel(() => vipModal.show());
  viewKnowledge.appendChild(knowledgePanel);

  // 5.6 批量合并专属主页面容器
  const viewMerge = document.createElement('main');
  viewMerge.style.cssText = 'flex: 1; padding: 1.5rem 2rem; display: none; max-width: 900px; margin: 0 auto; width: 100%;';

  const mergePanel = createMergePanel(
    async (files) => {
      // 合并执行回调
      window.alert('正在为您高速并发合并，请稍候...', 'success');
      try {
        const mergedData = await mergeMultipleExcels(files);
        exportProcessedExcel(mergedData, '智能多表批量合并总表.xlsx');
        window.alert('✅ 多表合并成功，已自动下载报表！', 'success');
        mergePanel.clear();
      } catch (err) {
        window.alert('❌ 合并失败: ' + err.message, 'error');
      }
    },
    () => vipModal.show()
  );
  viewMerge.appendChild(mergePanel.element);

  // 6. 创建顶部导航栏（此时所有页面与面板变量均已完成声明）
  navbar = createNavbar(
    activeTab,
    (tabKey) => {
      activeTab = tabKey;
      if (tabKey === 'merge') {
        viewBeautify.style.display = 'none';
        viewConvert.style.display = 'none';
        viewKnowledge.style.display = 'none';
        viewMerge.style.display = 'block';
      } else if (tabKey === 'knowledge') {
        viewBeautify.style.display = 'none';
        viewConvert.style.display = 'none';
        viewMerge.style.display = 'none';
        viewKnowledge.style.display = 'block';
      } else if (tabKey === 'convert') {
        viewBeautify.style.display = 'none';
        viewKnowledge.style.display = 'none';
        viewMerge.style.display = 'none';
        viewConvert.style.display = 'block';
      } else {
        viewConvert.style.display = 'none';
        viewKnowledge.style.display = 'none';
        viewMerge.style.display = 'none';
        viewBeautify.style.display = 'flex';

        if (tabKey === 'watermark') {
          controlPanel.setActiveSubtab('text');
          controlPanel.openWatermarkDialog();
        } else if (tabKey === 'beautify') {
          controlPanel.setActiveSubtab('text');
        }
      }
    },
    () => authModal.show('login'),
    () => vipModal.show(),
    () => adminModal.show()
  );

  appContainer.appendChild(navbar.element);
  appContainer.appendChild(viewBeautify);
  appContainer.appendChild(viewConvert);
  appContainer.appendChild(viewKnowledge);
  appContainer.appendChild(viewMerge);

  let activeApiBaseUrl = 'http://localhost:5001';
  window.getApiBaseUrl = () => activeApiBaseUrl;

  excelSheet.loadAoA(JSON.parse(JSON.stringify(DEFAULT_TEST_EXCEL)), '演示');
  showEditor('excel');
  checkLocalUser();
  checkServerHealth();

  async function checkServerHealth() {
    const candidatePorts = [5001, 5002, 3001, 3002];
    let foundOnline = false;

    for (const port of candidatePorts) {
      try {
        const controller = new AbortController();
        const timeoutId = setTimeout(() => controller.abort(), 1200);
        const url = `http://localhost:${port}/api/health`;
        const res = await fetch(url, { signal: controller.signal });
        clearTimeout(timeoutId);
        if (res.ok) {
          activeApiBaseUrl = `http://localhost:${port}`;
          foundOnline = true;
          if (navbar) navbar.updateServerStatus(true);
          break;
        }
      } catch (err) {
        // 继续尝试下一个端口
      }
    }

    if (!foundOnline && navbar) {
      navbar.updateServerStatus(false);
    }
  }

  async function checkLocalUser() {
    const session = await ensureSession();
    if (session.ok && session.user && navbar) {
      navbar.updateUserState(session.user);
    } else if (navbar) {
      const cached = getCachedUser();
      if (cached) navbar.updateUserState(cached);
    }
  }
});