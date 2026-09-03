/**
 * Office 功能区：命令作用在当前选区（单元格 / 文字），不是整篇开关。
 */

import { createWatermarkModal } from './WatermarkModal.js';

export function createControlPanel(onCommand) {
  const container = document.createElement('div');
  container.className = 'office-ribbon is-collapsed';
  let appMode = 'excel';

  container.innerHTML = `
    <div class="ribbon-appbar">
      <div class="ribbon-app-switch">
        <button type="button" class="ribbon-app-btn active" data-app="excel">Excel</button>
        <button type="button" class="ribbon-app-btn" data-app="word">Word</button>
      </div>
      <button type="button" class="btn-primary ribbon-run" data-cmd="beautify">一键排版</button>
      <button type="button" class="btn-secondary ribbon-reset" data-cmd="undo">撤销</button>
      <button type="button" class="ribbon-toggle" id="ribbonToggle">展开功能区</button>
    </div>
    <div class="ribbon-detail">
    <div class="ribbon-tabs" data-app-tabs="excel">
      <button type="button" class="ribbon-tab active" data-tab="ex-home">开始</button>
      <button type="button" class="ribbon-tab" data-tab="ex-insert">插入</button>
      <button type="button" class="ribbon-tab" data-tab="ex-formula">公式</button>
      <button type="button" class="ribbon-tab" data-tab="ex-data">数据</button>
    </div>
    <div class="ribbon-tabs" data-app-tabs="word" style="display:none;">
      <button type="button" class="ribbon-tab active" data-tab="wd-home">开始</button>
      <button type="button" class="ribbon-tab" data-tab="wd-insert">插入</button>
      <button type="button" class="ribbon-tab" data-tab="wd-design">设计</button>
      <button type="button" class="ribbon-tab" data-tab="wd-layout">布局</button>
      <button type="button" class="ribbon-tab" data-tab="wd-refs">引用</button>
    </div>
    <div class="ribbon-body">
      <div class="ribbon-panel" data-panel="ex-home">
        <div class="ribbon-group">
          <div class="ribbon-group-title">剪贴板</div>
          <button type="button" class="cmd" data-cmd="copy">复制</button>
          <button type="button" class="cmd" data-cmd="paste">粘贴</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">字体</div>
          <div class="ribbon-toggle-row">
            <button type="button" class="fmt-btn" data-cmd="bold" title="加粗选区 Ctrl+B">B</button>
            <button type="button" class="fmt-btn" data-cmd="italic" title="斜体选区"><i>I</i></button>
            <button type="button" class="fmt-btn" data-cmd="underline" title="下划线选区"><u>U</u></button>
            <button type="button" class="fmt-btn" data-cmd="strike" title="删除线选区"><s>S</s></button>
          </div>
          <select class="form-select compact" data-cmd-value="fontFamily">
            <option value="">字体</option>
            <option value="Microsoft YaHei">微软雅黑</option>
            <option value="SimSun">宋体</option>
            <option value="SimHei">黑体</option>
            <option value="KaiTi">楷体</option>
            <option value="Arial">Arial</option>
          </select>
          <select class="form-select compact" data-cmd-value="fontSize">
            <option value="">字号</option>
            <option value="10">10</option>
            <option value="12">12</option>
            <option value="14">14</option>
            <option value="16">16</option>
            <option value="18">18</option>
            <option value="24">24</option>
          </select>
          <select class="form-select compact" data-cmd-value="fontColor">
            <option value="">字体颜色</option>
            <option value="#000000">黑</option>
            <option value="#dc2626">红</option>
            <option value="#2563eb">蓝</option>
            <option value="#059669">绿</option>
            <option value="#ffffff">白</option>
          </select>
          <select class="form-select compact" data-cmd-value="fill">
            <option value="">填充</option>
            <option value="#ffffff">无</option>
            <option value="#217346">绿</option>
            <option value="#fff2cc">黄</option>
            <option value="#deebf7">蓝</option>
            <option value="#fce4d6">橙</option>
          </select>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">对齐</div>
          <button type="button" class="cmd" data-cmd="align" data-payload="left">左对齐</button>
          <button type="button" class="cmd" data-cmd="align" data-payload="center">居中</button>
          <button type="button" class="cmd" data-cmd="align" data-payload="right">右对齐</button>
          <button type="button" class="cmd" data-cmd="wrap">自动换行</button>
          <button type="button" class="cmd" data-cmd="merge">合并单元格</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">数字</div>
          <button type="button" class="cmd" data-cmd="numFmt" data-payload="general">常规</button>
          <button type="button" class="cmd" data-cmd="numFmt" data-payload="thousand">千分位</button>
          <button type="button" class="cmd" data-cmd="numFmt" data-payload="currency">￥货币</button>
          <button type="button" class="cmd" data-cmd="numFmt" data-payload="percent">百分比</button>
          <button type="button" class="cmd" data-cmd="numFmt" data-payload="int">整数</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">边框</div>
          <button type="button" class="cmd" data-cmd="border" data-payload="all">所有框线</button>
          <button type="button" class="cmd" data-cmd="border" data-payload="none">无框线</button>
          <button type="button" class="cmd" data-cmd="clearFormat">清除格式</button>
          <button type="button" class="cmd" data-cmd="clear">清除内容</button>
        </div>
      </div>

      <div class="ribbon-panel" data-panel="ex-insert" style="display:none;">
        <div class="ribbon-group">
          <div class="ribbon-group-title">单元格</div>
          <button type="button" class="cmd" data-cmd="insertRow">插入行</button>
          <button type="button" class="cmd" data-cmd="insertCol">插入列</button>
          <button type="button" class="cmd" data-cmd="deleteRow">删除行</button>
          <button type="button" class="cmd" data-cmd="deleteCol">删除列</button>
          <button type="button" class="cmd" data-cmd="slashHeader">斜线表头</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">图表</div>
          <button type="button" class="cmd" data-cmd="chart" data-payload="bar">柱形图</button>
          <button type="button" class="cmd" data-cmd="chart" data-payload="line">折线图</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">表单控件</div>
          <button type="button" class="cmd" data-cmd="checkbox">勾选框</button>
          <button type="button" class="cmd" data-cmd="yesno">是 / 否</button>
          <button type="button" class="cmd" data-cmd="radio">单选圈</button>
          <button type="button" class="cmd" data-cmd="dropdown">下拉单选</button>
          <button type="button" class="cmd" data-cmd="multiselect">多选</button>
        </div>
      </div>

      <div class="ribbon-panel" data-panel="ex-formula" style="display:none;">
        <div class="ribbon-group">
          <div class="ribbon-group-title">函数库 · 写入活动单元格下方</div>
          <button type="button" class="cmd primary" data-cmd="autosum">Σ 求和 SUM</button>
          <button type="button" class="cmd" data-cmd="average">平均值</button>
          <button type="button" class="cmd" data-cmd="max">最大值</button>
          <button type="button" class="cmd" data-cmd="min">最小值</button>
          <button type="button" class="cmd" data-cmd="count">计数</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">编辑栏</div>
          <p class="ribbon-hint">在格子或 fx 栏输入 =SUM(B2:B6)、=IF(C2&gt;0,"是","否")、=VLOOKUP(...) 后回车计算。</p>
        </div>
      </div>

      <div class="ribbon-panel" data-panel="ex-data" style="display:none;">
        <div class="ribbon-group">
          <div class="ribbon-group-title">排序和筛选</div>
          <button type="button" class="cmd" data-cmd="sortAsc">升序</button>
          <button type="button" class="cmd" data-cmd="sortDesc">降序</button>
          <button type="button" class="cmd" data-cmd="filter">自动筛选</button>
          <button type="button" class="cmd" data-cmd="unique">删除重复项</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">查找</div>
          <button type="button" class="cmd" data-cmd="find">查找 Ctrl+F</button>
          <button type="button" class="cmd" data-cmd="replace">替换 Ctrl+H</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">冻结窗格</div>
          <button type="button" class="cmd" data-cmd="freezeTop">冻结首行</button>
          <button type="button" class="cmd" data-cmd="freezeLeft">冻结首列</button>
          <button type="button" class="cmd" data-cmd="freeze">冻结窗格</button>
          <button type="button" class="cmd" data-cmd="unfreeze">取消冻结</button>
        </div>
      </div>

      <div class="ribbon-panel" data-panel="wd-home" style="display:none;">
        <div class="ribbon-group">
          <div class="ribbon-group-title">剪贴板</div>
          <button type="button" class="cmd" data-cmd="undo">撤销</button>
          <button type="button" class="cmd" data-cmd="redo">重做</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">字体</div>
          <div class="ribbon-toggle-row">
            <button type="button" class="fmt-btn" data-cmd="bold">B</button>
            <button type="button" class="fmt-btn" data-cmd="italic"><i>I</i></button>
            <button type="button" class="fmt-btn" data-cmd="underline"><u>U</u></button>
            <button type="button" class="fmt-btn" data-cmd="strike"><s>S</s></button>
          </div>
          <select class="form-select compact" data-cmd-value="fontFamily">
            <option value="">字体</option>
            <option value="Microsoft YaHei">微软雅黑</option>
            <option value="SimSun">宋体</option>
            <option value="SimHei">黑体</option>
            <option value="KaiTi">楷体</option>
          </select>
          <select class="form-select compact" data-cmd-value="fontSize">
            <option value="">字号</option>
            <option value="2">10</option>
            <option value="3">12</option>
            <option value="4">14</option>
            <option value="5">18</option>
            <option value="6">24</option>
          </select>
          <select class="form-select compact" data-cmd-value="fontColor">
            <option value="">颜色</option>
            <option value="#000000">黑</option>
            <option value="#dc2626">红</option>
            <option value="#2563eb">蓝</option>
          </select>
          <select class="form-select compact" data-cmd-value="highlight">
            <option value="">底纹</option>
            <option value="#fef08a">黄</option>
            <option value="#bbf7d0">绿</option>
            <option value="#bfdbfe">蓝</option>
          </select>
          <button type="button" class="cmd" data-cmd="removeFormat">清除格式</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">段落</div>
          <button type="button" class="cmd" data-cmd="align" data-payload="left">左对齐</button>
          <button type="button" class="cmd" data-cmd="align" data-payload="center">居中</button>
          <button type="button" class="cmd" data-cmd="align" data-payload="right">右对齐</button>
          <button type="button" class="cmd" data-cmd="align" data-payload="justify">两端</button>
          <button type="button" class="cmd" data-cmd="bullet">项目符号</button>
          <button type="button" class="cmd" data-cmd="number">编号</button>
          <button type="button" class="cmd" data-cmd="indent">增加缩进</button>
          <button type="button" class="cmd" data-cmd="outdent">减少缩进</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">样式</div>
          <button type="button" class="cmd" data-cmd="style" data-payload="P">正文</button>
          <button type="button" class="cmd" data-cmd="style" data-payload="H1">标题 1</button>
          <button type="button" class="cmd" data-cmd="style" data-payload="H2">标题 2</button>
          <button type="button" class="cmd" data-cmd="style" data-payload="H3">标题 3</button>
        </div>
      </div>

      <div class="ribbon-panel" data-panel="wd-insert" style="display:none;">
        <div class="ribbon-group">
          <div class="ribbon-group-title">表格 / 插图</div>
          <button type="button" class="cmd" data-cmd="table">插入表格</button>
          <button type="button" class="cmd" data-cmd="image">插入图片</button>
          <button type="button" class="cmd" data-cmd="hr">分隔线</button>
          <button type="button" class="cmd" data-cmd="pageBreak">分页符</button>
          <button type="button" class="cmd" data-cmd="link">超链接</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">页</div>
          <button type="button" class="cmd" data-cmd="redhead">红头</button>
        </div>
        <div class="ribbon-group">
          <div class="ribbon-group-title">表单控件</div>
          <button type="button" class="cmd" data-cmd="checkbox">勾选框</button>
          <button type="button" class="cmd" data-cmd="yesno">是 / 否</button>
          <button type="button" class="cmd" data-cmd="radio">单选</button>
          <button type="button" class="cmd" data-cmd="dropdown">下拉单选</button>
          <button type="button" class="cmd" data-cmd="multiselect">多选</button>
        </div>
      </div>

      <div class="ribbon-panel" data-panel="wd-design" style="display:none;">
        <div class="ribbon-group">
          <div class="ribbon-group-title">页面背景</div>
          <button type="button" class="cmd" data-cmd="pageColor" data-payload="#ffffff">白</button>
          <button type="button" class="cmd" data-cmd="pageColor" data-payload="#fffbeb">米黄</button>
          <button type="button" class="cmd" data-cmd="pageColor" data-payload="#f0fdf4">浅绿</button>
          <button type="button" class="cmd" data-cmd="pageBorder">页面边框</button>
        </div>
      </div>

      <div class="ribbon-panel" data-panel="wd-layout" style="display:none;">
        <div class="ribbon-group">
          <div class="ribbon-group-title">页面设置</div>
          <button type="button" class="cmd" data-cmd="orientation" data-payload="portrait">纵向</button>
          <button type="button" class="cmd" data-cmd="orientation" data-payload="landscape">横向</button>
          <button type="button" class="cmd" data-cmd="columns" data-payload="1">一栏</button>
          <button type="button" class="cmd" data-cmd="columns" data-payload="2">两栏</button>
        </div>
      </div>

      <div class="ribbon-panel" data-panel="wd-refs" style="display:none;">
        <div class="ribbon-group">
          <div class="ribbon-group-title">目录</div>
          <button type="button" class="cmd" data-cmd="toc">插入目录</button>
        </div>
      </div>
    </div>
    </div>
  `;

  function showApp(mode) {
    const next = mode === 'word' ? 'word' : 'excel';
    const same = appMode === next;
    appMode = next;
    container.querySelectorAll('.ribbon-app-btn').forEach((b) => b.classList.toggle('active', b.dataset.app === next));
    container.querySelectorAll('[data-app-tabs]').forEach((el) => {
      el.style.display = el.getAttribute('data-app-tabs') === next ? 'flex' : 'none';
    });
    const tab = container.querySelector(`[data-app-tabs="${next}"] .ribbon-tab.active`);
    showPanel(tab ? tab.dataset.tab : (next === 'word' ? 'wd-home' : 'ex-home'));
    if (!same) onCommand('appMode', next);
  }

  const wmModal = createWatermarkModal({
    onStamp: (payload) => onCommand('watermark', payload),
    onMask: () => onCommand('maskSensitive')
  });
  document.body.appendChild(wmModal.element);

  function showPanel(id) {
    container.querySelectorAll('.ribbon-tab').forEach((t) => t.classList.toggle('active', t.dataset.tab === id));
    container.querySelectorAll('.ribbon-panel').forEach((p) => {
      p.style.display = p.dataset.panel === id ? 'flex' : 'none';
    });
  }

  const toggleBtn = container.querySelector('#ribbonToggle');
  toggleBtn.addEventListener('click', () => {
    const collapsed = container.classList.toggle('is-collapsed');
    toggleBtn.textContent = collapsed ? '展开功能区' : '收起功能区';
  });

  container.querySelectorAll('.ribbon-app-btn').forEach((b) => b.addEventListener('click', (e) => {
    e.preventDefault();
    e.stopPropagation();
    showApp(b.dataset.app);
  }));
  container.querySelectorAll('.ribbon-tab').forEach((b) => b.addEventListener('click', () => showPanel(b.dataset.tab)));

  container.addEventListener('click', (e) => {
    const btn = e.target.closest('[data-cmd]');
    if (!btn || btn.tagName === 'SELECT') return;
    onCommand(btn.dataset.cmd, btn.dataset.payload);
  });
  container.querySelectorAll('[data-cmd-value]').forEach((sel) => {
    sel.addEventListener('change', () => {
      if (!sel.value) return;
      onCommand(sel.dataset.cmdValue, sel.value);
      sel.selectedIndex = 0;
    });
  });

  showApp('excel');

  const runBtn = container.querySelector('.ribbon-run');

  return {
    element: container,
    setRunBusy(busy) {
      if (!runBtn) return;
      runBtn.disabled = !!busy;
      runBtn.textContent = busy ? '正在排版…' : '一键排版';
    },
    setAppMode: showApp,
    setActiveSubtab(name) {
      const map = {
        text: appMode === 'word' ? 'wd-home' : 'ex-home',
        formula: 'ex-formula',
        data: 'ex-data'
      };
      showPanel(map[name] || 'ex-home');
    },
    openWatermarkDialog() {
      wmModal.show();
    },
    getOptions() { return { appMode }; },
    setOptions() {}
  };
}
