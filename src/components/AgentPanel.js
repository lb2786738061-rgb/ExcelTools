/**
 * 给不会用 Office 的人：说话，表格/文档自己改。
 */

export function createAgentPanel(onAsk) {
  const el = document.createElement('aside');
  el.className = 'agent-panel';
  el.innerHTML = `
    <div class="agent-head">
      <div class="agent-title">智能助手</div>
      <p class="agent-sub">不用找功能区。用平常话告诉我，例如「把工资加总」。我会直接改右边的表或文档。</p>
    </div>
    <div class="agent-chips" id="agentChips"></div>
    <div class="agent-log" id="agentLog"></div>
    <form class="agent-form" id="agentForm">
      <textarea id="agentInput" rows="2" placeholder="说你想做的事…"></textarea>
      <button type="submit" class="btn-primary" id="agentSend">去做</button>
    </form>
  `;

  const log = el.querySelector('#agentLog');
  const input = el.querySelector('#agentInput');
  const chips = el.querySelector('#agentChips');
  const sendBtn = el.querySelector('#agentSend');

  const excelChips = ['一键排版', '水印写上华润集团仅供张三使用', '只要销售部', '按部门汇总工资', '插入序号列', '去掉多余空格'];
  const wordChips = ['一键排版', '套成公文', '插入目录', '插入是/否', '插入勾选框', '插入表格'];

  function setMode(mode) {
    const list = mode === 'word' ? wordChips : excelChips;
    chips.innerHTML = list.map((t) => `<button type="button" class="agent-chip">${t}</button>`).join('');
    chips.querySelectorAll('.agent-chip').forEach((b) => {
      b.addEventListener('click', () => {
        input.value = b.textContent;
        el.querySelector('#agentForm').requestSubmit();
      });
    });
  }

  function push(role, text) {
    const item = document.createElement('div');
    item.className = `agent-msg ${role}`;
    item.textContent = text;
    log.appendChild(item);
    log.scrollTop = log.scrollHeight;
  }

  el.querySelector('#agentForm').addEventListener('submit', async (e) => {
    e.preventDefault();
    const text = input.value.trim();
    if (!text) return;
    input.value = '';
    push('user', text);
    sendBtn.disabled = true;
    push('bot', '正在改…');
    const pending = log.lastChild;
    try {
      const reply = await onAsk(text);
      pending.textContent = reply || '做完了。';
    } catch (err) {
      pending.textContent = '没做成：' + (err.message || '请再试一次');
    } finally {
      sendBtn.disabled = false;
      input.focus();
    }
  });

  setMode('excel');
  push('bot', '把文件丢进来，或直接用演示表。然后告诉我要算什么、怎么排。');

  return { element: el, setMode, push };
}
