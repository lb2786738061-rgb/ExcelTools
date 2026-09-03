/**
 * AdminModal 管理员后台系统组件
 * 包含平台数据看板、卡密批量生成器（自动发卡卖卡）、用户与日志查看
 */

import { getApiBaseUrl, getToken } from '../utils/authClient.js';

export function createAdminModal() {
  const container = document.createElement('div');
  container.className = 'modal-backdrop';
  container.style.display = 'none';

  container.innerHTML = `
    <div class="modal-card" style="max-width: 800px; width: 95%; max-height: 85vh; overflow-y: auto;">
      <div class="modal-header">
        <h3 class="modal-title">🛠️ 管理员后台系统 (商业化运营控制台)</h3>
        <button class="modal-close" id="closeAdminModal">&times;</button>
      </div>

      <!-- 数据概览 -->
      <div style="display: grid; grid-template-columns: repeat(4, 1fr); gap: 0.75rem; margin-top: 1rem;">
        <div style="background: rgba(59,130,246,0.15); border: 1px solid rgba(59,130,246,0.3); padding: 0.75rem; border-radius: 8px; text-align: center;">
          <div style="font-size: 0.8rem; color: #94a3b8;">总注册用户</div>
          <div style="font-size: 1.5rem; font-weight: 700; color: #3b82f6;" id="statUsers">0</div>
        </div>
        <div style="background: rgba(245,158,11,0.15); border: 1px solid rgba(245,158,11,0.3); padding: 0.75rem; border-radius: 8px; text-align: center;">
          <div style="font-size: 0.8rem; color: #94a3b8;">活跃 VIP 会员</div>
          <div style="font-size: 1.5rem; font-weight: 700; color: #f59e0b;" id="statVips">0</div>
        </div>
        <div style="background: rgba(16,185,129,0.15); border: 1px solid rgba(16,185,129,0.3); padding: 0.75rem; border-radius: 8px; text-align: center;">
          <div style="font-size: 0.8rem; color: #94a3b8;">处理文档日志</div>
          <div style="font-size: 1.5rem; font-weight: 700; color: #10b981;" id="statLogs">0</div>
        </div>
        <div style="background: rgba(139,92,246,0.15); border: 1px solid rgba(139,92,246,0.3); padding: 0.75rem; border-radius: 8px; text-align: center;">
          <div style="font-size: 0.8rem; color: #94a3b8;">未兑换卡密存量</div>
          <div style="font-size: 1.5rem; font-weight: 700; color: #8b5cf6;" id="statCards">0</div>
        </div>
      </div>

      <!-- 卡密生成器 -->
      <div style="margin-top: 1.5rem; background: rgba(15,23,42,0.6); border: 1px solid var(--border-color); padding: 1rem; border-radius: 8px;">
        <h4 style="font-size: 1rem; font-weight: 700; margin-bottom: 0.75rem; color: var(--text-main);">💳 VIP 充值卡密批量生成器 (可导出卖卡)</h4>
        
        <div style="display: flex; gap: 0.75rem; align-items: flex-end;">
          <div style="flex: 1;">
            <label class="form-label">卡密类型</label>
            <select class="form-select" id="genTypeSelect">
              <option value="month">VIP 月卡 (30天)</option>
              <option value="year">VIP 年卡 (365天)</option>
              <option value="count">50 次体验卡</option>
            </select>
          </div>

          <div style="flex: 1;">
            <label class="form-label">生成数量</label>
            <input type="number" class="form-select" id="genCountInput" value="5" min="1" max="50" />
          </div>

          <button class="btn-primary" id="btnGenCards" style="width: auto; padding: 0.55rem 1.25rem;">
            <span>✨ 立即生成</span>
          </button>
        </div>

        <div id="genResultArea" style="margin-top: 1rem; display: none;">
          <label class="form-label" style="color: #10b981;">已生成兑换卡密 (可复制直接发货):</label>
          <textarea class="form-select" id="genCardsText" rows="4" readonly style="font-family: monospace; font-size: 0.85rem; background: #0f172a;"></textarea>
        </div>
      </div>
    </div>
  `;

  const closeBtn = container.querySelector('#closeAdminModal');
  const statUsers = container.querySelector('#statUsers');
  const statVips = container.querySelector('#statVips');
  const statLogs = container.querySelector('#statLogs');
  const statCards = container.querySelector('#statCards');

  const genTypeSelect = container.querySelector('#genTypeSelect');
  const genCountInput = container.querySelector('#genCountInput');
  const btnGenCards = container.querySelector('#btnGenCards');
  const genResultArea = container.querySelector('#genResultArea');
  const genCardsText = container.querySelector('#genCardsText');

  function hide() {
    container.style.display = 'none';
  }

  async function show() {
    container.style.display = 'flex';
    genResultArea.style.display = 'none';
    await loadStats();
  }

  closeBtn.addEventListener('click', hide);

  async function loadStats() {
    const token = getToken();
    if (!token) return;

    const baseUrl = getApiBaseUrl();

    try {
      const res = await fetch(`${baseUrl}/api/admin/dashboard`, {
        headers: { 'Authorization': `Bearer ${token}` }
      });
      if (!res.ok) return;

      const data = await res.json();
      statUsers.textContent = data.totalUsers;
      statVips.textContent = data.activeVips;
      statLogs.textContent = data.totalLogs;
      statCards.textContent = data.unusedCards;
    } catch (err) {
      console.error(err);
    }
  }

  btnGenCards.addEventListener('click', async () => {
    const token = getToken();
    if (!token) return alert('未登录管理员账号！');

    const type = genTypeSelect.value;
    const count = parseInt(genCountInput.value, 10) || 5;
    const value = type === 'month' ? 30 : type === 'year' ? 365 : 50;

    const baseUrl = getApiBaseUrl();

    try {
      const res = await fetch(`${baseUrl}/api/admin/cards/generate`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({ count, type, value })
      });
      const data = await res.json();

      if (!res.ok) return alert(data.error || '生成失败');

      const cardLines = data.cards.map(c => `${c.code} (${c.type === 'month' ? '月卡30天' : c.type === 'year' ? '年卡365天' : '50次卡'})`).join('\n');
      genCardsText.value = cardLines;
      genResultArea.style.display = 'block';
      loadStats();
    } catch (err) {
      alert('连接失败');
    }
  });

  return {
    element: container,
    show,
    hide
  };
}
