/**
 * VipModal VIP 充值与卡密兑换模态框组件
 */

import { getApiBaseUrl, getCachedUser, getToken, persistUser } from '../utils/authClient.js';

export function createVipModal(onRedeemSuccess) {
  const container = document.createElement('div');
  container.className = 'modal-backdrop';
  container.style.display = 'none';

  container.innerHTML = `
    <div class="modal-card" style="max-width: 480px; width: 90%;">
      <div class="modal-header">
        <h3 class="modal-title">💎 VIP 会员充值与卡密兑换</h3>
        <button class="modal-close" id="closeVipModal">&times;</button>
      </div>

      <div style="background: linear-gradient(135deg, rgba(245,158,11,0.15), rgba(217,119,6,0.25)); border: 1px solid rgba(245,158,11,0.4); padding: 1rem; border-radius: 8px; margin-top: 1rem;">
        <div style="font-weight: 700; color: #fbbf24; font-size: 1.05rem; margin-bottom: 0.25rem;">💎 VIP 专属尊享权益</div>
        <ul style="font-size: 0.85rem; color: #fde68a; padding-left: 1.2rem; line-height: 1.6;">
          <li>⚡ 无限制无限次美化 Excel / Word 文档</li>
          <li>📊 独享全套高级公式与数值热力图标红</li>
          <li>🛡️ 导出高清晰无水印 PDF / 微信高清长图</li>
        </ul>
      </div>

      <!-- 🛒 在线购买直达链接入口区 -->
      <div style="margin-top: 1.25rem; background: rgba(15,23,42,0.6); border: 1px solid var(--border-color); padding: 1rem; border-radius: 8px;">
        <div style="font-size: 0.9rem; font-weight: 600; color: var(--text-main); margin-bottom: 0.6rem; display: flex; align-items: center; justify-content: space-between;">
          <span>🛒 还没有卡密？在线极速获取</span>
          <span style="font-size: 0.75rem; color: #f59e0b; background: rgba(245,158,11,0.15); padding: 0.1rem 0.4rem; border-radius: 4px;">自动发货 24h</span>
        </div>

        <button type="button" class="btn-primary" id="btnGoBuy" style="width: 100%; background: linear-gradient(135deg, #10b981, #059669); font-size: 0.95rem; padding: 0.65rem 1rem;">
          <span>🛒 点击前往在线商城购买卡密 (自动发货) ›</span>
        </button>
      </div>

      <!-- 卡密激活兑换区域 -->
      <form id="vipForm" style="margin-top: 1.25rem;">
        <div class="form-group">
          <label class="form-label" style="font-weight: 600;">🔑 已有卡密？输入/粘贴兑换码激活</label>
          <input type="text" class="form-select" id="cardCodeInput" placeholder="例如: VIP-MONTH-8888-9999" required style="background: rgba(15,23,42,0.8); font-family: monospace; font-size: 0.95rem;" />
        </div>

        <div id="vipMsg" style="font-size: 0.85rem; margin-top: 0.5rem; display: none;"></div>

        <button class="btn-primary" type="submit" style="margin-top: 1rem; width: 100%; background: linear-gradient(135deg, #f59e0b, #d97706); font-size: 0.95rem;">
          <span>✨ 确认激活兑换 VIP</span>
        </button>
      </form>
    </div>
  `;

  const closeBtn = container.querySelector('#closeVipModal');
  const btnGoBuy = container.querySelector('#btnGoBuy');
  const vipForm = container.querySelector('#vipForm');
  const cardCodeInput = container.querySelector('#cardCodeInput');
  const vipMsg = container.querySelector('#vipMsg');

  btnGoBuy.addEventListener('click', () => {
    // 可配置为您自己的发卡网、微店或购买网页 URL
    const buyUrl = 'https://example.com/buy-vip';
    alert('正在为您跳转至在线卡密发卡商城页面！在生产环境中将直接打开您的自动发卡网链接。');
    window.open(buyUrl, '_blank');
  });

  function hide() {
    container.style.display = 'none';
  }

  function show() {
    vipMsg.style.display = 'none';
    cardCodeInput.value = '';
    container.style.display = 'flex';
  }

  closeBtn.addEventListener('click', hide);

  vipForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    vipMsg.style.display = 'none';

    const token = getToken();
    const cached = getCachedUser();
    if (!token || (cached && cached.isGuest)) {
      alert('访客试用无法兑换卡密，请先注册或登录正式账号！');
      return;
    }

    const code = cardCodeInput.value.trim();

    const baseUrl = getApiBaseUrl();

    try {
      const res = await fetch(`${baseUrl}/api/vip/redeem`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          'Authorization': `Bearer ${token}`
        },
        body: JSON.stringify({ code })
      });
      const data = await res.json();

      if (!res.ok) {
        vipMsg.style.color = '#ef4444';
        vipMsg.textContent = data.error || '兑换失败';
        vipMsg.style.display = 'block';
        return;
      }

      vipMsg.style.color = '#10b981';
      vipMsg.textContent = data.message;
      vipMsg.style.display = 'block';

      persistUser(data.user);
      alert(data.message);
      hide();
      onRedeemSuccess(data.user);
    } catch (err) {
      vipMsg.style.color = '#ef4444';
      vipMsg.textContent = '无法连接到后端，请确认 C# (.NET) / Node.js 后端已启动！';
      vipMsg.style.display = 'block';
    }
  });

  return {
    element: container,
    show,
    hide
  };
}
