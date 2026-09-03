/**
 * Navbar 顶部导航栏组件
 * 包含品牌图标、模式选择标签页与【用户中心 / VIP充值 / 管理后台】快捷控制按钮
 */

export function createNavbar(activeTab, onTabChange, onAuthClick, onVipClick, onAdminClick) {
  const navbar = document.createElement('header');
  navbar.className = 'navbar';

  navbar.innerHTML = `
    <div class="nav-brand">
      <div class="brand-icon">
        <svg width="24" height="24" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2">
          <path d="M14 2H6 C4.9 2 4 2.9 4 4v16 C4 21.1 4.9 22 6 22h12 C19.1 22 20 21.1 20 20V8L14 2z"></path>
          <polyline points="14 2 14 8 20 8"></polyline>
          <line x1="16" y1="13" x2="8" y2="13"></line>
          <line x1="16" y1="17" x2="8" y2="17"></line>
          <polyline points="10 9 9 9 8 9"></polyline>
        </svg>
      </div>
      <div>
        <div style="display: flex; align-items: center; gap: 0.5rem;">
          <span class="brand-title">Office 智能效率工具箱</span>
          <span class="brand-badge">商业全栈 V2.0</span>
          <span id="serverStatusBadge" style="font-size: 0.75rem; padding: 0.15rem 0.45rem; border-radius: 12px; background: rgba(234,179,8,0.2); color: #eab308; border: 1px solid #eab308;" title="检测服务器联机状态">
            🟡 探测中
          </span>
        </div>
      </div>
    </div>

    <nav class="nav-tabs">
      <button class="tab-btn ${activeTab === 'beautify' ? 'active' : ''}" data-tab="beautify">
        <span>✨ 智能美化与规整</span>
      </button>
      <button class="tab-btn ${activeTab === 'watermark' ? 'active' : ''}" data-tab="watermark">
        <span>🛡️ 安全水印与脱敏</span>
      </button>
      <button class="tab-btn ${activeTab === 'convert' ? 'active' : ''}" data-tab="convert">
        <span>📄 格式转换与导出</span>
      </button>
      <button class="tab-btn ${activeTab === 'knowledge' ? 'active' : ''}" data-tab="knowledge">
        <span>📕 极速公式字典</span>
      </button>
      <button class="tab-btn ${activeTab === 'merge' ? 'active' : ''}" data-tab="merge" style="color: #f59e0b; border-color: #f59e0b;">
        <span>🗂️ 批量合并中心</span>
      </button>
    </nav>

    <!-- 右侧商业化用户控制中心 -->
    <div style="display: flex; align-items: center; gap: 0.5rem;">
      <button class="btn-secondary" id="navAuthBtn" style="padding: 0.4rem 0.75rem; font-size: 0.85rem;">
        <span id="navAuthText">👤 登录 / 注册</span>
      </button>
      <button class="btn-primary" id="navVipBtn" style="padding: 0.4rem 0.75rem; font-size: 0.85rem; background: linear-gradient(135deg, #f59e0b, #d97706); width: auto;">
        <span>💎 充值 VIP</span>
      </button>
      <button class="btn-secondary" id="navAdminBtn" style="padding: 0.4rem 0.75rem; font-size: 0.85rem; display: none;" title="管理后台">
        <span>🛠️ 后台</span>
      </button>
    </div>
  `;

  const serverStatusBadge = navbar.querySelector('#serverStatusBadge');

  const navAuthBtn = navbar.querySelector('#navAuthBtn');
  const navAuthText = navbar.querySelector('#navAuthText');
  const navVipBtn = navbar.querySelector('#navVipBtn');
  const navAdminBtn = navbar.querySelector('#navAdminBtn');

  function setActiveTab(tabKey) {
    navbar.querySelectorAll('.tab-btn').forEach(btn => {
      if (btn.dataset.tab === tabKey) {
        btn.classList.add('active');
      } else {
        btn.classList.remove('active');
      }
    });
  }

  // 绑定模式切换点击事件
  navbar.querySelectorAll('.tab-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      const tabKey = e.currentTarget.dataset.tab;
      setActiveTab(tabKey);
      onTabChange(tabKey);
    });
  });

  navAuthBtn.addEventListener('click', () => onAuthClick());
  navVipBtn.addEventListener('click', () => onVipClick());
  navAdminBtn.addEventListener('click', () => onAdminClick());

  function updateUserState(user) {
    if (user) {
      const vipTag = user.isVip ? ' [VIP会员]' : ` [试用剩 ${user.balance ?? 0} 次]`;
      const label = user.isGuest ? '访客' : user.username;
      navAuthText.textContent = `👤 ${label}${vipTag}`;
      if (user.role === 'admin') {
        navAdminBtn.style.display = 'inline-flex';
      } else {
        navAdminBtn.style.display = 'none';
      }
    } else {
      navAuthText.textContent = '👤 登录 / 注册';
      navAdminBtn.style.display = 'none';
    }
  }

  function updateServerStatus(isOnline) {
    if (!serverStatusBadge) return;
    if (isOnline) {
      serverStatusBadge.textContent = '🟢 云端已连接';
      serverStatusBadge.style.background = 'rgba(16,185,129,0.2)';
      serverStatusBadge.style.color = '#10b981';
      serverStatusBadge.style.borderColor = '#10b981';
      serverStatusBadge.title = '后端 API 服务在线，支持云端鉴权与 VIP 次数扣除';
    } else {
      serverStatusBadge.textContent = '🟡 离线模式';
      serverStatusBadge.style.background = 'rgba(245,158,11,0.2)';
      serverStatusBadge.style.color = '#f59e0b';
      serverStatusBadge.style.borderColor = '#f59e0b';
      serverStatusBadge.title = '后端服务未开启，已无缝降级为纯前端 100% 离线安全处理模式';
    }
  }

  return {
    element: navbar,
    setActiveTab,
    updateUserState,
    updateServerStatus
  };
}
