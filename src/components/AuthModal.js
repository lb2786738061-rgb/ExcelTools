/**
 * AuthModal 用户登录 / 注册模态框组件
 */

import { getApiBaseUrl, saveSession } from '../utils/authClient.js';

export function createAuthModal(onLoginSuccess) {
  const container = document.createElement('div');
  container.className = 'modal-backdrop';
  container.style.display = 'none';

  container.innerHTML = `
    <div class="modal-card" style="max-width: 400px; width: 90%;">
      <div class="modal-header">
        <h3 class="modal-title" id="authTitle">🔑 用户登录 / 注册</h3>
        <button class="modal-close" id="closeAuthModal">&times;</button>
      </div>

      <div class="panel-header-tabs" style="margin-top: 1rem;">
        <button class="panel-subtab active" id="tabLogin">用户登录</button>
        <button class="panel-subtab" id="tabRegister">新用户注册</button>
      </div>

      <form id="authForm" style="margin-top: 1rem;">
        <div class="form-group">
          <label class="form-label">👤 用户名</label>
          <input type="text" class="form-select" id="authUsername" placeholder="请输入用户名" required style="background: rgba(15,23,42,0.8);" />
        </div>

        <div class="form-group">
          <label class="form-label">🔒 密码</label>
          <input type="password" class="form-select" id="authPassword" placeholder="请输入密码" required style="background: rgba(15,23,42,0.8);" />
        </div>

        <div id="authErrorMsg" style="color: #ef4444; font-size: 0.85rem; margin-top: 0.5rem; display: none;"></div>

        <button class="btn-primary" type="submit" id="authSubmitBtn" style="margin-top: 1.25rem; width: 100%;">
          <span>🔑 立即登录</span>
        </button>
      </form>
    </div>
  `;

  let mode = 'login'; // login | register

  const closeBtn = container.querySelector('#closeAuthModal');
  const tabLogin = container.querySelector('#tabLogin');
  const tabRegister = container.querySelector('#tabRegister');
  const authTitle = container.querySelector('#authTitle');
  const authForm = container.querySelector('#authForm');
  const authUsername = container.querySelector('#authUsername');
  const authPassword = container.querySelector('#authPassword');
  const authErrorMsg = container.querySelector('#authErrorMsg');
  const authSubmitBtn = container.querySelector('#authSubmitBtn');

  function hide() {
    container.style.display = 'none';
  }

  function show(defaultMode = 'login') {
    mode = defaultMode;
    updateTab();
    authErrorMsg.style.display = 'none';
    container.style.display = 'flex';
  }

  function updateTab() {
    if (mode === 'login') {
      tabLogin.classList.add('active');
      tabRegister.classList.remove('active');
      authTitle.textContent = '🔑 用户登录';
      authSubmitBtn.querySelector('span').textContent = '🔑 立即登录';
    } else {
      tabRegister.classList.add('active');
      tabLogin.classList.remove('active');
      authTitle.textContent = '📝 新用户注册';
      authSubmitBtn.querySelector('span').textContent = '📝 注册并领取 3 次免费试用';
    }
  }

  closeBtn.addEventListener('click', hide);
  tabLogin.addEventListener('click', () => { mode = 'login'; updateTab(); });
  tabRegister.addEventListener('click', () => { mode = 'register'; updateTab(); });

  authForm.addEventListener('submit', async (e) => {
    e.preventDefault();
    authErrorMsg.style.display = 'none';

    const username = authUsername.value.trim();
    const password = authPassword.value.trim();

    const baseUrl = getApiBaseUrl();
    const endpoint = mode === 'login' ? `${baseUrl}/api/auth/login` : `${baseUrl}/api/auth/register`;

    try {
      const res = await fetch(endpoint, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ username, password })
      });
      const data = await res.json();

      if (!res.ok) {
        authErrorMsg.textContent = data.error || '请求失败';
        authErrorMsg.style.display = 'block';
        return;
      }

      saveSession(data.token, data.user);

      alert(data.message || '操作成功！');
      hide();
      onLoginSuccess(data.user);
    } catch (err) {
      authErrorMsg.textContent = '无法连接到后端服务器，请检查 C# (.NET) 或 Node.js 后端服务是否已启动！';
      authErrorMsg.style.display = 'block';
    }
  });

  return {
    element: container,
    show,
    hide
  };
}
