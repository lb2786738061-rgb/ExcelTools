/**
 * 前端会话：签名 Token 存 localStorage 仅作凭证；
 * VIP / 余额以 /api/user/profile 为准。
 */

const TOKEN_KEY = 'auth_token';
const USER_KEY = 'user_info';

export function getApiBaseUrl() {
  return window.getApiBaseUrl ? window.getApiBaseUrl() : 'http://localhost:5001';
}

export function getToken() {
  return localStorage.getItem(TOKEN_KEY);
}

export function getCachedUser() {
  try {
    return JSON.parse(localStorage.getItem(USER_KEY) || 'null');
  } catch {
    return null;
  }
}

export function saveSession(token, user) {
  if (token) localStorage.setItem(TOKEN_KEY, token);
  if (user) localStorage.setItem(USER_KEY, JSON.stringify(user));
}

export function persistUser(user) {
  if (!user) return;
  const prev = getCachedUser() || {};
  localStorage.setItem(USER_KEY, JSON.stringify({ ...prev, ...user }));
}

export function clearSession() {
  localStorage.removeItem(TOKEN_KEY);
  localStorage.removeItem(USER_KEY);
}

export function authHeaders() {
  const token = getToken();
  return token ? { Authorization: `Bearer ${token}` } : {};
}

export async function fetchProfile() {
  const token = getToken();
  if (!token) return { ok: false, status: 401 };
  const res = await fetch(`${getApiBaseUrl()}/api/user/profile`, {
    headers: authHeaders()
  });
  if (!res.ok) return { ok: false, status: res.status };
  const data = await res.json();
  persistUser(data.user);
  return { ok: true, user: data.user };
}

export async function ensureSession() {
  try {
    if (getToken()) {
      const profile = await fetchProfile();
      if (profile.ok) {
        return { ok: true, user: profile.user, serverOnline: true };
      }
      if (profile.status === 401) {
        clearSession();
      } else {
        return { ok: false, serverOnline: true, error: '无法读取账户资料' };
      }
    }

    const res = await fetch(`${getApiBaseUrl()}/api/auth/guest`, { method: 'POST' });
    if (!res.ok) {
      return { ok: false, serverOnline: true, error: '无法创建试用会话' };
    }
    const data = await res.json();
    saveSession(data.token, data.user);
    return { ok: true, user: data.user, serverOnline: true, isGuest: true };
  } catch {
    return { ok: false, serverOnline: false, error: '后端离线' };
  }
}

export async function checkVipStatus() {
  const session = await ensureSession();
  if (!session.serverOnline) {
    return { serverOnline: false, isVip: false, user: getCachedUser() };
  }
  if (!session.ok || !session.user) {
    return { serverOnline: true, isVip: false, user: null };
  }
  return {
    serverOnline: true,
    isVip: !!session.user.isVip,
    user: session.user
  };
}

export async function consumeProcessQuota(fileName) {
  const session = await ensureSession();
  if (!session.serverOnline) {
    return { serverOnline: false, allowed: false, offline: true };
  }
  if (!session.ok) {
    return { serverOnline: true, allowed: false, code: 'NEED_LOGIN', error: session.error };
  }

  const res = await fetch(`${getApiBaseUrl()}/api/document/process`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      ...authHeaders()
    },
    body: JSON.stringify({ fileName: fileName || '未命名文档.xlsx' })
  });

  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    return {
      serverOnline: true,
      allowed: false,
      code: data.code || (res.status === 401 ? 'NEED_LOGIN' : 'NEED_RECHARGE'),
      error: data.error
    };
  }

  if (data.quota) {
    persistUser({
      ...(getCachedUser() || {}),
      balance: data.quota.balance,
      isVip: data.quota.isVip
    });
  }

  return {
    serverOnline: true,
    allowed: true,
    quota: data.quota,
    user: getCachedUser()
  };
}
