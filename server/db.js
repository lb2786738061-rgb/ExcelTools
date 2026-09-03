/**
 * JSON 文件存储：用户、VIP 卡密、消费日志。
 * 密码为 PBKDF2；旧版无盐 SHA256 在登录成功后自动升级。
 */

import fs from 'fs';
import path from 'path';
import crypto from 'crypto';

const DATA_DIR = path.resolve(process.cwd(), 'server/data');
const STORE_FILE = path.join(DATA_DIR, 'db.json');

const PBKDF2_ITERATIONS = 100000;
const PBKDF2_KEYLEN = 32;
const PBKDF2_DIGEST = 'sha256';

function envOr(name, fallback) {
  const v = process.env[name];
  return v && v.trim() ? v.trim() : fallback;
}

function hashPassword(pwd) {
  const salt = crypto.randomBytes(16);
  const hash = crypto.pbkdf2Sync(String(pwd), salt, PBKDF2_ITERATIONS, PBKDF2_KEYLEN, PBKDF2_DIGEST);
  return `pbkdf2$${PBKDF2_ITERATIONS}$${salt.toString('hex')}$${hash.toString('hex')}`;
}

function verifyPassword(pwd, stored) {
  if (!stored || !pwd) return false;
  if (stored.startsWith('pbkdf2$')) {
    const parts = stored.split('$');
    if (parts.length !== 4) return false;
    const iterations = parseInt(parts[1], 10);
    const salt = Buffer.from(parts[2], 'hex');
    const expected = Buffer.from(parts[3], 'hex');
    const actual = crypto.pbkdf2Sync(String(pwd), salt, iterations, expected.length, PBKDF2_DIGEST);
    if (actual.length !== expected.length) return false;
    return crypto.timingSafeEqual(actual, expected);
  }

  const legacy = crypto.createHash('sha256').update(String(pwd)).digest('hex');
  const a = Buffer.from(legacy, 'utf8');
  const b = Buffer.from(String(stored), 'utf8');
  if (a.length !== b.length) return false;
  return crypto.timingSafeEqual(a, b);
}

function publicUser(user, extra = {}) {
  if (!user) return null;
  const now = Date.now();
  const isVip = !!(user.isVip && user.vipExpireAt > now);
  return {
    id: user.id,
    username: user.username,
    role: user.role,
    isVip,
    isGuest: user.role === 'guest',
    vipExpireAt: user.vipExpireAt,
    balance: user.balance,
    createdAt: user.createdAt,
    ...extra
  };
}

function defaultStore() {
  const adminUser = envOr('ADMIN_USERNAME', 'admin');
  const adminPass = envOr('ADMIN_PASSWORD', 'admin123');
  const demoUser = envOr('DEMO_USERNAME', 'testuser');
  const demoPass = envOr('DEMO_PASSWORD', '123456');

  return {
    users: [
      {
        id: 'admin-001',
        username: adminUser,
        passwordHash: hashPassword(adminPass),
        role: 'admin',
        isVip: true,
        vipExpireAt: 2524608000000,
        balance: 99999,
        createdAt: Date.now()
      },
      {
        id: 'user-demo',
        username: demoUser,
        passwordHash: hashPassword(demoPass),
        role: 'user',
        isVip: true,
        vipExpireAt: 2524608000000,
        balance: 99999,
        createdAt: Date.now()
      }
    ],
    vipCards: [
      {
        code: 'VIP-MONTH-8888-9999',
        type: 'month',
        value: 30,
        status: 'unused',
        usedBy: null,
        usedAt: null,
        createdAt: Date.now()
      },
      {
        code: 'VIP-YEAR-6666-8888',
        type: 'year',
        value: 365,
        status: 'unused',
        usedBy: null,
        usedAt: null,
        createdAt: Date.now()
      }
    ],
    usageLogs: [],
    templates: []
  };
}

function initDb() {
  if (!fs.existsSync(DATA_DIR)) {
    fs.mkdirSync(DATA_DIR, { recursive: true });
  }
  if (!fs.existsSync(STORE_FILE)) {
    fs.writeFileSync(STORE_FILE, JSON.stringify(defaultStore(), null, 2), 'utf-8');
  }
}

initDb();

function getStore() {
  try {
    const data = fs.readFileSync(STORE_FILE, 'utf-8');
    return JSON.parse(data);
  } catch (err) {
    return defaultStore();
  }
}

function saveStore(store) {
  fs.writeFileSync(STORE_FILE, JSON.stringify(store, null, 2), 'utf-8');
}

export const db = {
  hashPassword,
  verifyPassword,
  publicUser,

  findUserByUsername(username) {
    const store = getStore();
    return store.users.find(u => u.username === username);
  },

  findUserById(id) {
    const store = getStore();
    return store.users.find(u => u.id === id);
  },

  createUser(username, password) {
    const store = getStore();
    const newUser = {
      id: 'usr-' + Date.now() + '-' + Math.floor(Math.random() * 1000),
      username,
      passwordHash: hashPassword(password),
      role: 'user',
      isVip: false,
      vipExpireAt: 0,
      balance: 3,
      createdAt: Date.now()
    };
    store.users.push(newUser);
    saveStore(store);
    return newUser;
  },

  createGuestUser() {
    const store = getStore();
    const id = 'gst-' + crypto.randomUUID();
    const trial = Number(envOr('GUEST_TRIAL_BALANCE', '3')) || 3;
    const guest = {
      id,
      username: 'guest_' + id.replace(/-/g, '').slice(-8),
      passwordHash: hashPassword(crypto.randomBytes(24).toString('hex')),
      role: 'guest',
      isVip: false,
      vipExpireAt: 0,
      balance: trial,
      createdAt: Date.now()
    };
    store.users.push(guest);
    saveStore(store);
    return guest;
  },

  upgradePasswordHash(userId, password) {
    const store = getStore();
    const user = store.users.find(u => u.id === userId);
    if (!user) return;
    user.passwordHash = hashPassword(password);
    saveStore(store);
  },

  consumeQuota(userId, fileName) {
    const store = getStore();
    const user = store.users.find(u => u.id === userId);
    if (!user) return { success: false, message: '用户不存在', code: 'NEED_LOGIN' };

    const now = Date.now();
    const isVipValid = user.isVip && user.vipExpireAt > now;

    if (isVipValid) {
      store.usageLogs.push({
        id: 'log-' + Date.now(),
        userId,
        username: user.username,
        fileName,
        type: 'VIP免扣费',
        cost: 0,
        createdAt: now
      });
      saveStore(store);
      return { success: true, isVip: true, balance: user.balance, expireAt: user.vipExpireAt };
    }

    if (user.balance <= 0) {
      return {
        success: false,
        message: '您的免费试用次数已用完，请充值兑换 VIP！',
        code: 'NEED_RECHARGE'
      };
    }

    user.balance -= 1;
    store.usageLogs.push({
      id: 'log-' + Date.now(),
      userId,
      username: user.username,
      fileName,
      type: '扣除1次次数',
      cost: 1,
      createdAt: now
    });
    saveStore(store);

    return { success: true, isVip: false, balance: user.balance, expireAt: 0 };
  },

  redeemCard(userId, code) {
    const store = getStore();
    const user = store.users.find(u => u.id === userId);
    if (!user) return { success: false, message: '用户不存在' };
    if (user.role === 'guest') {
      return { success: false, message: '访客试用账号无法兑换卡密，请先注册或登录正式账号' };
    }

    const card = store.vipCards.find(c => c.code === code && c.status === 'unused');
    if (!card) return { success: false, message: '无效或已被使用的兑换卡密！' };

    const now = Date.now();
    card.status = 'redeemed';
    card.usedBy = user.username;
    card.usedAt = now;

    if (card.type === 'count') {
      user.balance += card.value;
    } else {
      user.isVip = true;
      const currentExpire = user.vipExpireAt > now ? user.vipExpireAt : now;
      user.vipExpireAt = currentExpire + card.value * 24 * 60 * 60 * 1000;
    }

    saveStore(store);
    return { success: true, user, card };
  },

  generateVipCards(count = 5, type = 'month', value = 30) {
    const store = getStore();
    const generated = [];
    const safeCount = Math.min(Math.max(Number(count) || 0, 1), 100);

    for (let i = 0; i < safeCount; i++) {
      const randomKey = crypto.randomBytes(4).toString('hex').toUpperCase();
      const code = `VIP-${String(type).toUpperCase()}-${Date.now().toString().slice(-4)}-${randomKey}`;
      const card = {
        code,
        type,
        value,
        status: 'unused',
        usedBy: null,
        usedAt: null,
        createdAt: Date.now()
      };
      store.vipCards.push(card);
      generated.push(card);
    }

    saveStore(store);
    return generated;
  },

  getStats() {
    const store = getStore();
    const now = Date.now();
    const totalUsers = store.users.filter(u => u.role !== 'guest').length;
    const activeVips = store.users.filter(u => u.isVip && u.vipExpireAt > now).length;
    const totalLogs = store.usageLogs.length;
    const unusedCards = store.vipCards.filter(c => c.status === 'unused').length;

    return {
      totalUsers,
      activeVips,
      totalLogs,
      unusedCards,
      cards: store.vipCards,
      users: store.users.map(u => publicUser(u)),
      logs: store.usageLogs.slice(-50)
    };
  }
};
