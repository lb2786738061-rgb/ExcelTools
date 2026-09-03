/**
 * Office 智能美化工具箱 - Node.js Express 后端
 * 签名 JWT、访客试用额度、VIP 卡密、管理后台
 */

import express from 'express';
import cors from 'cors';
import { db } from './db.js';
import { signToken, readBearerToken } from './token.js';

const app = express();
const PORT = process.env.PORT || 5001;

app.use(cors());
app.use(express.json());

function requireAuth(req, res, next) {
  const payload = readBearerToken(req);
  if (!payload) {
    return res.status(401).json({ error: '未登录或登录凭证无效', code: 'NEED_LOGIN' });
  }
  const user = db.findUserById(payload.id);
  if (!user) {
    return res.status(401).json({ error: '用户不存在或凭证已失效', code: 'NEED_LOGIN' });
  }
  req.auth = payload;
  req.dbUser = user;
  next();
}

function requireAdmin(req, res, next) {
  requireAuth(req, res, () => {
    if (req.dbUser.role !== 'admin') {
      return res.status(403).json({ error: '越权操作，需要管理员权限！' });
    }
    next();
  });
}

app.get('/api/health', (req, res) => {
  res.json({
    status: 'ok',
    service: 'Office 智能美化工具箱后端 API 服务',
    timestamp: new Date().toISOString(),
    uptime: process.uptime()
  });
});

app.post('/api/auth/register', (req, res) => {
  const { username, password } = req.body;
  if (!username || !password) {
    return res.status(400).json({ error: '用户名和密码不能为空' });
  }

  const existing = db.findUserByUsername(username);
  if (existing) {
    return res.status(400).json({ error: '该用户名已被注册，请更换！' });
  }

  const newUser = db.createUser(username, password);
  const token = signToken(newUser);

  res.json({
    message: '注册成功！已默认赠送 3 次免费试用体验',
    token,
    user: db.publicUser(newUser)
  });
});

app.post('/api/auth/login', (req, res) => {
  const { username, password } = req.body;
  const user = db.findUserByUsername(username);

  if (!user || user.role === 'guest' || !db.verifyPassword(password, user.passwordHash)) {
    return res.status(401).json({ error: '用户名或密码不正确' });
  }

  if (!user.passwordHash.startsWith('pbkdf2$')) {
    db.upgradePasswordHash(user.id, password);
  }

  const token = signToken(user);
  res.json({
    message: '登录成功！',
    token,
    user: db.publicUser(user)
  });
});

app.post('/api/auth/guest', (req, res) => {
  const guest = db.createGuestUser();
  const token = signToken(guest);
  res.json({
    message: '已创建访客试用会话',
    token,
    user: db.publicUser(guest)
  });
});

app.get('/api/user/profile', requireAuth, (req, res) => {
  res.json({ user: db.publicUser(req.dbUser) });
});

app.post('/api/vip/redeem', requireAuth, (req, res) => {
  const { code } = req.body;
  if (!code) return res.status(400).json({ error: '请输入卡密兑换码！' });

  const result = db.redeemCard(req.dbUser.id, code.trim());
  if (!result.success) {
    return res.status(400).json({ error: result.message });
  }

  res.json({
    message: '🎉 成功兑换 VIP / 充值使用次数！',
    user: db.publicUser(result.user)
  });
});

app.post('/api/document/process', requireAuth, (req, res) => {
  const { fileName = '未命名文档.xlsx' } = req.body;
  const result = db.consumeQuota(req.dbUser.id, fileName);
  if (!result.success) {
    return res.status(403).json({
      error: result.message,
      code: result.code || 'NEED_RECHARGE'
    });
  }

  res.json({
    message: '鉴权通过，已扣除额度并授权处理！',
    quota: {
      success: true,
      isVip: result.isVip,
      balance: result.balance,
      expireAt: result.expireAt
    }
  });
});

app.get('/api/admin/dashboard', requireAdmin, (req, res) => {
  res.json(db.getStats());
});

app.post('/api/admin/cards/generate', requireAdmin, (req, res) => {
  const { count = 5, type = 'month', value = 30 } = req.body;
  const cards = db.generateVipCards(count, type, value);
  res.json({ message: `成功批量生成 ${cards.length} 张 VIP 兑换卡密！`, cards });
});

app.use((err, req, res, next) => {
  console.error('❌ 服务端未捕获异常:', err);
  res.status(500).json({ error: '服务器内部错误：' + (err.message || '未知异常') });
});

function startServer(portToTry) {
  const server = app.listen(portToTry, () => {
    console.log(`✅ 商业化 Node.js 后端服务器已成功启动：http://localhost:${portToTry}`);
  });

  server.on('error', (err) => {
    if (err.code === 'EACCES' || err.code === 'EADDRINUSE') {
      console.warn(`⚠️ 端口 ${portToTry} 被占用或受到系统权限限制，尝试平滑切换至端口 ${portToTry + 1}...`);
      startServer(portToTry + 1);
    } else {
      console.error('❌ 服务器启动异常:', err);
    }
  });
}

startServer(Number(PORT));
