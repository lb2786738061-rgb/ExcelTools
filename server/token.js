/**
 * HS256 JWT（header.payload.signature），Node 与 .NET TokenService 使用同一算法。
 * 密钥来自 JWT_SECRET，缺省仅用于本地开发。
 */

import crypto from 'crypto';

const DEFAULT_SECRET = 'exceltools-dev-hmac-secret-change-in-production';
const DEFAULT_EXPIRES = 7 * 24 * 60 * 60;

export function getJwtSecret() {
  return process.env.JWT_SECRET || DEFAULT_SECRET;
}

export function getJwtExpiresSeconds() {
  const raw = Number(process.env.JWT_EXPIRES_SECONDS);
  return Number.isFinite(raw) && raw > 0 ? raw : DEFAULT_EXPIRES;
}

function b64url(input) {
  const buf = Buffer.isBuffer(input) ? input : Buffer.from(input);
  return buf.toString('base64url');
}

function timingSafeEqualStr(a, b) {
  const ba = Buffer.from(String(a));
  const bb = Buffer.from(String(b));
  if (ba.length !== bb.length) return false;
  return crypto.timingSafeEqual(ba, bb);
}

export function signToken(user) {
  const header = b64url(JSON.stringify({ alg: 'HS256', typ: 'JWT' }));
  const now = Math.floor(Date.now() / 1000);
  const payload = b64url(JSON.stringify({
    id: user.id,
    username: user.username,
    role: user.role,
    iat: now,
    exp: now + getJwtExpiresSeconds()
  }));
  const data = `${header}.${payload}`;
  const sig = crypto.createHmac('sha256', getJwtSecret()).update(data).digest('base64url');
  return `${data}.${sig}`;
}

export function verifyToken(token) {
  if (!token || typeof token !== 'string') return null;
  const parts = token.split('.');
  if (parts.length !== 3) return null;

  const [header, payload, sig] = parts;
  const data = `${header}.${payload}`;
  const expected = crypto.createHmac('sha256', getJwtSecret()).update(data).digest('base64url');
  if (!timingSafeEqualStr(sig, expected)) return null;

  try {
    const json = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
    if (!json || typeof json !== 'object') return null;
    if (typeof json.exp !== 'number' || json.exp < Math.floor(Date.now() / 1000)) return null;
    if (!json.id) return null;
    return {
      id: String(json.id),
      username: String(json.username || ''),
      role: String(json.role || 'user')
    };
  } catch {
    return null;
  }
}

export function readBearerToken(req) {
  const authHeader = req.headers.authorization;
  if (!authHeader) return null;
  const parts = authHeader.split(' ');
  if (parts.length < 2 || parts[0] !== 'Bearer') return null;
  return verifyToken(parts[1]);
}
