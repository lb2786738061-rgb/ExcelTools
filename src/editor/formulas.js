/**
 * Excel 公式求值：选区写入的公式在格子里真实计算。
 */

export function colName(index) {
  let n = index + 1;
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

export function parseA1(ref) {
  const m = String(ref).trim().toUpperCase().match(/^\$?([A-Z]+)\$?(\d+)$/);
  if (!m) return null;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
  return { r: Number(m[2]) - 1, c: c - 1 };
}

export function a1(r, c) {
  return `${colName(c)}${r + 1}`;
}

export function letterToIndex(letters) {
  let c = 0;
  for (const ch of String(letters).toUpperCase()) {
    if (ch < 'A' || ch > 'Z') return -1;
    c = c * 26 + (ch.charCodeAt(0) - 64);
  }
  return c - 1;
}

function splitArgs(src) {
  const args = [];
  let cur = '';
  let depth = 0;
  let quote = false;
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (ch === '"') quote = !quote;
    else if (!quote && ch === '(') depth++;
    else if (!quote && ch === ')') depth--;
    else if (!quote && depth === 0 && ch === ',') {
      args.push(cur.trim());
      cur = '';
      continue;
    }
    cur += ch;
  }
  if (cur.trim()) args.push(cur.trim());
  return args;
}

function toNumber(v) {
  if (typeof v === 'number' && Number.isFinite(v)) return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  const n = Number(String(v).replace(/,/g, ''));
  return Number.isFinite(n) ? n : 0;
}

function expandRange(a, b) {
  const p = parseA1(a);
  const q = parseA1(b);
  if (!p || !q) return [];
  const cells = [];
  for (let r = Math.min(p.r, q.r); r <= Math.max(p.r, q.r); r++) {
    for (let c = Math.min(p.c, q.c); c <= Math.max(p.c, q.c); c++) {
      cells.push(a1(r, c));
    }
  }
  return cells;
}

function collect(arg, getValue) {
  if (arg.includes(':')) {
    const [a, b] = arg.split(':');
    return expandRange(a.trim(), b.trim()).map((addr) => getValue(addr));
  }
  if (parseA1(arg)) return [getValue(arg)];
  return [evalExpr(arg, getValue)];
}

function matchCriteria(value, criteria) {
  const c = String(criteria).trim();
  const m = c.match(/^(>=|<=|<>|>|<|=)(.*)$/);
  if (m) {
    const op = m[1];
    const rhs = m[2];
    const nv = toNumber(value);
    const nr = toNumber(rhs);
    if (op === '>') return nv > nr;
    if (op === '<') return nv < nr;
    if (op === '>=') return nv >= nr;
    if (op === '<=') return nv <= nr;
    if (op === '<>') return String(value) !== rhs;
    return String(value) === rhs || nv === nr;
  }
  return String(value) === c || toNumber(value) === toNumber(c);
}

export function evalExpr(expr, getValue) {
  if (expr == null) return 0;
  let src = String(expr).trim();
  if (src.startsWith('=')) src = src.slice(1).trim();
  if (!src) return 0;

  if (src.startsWith('"') && src.endsWith('"')) return src.slice(1, -1);
  if (/^-?\d+(\.\d+)?$/.test(src)) return Number(src);

  const fn = src.match(/^([A-Z_][A-Z0-9_]*)\((.*)\)\s*$/i);
  if (fn) {
    const name = fn[1].toUpperCase();
    const args = splitArgs(fn[2]);
    const nums = () => args.flatMap((a) => collect(a, getValue)).map(toNumber);
    if (name === 'SUM') return nums().reduce((a, b) => a + b, 0);
    if (name === 'AVERAGE') {
      const n = nums();
      return n.length ? n.reduce((a, b) => a + b, 0) / n.length : 0;
    }
    if (name === 'MAX') return Math.max(...nums(), 0);
    if (name === 'MIN') return Math.min(...nums());
    if (name === 'COUNT') return args.flatMap((a) => collect(a, getValue)).filter((v) => v !== '' && Number.isFinite(toNumber(v))).length;
    if (name === 'COUNTA') return args.flatMap((a) => collect(a, getValue)).filter((v) => v !== '' && v != null).length;
    if (name === 'ROUND') return Math.round(toNumber(evalExpr(args[0], getValue)) * 10 ** toNumber(evalExpr(args[1] || '0', getValue))) / 10 ** toNumber(evalExpr(args[1] || '0', getValue));
    if (name === 'ABS') return Math.abs(toNumber(evalExpr(args[0], getValue)));
    if (name === 'TRIM') return String(evalExpr(args[0], getValue) ?? '').trim();
    if (name === 'LEN') return String(evalExpr(args[0], getValue) ?? '').length;
    if (name === 'LEFT') return String(evalExpr(args[0], getValue) ?? '').slice(0, toNumber(evalExpr(args[1] || '1', getValue)));
    if (name === 'RIGHT') {
      const s = String(evalExpr(args[0], getValue) ?? '');
      return s.slice(-toNumber(evalExpr(args[1] || '1', getValue)));
    }
    if (name === 'CONCAT' || name === 'CONCATENATE') {
      return args.map((a) => String(evalExpr(a, getValue) ?? '')).join('');
    }
    if (name === 'IF') {
      const cond = evalCondition(args[0], getValue);
      return evalExpr(cond ? (args[1] || '0') : (args[2] || '0'), getValue);
    }
    if (name === 'IFERROR') {
      try {
        const v = evalExpr(args[0], getValue);
        if (v === '#DIV/0!' || v === '#N/A' || Number.isNaN(v)) return evalExpr(args[1] || '0', getValue);
        return v;
      } catch {
        return evalExpr(args[1] || '0', getValue);
      }
    }
    if (name === 'SUMIF') {
      const keys = collect(args[0], getValue);
      const crit = evalExpr(args[1], getValue);
      const sums = collect(args[2] || args[0], getValue);
      let t = 0;
      keys.forEach((k, i) => { if (matchCriteria(k, crit)) t += toNumber(sums[i]); });
      return t;
    }
    if (name === 'COUNTIF') {
      const keys = collect(args[0], getValue);
      const crit = evalExpr(args[1], getValue);
      return keys.filter((k) => matchCriteria(k, crit)).length;
    }
    if (name === 'VLOOKUP') {
      const lookup = evalExpr(args[0], getValue);
      const [a, b] = String(args[1]).split(':');
      const colIndex = toNumber(evalExpr(args[2], getValue)) - 1;
      const start = parseA1(a);
      const end = parseA1(b);
      if (!start || !end) return '#N/A';
      for (let r = Math.min(start.r, end.r); r <= Math.max(start.r, end.r); r++) {
        const key = getValue(a1(r, Math.min(start.c, end.c)));
        if (String(key) === String(lookup) || toNumber(key) === toNumber(lookup)) {
          return getValue(a1(r, Math.min(start.c, end.c) + colIndex));
        }
      }
      return '#N/A';
    }
    return '#NAME?';
  }

  if (src.includes(':') && /^[A-Z]+\d+:[A-Z]+\d+$/i.test(src.replace(/\$/g, ''))) {
    return collect(src, getValue).map(toNumber).reduce((a, b) => a + b, 0);
  }

  const cell = parseA1(src);
  if (cell) return getValue(src);

  return evalCondition(src, getValue);
}

function evalCondition(src, getValue) {
  const m = String(src).match(/(.+?)(>=|<=|<>|>|<|=)(.+)/);
  if (!m) {
    const v = parseA1(src) ? getValue(src) : src;
    return v === true || v === 1 || String(v).toLowerCase() === 'true';
  }
  const left = evalExpr(m[1].trim(), getValue);
  const right = evalExpr(m[3].trim(), getValue);
  const op = m[2];
  const nl = toNumber(left);
  const nr = toNumber(right);
  if (op === '>') return nl > nr;
  if (op === '<') return nl < nr;
  if (op === '>=') return nl >= nr;
  if (op === '<=') return nl <= nr;
  if (op === '<>') return String(left) !== String(right);
  return String(left) === String(right) || nl === nr;
}

export function evaluateCellFormula(formula, getRaw) {
  const seen = new Set();
  const getValue = (addr) => {
    const key = String(addr).toUpperCase().replace(/\$/g, '');
    if (seen.has(key)) return 0;
    seen.add(key);
    const raw = getRaw(key);
    if (raw && typeof raw === 'object' && raw.f) {
      try { return evalExpr(raw.f, getValue); } catch { return 0; }
    }
    if (raw && typeof raw === 'object') return raw.v;
    return raw ?? '';
  };
  try {
    return evalExpr(formula, getValue);
  } catch {
    return '#VALUE!';
  }
}
