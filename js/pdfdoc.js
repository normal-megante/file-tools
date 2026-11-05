'use strict';
// =========================================================
// PDF を読んで、ページを組み替えて書き出す部品
//   ・PDF の中身（オブジェクト）を読む … xref 表・xref ストリーム・オブジェクトストリーム・追記された更新
//   ・こわれた PDF は、ファイル全体を探して立て直す
//   ・保護（暗号化）された PDF … パスワードなしで開けるもの（印刷禁止などの制限だけのもの）は解除して読む
//   ・ページを選んで新しい PDF にする（結合・抜き出し・並べ替え・回転）
// ページの中身（文字・図・画像）はそのままコピーするので、画質は落ちず、文字も選べるまま。
// 画面には依存しない（Node でもテストできる）
// =========================================================
const PdfDoc = (() => {

class PName { constructor(n) { this.n = n; } }
class PRef { constructor(num, gen) { this.num = num; this.gen = gen; } }
class PStr { constructor(b, hex) { this.b = b; this.hex = hex; } }
class PStream { constructor(dict, data) { this.dict = dict; this.data = data; } }
class PKw { constructor(k) { this.k = k; } }      // キーワード（obj, endobj, R など）
class ORef { constructor(id) { this.id = id; } }  // 書き出す側の参照

class PdfError extends Error { constructor(code, msg) { super(msg || code); this.code = code; } }

const nameOf = v => (v instanceof PName ? v.n : null);
const enc = s => { const u = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) u[i] = s.charCodeAt(i) & 255; return u; };
function concat(parts) {
  let n = 0; for (const p of parts) n += p.length;
  const out = new Uint8Array(n); let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}
function indexOf(b, str, from = 0) {
  const c0 = str.charCodeAt(0), n = str.length;
  outer: for (let i = b.indexOf(c0, from); i >= 0 && i + n <= b.length; i = b.indexOf(c0, i + 1)) {
    for (let k = 1; k < n; k++) if (b[i + k] !== str.charCodeAt(k)) continue outer;
    return i;
  }
  return -1;
}
const startsWith = (b, p, str) => { for (let k = 0; k < str.length; k++) if (b[p + k] !== str.charCodeAt(k)) return false; return true; };

/* ---------- 字句の読み取り ---------- */
const WS = new Uint8Array(256); for (const c of [0, 9, 10, 12, 13, 32]) WS[c] = 1;
const DL = new Uint8Array(256); for (const c of '()<>[]{}/%') DL[c.charCodeAt(0)] = 1;
const hexVal = c => (c >= 48 && c <= 57 ? c - 48 : c >= 65 && c <= 70 ? c - 55 : c >= 97 && c <= 102 ? c - 87 : -1);
const NUM_RE = /^[+-]?(\d+\.?\d*|\.\d+)$/;

class Parser {
  constructor(b, p = 0) { this.b = b; this.p = p; }
  skipWs() {
    const b = this.b; let p = this.p;
    for (;;) {
      while (p < b.length && WS[b[p]]) p++;
      if (b[p] === 37) { while (p < b.length && b[p] !== 10 && b[p] !== 13) p++; continue; } // % コメント
      break;
    }
    this.p = p;
  }
  // 値を1つ読む（数・名前・文字列・配列・辞書・参照）。キーワードなら PKw
  next() {
    this.skipWs();
    const b = this.b; let p = this.p;
    if (p >= b.length) throw new PdfError('broken', 'ファイルの終わりに来ました');
    const c = b[p];
    if (c === 47) { // /名前
      p++; let s = '';
      while (p < b.length && !WS[b[p]] && !DL[b[p]]) {
        if (b[p] === 35 && hexVal(b[p + 1]) >= 0 && hexVal(b[p + 2]) >= 0) { s += String.fromCharCode(hexVal(b[p + 1]) * 16 + hexVal(b[p + 2])); p += 3; }
        else s += String.fromCharCode(b[p++]);
      }
      this.p = p; return new PName(s);
    }
    if (c === 40) return this.litString();
    if (c === 60) {
      if (b[p + 1] === 60) { this.p = p + 2; return this.dict(); }
      return this.hexString();
    }
    if (c === 91) {
      this.p = p + 1; const a = [];
      for (;;) {
        this.skipWs();
        if (this.p >= b.length) throw new PdfError('broken', '配列が閉じていません');
        if (b[this.p] === 93) { this.p++; return a; }
        const v = this.value();
        if (v instanceof PKw) { if (v.k === 'endobj' || v.k === '>>') { this.p -= v.k.length; return a; } continue; }
        a.push(v);
      }
    }
    if (c === 62 && b[p + 1] === 62) { this.p = p + 2; return new PKw('>>'); }
    if (DL[c]) { this.p = p + 1; return new PKw(String.fromCharCode(c)); } // 余計な記号
    let s = '';
    while (p < b.length && !WS[b[p]] && !DL[b[p]]) s += String.fromCharCode(b[p++]);
    this.p = p;
    if (NUM_RE.test(s)) return +s;
    if (/^[+\-.\d]+$/.test(s)) { const v = parseFloat(s.replace(/^[+-]+/, m => (m.includes('-') ? '-' : ''))); return isNaN(v) ? 0 : v; }
    if (s === 'true') return true;
    if (s === 'false') return false;
    if (s === 'null') return null;
    return new PKw(s);
  }
  // 整数のあとに「数 R」が続けば参照
  value() {
    const v = this.next();
    if (typeof v === 'number' && v >= 0 && Number.isInteger(v)) {
      const b = this.b; let p = this.p;
      if (WS[b[p]]) {
        while (WS[b[p]]) p++;
        let g = 0, d = 0;
        while (b[p] >= 48 && b[p] <= 57) { g = g * 10 + b[p] - 48; p++; d++; }
        if (d && WS[b[p]]) {
          while (WS[b[p]]) p++;
          if (b[p] === 82 && (p + 1 >= b.length || WS[b[p + 1]] || DL[b[p + 1]])) { this.p = p + 1; return new PRef(v, g); }
        }
      }
    }
    return v;
  }
  dict() {
    const d = new Map(), b = this.b;
    for (;;) {
      this.skipWs();
      if (this.p >= b.length) throw new PdfError('broken', '辞書が閉じていません');
      if (b[this.p] === 62 && b[this.p + 1] === 62) { this.p += 2; return d; }
      const k = this.next();
      if (!(k instanceof PName)) {
        if (k instanceof PKw && (k.k === 'endobj' || k.k === 'stream' || k.k === 'endstream')) { this.p -= k.k.length; return d; }
        continue; // 変なものは読み飛ばす
      }
      this.skipWs();
      if (b[this.p] === 62 && b[this.p + 1] === 62) continue; // 値のないキー
      const v = this.value();
      if (v instanceof PKw) {
        if (v.k === 'endobj' || v.k === 'stream' || v.k === 'endstream') { this.p -= v.k.length; return d; }
        continue;
      }
      d.set(k.n, v);
    }
  }
  litString() {
    const b = this.b; let p = this.p + 1, depth = 1;
    const out = [];
    while (p < b.length) {
      const c = b[p++];
      if (c === 92) {
        const e = b[p++];
        switch (e) {
          case 110: out.push(10); break;
          case 114: out.push(13); break;
          case 116: out.push(9); break;
          case 98: out.push(8); break;
          case 102: out.push(12); break;
          case 13: if (b[p] === 10) p++; break; // 行の続き
          case 10: break;
          default:
            if (e >= 48 && e <= 55) {
              let v = e - 48;
              for (let k = 0; k < 2 && b[p] >= 48 && b[p] <= 55; k++) v = v * 8 + b[p++] - 48;
              out.push(v & 255);
            } else if (e !== undefined) out.push(e);
        }
      } else if (c === 40) { depth++; out.push(c); }
      else if (c === 41) { if (--depth === 0) break; out.push(c); }
      else out.push(c);
    }
    this.p = p;
    return new PStr(Uint8Array.from(out), false);
  }
  hexString() {
    const b = this.b; let p = this.p + 1, hi = -1;
    const out = [];
    while (p < b.length && b[p] !== 62) {
      const v = hexVal(b[p++]);
      if (v < 0) continue;
      if (hi < 0) hi = v; else { out.push(hi * 16 + v); hi = -1; }
    }
    if (hi >= 0) out.push(hi * 16);
    this.p = p + 1;
    return new PStr(Uint8Array.from(out), true);
  }
}

/* ---------- ストリームの展開（xref ストリーム・オブジェクトストリーム用） ---------- */
async function inflate(data, raw = false) {
  const ds = new DecompressionStream(raw ? 'deflate-raw' : 'deflate');
  const w = ds.writable.getWriter();
  w.write(data).catch(() => {}); w.close().catch(() => {});
  const r = ds.readable.getReader(), chunks = [];
  try { for (;;) { const { done, value } = await r.read(); if (done) break; chunks.push(value); } }
  catch (e) { if (!chunks.length) { if (!raw && data.length > 2) return inflate(data.subarray(2), true); throw e; } }
  return concat(chunks);
}
function unpredict(data, parms) {
  const pred = (parms && parms.get('Predictor')) || 1;
  if (pred < 2) return data;
  const colors = parms.get('Colors') || 1, bpc = parms.get('BitsPerComponent') || 8, cols = parms.get('Columns') || 1;
  const bpp = Math.max(1, Math.ceil(colors * bpc / 8)), rl = Math.ceil(colors * bpc * cols / 8);
  if (pred === 2) {
    const out = data.slice();
    if (bpc === 8) for (let o = 0; o + rl <= out.length; o += rl) for (let i = bpp; i < rl; i++) out[o + i] = (out[o + i] + out[o + i - bpp]) & 255;
    return out;
  }
  const rows = Math.floor(data.length / (rl + 1)), out = new Uint8Array(rows * rl);
  let prev = new Uint8Array(rl);
  for (let y = 0; y < rows; y++) {
    const ft = data[y * (rl + 1)], src = data.subarray(y * (rl + 1) + 1, (y + 1) * (rl + 1)), cur = out.subarray(y * rl, (y + 1) * rl);
    for (let i = 0; i < rl; i++) {
      const a = i >= bpp ? cur[i - bpp] : 0, b = prev[i], c = i >= bpp ? prev[i - bpp] : 0;
      let p = 0;
      if (ft === 1) p = a;
      else if (ft === 2) p = b;
      else if (ft === 3) p = (a + b) >> 1;
      else if (ft === 4) { const q = a + b - c, pa = Math.abs(q - a), pb = Math.abs(q - b), pc = Math.abs(q - c); p = pa <= pb && pa <= pc ? a : pb <= pc ? b : c; }
      cur[i] = (src[i] + p) & 255;
    }
    prev = cur;
  }
  return out;
}
async function decodeStream(st) {
  let filters = st.dict.get('Filter'), parms = st.dict.get('DecodeParms');
  if (!Array.isArray(filters)) filters = filters ? [filters] : [];
  if (!Array.isArray(parms)) parms = [parms];
  let data = st.data;
  for (let i = 0; i < filters.length; i++) {
    const f = nameOf(filters[i]);
    if (f === 'FlateDecode' || f === 'Fl') data = unpredict(await inflate(data), parms[i] instanceof Map ? parms[i] : null);
    else if (f === 'Crypt') continue;
    else throw new PdfError('broken', '読めない圧縮形式です: ' + f);
  }
  return data;
}

/* ---------- 暗号（保護の解除用）：MD5・RC4・AES ---------- */
const MD5_S = [7, 12, 17, 22, 5, 9, 14, 20, 4, 11, 16, 23, 6, 10, 15, 21];
const MD5_K = Int32Array.from({ length: 64 }, (_, i) => Math.floor(Math.abs(Math.sin(i + 1)) * 4294967296) | 0);
function md5(data) {
  const n = data.length, len = ((n + 72) >>> 6) << 6, buf = new Uint8Array(len);
  buf.set(data); buf[n] = 0x80;
  const lo = (n * 8) >>> 0, hi = Math.floor(n / 0x20000000);
  for (let k = 0; k < 4; k++) { buf[len - 8 + k] = (lo >>> (8 * k)) & 255; buf[len - 4 + k] = (hi >>> (8 * k)) & 255; }
  let a0 = 0x67452301, b0 = 0xefcdab89 | 0, c0 = 0x98badcfe | 0, d0 = 0x10325476;
  const M = new Int32Array(16);
  for (let off = 0; off < len; off += 64) {
    for (let j = 0; j < 16; j++) M[j] = buf[off + 4 * j] | (buf[off + 4 * j + 1] << 8) | (buf[off + 4 * j + 2] << 16) | (buf[off + 4 * j + 3] << 24);
    let A = a0, B = b0, C = c0, D = d0;
    for (let i = 0; i < 64; i++) {
      let F, g;
      if (i < 16) { F = (B & C) | (~B & D); g = i; }
      else if (i < 32) { F = (D & B) | (~D & C); g = (5 * i + 1) & 15; }
      else if (i < 48) { F = B ^ C ^ D; g = (3 * i + 5) & 15; }
      else { F = C ^ (B | ~D); g = (7 * i) & 15; }
      F = (F + A + MD5_K[i] + M[g]) | 0;
      const s = MD5_S[(i >> 4) * 4 + (i & 3)];
      A = D; D = C; C = B;
      B = (B + ((F << s) | (F >>> (32 - s)))) | 0;
    }
    a0 = (a0 + A) | 0; b0 = (b0 + B) | 0; c0 = (c0 + C) | 0; d0 = (d0 + D) | 0;
  }
  const out = new Uint8Array(16);
  [a0, b0, c0, d0].forEach((v, k) => { for (let j = 0; j < 4; j++) out[4 * k + j] = (v >>> (8 * j)) & 255; });
  return out;
}
function rc4(key, data) {
  const s = new Uint8Array(256);
  for (let i = 0; i < 256; i++) s[i] = i;
  for (let i = 0, j = 0; i < 256; i++) { j = (j + s[i] + key[i % key.length]) & 255; const t = s[i]; s[i] = s[j]; s[j] = t; }
  const out = new Uint8Array(data.length);
  for (let k = 0, i = 0, j = 0; k < data.length; k++) {
    i = (i + 1) & 255; j = (j + s[i]) & 255;
    const t = s[i]; s[i] = s[j]; s[j] = t;
    out[k] = data[k] ^ s[(s[i] + s[j]) & 255];
  }
  return out;
}
const SBOX = new Uint8Array(256), INV = new Uint8Array(256);
const xtime = a => ((a << 1) ^ (a & 0x80 ? 0x1b : 0)) & 255;
(() => {
  const rotl = (x, s) => ((x << s) | (x >> (8 - s))) & 255;
  let p = 1, q = 1;
  do {
    p = (p ^ (p << 1) ^ (p & 0x80 ? 0x1b : 0)) & 255;
    q = (q ^ (q << 1)) & 255; q = (q ^ (q << 2)) & 255; q = (q ^ (q << 4)) & 255;
    if (q & 0x80) q ^= 0x09;
    SBOX[p] = (q ^ rotl(q, 1) ^ rotl(q, 2) ^ rotl(q, 3) ^ rotl(q, 4) ^ 0x63) & 255;
  } while (p !== 1);
  SBOX[0] = 0x63;
  for (let i = 0; i < 256; i++) INV[SBOX[i]] = i;
})();
const gmul = (a, b) => { let r = 0; while (b) { if (b & 1) r ^= a; a = xtime(a); b >>= 1; } return r; };
const [M2, M3, M9, M11, M13, M14] = [2, 3, 9, 11, 13, 14].map(m => Uint8Array.from({ length: 256 }, (_, i) => gmul(i, m)));
function aesKey(key) {
  const nk = key.length / 4, nr = nk + 6, w = new Uint8Array(16 * (nr + 1));
  w.set(key);
  let rcon = 1;
  for (let i = nk; i < 4 * (nr + 1); i++) {
    let t0 = w[4 * i - 4], t1 = w[4 * i - 3], t2 = w[4 * i - 2], t3 = w[4 * i - 1];
    if (i % nk === 0) { const t = t0; t0 = SBOX[t1] ^ rcon; t1 = SBOX[t2]; t2 = SBOX[t3]; t3 = SBOX[t]; rcon = xtime(rcon); }
    else if (nk > 6 && i % nk === 4) { t0 = SBOX[t0]; t1 = SBOX[t1]; t2 = SBOX[t2]; t3 = SBOX[t3]; }
    const j = 4 * (i - nk);
    w[4 * i] = w[j] ^ t0; w[4 * i + 1] = w[j + 1] ^ t1; w[4 * i + 2] = w[j + 2] ^ t2; w[4 * i + 3] = w[j + 3] ^ t3;
  }
  return { w, nr };
}
const S = new Uint8Array(16), T = new Uint8Array(16);
function aesEncBlock(k, inp, out) {
  const w = k.w;
  for (let i = 0; i < 16; i++) S[i] = inp[i] ^ w[i];
  for (let r = 1; r <= k.nr; r++) {
    for (let c = 0; c < 4; c++) for (let row = 0; row < 4; row++) T[row + 4 * c] = SBOX[S[row + 4 * ((c + row) & 3)]];
    if (r < k.nr) {
      for (let c = 0; c < 16; c += 4) {
        const a0 = T[c], a1 = T[c + 1], a2 = T[c + 2], a3 = T[c + 3];
        S[c] = M2[a0] ^ M3[a1] ^ a2 ^ a3; S[c + 1] = a0 ^ M2[a1] ^ M3[a2] ^ a3;
        S[c + 2] = a0 ^ a1 ^ M2[a2] ^ M3[a3]; S[c + 3] = M3[a0] ^ a1 ^ a2 ^ M2[a3];
      }
    } else S.set(T);
    for (let i = 0; i < 16; i++) S[i] ^= w[16 * r + i];
  }
  out.set(S);
}
function aesDecBlock(k, inp, out) {
  const w = k.w;
  for (let i = 0; i < 16; i++) S[i] = inp[i] ^ w[16 * k.nr + i];
  for (let r = k.nr - 1; r >= 0; r--) {
    for (let c = 0; c < 4; c++) for (let row = 0; row < 4; row++) T[row + 4 * ((c + row) & 3)] = INV[S[row + 4 * c]];
    for (let i = 0; i < 16; i++) T[i] ^= w[16 * r + i];
    if (r > 0) {
      for (let c = 0; c < 16; c += 4) {
        const a0 = T[c], a1 = T[c + 1], a2 = T[c + 2], a3 = T[c + 3];
        S[c] = M14[a0] ^ M11[a1] ^ M13[a2] ^ M9[a3]; S[c + 1] = M9[a0] ^ M14[a1] ^ M11[a2] ^ M13[a3];
        S[c + 2] = M13[a0] ^ M9[a1] ^ M14[a2] ^ M11[a3]; S[c + 3] = M11[a0] ^ M13[a1] ^ M9[a2] ^ M14[a3];
      }
    } else S.set(T);
  }
  out.set(S);
}
function aesCbcDecrypt(key, iv, ct, unpad) {
  const k = aesKey(key), n = ct.length & ~15, out = new Uint8Array(n), blk = new Uint8Array(16);
  for (let o = 0; o < n; o += 16) {
    aesDecBlock(k, ct.subarray(o, o + 16), blk);
    for (let i = 0; i < 16; i++) out[o + i] = blk[i] ^ iv[i];
    iv = ct.subarray(o, o + 16);
  }
  if (unpad && n) { const p = out[n - 1]; if (p >= 1 && p <= 16) return out.subarray(0, n - p); }
  return out;
}
function aesCbcEncrypt(key, iv, data) { // 詰め物なし（長さは 16 の倍数）
  const k = aesKey(key), out = new Uint8Array(data.length), blk = new Uint8Array(16);
  for (let o = 0; o < data.length; o += 16) {
    for (let i = 0; i < 16; i++) blk[i] = data[o + i] ^ iv[i];
    aesEncBlock(k, blk, out.subarray(o, o + 16));
    iv = out.subarray(o, o + 16);
  }
  return out;
}
async function sha(alg, data) { return new Uint8Array(await crypto.subtle.digest(alg, data)); }
const eqBytes = (a, b) => a.length === b.length && a.every((v, i) => v === b[i]);
const PAD = Uint8Array.from([0x28, 0xBF, 0x4E, 0x5E, 0x4E, 0x75, 0x8A, 0x41, 0x64, 0x00, 0x4E, 0x56, 0xFF, 0xFA, 0x01, 0x08,
  0x2E, 0x2E, 0x00, 0xB6, 0xD0, 0x68, 0x3E, 0x80, 0x2F, 0x0C, 0xA9, 0xFE, 0x64, 0x53, 0x69, 0x7A]);

// R6 のハッシュ（ISO 32000-2 の方法 2.B）
async function hash2B(pw, salt, u) {
  let K = await sha('SHA-256', concat([pw, salt, u])), E, i = 0;
  while (i < 64 || E[E.length - 1] > i - 32) {
    const k1 = concat([pw, K, u]), K1 = new Uint8Array(k1.length * 64);
    for (let j = 0; j < 64; j++) K1.set(k1, j * k1.length);
    E = aesCbcEncrypt(K.subarray(0, 16), K.subarray(16, 32), K1);
    let sum = 0; for (let j = 0; j < 16; j++) sum += E[j];
    K = await sha(['SHA-256', 'SHA-384', 'SHA-512'][sum % 3], E);
    i++;
  }
  return K.subarray(0, 32);
}

class Crypt {
  static async create(encDict, id0, get) {
    if (nameOf(await get(encDict.get('Filter'))) !== 'Standard') throw new PdfError('unsupported', '対応していない保護のしかたです');
    const c = new Crypt(), num = async k => { const v = await get(encDict.get(k)); return typeof v === 'number' ? v : null; };
    const bytes = async k => { const v = await get(encDict.get(k)); return v instanceof PStr ? v.b : new Uint8Array(0); };
    const V = (await num('V')) || 0, R = (await num('R')) || 2, P = (await num('P')) | 0;
    const O = await bytes('O'), U = await bytes('U');
    c.encMeta = (await get(encDict.get('EncryptMetadata'))) !== false;
    let bits = (await num('Length')) || 40;
    c.stm = c.str = 'rc4';
    if (V >= 4) {
      const cf = await get(encDict.get('CF')), method = async nm => {
        if (!nm || nm === 'Identity') return 'none';
        const f = cf instanceof Map ? await get(cf.get(nm)) : null;
        const cfm = f instanceof Map ? nameOf(await get(f.get('CFM'))) : null;
        if (f instanceof Map && typeof f.get('Length') === 'number') bits = f.get('Length') <= 32 ? f.get('Length') * 8 : f.get('Length');
        return cfm === 'V2' ? 'rc4' : cfm === 'AESV2' ? 'aes' : cfm === 'AESV3' ? 'aes256' : 'none';
      };
      c.stm = await method(nameOf(await get(encDict.get('StmF'))));
      c.str = await method(nameOf(await get(encDict.get('StrF'))));
      if (V === 4) bits = 128;
    }
    if (R >= 5) {
      // AES-256。空のパスワードで開けるか試す
      const pw = new Uint8Array(0), ue = await bytes('UE'), oe = await bytes('OE');
      const h = (salt, u) => (R === 5 ? sha('SHA-256', concat([pw, salt, u])) : hash2B(pw, salt, u));
      let key = null;
      if (eqBytes((await h(U.subarray(32, 40), new Uint8Array(0))).subarray(0, 32), U.subarray(0, 32)))
        key = aesCbcDecrypt(await h(U.subarray(40, 48), new Uint8Array(0)), new Uint8Array(16), ue, false);
      else if (eqBytes((await h(O.subarray(32, 40), U.subarray(0, 48))).subarray(0, 32), O.subarray(0, 32)))
        key = aesCbcDecrypt(await h(O.subarray(40, 48), U.subarray(0, 48)), new Uint8Array(16), oe, false);
      if (!key) throw new PdfError('password', 'パスワードが必要な PDF です');
      c.key = key.subarray(0, 32); c.aes256 = true;
      return c;
    }
    const n = R === 2 ? 5 : Math.max(5, Math.min(16, bits >> 3));
    const makeKey = pwPadded => {
      const parts = [pwPadded, O.subarray(0, 32), Uint8Array.of(P & 255, (P >> 8) & 255, (P >> 16) & 255, (P >>> 24) & 255), id0];
      if (R >= 4 && !c.encMeta) parts.push(Uint8Array.of(255, 255, 255, 255));
      let hh = md5(concat(parts));
      if (R >= 3) for (let i = 0; i < 50; i++) hh = md5(hh.subarray(0, n));
      return hh.subarray(0, n);
    };
    const checkUser = key => {
      if (R === 2) return eqBytes(rc4(key, PAD), U.subarray(0, 32));
      let x = rc4(key, md5(concat([PAD, id0])));
      for (let i = 1; i <= 19; i++) x = rc4(key.map(b => b ^ i), x);
      return eqBytes(x.subarray(0, 16), U.subarray(0, 16));
    };
    let key = makeKey(PAD);
    if (!checkUser(key)) {
      // 所有者パスワードが空の場合（まれ）
      let hh = md5(PAD);
      if (R >= 3) for (let i = 0; i < 50; i++) hh = md5(hh.subarray(0, n));
      const ok = hh.subarray(0, n);
      let up = O.subarray(0, 32);
      if (R === 2) up = rc4(ok, up);
      else for (let i = 19; i >= 0; i--) up = rc4(ok.map(b => b ^ i), up);
      key = makeKey(up);
      if (!checkUser(key)) throw new PdfError('password', 'パスワードが必要な PDF です');
    }
    c.key = key;
    return c;
  }
  bytes(data, num, gen, method) {
    if (method === 'none' || !data.length) return data;
    if (method === 'aes256') return data.length < 16 ? new Uint8Array(0) : aesCbcDecrypt(this.key, data.subarray(0, 16), data.subarray(16), true);
    const aes = method === 'aes';
    const k = md5(concat([this.key, Uint8Array.of(num & 255, (num >> 8) & 255, (num >> 16) & 255, gen & 255, (gen >> 8) & 255), aes ? enc('sAlT') : new Uint8Array(0)]))
      .subarray(0, Math.min(this.key.length + 5, 16));
    if (!aes) return rc4(k, data);
    return data.length < 16 ? new Uint8Array(0) : aesCbcDecrypt(k, data.subarray(0, 16), data.subarray(16), true);
  }
  value(v, num, gen) {
    if (v instanceof PStr) return new PStr(this.bytes(v.b, num, gen, this.aes256 && this.str !== 'none' ? 'aes256' : this.str), v.hex);
    if (Array.isArray(v)) return v.map(x => this.value(x, num, gen));
    if (v instanceof Map) { const m = new Map(); for (const [k, x] of v) m.set(k, this.value(x, num, gen)); return m; }
    return v;
  }
  object(v, num, gen) {
    if (!(v instanceof PStream)) return this.value(v, num, gen);
    const type = nameOf(v.dict.get('Type'));
    if (type === 'XRef') return v;
    const dict = this.value(v.dict, num, gen);
    let method = this.aes256 && this.stm !== 'none' ? 'aes256' : this.stm;
    if (type === 'Metadata' && !this.encMeta) method = 'none';
    // ストリーム自身に Crypt フィルタが付いている場合（Identity なら暗号なし）
    let f = dict.get('Filter');
    if (Array.isArray(f) ? f.some(x => nameOf(x) === 'Crypt') : nameOf(f) === 'Crypt') {
      method = 'none';
      if (Array.isArray(f)) {
        const parms = dict.get('DecodeParms'), keep = f.map((x, i) => nameOf(x) !== 'Crypt' ? i : -1).filter(i => i >= 0);
        dict.set('Filter', keep.map(i => f[i]));
        if (Array.isArray(parms)) dict.set('DecodeParms', keep.map(i => parms[i] ?? null));
      } else { dict.delete('Filter'); dict.delete('DecodeParms'); }
    }
    return new PStream(dict, this.bytes(v.data, num, gen, method));
  }
}

/* ---------- PDF を読む ---------- */
class Doc {
  constructor(buf) {
    this.buf = buf; this.xref = new Map(); this.cache = new Map(); this.objStms = new Map();
    this.trailer = new Map(); this.crypt = null; this.encryptNum = -1; this.scanned = null; this.repaired = false;
  }
  async get(v) { return v instanceof PRef ? this.getObj(v.num) : v; }
  getObj(num) {
    let p = this.cache.get(num);
    if (!p) { p = this.loadObj(num).catch(() => null); this.cache.set(num, p); }
    return p;
  }
  async loadObj(num) {
    const e = this.xref.get(num);
    if (e && e.type === 2) {
      const s = await this.loadObjStm(e.stm);
      let k = e.idx;
      if (!s.offs[k] || s.offs[k][0] !== num) k = s.offs.findIndex(o => o[0] === num);
      if (k < 0) return null;
      return new Parser(s.data, s.first + s.offs[k][1]).value();
    }
    if (e) {
      try { return await this.readObjAt(e.off, num); }
      catch (err) { /* 位置がずれている → 下で探し直す */ }
    }
    const off = this.scanIndex().get(num);
    if (off == null || (e && off === e.off)) return null;
    return this.readObjAt(off, num);
  }
  async loadObjStm(num) {
    let p = this.objStms.get(num);
    if (!p) {
      p = (async () => {
        const st = await this.getObj(num);
        if (!(st instanceof PStream)) throw new PdfError('broken');
        const data = await decodeStream(st), n = st.dict.get('N') || 0, first = st.dict.get('First') || 0;
        const ps = new Parser(data, 0), offs = [];
        for (let i = 0; i < n; i++) { const a = ps.next(), b = ps.next(); if (typeof a !== 'number' || typeof b !== 'number') break; offs.push([a, b]); }
        return { data, first, offs };
      })();
      this.objStms.set(num, p);
    }
    return p;
  }
  async readObjAt(off, num, decrypt = true) {
    const b = this.buf, ps = new Parser(b, off);
    const n = ps.next(), g = ps.next(), kw = ps.next();
    if ((num != null && n !== num) || !(kw instanceof PKw && kw.k === 'obj')) throw new PdfError('broken', 'オブジェクトの位置が違います');
    let v = ps.value();
    if (v instanceof PKw) v = null;
    if (v instanceof Map) {
      ps.skipWs();
      if (startsWith(b, ps.p, 'stream')) {
        let p = ps.p + 6;
        if (b[p] === 13) p++;
        if (b[p] === 10) p++;
        let len = v.get('Length');
        if (len instanceof PRef) len = await this.getObj(len.num);
        let end = -1;
        if (typeof len === 'number' && len >= 0 && p + len <= b.length) {
          const q = new Parser(b, p + len); q.skipWs();
          if (startsWith(b, q.p, 'endstream')) end = p + len;
        }
        if (end < 0) { // 長さが合わないときは endstream を探す
          end = indexOf(b, 'endstream', p);
          if (end < 0) throw new PdfError('broken', 'ストリームの終わりが見つかりません');
          if (b[end - 1] === 10) end--;
          if (b[end - 1] === 13) end--;
        }
        v = new PStream(v, b.subarray(p, end));
      }
    }
    if (decrypt && this.crypt && n !== this.encryptNum) v = this.crypt.object(v, n, g);
    return v;
  }
  // ファイル全体から「番号 世代 obj」を探す（こわれた PDF 用。あとに出てきたものが新しい）
  scanIndex() {
    if (this.scanned) return this.scanned;
    const b = this.buf, m = new Map();
    for (let i = indexOf(b, 'obj'); i >= 0; i = indexOf(b, 'obj', i + 3)) {
      if (b[i + 3] !== undefined && !WS[b[i + 3]] && !DL[b[i + 3]]) continue;
      let p = i - 1;
      if (!WS[b[p]]) continue;
      while (p >= 0 && WS[b[p]]) p--;
      let e = p; while (p >= 0 && b[p] >= 48 && b[p] <= 57) p--;
      if (p === e || !WS[b[p]]) continue;
      while (p >= 0 && WS[b[p]]) p--;
      e = p; while (p >= 0 && b[p] >= 48 && b[p] <= 57) p--;
      if (p === e || (p >= 0 && !WS[b[p]] && !DL[b[p]])) continue;
      let num = 0; for (let k = p + 1; k <= e; k++) num = num * 10 + b[k] - 48;
      m.set(num, p + 1);
    }
    return (this.scanned = m);
  }
  findStartxref() {
    const b = this.buf;
    for (let back = 1024; ; back *= 4) {
      const from = Math.max(0, b.length - back);
      let i = -1;
      for (let k = indexOf(b, 'startxref', from); k >= 0; k = indexOf(b, 'startxref', k + 9)) i = k;
      if (i >= 0) { const v = new Parser(b, i + 9).next(); if (typeof v === 'number') return v; }
      if (from === 0) return null;
    }
  }
  async readXrefAt(off) {
    const b = this.buf, ps = new Parser(b, off);
    ps.skipWs();
    if (startsWith(b, ps.p, 'xref')) {
      ps.p += 4;
      for (;;) {
        ps.skipWs();
        if (startsWith(b, ps.p, 'trailer')) { ps.p += 7; const t = ps.value(); if (!(t instanceof Map)) throw new PdfError('broken'); return t; }
        let start = ps.next(); const count = ps.next();
        if (typeof start !== 'number' || typeof count !== 'number') throw new PdfError('broken', 'xref 表が読めません');
        for (let i = 0; i < count; i++) {
          const o = ps.next(), g = ps.next(), t = ps.next();
          if (!(t instanceof PKw)) throw new PdfError('broken', 'xref 表が読めません');
          if (i === 0 && start === 1 && t.k === 'f' && o === 0 && g === 65535) start = 0; // よくあるずれ
          const num = start + i;
          if (t.k === 'n' && o > 0 && !this.xref.has(num)) this.xref.set(num, { type: 1, off: o });
        }
      }
    }
    // xref ストリーム
    const st = await this.readObjAt(ps.p, null, false);
    if (!(st instanceof PStream)) throw new PdfError('broken', 'xref が見つかりません');
    const data = await decodeStream(st), W = st.dict.get('W'), size = st.dict.get('Size') || 0;
    const index = st.dict.get('Index') || [0, size], ew = W[0] + W[1] + W[2];
    const field = (o, w, def) => { if (!w) return def; let v = 0; for (let k = 0; k < w; k++) v = v * 256 + data[o + k]; return v; };
    let o = 0;
    for (let s = 0; s + 1 < index.length; s += 2) {
      for (let i = 0; i < index[s + 1] && o + ew <= data.length; i++, o += ew) {
        const type = field(o, W[0], 1), f2 = field(o + W[0], W[1], 0), f3 = field(o + W[0] + W[1], W[2], 0), num = index[s] + i;
        if (this.xref.has(num)) continue;
        if (type === 1 && f2 > 0) this.xref.set(num, { type: 1, off: f2 });
        else if (type === 2) this.xref.set(num, { type: 2, stm: f2, idx: f3 });
      }
    }
    return st.dict;
  }
  async readXrefChain() {
    let off = this.findStartxref();
    if (off == null) throw new PdfError('broken', 'startxref がありません');
    const seen = new Set(), trailers = [];
    while (off != null && !seen.has(off)) {
      seen.add(off);
      const t = await this.readXrefAt(off);
      trailers.push(t);
      const xs = t.get('XRefStm');
      if (typeof xs === 'number' && !seen.has(xs)) { seen.add(xs); try { await this.readXrefAt(xs); } catch (e) { /* なくても読める */ } }
      const prev = t.get('Prev');
      off = typeof prev === 'number' ? prev : null;
    }
    for (const t of trailers) for (const [k, v] of t) if (!this.trailer.has(k) && k !== 'Prev' && k !== 'XRefStm') this.trailer.set(k, v);
  }
  // xref が使えないとき：ファイル全体を調べて立て直す
  async rebuild() {
    this.repaired = true;
    this.xref.clear(); this.cache.clear(); this.objStms.clear(); this.trailer = new Map();
    const b = this.buf, idx = this.scanIndex();
    for (const [num, off] of idx) this.xref.set(num, { type: 1, off });
    // trailer を探す（あとのものほど新しい）
    const trailers = [];
    for (let i = indexOf(b, 'trailer'); i >= 0; i = indexOf(b, 'trailer', i + 7)) {
      try { const t = new Parser(b, i + 7).value(); if (t instanceof Map) trailers.unshift(t); } catch (e) { /* 次へ */ }
    }
    // オブジェクトストリームの中身と、xref ストリームの辞書（trailer の代わり）
    const head = off => new TextDecoder('latin1').decode(b.subarray(off, Math.min(b.length, off + 400)));
    for (const [num, off] of [...idx].sort((x, y) => y[1] - x[1])) {
      const h = head(off);
      if (/\/Type\s*\/XRef\b/.test(h)) {
        try { const st = await this.readObjAt(off, num, false); if (st instanceof PStream) trailers.push(st.dict); } catch (e) { /* 次へ */ }
      }
    }
    for (const t of trailers) for (const [k, v] of t) if (!this.trailer.has(k)) this.trailer.set(k, v);
    await this.setupCrypt();
    for (const [num, off] of idx) {
      if (!/\/Type\s*\/ObjStm\b/.test(head(off))) continue;
      try {
        const s = await this.loadObjStm(num);
        s.offs.forEach(([n], k) => { if (!idx.has(n)) this.xref.set(n, { type: 2, stm: num, idx: k }); });
      } catch (e) { /* 次へ */ }
    }
    if (!(await this.catalogOk())) {
      // trailer がない → Catalog を自分で探す
      let best = null;
      for (const [num, off] of idx) if (/\/Type\s*\/Catalog\b/.test(head(off)) && (!best || off > best[1])) best = [num, off];
      if (best) this.trailer.set('Root', new PRef(best[0], 0));
    }
  }
  async setupCrypt() {
    this.crypt = null; this.encryptNum = -1;
    let e = this.trailer.get('Encrypt');
    if (!e) return;
    if (e instanceof PRef) {
      this.encryptNum = e.num;
      const x = this.xref.get(e.num);
      const off = x && x.type === 1 ? x.off : this.scanIndex().get(e.num);
      e = off != null ? await this.readObjAt(off, e.num, false) : null;
    }
    if (!(e instanceof Map)) throw new PdfError('broken', '保護の情報が読めません');
    const id = this.trailer.get('ID'), id0 = Array.isArray(id) && id[0] instanceof PStr ? id[0].b : new Uint8Array(0);
    this.crypt = await Crypt.create(e, id0, v => this.get(v));
    this.encrypted = true;
    this.cache.clear(); this.objStms.clear();
  }
  // 名前つきのリンク先（目次のリンクなどで使われる）を探す。見つかればリンク先の配列か辞書
  async namedDest(key) {
    const k = key instanceof PName ? key.n : String.fromCharCode(...key.b);
    const cat = await this.get(this.trailer.get('Root'));
    const dests = await this.get(cat.get('Dests'));
    if (dests instanceof Map && dests.has(k)) return dests.get(k);
    const names = await this.get(cat.get('Names')), tree = names instanceof Map ? names.get('Dests') : null;
    const str = v => (v instanceof PStr ? String.fromCharCode(...v.b) : null);
    const find = async (node, depth) => {
      node = await this.get(node);
      if (!(node instanceof Map) || depth > 40) return null;
      const lim = await this.get(node.get('Limits'));
      if (Array.isArray(lim) && lim.length === 2 && str(lim[0]) != null && (k < str(lim[0]) || k > str(lim[1]))) return null;
      const arr = await this.get(node.get('Names'));
      if (Array.isArray(arr)) for (let i = 0; i + 1 < arr.length; i += 2) if (str(await this.get(arr[i])) === k) return arr[i + 1];
      const kids = await this.get(node.get('Kids'));
      if (Array.isArray(kids)) for (const kid of kids) { const r = await find(kid, depth + 1); if (r) return r; }
      return null;
    };
    return tree ? find(tree, 0) : null;
  }
  async catalogOk() {
    try {
      const cat = await this.get(this.trailer.get('Root'));
      return cat instanceof Map && (await this.get(cat.get('Pages'))) instanceof Map;
    } catch (e) { return false; }
  }
  async init() {
    if (indexOf(this.buf.subarray(0, 1024), '%PDF-') < 0) throw new PdfError('notpdf', 'PDF ではありません');
    let ok = false;
    try { await this.readXrefChain(); await this.setupCrypt(); ok = await this.catalogOk(); }
    catch (e) { if (e instanceof PdfError && (e.code === 'password' || e.code === 'unsupported')) throw e; }
    if (!ok) {
      await this.rebuild();
      if (!(await this.catalogOk())) throw new PdfError('broken', 'PDF の中身が読めません');
    }
    this.pageList = await this.readPages();
    if (!this.pageList.length) throw new PdfError('broken', 'ページが見つかりません');
    this.numPages = this.pageList.length;
  }
  // ページの一覧。親から受け継ぐ設定（Resources・MediaBox・CropBox・Rotate）も集める
  async readPages() {
    const cat = await this.get(this.trailer.get('Root')), out = [], seen = new Set();
    const walk = async (ref, inh, depth) => {
      if (depth > 100) return;
      if (ref instanceof PRef) { if (seen.has(ref.num)) return; seen.add(ref.num); }
      const node = await this.get(ref);
      if (!(node instanceof Map)) return;
      const inh2 = { ...inh };
      for (const k of ['Resources', 'MediaBox', 'CropBox', 'Rotate']) if (node.has(k)) inh2[k] = node.get(k);
      const kids = await this.get(node.get('Kids')), type = nameOf(node.get('Type'));
      if (type === 'Pages' || (type !== 'Page' && Array.isArray(kids))) { if (Array.isArray(kids)) for (const k of kids) await walk(k, inh2, depth + 1); }
      else out.push({ ref: ref instanceof PRef ? ref : null, dict: node, inh: inh2 });
    };
    await walk(cat.get('Pages'), {}, 0);
    return out;
  }
}

async function open(bytes) {
  const doc = new Doc(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
  await doc.init();
  return doc;
}

/* ---------- 書き出し ---------- */
function fmtNum(v) {
  if (!isFinite(v)) return '0';
  if (Number.isInteger(v)) return String(v);
  let s = String(+v.toPrecision(10));
  if (s.includes('e')) s = v.toFixed(10).replace(/0+$/, '').replace(/\.$/, '');
  return s;
}
function escName(n) {
  let s = '';
  for (let i = 0; i < n.length; i++) {
    const c = n.charCodeAt(i);
    s += c < 0x21 || c > 0x7e || c === 35 || DL[c] ? '#' + c.toString(16).padStart(2, '0') : n[i];
  }
  return s;
}
function serStr(v) {
  const b = v.b;
  if (v.hex) { let s = '<'; for (let i = 0; i < b.length; i++) s += b[i].toString(16).padStart(2, '0'); return s + '>'; }
  let s = '(';
  for (let i = 0; i < b.length; i++) {
    const c = b[i];
    s += c === 40 || c === 41 || c === 92 ? '\\' + String.fromCharCode(c) : c === 13 ? '\\r' : String.fromCharCode(c);
  }
  return s + ')';
}
function ser(v) {
  if (v === null || v === undefined) return 'null';
  if (v === true) return 'true';
  if (v === false) return 'false';
  if (typeof v === 'number') return fmtNum(v);
  if (v instanceof PName) return '/' + escName(v.n);
  if (v instanceof ORef) return v.id + ' 0 R';
  if (v instanceof PStr) return serStr(v);
  if (Array.isArray(v)) return '[' + v.map(ser).join(' ') + ']';
  if (v instanceof Map) { let s = '<<'; for (const [k, x] of v) s += '/' + escName(k) + ' ' + ser(x); return s + '>>'; }
  return 'null';
}
function textString(s) { // 日本語も入る文字列（UTF-16BE）
  const b = [0xFE, 0xFF];
  for (let i = 0; i < s.length; i++) { const c = s.charCodeAt(i); b.push(c >> 8, c & 255); }
  return new PStr(Uint8Array.from(b), true);
}
function pdfDateNow(d = new Date()) {
  const p = n => String(n).padStart(2, '0'), tz = -d.getTimezoneOffset(), a = Math.abs(tz);
  return `D:${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}${tz >= 0 ? '+' : '-'}${p(a / 60 | 0)}'${p(a % 60)}'`;
}

// pages: [{ doc, index, rotate }]  rotate は「さらに右へ回す角度」（90 の倍数）
// 返り値は Blob（application/pdf）
async function writePages(pages, { title, onProgress } = {}) {
  const parts = [], offsets = [];
  let pos = 0, nextId = 1;
  const raw = x => { const u = typeof x === 'string' ? enc(x) : x; parts.push(u); pos += u.length; };
  const alloc = () => nextId++;
  const writeObj = (id, v) => {
    offsets[id] = pos;
    if (v instanceof PStream) {
      v.dict.set('Length', v.data.length);
      raw(`${id} 0 obj\n${ser(v.dict)}\nstream\n`); raw(v.data); raw('\nendstream\nendobj\n');
    } else raw(`${id} 0 obj\n${ser(v)}\nendobj\n`);
  };
  raw('%PDF-1.7\n%\xE2\xE3\xCF\xD3\n');
  const catalogId = alloc(), pagesId = alloc(), infoId = alloc();

  // 元の PDF ごとに「元の番号 → 新しい番号」を覚える（同じフォントや画像は1回だけ入れる）
  const ctxs = new Map(), queue = [];
  const ctxOf = doc => {
    let c = ctxs.get(doc);
    if (!c) {
      c = { doc, map: new Map(), pageMap: new Map(), override: new Map(), pageNums: new Set(doc.pageList.filter(p => p.ref).map(p => p.ref.num)) };
      ctxs.set(doc, c);
    }
    return c;
  };
  const copy = (c, v) => {
    if (v instanceof PRef) {
      if (c.pageNums.has(v.num)) { const id = c.pageMap.get(v.num); return id ? new ORef(id) : null; } // 入れないページへのリンクは消す
      let id = c.map.get(v.num);
      if (id === undefined) { id = alloc(); c.map.set(v.num, id); queue.push([c, v.num, id]); }
      return new ORef(id);
    }
    if (Array.isArray(v)) return v.map(x => copy(c, x));
    if (v instanceof Map) { const m = new Map(); for (const [k, x] of v) m.set(k, copy(c, x)); return m; }
    return v;
  };
  const drain = async () => {
    while (queue.length) {
      const [c, num, id] = queue.shift();
      let v = c.override.has(num) ? c.override.get(num) : await c.doc.getObj(num);
      if (v instanceof Map) {
        const t = nameOf(v.get('Type'));
        if (t === 'Catalog' || t === 'Pages' || t === 'Page') v = null; // 文書全体を引き込まない
        else if (nameOf(v.get('Subtype')) === 'Widget' || t === 'Annot') { v = new Map(v); if (nameOf(v.get('Subtype')) === 'Widget') v.delete('Parent'); }
      }
      if (v instanceof PStream) {
        const d = new Map(v.dict); d.delete('Length');
        writeObj(id, new PStream(copy(c, d), v.data));
      } else writeObj(id, copy(c, v));
    }
  };

  // ページ内リンクの手入れ：入れないページへのリンクは消し、名前つきのリンク先は直接の行き先に書き換える
  const fixAnnot = async (c, a) => {
    const doc = c.doc, an = await doc.get(a);
    if (!(an instanceof Map) || nameOf(an.get('Subtype')) !== 'Link') return a;
    let dest = await doc.get(an.get('Dest'));
    if (dest == null) {
      const act = await doc.get(an.get('A'));
      if (!(act instanceof Map) || nameOf(act.get('S')) !== 'GoTo') return a;
      dest = await doc.get(act.get('D'));
    }
    let named = false;
    if (dest instanceof PStr || dest instanceof PName) { named = true; dest = await doc.get(await doc.namedDest(dest)); }
    if (dest instanceof Map) dest = await doc.get(dest.get('D'));
    if (!Array.isArray(dest) || !(dest[0] instanceof PRef) || !c.pageNums.has(dest[0].num)) return named ? undefined : a;
    if (!c.pageMap.has(dest[0].num)) return undefined;
    if (!named) return a;
    const m = new Map(an); m.delete('A'); m.set('Dest', dest);
    if (a instanceof PRef) { c.override.set(a.num, m); return a; }
    return m;
  };

  const ids = pages.map(() => alloc());
  pages.forEach((it, i) => {
    const c = ctxOf(it.doc), pg = it.doc.pageList[it.index];
    if (pg.ref && !c.pageMap.has(pg.ref.num)) c.pageMap.set(pg.ref.num, ids[i]);
  });
  for (let i = 0; i < pages.length; i++) {
    if (onProgress) onProgress(i, pages.length);
    const it = pages[i], c = ctxOf(it.doc), pg = it.doc.pageList[it.index], doc = it.doc;
    const d = new Map(pg.dict);
    for (const k of ['Parent', 'B', 'Thumb']) d.delete(k);
    d.set('Type', new PName('Page'));
    for (const k of ['Resources', 'MediaBox', 'CropBox']) if (!d.has(k) && pg.inh[k] !== undefined) d.set(k, pg.inh[k]);
    for (const k of ['MediaBox', 'CropBox']) if (d.has(k)) d.set(k, await doc.get(d.get(k)));
    if (!Array.isArray(d.get('MediaBox'))) d.set('MediaBox', [0, 0, 595.28, 841.89]);
    if (!d.has('Resources')) d.set('Resources', new Map());
    let rot = await doc.get(d.has('Rotate') ? d.get('Rotate') : pg.inh.Rotate);
    rot = typeof rot === 'number' ? Math.round(rot / 90) * 90 : 0;
    rot = (((rot + (it.rotate || 0)) % 360) + 360) % 360;
    if (rot) d.set('Rotate', rot); else d.delete('Rotate');
    const annots = await doc.get(d.get('Annots'));
    if (Array.isArray(annots)) {
      const keep = [];
      for (const a of annots) { const v = await fixAnnot(c, a); if (v !== undefined) keep.push(v); }
      if (keep.length) d.set('Annots', keep); else d.delete('Annots');
    }
    const out = copy(c, d);
    out.set('Parent', new ORef(pagesId));
    writeObj(ids[i], out);
    await drain();
  }
  writeObj(pagesId, new Map([['Type', new PName('Pages')], ['Kids', ids.map(id => new ORef(id))], ['Count', ids.length]]));
  writeObj(catalogId, new Map([['Type', new PName('Catalog')], ['Pages', new ORef(pagesId)]]));
  const info = new Map([['Producer', new PStr(enc('file-tools'), false)], ['CreationDate', new PStr(enc(pdfDateNow()), false)]]);
  if (title) info.set('Title', textString(title));
  writeObj(infoId, info);
  const idBytes = new Uint8Array(16);
  if (typeof crypto !== 'undefined' && crypto.getRandomValues) crypto.getRandomValues(idBytes);
  const idStr = ser(new PStr(idBytes, true));
  const xref = pos, n = nextId;
  let x = `xref\n0 ${n}\n0000000000 65535 f \n`;
  for (let id = 1; id < n; id++) x += String(offsets[id] || 0).padStart(10, '0') + (offsets[id] ? ' 00000 n \n' : ' 00000 f \n');
  x += `trailer\n<< /Size ${n} /Root ${catalogId} 0 R /Info ${infoId} 0 R /ID [${idStr} ${idStr}] >>\nstartxref\n${xref}\n%%EOF\n`;
  raw(x);
  return new Blob(parts, { type: 'application/pdf' });
}

return { open, writePages, PdfError, _test: { md5, rc4, aesCbcDecrypt, aesCbcEncrypt, Parser } };
})();
if (typeof module !== 'undefined') module.exports = PdfDoc;
