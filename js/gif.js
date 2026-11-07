'use strict';
// =========================================================
// GIF アニメを作る（自作。画面に依存しないので Node でもテストできる）
//   GIF.palette(samples, n)  … 色を n 色に減らした「パレット」を作る（全部のコマで同じものを使う）
//   new GIF.Encoder(w, h, { palette, loop }) → addFrame(rgba, delay) をコマの数だけ → finish() で Uint8Array
//   ・色が足りない所は、決まった模様で色を混ぜる（ディザ）。コマごとに模様が変わらないので、ちらつかず、軽い
//   ・前のコマと同じ所（ほぼ同じ所）は「透明＝前のまま」にして、変わった所の四角だけを書く（軽くなる）
//   ・圧縮は GIF の決まりの LZW
// =========================================================
const GIF = (() => {
  /* ---------- 色を減らす（パレット） ---------- */
  // samples: [{ data: RGBA, width, height }]（いくつかのコマの見本）。間引いて色を数え、
  // 「色の散らばりがいちばん大きい箱を2つに分ける」をくり返して n 色にし、少し整える（k-means）
  function palette(samples, n, maxPixels = 400000) {
    n = Math.max(2, Math.min(256, n | 0));
    const NB = 32768, cnt = new Float64Array(NB), sr = new Float64Array(NB), sg = new Float64Array(NB), sb = new Float64Array(NB);
    let total = 0;
    for (const s of samples) total += s.width * s.height;
    const step = Math.max(1, Math.ceil(Math.sqrt(total / maxPixels)));
    for (const s of samples) {
      const d = s.data, W = s.width, H = s.height;
      for (let y = 0; y < H; y += step) {
        for (let x = (y * 7) % step; x < W; x += step) {
          const j = (y * W + x) * 4;
          if (d[j + 3] < 128) continue;
          const r = d[j], g = d[j + 1], b = d[j + 2], k = ((r >> 3) << 10) | ((g >> 3) << 5) | (b >> 3);
          cnt[k]++; sr[k] += r; sg[k] += g; sb[k] += b;
        }
      }
    }
    // 色のある箱（ビン）の一覧と、その平均の色
    const bins = [];
    for (let k = 0; k < NB; k++) if (cnt[k]) bins.push(k);
    const m = bins.length, w = new Float64Array(m), mr = new Float64Array(m), mg = new Float64Array(m), mb = new Float64Array(m);
    for (let i = 0; i < m; i++) { const k = bins[i]; w[i] = cnt[k]; mr[i] = sr[k] / cnt[k]; mg[i] = sg[k] / cnt[k]; mb[i] = sb[k] / cnt[k]; }
    if (!m) return { colors: new Uint8Array(n * 3), count: 1 };
    const stats = idx => {
      let W = 0, R = 0, G = 0, B = 0, Q = 0, RR = 0, GG = 0, BB = 0;
      for (const i of idx) { const a = w[i]; W += a; R += a * mr[i]; G += a * mg[i]; B += a * mb[i]; RR += a * mr[i] * mr[i]; GG += a * mg[i] * mg[i]; BB += a * mb[i] * mb[i]; }
      Q = RR + GG + BB;
      return { idx, W, R, G, B, RR, GG, BB, sse: Q - (R * R + G * G + B * B) / W };
    };
    const boxes = [stats(Array.from({ length: m }, (_, i) => i))];
    while (boxes.length < n) {
      let bi = -1, best = 1e-9;
      for (let i = 0; i < boxes.length; i++) if (boxes[i].idx.length > 1 && boxes[i].sse > best) { best = boxes[i].sse; bi = i; }
      if (bi < 0) break;
      const bx = boxes[bi];
      // いちばん散らばっている向き（R・G・B）で、2つの散らばりの合計がいちばん小さくなる所で分ける
      const vr = bx.RR - bx.R * bx.R / bx.W, vg = bx.GG - bx.G * bx.G / bx.W, vb = bx.BB - bx.B * bx.B / bx.W;
      const ch = vr >= vg && vr >= vb ? mr : vg >= vb ? mg : mb;
      const idx = bx.idx.slice().sort((a, b) => ch[a] - ch[b]);
      let W = 0, R = 0, G = 0, B = 0, Q = 0, cut = 1, bestE = Infinity;
      const QT = bx.RR + bx.GG + bx.BB;
      for (let k = 0; k < idx.length - 1; k++) {
        const i = idx[k], a = w[i];
        W += a; R += a * mr[i]; G += a * mg[i]; B += a * mb[i]; Q += a * (mr[i] * mr[i] + mg[i] * mg[i] + mb[i] * mb[i]);
        const W2 = bx.W - W, R2 = bx.R - R, G2 = bx.G - G, B2 = bx.B - B;
        const e = Q - (R * R + G * G + B * B) / W + (QT - Q) - (R2 * R2 + G2 * G2 + B2 * B2) / W2;
        if (e < bestE) { bestE = e; cut = k + 1; }
      }
      boxes.splice(bi, 1, stats(idx.slice(0, cut)), stats(idx.slice(cut)));
    }
    const P = boxes.map(b => [b.R / b.W, b.G / b.W, b.B / b.W]);
    // 少し整える：色ごとに、いちばん近い色の平均に動かす（2回）
    const near = new Int32Array(m);
    for (let it = 0; it < 2; it++) {
      const acc = P.map(() => [0, 0, 0, 0]);
      for (let i = 0; i < m; i++) {
        let bj = 0, bd = Infinity;
        for (let j = 0; j < P.length; j++) {
          const dr = mr[i] - P[j][0], dg = mg[i] - P[j][1], db = mb[i] - P[j][2], dd = dr * dr + dg * dg + db * db;
          if (dd < bd) { bd = dd; bj = j; }
        }
        near[i] = bj;
        const a = acc[bj]; a[0] += w[i] * mr[i]; a[1] += w[i] * mg[i]; a[2] += w[i] * mb[i]; a[3] += w[i];
      }
      for (let j = 0; j < P.length; j++) if (acc[j][3]) P[j] = [acc[j][0] / acc[j][3], acc[j][1] / acc[j][3], acc[j][2] / acc[j][3]];
    }
    const colors = new Uint8Array(n * 3);
    P.forEach((p, j) => { colors[j * 3] = Math.round(p[0]); colors[j * 3 + 1] = Math.round(p[1]); colors[j * 3 + 2] = Math.round(p[2]); });
    return { colors, count: P.length };
  }

  // いちばん近いパレットの色を探す（一度調べた色は覚えておく。R・G・B を 6 ビットずつにまとめて）
  class Mapper {
    constructor(colors, count) { this.c = colors; this.n = count; this.lut = new Int16Array(1 << 18).fill(-1); }
    // その色と、いちばん近いパレットの色との距離（の2乗）
    err(r, g, b, v) { const c = this.c, dr = r - c[v * 3], dg = g - c[v * 3 + 1], db = b - c[v * 3 + 2]; return dr * dr + dg * dg + db * db; }
    index(r, g, b) {
      const key = ((r >> 2) << 12) | ((g >> 2) << 6) | (b >> 2);
      let v = this.lut[key];
      if (v >= 0) return v;
      const R = (r & 0xFC) + 2, G = (g & 0xFC) + 2, B = (b & 0xFC) + 2, c = this.c;
      let bd = Infinity;
      for (let j = 0; j < this.n; j++) {
        const dr = R - c[j * 3], dg = G - c[j * 3 + 1], db = B - c[j * 3 + 2], dd = 2 * dr * dr + 4 * dg * dg + db * db;
        if (dd < bd) { bd = dd; v = j; }
      }
      this.lut[key] = v;
      return v;
    }
  }
  // パレットの色どうしの近さ（ディザの強さに使う）。それぞれの色から、いちばん近いほかの色までの距離の中央値
  function spacing(colors, count) {
    const ds = [];
    for (let i = 0; i < count; i++) {
      let bd = Infinity;
      for (let j = 0; j < count; j++) {
        if (i === j) continue;
        const dr = colors[i * 3] - colors[j * 3], dg = colors[i * 3 + 1] - colors[j * 3 + 1], db = colors[i * 3 + 2] - colors[j * 3 + 2];
        bd = Math.min(bd, dr * dr + dg * dg + db * db);
      }
      if (isFinite(bd)) ds.push(Math.sqrt(bd));
    }
    ds.sort((a, b) => a - b);
    return ds.length ? ds[ds.length >> 1] : 32;
  }
  // 8×8 の決まった模様（Bayer）。0〜63
  const BAYER = (() => {
    const m = [[0]];
    let cur = m;
    for (let s = 1; s < 8; s *= 2) {
      const n = [];
      for (let y = 0; y < s * 2; y++) { n.push([]); for (let x = 0; x < s * 2; x++) n[y].push(4 * cur[y % s][x % s] + [0, 2, 3, 1][(y >= s) * 2 + (x >= s)]); }
      cur = n;
    }
    return Int8Array.from(cur.flat());
  })();

  /* ---------- 書き出し ---------- */
  class Bytes {
    constructor() { this.parts = []; this.buf = new Uint8Array(1 << 16); this.len = 0; this.total = 0; }
    byte(b) { if (this.len === this.buf.length) this.flush(); this.buf[this.len++] = b; }
    u16(v) { this.byte(v & 255); this.byte((v >> 8) & 255); }
    bytes(a) {
      if (a.length > 512) { this.flush(); this.parts.push(a.slice()); this.total += a.length; return; }
      for (let i = 0; i < a.length; i++) this.byte(a[i]);
    }
    str(s) { for (let i = 0; i < s.length; i++) this.byte(s.charCodeAt(i)); }
    flush() { if (this.len) { this.parts.push(this.buf.slice(0, this.len)); this.total += this.len; this.len = 0; } }
    result() {
      this.flush();
      const out = new Uint8Array(this.total);
      let o = 0; for (const p of this.parts) { out.set(p, o); o += p.length; }
      return out;
    }
  }

  // LZW 圧縮（GIF の決まり。昔からある compress のやり方を、そのまま写したもの）
  function lzw(px, minCode, out) {
    const HSIZE = 5003, HSHIFT = 4, MAXBITS = 12, MAXMAX = 1 << MAXBITS;
    const htab = new Int32Array(HSIZE).fill(-1), codetab = new Int32Array(HSIZE);
    const initBits = minCode + 1, clearCode = 1 << minCode, eofCode = clearCode + 1;
    let nBits = initBits, maxcode = (1 << nBits) - 1, freeEnt = clearCode + 2, clearFlg = false;
    let cur = 0, curBits = 0;
    const blk = new Uint8Array(255); let blen = 0;
    const put = b => { blk[blen++] = b; if (blen === 255) { out.byte(255); out.bytes(blk); blen = 0; } };
    const output = code => {
      cur |= code << curBits; curBits += nBits;
      while (curBits >= 8) { put(cur & 255); cur >>>= 8; curBits -= 8; }
      if (freeEnt > maxcode || clearFlg) {
        if (clearFlg) { nBits = initBits; maxcode = (1 << nBits) - 1; clearFlg = false; }
        else { nBits++; maxcode = nBits === MAXBITS ? MAXMAX : (1 << nBits) - 1; }
      }
      if (code === eofCode) {
        while (curBits > 0) { put(cur & 255); cur >>>= 8; curBits -= 8; }
        if (blen) { out.byte(blen); out.bytes(blk.subarray(0, blen)); blen = 0; }
      }
    };
    out.byte(minCode);
    output(clearCode);
    let ent = px[0];
    outer: for (let i = 1; i < px.length; i++) {
      const c = px[i], fcode = (c << MAXBITS) + ent;
      let h = (c << HSHIFT) ^ ent;
      if (htab[h] === fcode) { ent = codetab[h]; continue; }
      if (htab[h] >= 0) {
        const disp = h === 0 ? 1 : HSIZE - h;
        do {
          if ((h -= disp) < 0) h += HSIZE;
          if (htab[h] === fcode) { ent = codetab[h]; continue outer; }
        } while (htab[h] >= 0);
      }
      output(ent);
      ent = c;
      if (freeEnt < MAXMAX) { codetab[h] = freeEnt++; htab[h] = fcode; }
      else { htab.fill(-1); freeEnt = clearCode + 2; clearFlg = true; output(clearCode); }
    }
    output(ent);
    output(eofCode);
    out.byte(0); // ブロックの終わり
  }

  // w×h の GIF を作る。palette: GIF.palette の結果。loop: 0 = ずっとくり返す、null = 1回だけ
  class Encoder {
    constructor(w, h, { palette: pal, loop = 0 }) {
      this.w = w; this.h = h;
      this.n = Math.min(255, pal.count); // 最後の1つは「透明（前のコマのまま）」に使う
      this.T = this.n;
      let bits = 1; while ((1 << bits) < this.n + 1) bits++;
      this.bits = bits;
      this.pal = pal.colors;
      this.map = new Mapper(pal.colors, this.n);
      // 色を混ぜる強さ：パレットの色の間かくくらい（色の少ない絵では大きすぎるので上限あり）
      this.spread = Math.min(40, spacing(pal.colors, this.n));
      this.prev = new Uint8Array(w * h); this.has = false;
      this.idx = new Uint8Array(w * h);
      this.pending = null; this.frames = 0; this.time = 0; // time: ここまでの時間（1/100 秒、端数つき）
      const o = this.out = new Bytes();
      o.str('GIF89a'); o.u16(w); o.u16(h);
      o.byte(0x80 | (7 << 4) | (bits - 1)); o.byte(0); o.byte(0);
      for (let i = 0; i < (1 << bits); i++) for (let c = 0; c < 3; c++) o.byte(i < this.n ? this.pal[i * 3 + c] : 0);
      if (loop != null) { o.byte(0x21); o.byte(0xFF); o.byte(11); o.str('NETSCAPE2.0'); o.byte(3); o.byte(1); o.u16(loop); o.byte(0); }
    }
    // rgba: w×h の RGBA。delay: 表示する時間（1/100 秒。端数があってもよい。全体の時間がずれないように整数にする）。
    // dither: 色を混ぜる強さ（0〜1）。keep: 前のコマとの色の差がこれ以下なら、前のまま（0〜255）
    addFrame(rgba, delay, { dither = 1, keep = 0 } = {}) {
      const { w, h, prev, idx, T, map } = this, pal = this.pal, keep2 = keep * keep, has = this.has;
      const amp = dither > 0 ? this.spread * dither : 0, near2 = (this.spread * 0.25) ** 2;
      let x0 = w, y0 = h, x1 = -1, y1 = -1;
      for (let y = 0; y < h; y++) {
        const by = (y & 7) * 8;
        for (let x = 0, i = y * w; x < w; x++, i++) {
          const j = i * 4;
          let r = rgba[j], g = rgba[j + 1], b = rgba[j + 2];
          if (has && keep2) {
            const p = prev[i] * 3, dr = r - pal[p], dg = g - pal[p + 1], db = b - pal[p + 2];
            if (dr * dr + dg * dg + db * db <= keep2) { idx[i] = T; continue; }
          }
          let c = map.index(r, g, b);
          // パレットにほぼ同じ色があれば、そのまま（ベタ塗りの所に点々が出ないように）。なければ模様で色を混ぜる
          if (amp && map.err(r, g, b, c) > near2) {
            const o = ((BAYER[by + (x & 7)] + 0.5) / 64 - 0.5) * amp;
            r = r + o; g = g + o; b = b + o;
            r = r < 0 ? 0 : r > 255 ? 255 : r | 0; g = g < 0 ? 0 : g > 255 ? 255 : g | 0; b = b < 0 ? 0 : b > 255 ? 255 : b | 0;
            c = map.index(r, g, b);
          }
          if (has && c === prev[i]) { idx[i] = T; continue; }
          idx[i] = c; prev[i] = c;
          if (x < x0) x0 = x; if (x > x1) x1 = x; if (y < y0) y0 = y; if (y > y1) y1 = y;
        }
      }
      const t0 = Math.round(this.time); this.time += delay;
      delay = Math.max(2, Math.round(this.time) - t0);
      if (x1 < 0) { // 前のコマとまったく同じ：前のコマを長く見せる
        if (this.pending) { this.pending.delay += delay; return; }
        x0 = y0 = 0; x1 = y1 = 0;
      }
      // 変わった所の四角だけを書く
      const fw = x1 - x0 + 1, fh = y1 - y0 + 1, sub = new Uint8Array(fw * fh);
      for (let y = 0; y < fh; y++) sub.set(idx.subarray((y0 + y) * w + x0, (y0 + y) * w + x0 + fw), y * fw);
      this.flushPending();
      const data = new Bytes();
      lzw(sub, Math.max(2, this.bits), data);
      this.pending = { x0, y0, fw, fh, delay, transparent: has, data: data.result() };
      this.has = true; this.frames++;
    }
    flushPending() {
      const p = this.pending, o = this.out;
      if (!p) return;
      o.byte(0x21); o.byte(0xF9); o.byte(4);
      o.byte((1 << 2) | (p.transparent ? 1 : 0)); // 次のコマを描くとき、このコマは残す
      o.u16(Math.min(65535, p.delay)); o.byte(p.transparent ? this.T : 0); o.byte(0);
      o.byte(0x2C); o.u16(p.x0); o.u16(p.y0); o.u16(p.fw); o.u16(p.fh); o.byte(0);
      o.bytes(p.data);
      this.pending = null;
    }
    // できあがりの大きさ（まだ書いていないコマを除く、だいたい）
    get size() { return this.out.total + this.out.len + (this.pending ? this.pending.data.length : 0); }
    finish() {
      this.flushPending();
      this.out.byte(0x3B);
      return this.out.result();
    }
  }

  return { palette, Encoder };
})();
if (typeof module === 'object' && module.exports) module.exports = GIF; // テスト用
