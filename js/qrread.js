'use strict';
// =========================================================
// QRコードを読み取るしくみ（規格 JIS X 0510 / ISO 18004 に沿って自前で計算）
//
//   const found = QRRead.scan(ctx.getImageData(0, 0, w, h));
//   found[i].text     … 読み取った文字
//   found[i].corners  … 画像の中の QR の四隅 [{x,y} ×4]（左上・右上・右下・左下）
//   found[i].version  … 大きさの段階（1〜40）、found[i].ecl … 誤り訂正（L/M/Q/H）
//
// 流れ：白黒にする → 3つの角の目印（ファインダー）を探す → 3つ組ごとにマス目を読み取る
//       → 形式情報（誤り訂正とマスク）→ マスクを外す → 誤り訂正（リード・ソロモン）→ 文字に戻す
// 1枚の画像にいくつ QR があっても全部読みます。白黒反転・鏡写しの QR も読めます。
// 規格の表は js/qrcode.js（作るしくみ）と共通です（先に読み込んでおく）。
// =========================================================
const QRRead = (() => {
  const S = (typeof QR !== 'undefined' ? QR : require('./qrcode.js')).spec;

  /* ---------- 白黒にする ---------- */
  // 明るさ（0〜255）。透明な所は白とみなす
  function toGray(img) {
    const { data, width: w, height: h } = img, n = w * h, g = new Uint8Array(n);
    for (let i = 0, j = 0; i < n; i++, j += 4) {
      const l = (data[j] * 77 + data[j + 1] * 150 + data[j + 2] * 29) >> 8, a = data[j + 3];
      g[i] = a === 255 ? l : (l * a + 255 * (255 - a)) / 255;
    }
    return g;
  }

  // 場所ごとのしきい値（8×8 画素のかたまりごとに明るさを測り、まわり 5×5 かたまりの平均を使う）。
  // 影やむらがある写真でも白黒に分けられる。1 = 黒
  function binarizeLocal(g, w, h) {
    if (w < 40 || h < 40) return binarizeGlobal(g, w, h);
    const sw = Math.ceil(w / 8), sh = Math.ceil(h / 8), bp = new Int32Array(sw * sh);
    for (let by = 0; by < sh; by++) {
      const y0 = Math.min(by * 8, h - 8);
      for (let bx = 0; bx < sw; bx++) {
        const x0 = Math.min(bx * 8, w - 8);
        let sum = 0, min = 255, max = 0;
        for (let yy = 0, o = y0 * w + x0; yy < 8; yy++, o += w) {
          for (let xx = 0; xx < 8; xx++) { const p = g[o + xx]; sum += p; if (p < min) min = p; if (p > max) max = p; }
        }
        let avg;
        if (max - min > 24) avg = sum >> 6;
        else {
          // 明るさがそろったかたまり：ふつうは白（背景）とみなす。となりが黒っぽい所の続きなら黒
          avg = min >> 1;
          if (by > 0 && bx > 0) {
            const nb = (bp[(by - 1) * sw + bx] + 2 * bp[by * sw + bx - 1] + bp[(by - 1) * sw + bx - 1]) >> 2;
            if (min < nb) avg = nb;
          }
        }
        bp[by * sw + bx] = avg;
      }
    }
    const out = new Uint8Array(w * h);
    for (let by = 0; by < sh; by++) {
      const y0 = Math.min(by * 8, h - 8), cy = Math.min(Math.max(by, 2), sh - 3);
      for (let bx = 0; bx < sw; bx++) {
        const x0 = Math.min(bx * 8, w - 8), cx = Math.min(Math.max(bx, 2), sw - 3);
        let sum = 0;
        for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) sum += bp[(cy + dy) * sw + cx + dx];
        const t = sum / 25;
        for (let yy = 0, o = y0 * w + x0; yy < 8; yy++, o += w) {
          for (let xx = 0; xx < 8; xx++) out[o + xx] = g[o + xx] <= t ? 1 : 0;
        }
      }
    }
    return out;
  }

  // 画像全体で1つのしきい値（大津の方法）。くっきりした画像向け
  function binarizeGlobal(g, w, h) {
    const n = w * h, hist = new Float64Array(256);
    for (let i = 0; i < n; i++) hist[g[i]]++;
    let all = 0; for (let i = 0; i < 256; i++) all += i * hist[i];
    let wb = 0, sb = 0, best = -1, t = 127;
    for (let i = 0; i < 256; i++) {
      wb += hist[i]; if (!wb) continue;
      const wf = n - wb; if (!wf) break;
      sb += i * hist[i];
      const v = wb * wf * (sb / wb - (all - sb) / wf) ** 2;
      if (v > best) { best = v; t = i; }
    }
    const out = new Uint8Array(n);
    for (let i = 0; i < n; i++) out[i] = g[i] <= t ? 1 : 0;
    return out;
  }

  // 縦横半分に縮める（ぼやけた大きな写真用）
  function half(g, w, h) {
    const W = w >> 1, H = h >> 1, o = new Uint8Array(W * H);
    for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
      const i = 2 * y * w + 2 * x;
      o[y * W + x] = (g[i] + g[i + 1] + g[i + w] + g[i + w + 1] + 2) >> 2;
    }
    return { g: o, w: W, h: H };
  }

  /* ---------- 角の目印（ファインダー：黒白黒白黒 = 1:1:3:1:1）を探す ---------- */
  function ratio5(a, b, c, d, e) {
    const t = a + b + c + d + e;
    if (t < 7) return false;
    const m = t / 7, v = m / 2 + 0.5;
    return Math.abs(a - m) < v && Math.abs(b - m) < v && Math.abs(c - 3 * m) < 3 * m / 2 + 0.5 && Math.abs(d - m) < v && Math.abs(e - m) < v;
  }

  // (x, y) から (dx, dy) の向きと逆向きに「黒・白・黒・白・黒」をたどる。
  // 中心（(x, y) の画素の左上を 0 とした連続座標でのずれ）と全長を返す
  function cross(bin, w, h, x, y, dx, dy, maxRun) {
    const at = i => {
      const px = x + dx * i, py = y + dy * i;
      return px >= 0 && py >= 0 && px < w && py < h ? bin[py * w + px] : -1;
    };
    if (at(0) !== 1) return null;
    let i = 0, b2 = 0, c1 = 0, c0 = 0;
    while (at(i) === 1) { if (++b2 > maxRun) return null; i--; }
    while (at(i) === 0) { if (++c1 > maxRun) return null; i--; }
    if (at(i) !== 1) return null;
    while (at(i) === 1 && c0 <= maxRun) { c0++; i--; }
    let j = 1, f2 = 0, c3 = 0, c4 = 0;
    while (at(j) === 1) { if (++f2 + b2 > maxRun) return null; j++; }
    while (at(j) === 0) { if (++c3 > maxRun) return null; j++; }
    if (at(j) !== 1) return null;
    while (at(j) === 1 && c4 <= maxRun) { c4++; j++; }
    if (!ratio5(c0, c1, b2 + f2, c3, c4)) return null;
    return { center: (f2 - b2 + 2) / 2, total: c0 + c1 + b2 + f2 + c3 + c4 };
  }

  // 同じ目印を何行も見つけるので、近くの候補にまとめる（ざらついた写真では候補がとても多いので、場所ごとの箱で探す）
  function candidateList() {
    const cands = [], boxes = new Map(), CELL = 32, key = (cx, cy) => cx * 100003 + cy;
    cands.add = (x, y, size) => {
      const R = Math.max(size, 2) * 1.5;
      for (let cy = Math.floor((y - R) / CELL); cy <= Math.floor((y + R) / CELL); cy++) {
        for (let cx = Math.floor((x - R) / CELL); cx <= Math.floor((x + R) / CELL); cx++) {
          for (const c of boxes.get(key(cx, cy)) || []) {
            const r = Math.max(size, c.size), d = Math.abs(size - c.size);
            if (Math.abs(y - c.y) > r || Math.abs(x - c.x) > r || (d > 1 && d > c.size)) continue;
            const n = c.count;
            c.x = (c.x * n + x) / (n + 1); c.y = (c.y * n + y) / (n + 1); c.size = (c.size * n + size) / (n + 1); c.count++;
            return;
          }
        }
      }
      const c = { x, y, size, count: 1 }, k = key(Math.floor(x / CELL), Math.floor(y / CELL));
      cands.push(c);
      if (!boxes.has(k)) boxes.set(k, []);
      boxes.get(k).push(c);
    };
    return cands;
  }

  function findFinders(bin, w, h) {
    const cands = candidateList(), st = new Int32Array(w + 1), len = new Int32Array(w + 1);
    for (let y = 0; y < h; y++) {
      const row = y * w;
      let n = 0, s = 0;
      for (let x = 1; x <= w; x++) if (x === w || bin[row + x] !== bin[row + s]) { st[n] = s; len[n++] = x - s; s = x; }
      for (let k = bin[row] ? 0 : 1; k + 4 < n; k += 2) {
        if (!ratio5(len[k], len[k + 1], len[k + 2], len[k + 3], len[k + 4])) continue;
        const total = len[k] + len[k + 1] + len[k + 2] + len[k + 3] + len[k + 4];
        const ix = Math.floor(st[k + 2] + len[k + 2] / 2);
        const v = cross(bin, w, h, ix, y, 0, 1, total);
        if (!v || Math.abs(v.total - total) * 5 >= total * 2) continue;
        const cy = y + v.center;
        const hz = cross(bin, w, h, ix, Math.floor(cy), 1, 0, total);
        if (!hz || Math.abs(hz.total - total) * 5 >= total * 2) continue;
        cands.add(ix + hz.center, cy, (v.total + hz.total) / 14);
      }
    }
    // 何行にもわたって見つかったものだけ（たまたまの模様を除く）。足りなければ1行だけのものも使う
    let good = cands.filter(c => c.count >= 2);
    if (good.length < 3) good = cands;
    return good.sort((a, b) => b.count - a.count).slice(0, 40);
  }

  /* ---------- 3つ組（左上・右上・左下）を作る ---------- */
  function makeTriples(c) {
    const out = [], d2 = (p, q) => (p.x - q.x) ** 2 + (p.y - q.y) ** 2;
    for (let i = 0; i < c.length; i++) for (let j = i + 1; j < c.length; j++) for (let k = j + 1; k < c.length; k++) {
      const A = c[i], B = c[j], C = c[k];
      const smin = Math.min(A.size, B.size, C.size), smax = Math.max(A.size, B.size, C.size);
      if (smax / smin > 1.7) continue;
      const ab = d2(A, B), bc = d2(B, C), ac = d2(A, C);
      // 直角の角（＝左上）は、いちばん長い辺の向かい
      let tl, p, q;
      if (bc >= ab && bc >= ac) { tl = A; p = B; q = C; }
      else if (ac >= ab) { tl = B; p = A; q = C; }
      else { tl = C; p = A; q = B; }
      const l1 = Math.sqrt(d2(tl, p)), l2 = Math.sqrt(d2(tl, q));
      if (Math.min(l1, l2) / Math.max(l1, l2) < 0.6) continue;
      const cos = ((p.x - tl.x) * (q.x - tl.x) + (p.y - tl.y) * (q.y - tl.y)) / (l1 * l2);
      if (Math.abs(cos) > 0.5) continue;
      const m = (A.size + B.size + C.size) / 3, dim = (l1 + l2) / 2 / m + 7;
      if (dim < 17 || dim > 190) continue;
      // 右上と左下の区別：画面の座標（下が +y）で外積が正なら p が右上
      const crs = (p.x - tl.x) * (q.y - tl.y) - (p.y - tl.y) * (q.x - tl.x);
      const [tr, bl] = crs > 0 ? [p, q] : [q, p];
      const score = Math.abs(cos) + Math.abs(Math.log(l1 / l2)) + Math.log(smax / smin);
      out.push({ tl, tr, bl, ids: [i, j, k], score });
    }
    return out.sort((a, b) => a.score - b.score);
  }

  /* ---------- 画像の座標とマス目の対応（射影変換） ---------- */
  function solve8(A, B) {
    const n = 8;
    for (let c = 0; c < n; c++) {
      let piv = c;
      for (let r = c + 1; r < n; r++) if (Math.abs(A[r][c]) > Math.abs(A[piv][c])) piv = r;
      if (Math.abs(A[piv][c]) < 1e-12) return null;
      [A[c], A[piv]] = [A[piv], A[c]]; [B[c], B[piv]] = [B[piv], B[c]];
      for (let r = 0; r < n; r++) {
        if (r === c) continue;
        const f = A[r][c] / A[c][c];
        if (!f) continue;
        for (let k = c; k < n; k++) A[r][k] -= f * A[c][k];
        B[r] -= f * B[c];
      }
    }
    return B.map((b, i) => b / A[i][i]);
  }
  // src: マス目の座標 [u, v] ×4、dst: 画像の座標 {x, y} ×4 → (u, v) から画像の座標を返す関数
  function homography(src, dst) {
    const A = [], B = [];
    for (let i = 0; i < 4; i++) {
      const [u, v] = src[i], { x, y } = dst[i];
      A.push([u, v, 1, 0, 0, 0, -u * x, -v * x]); B.push(x);
      A.push([0, 0, 0, u, v, 1, -u * y, -v * y]); B.push(y);
    }
    const s = solve8(A, B);
    if (!s) return null;
    const [a, b, c, d, e, f, g, k] = s;
    return (u, v) => { const z = g * u + k * v + 1; return { x: (a * u + b * v + c) / z, y: (d * u + e * v + f) / z }; };
  }
  // マス (mx, my) の中心の色（1 = 黒）。画像の外は白
  const sampler = (bin, w, h, H) => (mx, my) => {
    const p = H(mx + 0.5, my + 0.5), px = Math.floor(p.x), py = Math.floor(p.y);
    return px >= 0 && py >= 0 && px < w && py < h ? bin[py * w + px] : 0;
  };

  /* ---------- 形式情報・型番情報（誤りがあっても近いものを選ぶ） ---------- */
  const popcount = v => { let c = 0; while (v) { v &= v - 1; c++; } return c; };
  const FORMAT_CODES = Array.from({ length: 32 }, (_, d) => {
    let rem = d;
    for (let i = 0; i < 10; i++) rem = (rem << 1) ^ ((rem >>> 9) * 0x537);
    return ((d << 10) | rem) ^ 0x5412;
  });
  const VERSION_CODES = Array.from({ length: 41 }, (_, v) => {
    let rem = v;
    for (let i = 0; i < 12; i++) rem = (rem << 1) ^ ((rem >>> 11) * 0x1F25);
    return (v << 12) | rem;
  });
  const ECL_OF = ['M', 'L', 'H', 'Q']; // 形式情報の2ビット → 誤り訂正

  function readFormat(get, dim) {
    let a = 0, b = 0;
    for (let i = 0; i <= 5; i++) a |= get(8, i) << i;
    a |= get(8, 7) << 6; a |= get(8, 8) << 7; a |= get(7, 8) << 8;
    for (let i = 9; i < 15; i++) a |= get(14 - i, 8) << i;
    for (let i = 0; i < 8; i++) b |= get(dim - 1 - i, 8) << i;
    for (let i = 8; i < 15; i++) b |= get(8, dim - 15 + i) << i;
    return bestFormat(a, b);
  }
  function bestFormat(...codes) {
    let best = -1, bestD = 99;
    for (let d = 0; d < 32; d++) for (const c of codes) {
      const e = popcount(FORMAT_CODES[d] ^ c);
      if (e < bestD) { bestD = e; best = d; }
    }
    return bestD <= 3 ? { ecl: ECL_OF[best >> 3], mask: best & 7 } : null;
  }
  function versionFromBits(bits) {
    let best = 0, bestD = 99;
    for (let v = 7; v <= 40; v++) { const d = popcount(VERSION_CODES[v] ^ bits); if (d < bestD) { bestD = d; best = v; } }
    return bestD <= 3 ? best : 0;
  }
  // 目印の中心から向き (ux, uy) とその逆に、「黒・白・黒」の外の端までの長さ（= 7 マス分、画素）を測る。測れなければ 0
  function chord(bin, w, h, c, ux, uy, maxLen) {
    const at = t => {
      const px = Math.floor(c.x + ux * t), py = Math.floor(c.y + uy * t);
      return px >= 0 && py >= 0 && px < w && py < h ? bin[py * w + px] : 0;
    };
    const side = sgn => {
      const r = [0, 0, 0];
      let st = 0;
      for (let t = 0; t <= maxLen; t += 0.5) {
        const v = at(sgn * t);
        if (st === 0) { if (v) r[0] += 0.5; else { st = 1; r[1] += 0.5; } }
        else if (st === 1) { if (!v) r[1] += 0.5; else { st = 2; r[2] += 0.5; } }
        else if (v) r[2] += 0.5;
        else return r;
      }
      return null;
    };
    const a = side(1), b = side(-1);
    if (!a || !b) return 0;
    const total = a[0] + a[1] + a[2] + b[0] + b[1] + b[2], m = total / 7;
    const ok = (x, k) => Math.abs(x - k * m) < k * m * 0.6 + 1;
    return ok(a[0] + b[0], 3) && ok(a[1], 1) && ok(b[1], 1) && ok(a[2], 1) && ok(b[2], 1) ? total : 0;
  }
  // 形式情報：左上のまわりに1つ、右上と左下に分けてもう1つ
  function formatNear(at, tl, tr, bl) {
    let a = 0, b = 0;
    for (let i = 0; i <= 5; i++) a |= at(tl, 5, i - 3) << i;
    a |= at(tl, 5, 4) << 6; a |= at(tl, 5, 5) << 7; a |= at(tl, 4, 5) << 8;
    for (let i = 9; i < 15; i++) a |= at(tl, 11 - i, 5) << i;
    for (let i = 0; i < 8; i++) b |= at(tr, 3 - i, 5) << i;
    for (let i = 8; i < 15; i++) b |= at(bl, 5, i - 11) << i;
    return bestFormat(a, b);
  }
  // 型番情報（7 以上）：右上と左下の目印のそば
  function versionNear(at, tr, bl) {
    for (const [c, horiz] of [[tr, true], [bl, false]]) {
      let bits = 0;
      for (let i = 0; i < 18; i++) {
        const a = -7 + i % 3, b = Math.floor(i / 3) - 3;
        bits |= (horiz ? at(c, a, b) : at(c, b, a)) << i;
      }
      const v = versionFromBits(bits);
      if (v) return v;
    }
    return 0;
  }
  // タイミングパターン（6 行目・6 列目の白黒交互）がそろっているか。マス目の当てはめが合っているかの目安
  function timingOk(get, dim) {
    let ok = 0, n = 0;
    for (let i = 8; i < dim - 8; i++) {
      const want = i % 2 === 0 ? 1 : 0;
      ok += (get(i, 6) === want ? 1 : 0) + (get(6, i) === want ? 1 : 0);
      n += 2;
    }
    return ok >= n * 0.7;
  }

  /* ---------- 右下の位置合わせの模様（5×5）を探す ---------- */
  function findAlignment(bin, w, h, est, ux, uy, vx, vy, m) {
    const tol = m / 2 + 1, near = r => Math.abs(r - m) < tol;
    const at = (x, y) => { const px = Math.floor(x), py = Math.floor(y); return px >= 0 && py >= 0 && px < w && py < h ? bin[py * w + px] : 0; };
    // 中心の黒・まわりの白・外側の黒の形にどれだけ合うか（25 マス中）
    const score = (cx, cy) => {
      let s = 0;
      for (let j = -2; j <= 2; j++) for (let i = -2; i <= 2; i++) {
        const want = Math.max(Math.abs(i), Math.abs(j)) !== 1 ? 1 : 0;
        if (at(cx + ux * i + vx * j, cy + uy * i + vy * j) === want) s++;
      }
      return s;
    };
    // 縦か横の「白・黒・白」（それぞれ 1 マス）の黒の中心
    const line = (x, y, dx, dy) => {
      const g = i => { const px = x + dx * i, py = y + dy * i; return px >= 0 && py >= 0 && px < w && py < h ? bin[py * w + px] : -1; };
      if (g(0) !== 1) return null;
      let i = 0, b = 0, lb = 0, j = 1, f = 0, lf = 0;
      while (g(i) === 1) { if (++b > 3 * m) return null; i--; }
      while (g(i) === 0) { if (++lb > 3 * m) return null; i--; }
      if (g(i) !== 1) return null;
      while (g(j) === 1) { if (++f > 3 * m) return null; j++; }
      while (g(j) === 0) { if (++lf > 3 * m) return null; j++; }
      if (g(j) !== 1) return null;
      return near(b + f) && near(lb) && near(lf) ? (f - b + 2) / 2 : null;
    };
    for (const rad of [3, 6, 12]) {
      const r = rad * m;
      const x0 = Math.max(0, Math.floor(est.x - r)), x1 = Math.min(w, Math.ceil(est.x + r));
      const y0 = Math.max(0, Math.floor(est.y - r)), y1 = Math.min(h, Math.ceil(est.y + r));
      let best = null;
      for (let y = y0; y < y1; y++) {
        const row = y * w, runs = []; // [始まり, 長さ, 色] の並び
        let s = x0;
        for (let x = x0 + 1; x <= x1; x++) if (x === x1 || bin[row + x] !== bin[row + s]) { runs.push(s, x - s, bin[row + s]); s = x; }
        for (let k = 3; k + 3 < runs.length; k += 3) {
          if (runs[k + 2] !== 1 || !near(runs[k + 1]) || !near(runs[k - 2]) || !near(runs[k + 4])) continue;
          let cx = runs[k] + runs[k + 1] / 2;
          const dyc = line(Math.floor(cx), y, 0, 1);
          if (dyc == null) continue;
          const cy = y + dyc, dxc = line(Math.floor(cx), Math.floor(cy), 1, 0);
          if (dxc != null) cx = Math.floor(cx) + dxc;
          const sc = score(cx, cy), d = Math.hypot(cx - est.x, cy - est.y);
          if (sc >= 22 && (!best || sc > best.sc || (sc === best.sc && d < best.d))) best = { x: cx, y: cy, sc, d };
        }
      }
      if (best) return best;
    }
    return null;
  }

  /* ---------- リード・ソロモン符号の誤り訂正（GF(256)、生成多項式の根は α^0 から） ---------- */
  const EXP = new Uint8Array(512), LOG = new Uint8Array(256);
  for (let i = 0, x = 1; i < 255; i++) { EXP[i] = x; LOG[x] = i; x <<= 1; if (x & 256) x ^= 0x11D; }
  for (let i = 255; i < 512; i++) EXP[i] = EXP[i - 255];
  const gmul = (a, b) => (a && b ? EXP[LOG[a] + LOG[b]] : 0);
  const ginv = a => EXP[255 - LOG[a]];
  // 多項式は「次数の高い順」の配列
  const pnorm = p => { let i = 0; while (i < p.length - 1 && p[i] === 0) i++; return i ? p.slice(i) : p; };
  const pdeg = p => p.length - 1;
  const pzero = p => p[0] === 0;
  const pcoef = (p, d) => p[p.length - 1 - d];
  function padd(a, b) {
    if (a.length < b.length) [a, b] = [b, a];
    const r = a.slice(), o = a.length - b.length;
    for (let i = 0; i < b.length; i++) r[o + i] ^= b[i];
    return pnorm(r);
  }
  function pmul(a, b) {
    if (pzero(a) || pzero(b)) return [0];
    const r = new Array(a.length + b.length - 1).fill(0);
    for (let i = 0; i < a.length; i++) for (let j = 0; j < b.length; j++) r[i + j] ^= gmul(a[i], b[j]);
    return pnorm(r);
  }
  const pmono = (d, c) => { if (!c) return [0]; const r = new Array(d + 1).fill(0); r[0] = c; return r; };
  const pmulMono = (a, d, c) => { if (!c) return [0]; const r = a.map(x => gmul(x, c)); for (let i = 0; i < d; i++) r.push(0); return pnorm(r); };
  const pscale = (a, c) => (c ? pnorm(a.map(x => gmul(x, c))) : [0]);
  const peval = (p, x) => { let r = 0; for (let i = 0; i < p.length; i++) r = gmul(r, x) ^ p[i]; return r; };

  // block（データ＋誤り訂正コード）を直す。直した数を返す。直せなければ -1
  function rsCorrect(block, ecc) {
    const synd = new Array(ecc);
    let any = false;
    for (let i = 0; i < ecc; i++) { const s = peval(block, EXP[i]); synd[ecc - 1 - i] = s; if (s) any = true; }
    if (!any) return 0;
    // ユークリッドの互除法で、誤りの位置の多項式（sigma）と値の多項式（omega）を求める
    let rLast = pmono(ecc, 1), r = pnorm(synd), tLast = [0], t = [1];
    if (pdeg(rLast) < pdeg(r)) [rLast, r] = [r, rLast];
    while (2 * pdeg(r) >= ecc) {
      const rLastLast = rLast, tLastLast = tLast;
      rLast = r; tLast = t;
      if (pzero(rLast)) return -1;
      r = rLastLast;
      let q = [0];
      const dltInv = ginv(rLast[0]);
      while (pdeg(r) >= pdeg(rLast) && !pzero(r)) {
        const dd = pdeg(r) - pdeg(rLast), sc = gmul(r[0], dltInv);
        q = padd(q, pmono(dd, sc));
        r = padd(r, pmulMono(rLast, dd, sc));
      }
      t = padd(pmul(q, tLast), tLastLast);
      if (pdeg(r) >= pdeg(rLast)) return -1;
    }
    const s0 = pcoef(t, 0);
    if (!s0) return -1;
    const iv = ginv(s0), sigma = pscale(t, iv), omega = pscale(r, iv);
    // 位置（チェンの探索）
    const ne = pdeg(sigma), locs = [];
    if (ne === 1) locs.push(pcoef(sigma, 1));
    else for (let i = 1; i < 256 && locs.length < ne; i++) if (peval(sigma, i) === 0) locs.push(ginv(i));
    if (locs.length !== ne) return -1;
    // 値（フォーニーの式）
    for (let i = 0; i < ne; i++) {
      const xi = ginv(locs[i]);
      let den = 1;
      for (let j = 0; j < ne; j++) if (j !== i) den = gmul(den, 1 ^ gmul(locs[j], xi));
      const pos = block.length - 1 - LOG[locs[i]];
      if (pos < 0 || !den) return -1;
      block[pos] ^= gmul(peval(omega, xi), ginv(den));
    }
    for (let i = 0; i < ecc; i++) if (peval(block, EXP[i])) return -1;
    return ne;
  }

  /* ---------- マス目からデータを取り出す ---------- */
  const fnCache = {};
  // 目印・タイミング・形式情報など、データでないマス
  function functionMask(ver) {
    if (fnCache[ver]) return fnCache[ver];
    const size = ver * 4 + 17, f = new Uint8Array(size * size);
    const set = (x, y) => { if (x >= 0 && y >= 0 && x < size && y < size) f[y * size + x] = 1; };
    for (let i = 0; i < size; i++) { set(6, i); set(i, 6); }
    for (const [cx, cy] of [[3, 3], [size - 4, 3], [3, size - 4]]) {
      for (let dy = -4; dy <= 4; dy++) for (let dx = -4; dx <= 4; dx++) set(cx + dx, cy + dy);
    }
    const pos = S.alignPositions(ver, size), n = pos.length;
    for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) {
      if ((i === 0 && j === 0) || (i === 0 && j === n - 1) || (i === n - 1 && j === 0)) continue;
      for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) set(pos[i] + dx, pos[j] + dy);
    }
    for (let i = 0; i <= 8; i++) { set(8, i); set(i, 8); }
    for (let i = 0; i < 8; i++) { set(size - 1 - i, 8); set(8, size - 1 - i); }
    if (ver >= 7) for (let i = 0; i < 18; i++) { const a = size - 11 + i % 3, c = Math.floor(i / 3); set(a, c); set(c, a); }
    return (fnCache[ver] = f);
  }
  function maskBit(m, x, y) {
    switch (m) {
      case 0: return (x + y) % 2 === 0;
      case 1: return y % 2 === 0;
      case 2: return x % 3 === 0;
      case 3: return (x + y) % 3 === 0;
      case 4: return (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0;
      case 5: return x * y % 2 + x * y % 3 === 0;
      case 6: return (x * y % 2 + x * y % 3) % 2 === 0;
      default: return ((x + y) % 2 + x * y % 3) % 2 === 0;
    }
  }

  // get(x, y) → 1 = 黒。読めれば { text, bytes, version, ecl, mask, errors, ... }
  function decodeGrid(get, dim, fmt) {
    const ver = (dim - 17) / 4;
    if (ver < 1 || ver > 40 || ver % 1) return null;
    const fn = functionMask(ver), raw = Math.floor(S.rawDataModules(ver) / 8), cw = new Uint8Array(raw);
    let k = 0;
    for (let right = dim - 1; right >= 1; right -= 2) {
      if (right === 6) right = 5;
      for (let v = 0; v < dim; v++) for (let j = 0; j < 2; j++) {
        const x = right - j, up = ((right + 1) & 2) === 0, y = up ? dim - 1 - v : v;
        if (fn[y * dim + x] || k >= raw * 8) continue;
        if (get(x, y) ^ (maskBit(fmt.mask, x, y) ? 1 : 0)) cw[k >>> 3] |= 0x80 >>> (k & 7);
        k++;
      }
    }
    // ブロックごとに分け直して、誤りを直す
    const nb = S.NUM_BLOCKS[fmt.ecl][ver], ecc = S.ECC_PER_BLOCK[fmt.ecl][ver];
    const numShort = nb - raw % nb, shortLen = Math.floor(raw / nb);
    const blocks = Array.from({ length: nb }, (_, j) => new Uint8Array(j < numShort ? shortLen : shortLen + 1));
    let p = 0;
    for (let i = 0; i <= shortLen; i++) for (let j = 0; j < nb; j++) {
      // 短いブロックは、データの終わり（shortLen - ecc の所）が1つ少ない
      if (j < numShort && i === shortLen - ecc) continue;
      blocks[j][j < numShort && i > shortLen - ecc ? i - 1 : i] = cw[p++];
    }
    const data = [];
    let errors = 0;
    for (const b of blocks) {
      const e = rsCorrect(b, ecc);
      if (e < 0) return null;
      errors += e;
      for (let i = 0; i < b.length - ecc; i++) data.push(b[i]);
    }
    const parsed = parseData(data, ver);
    if (!parsed) return null;
    return Object.assign(parsed, { version: ver, ecl: fmt.ecl, mask: fmt.mask, errors });
  }

  /* ---------- データの並び（数字・英数字・バイト・漢字）を文字に戻す ---------- */
  const ALNUM = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ $%*+-./:';
  const ECI_CHARSET = {
    1: 'iso-8859-1', 3: 'iso-8859-1', 4: 'iso-8859-2', 5: 'iso-8859-3', 6: 'iso-8859-4', 7: 'iso-8859-5', 8: 'iso-8859-6',
    9: 'iso-8859-7', 10: 'iso-8859-8', 11: 'iso-8859-9', 12: 'iso-8859-10', 13: 'iso-8859-11', 15: 'iso-8859-13',
    16: 'iso-8859-14', 17: 'iso-8859-15', 18: 'iso-8859-16', 20: 'shift_jis', 21: 'windows-1250', 22: 'windows-1251',
    23: 'windows-1252', 24: 'windows-1256', 25: 'utf-16be', 26: 'utf-8', 27: 'utf-8', 28: 'big5', 29: 'gb18030', 30: 'euc-kr', 170: 'utf-8',
  };
  function tryDecode(bytes, label) {
    try { return new TextDecoder(label, { fatal: true }).decode(bytes); } catch (e) { return null; }
  }
  // 文字コードの指定がないバイト列：UTF-8 → シフト JIS（昔の日本の QR）→ Latin-1 の順に試す
  function guessText(bytes) {
    const u = tryDecode(bytes, 'utf-8');
    if (u != null) return { text: u, charset: 'UTF-8' };
    const s = tryDecode(bytes, 'shift_jis');
    if (s != null) return { text: s, charset: 'Shift_JIS' };
    return { text: new TextDecoder('windows-1252').decode(bytes), charset: 'ISO-8859-1' };
  }

  function parseData(bytes, ver) {
    const nbits = bytes.length * 8;
    let pos = 0;
    const read = n => {
      let v = 0;
      for (let i = 0; i < n; i++, pos++) v = (v << 1) | (pos < nbits ? (bytes[pos >> 3] >> (7 - (pos & 7))) & 1 : 0);
      return v;
    };
    const g = ver <= 9 ? 0 : ver <= 26 ? 1 : 2;
    const parts = []; // { cs: 文字コード（null は推測）, b: [バイト] }
    const add = (arr, cs, ascii) => {
      const last = parts[parts.length - 1];
      if (last && (last.cs === cs || ascii)) last.b.push(...arr);
      else parts.push({ cs, b: [...arr] });
    };
    let eci = null, sa = null, gs1 = false;
    while (pos + 4 <= nbits) {
      const mode = read(4);
      if (mode === 0) break;
      if (mode === 7) { // ECI（文字コードの指定）
        let v = read(8);
        if ((v & 0x80) === 0) { /* 1 バイト */ } else if ((v & 0xC0) === 0x80) v = ((v & 0x3F) << 8) | read(8);
        else if ((v & 0xE0) === 0xC0) v = ((v & 0x1F) << 16) | read(16);
        else return null;
        eci = ECI_CHARSET[v] || null;
        continue;
      }
      if (mode === 3) { sa = { index: read(4), total: read(4) + 1, parity: read(8) }; continue; } // 分割 QR
      if (mode === 5) { gs1 = true; continue; }
      if (mode === 9) { read(8); continue; }
      if (mode === 1) { // 数字
        let n = read([10, 12, 14][g]);
        const out = [];
        while (n >= 3) { const v = read(10); if (v > 999) return null; out.push(...String(v).padStart(3, '0')); n -= 3; }
        if (n === 2) { const v = read(7); if (v > 99) return null; out.push(...String(v).padStart(2, '0')); }
        else if (n === 1) { const v = read(4); if (v > 9) return null; out.push(String(v)); }
        add(out.map(c => c.charCodeAt(0)), null, true);
      } else if (mode === 2) { // 英数字
        let n = read([9, 11, 13][g]);
        let s = '';
        while (n >= 2) { const v = read(11); if (v >= 45 * 45) return null; s += ALNUM[Math.floor(v / 45)] + ALNUM[v % 45]; n -= 2; }
        if (n === 1) { const v = read(6); if (v >= 45) return null; s += ALNUM[v]; }
        if (gs1) s = s.replace(/%%|%/g, m => (m === '%%' ? '%' : '\x1D'));
        add([...s].map(c => c.charCodeAt(0)), null, true);
      } else if (mode === 4) { // バイト
        const n = read([8, 16, 16][g]);
        if (pos + n * 8 > nbits) return null;
        const out = new Array(n);
        for (let i = 0; i < n; i++) out[i] = read(8);
        add(out, eci, false);
      } else if (mode === 8) { // 漢字（シフト JIS を 13 ビットに詰めたもの）
        const n = read([8, 10, 12][g]);
        if (pos + n * 13 > nbits) return null;
        const out = [];
        for (let i = 0; i < n; i++) {
          const v = read(13);
          let c = (Math.floor(v / 0xC0) << 8) | (v % 0xC0);
          c += c < 0x1F00 ? 0x8140 : 0xC140;
          out.push(c >> 8, c & 0xFF);
        }
        add(out, 'shift_jis', false);
      } else return null;
      if (pos > nbits) return null;
    }
    let text = '';
    const charsets = new Set(), all = [];
    for (const p of parts) {
      const b = Uint8Array.from(p.b);
      all.push(...p.b);
      let t = p.cs ? tryDecode(b, p.cs) : null;
      if (t != null) charsets.add(p.cs === 'shift_jis' ? 'Shift_JIS' : p.cs.toUpperCase());
      else { const gt = guessText(b); t = gt.text; charsets.add(gt.charset); }
      text += t;
    }
    return { text, bytes: Uint8Array.from(all), charset: [...charsets].join(' + ') || 'UTF-8', sa, gs1 };
  }

  /* ---------- 3つ組から読む ---------- */
  function decodeTriple(bin, w, h, t) {
    const unit = (a, b) => { const l = Math.hypot(b.x - a.x, b.y - a.y); return { x: (b.x - a.x) / l, y: (b.y - a.y) / l }; };
    const d1 = unit(t.tl, t.tr), d2 = unit(t.tl, t.bl);
    // 目印の大きさ（1 マスの画素数）を、QR の2つの向き（左上→右上、左上→左下）に沿って測り直す。傾きや、ななめから撮った歪みがあっても測れる。
    // 測れなければ、縦横の線で測った大きさを傾きの分だけ戻して使う（傾いていると大きめに出る。45° で 1.4 倍）
    const tilt = Math.max(Math.abs(d1.x), Math.abs(d1.y));
    const measure = c => {
      const a = chord(bin, w, h, c, d1.x, d1.y, c.size * 6), b = chord(bin, w, h, c, d2.x, d2.y, c.size * 6);
      return { x: c.x, y: c.y, s1: a ? a / 7 : c.size * tilt, s2: b ? b / 7 : c.size * tilt };
    };
    const TL = measure(t.tl), A = measure(t.tr), B = measure(t.bl);
    for (const mirror of [false, true]) {
      // 鏡写しのときは右上と左下が入れ替わる。su = 横（右上の向き）の 1 マス、sv = 縦（左下の向き）の 1 マス
      const U = mirror ? d2 : d1, V = mirror ? d1 : d2;
      const pick = c => (mirror ? { x: c.x, y: c.y, su: c.s2, sv: c.s1 } : { x: c.x, y: c.y, su: c.s1, sv: c.s2 });
      const tl = pick(TL), tr = pick(mirror ? B : A), bl = pick(mirror ? A : B);
      const gm = c => Math.sqrt(c.su * c.sv), m = (gm(tl) + gm(tr) + gm(bl)) / 3;
      const ltr = Math.hypot(tr.x - tl.x, tr.y - tl.y), lbl = Math.hypot(bl.x - tl.x, bl.y - tl.y);
      // 目印のそばのマス（at(目印, 横に何マス, 縦に何マス)）。マス目全体の大きさが分からなくても読める
      const at = (c, mx, my) => {
        const px = Math.floor(c.x + U.x * mx * c.su + V.x * my * c.sv), py = Math.floor(c.y + U.y * mx * c.su + V.y * my * c.sv);
        return px >= 0 && py >= 0 && px < w && py < h ? bin[py * w + px] : 0;
      };
      // まず目印のそばの形式情報が読めるかで、この3つ組が本物かを手早く確かめる
      const fmtNear = formatNear(at, tl, tr, bl);
      if (!fmtNear) continue;
      // 大きさの候補：見積もりに近いもの。型番情報（7 以上）が読めて、見積もりとかけ離れていなければ、それを先に
      const vEst = ((ltr / ((tl.su + tr.su) / 2) + lbl / ((tl.sv + bl.sv) / 2)) / 2 + 7 - 17) / 4, v0 = Math.round(vEst), dims = [];
      if (vEst >= 5.5) {
        const vi = versionNear(at, tr, bl);
        if (vi && Math.abs(vi - vEst) <= Math.max(2, vEst * 0.25)) dims.push(17 + 4 * vi);
      }
      for (const v of [v0, v0 - 1, v0 + 1]) if (v >= 1 && v <= 40 && !dims.includes(17 + 4 * v)) dims.push(17 + 4 * v);
      // 4つ目の角（右下）の見積もり。ななめから撮ると、近い側の目印ほど大きく写るので、その比から奥行きを見積もる。
      // （射影変換では「同次座標 × 奥行き w」がマス目の座標の一次式になる。目印の大きさ ∝ w^(-3/2)）
      const wTR = Math.min(2, Math.max(0.5, (gm(tr) / gm(tl)) ** (-2 / 3))), wBL = Math.min(2, Math.max(0.5, (gm(bl) / gm(tl)) ** (-2 / 3)));
      const wBR = wTR + wBL - 1;
      const brAff = { x: tr.x + bl.x - tl.x, y: tr.y + bl.y - tl.y };
      const brPer = wBR > 0.3 ? { x: (wTR * tr.x + wBL * bl.x - tl.x) / wBR, y: (wTR * tr.y + wBL * bl.y - tl.y) / wBR } : brAff;
      const persp = Math.hypot(brPer.x - brAff.x, brPer.y - brAff.y) > m;
      const make = (dim, br, brMod) => homography([[3.5, 3.5], [dim - 3.5, 3.5], [brMod, brMod], [3.5, dim - 3.5]], [tl, tr, br, bl]);
      for (const dim of dims) {
        const ver = (dim - 17) / 4, tries = [];
        if (ver >= 2) {
          // 位置合わせの模様（マス目で右下の角から 3 マス内側）の見積もり
          const kk = (dim - 10) / (dim - 7), wA = 1 + (wBR - 1) * kk;
          const est = persp ? { x: (tl.x + (wBR * brPer.x - tl.x) * kk) / wA, y: (tl.y + (wBR * brPer.y - tl.y) * kk) / wA }
            : { x: tl.x + (brAff.x - tl.x) * kk, y: tl.y + (brAff.y - tl.y) * kk };
          const ux = (tr.x - tl.x) / (dim - 7), uy = (tr.y - tl.y) / (dim - 7), vx = (bl.x - tl.x) / (dim - 7), vy = (bl.y - tl.y) / (dim - 7);
          const ap = findAlignment(bin, w, h, est, ux, uy, vx, vy, m);
          if (ap) tries.push([ap, dim - 6.5]);
        }
        if (persp) tries.push([brPer, dim - 3.5]);
        tries.push([brAff, dim - 3.5]);
        for (const [br, brMod] of tries) {
          const H = make(dim, br, brMod);
          if (!H) continue;
          const get = sampler(bin, w, h, H);
          if (!timingOk(get, dim)) continue;
          const fmts = [fmtNear], f2 = readFormat(get, dim);
          if (f2 && (f2.ecl !== fmtNear.ecl || f2.mask !== fmtNear.mask)) fmts.push(f2);
          for (const fmt of fmts) {
            const r = decodeGrid(get, dim, fmt);
            if (r) {
              r.corners = [H(0, 0), H(dim, 0), H(dim, dim), H(0, dim)];
              r.mirrored = mirror;
              return r;
            }
          }
        }
      }
    }
    return null;
  }

  function detect(bin, w, h, maxTries) {
    const cands = findFinders(bin, w, h), out = [];
    if (cands.length < 3) return out;
    const used = new Set();
    let tries = 0;
    for (const t of makeTriples(cands)) {
      if (t.ids.some(i => used.has(i))) continue;
      if (++tries > maxTries) break;
      const r = decodeTriple(bin, w, h, t);
      if (r) { out.push(r); t.ids.forEach(i => used.add(i)); }
    }
    return out;
  }

  // 同じ QR を二重に数えない（同じ中身で、中心が近い）
  const centerOf = r => ({ x: r.corners.reduce((s, p) => s + p.x, 0) / 4, y: r.corners.reduce((s, p) => s + p.y, 0) / 4 });
  function addResult(list, r) {
    const c = centerOf(r), size = Math.hypot(r.corners[0].x - r.corners[2].x, r.corners[0].y - r.corners[2].y);
    for (const o of list) {
      const oc = centerOf(o);
      if (o.text === r.text && Math.hypot(oc.x - c.x, oc.y - c.y) < size / 2) return;
    }
    list.push(r);
  }

  // img: getImageData の結果（{ data, width, height }）か { gray, width, height }
  // quick: カメラ用（速さ優先。見つかった時点でやめる）
  function scan(img, { quick = false } = {}) {
    const w = img.width, h = img.height, gray = img.gray || toGray(img), found = [];
    if (w < 21 || h < 21) return found;
    const maxTries = quick ? 12 : 80;
    const invert = b => { const o = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) o[i] = b[i] ^ 1; return o; };
    // 白黒の分け方：場所ごとのしきい値 → その白黒反転（白い QR）→ だめなら画像全体で1つのしきい値（と反転）
    const run = (g, W, Hh, scale) => {
      const before = found.length;
      const local = binarizeLocal(g, W, Hh);
      const passes = [() => local, () => invert(local)];
      if (!quick) { let gl; passes.push(() => (gl = binarizeGlobal(g, W, Hh)), () => invert(gl)); }
      for (let i = 0; i < passes.length; i++) {
        if (i === 2 && found.length > before) break; // 場所ごとのしきい値で見つかったら、全体のしきい値は試さない
        for (const r of detect(passes[i](), W, Hh, maxTries)) {
          if (scale !== 1) r.corners = r.corners.map(p => ({ x: p.x * scale, y: p.y * scale }));
          addResult(found, r);
        }
        if (quick && found.length) return;
      }
    };
    // 大きな画像は、まず縮めて読む（速く、写真のざらつきにも強い）。足りなければ大きくしていく。
    // 4M 画素までの画像（スクショなど）は、縮めて見つかっても元の大きさでも読む（小さな QR を見落とさない）
    const levels = [{ g: gray, w, h, scale: 1 }];
    while (Math.max(levels[0].w, levels[0].h) > 1600) {
      const L = levels[0], s = half(L.g, L.w, L.h);
      levels.unshift({ g: s.g, w: s.w, h: s.h, scale: L.scale * 2 });
    }
    for (const L of levels) {
      if (found.length && (quick || L.w * L.h > 4e6)) break;
      run(L.g, L.w, L.h, L.scale);
    }
    // ぼやけた画像は、縮めると読めることがある
    if (!found.length && !quick && levels.length === 1 && Math.min(w, h) >= 300) {
      const s = half(gray, w, h);
      run(s.g, s.w, s.h, 2);
    }
    // 上から下、左から右の順に並べる
    return found.sort((a, b) => { const ca = centerOf(a), cb = centerOf(b); return Math.abs(ca.y - cb.y) > 20 ? ca.y - cb.y : ca.x - cb.x; });
  }

  return { scan, _test: { decodeGrid, readFormat, rsCorrect, parseData, toGray, binarizeLocal, binarizeGlobal, findFinders, makeTriples, decodeTriple, detect } };
})();

if (typeof module === 'object' && module.exports) module.exports = QRRead; // テスト用
