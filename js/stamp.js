'use strict';
// =========================================================
// ハンコ・サインの部品
//   ・名字から赤い丸の印鑑を作る（Stamp.seal）
//   ・写真やスキャンから、白い紙の部分を透明にして切り抜く（Stamp.fromImage）
//   ・手書きの線を画像にする（Stamp.strokesToCanvas）、まわりの透明な余白を切る（Stamp.trim）
//   ・作ったものを、この端末のこのブラウザの中に保存しておく（StampStore。IndexedDB）
// どれも「透明つきの canvas」を返す。common.js のあとに読み込む
// =========================================================
const Stamp = (() => {
  const SEAL_COLOR = '#d9312b'; // 朱色
  const INKS = { red: SEAL_COLOR, black: '#1a1a1a', blue: '#1f4fd1' };
  const D = 600;                // 印鑑の画像の大きさ（画素）
  const hexRgb = h => [1, 3, 5].map(i => parseInt(h.slice(i, i + 2), 16));

  /* ---------- 書体（印鑑用。その端末にあるものだけ使う） ---------- */
  const SEAL_FONTS = {
    kaisho: { label: '楷書', families: ['HG正楷書体-PRO', 'HGSeikaishotaiPRO', 'UD デジタル 教科書体 N-B', 'UD Digi Kyokasho N-B', 'Klee'], weight: 'normal', thick: 0.05 },
    mincho: { label: '明朝', families: ['Yu Mincho', 'YuMincho', 'Hiragino Mincho ProN', 'Noto Serif CJK JP', 'Noto Serif JP'], weight: 'bold', thick: 0.035 },
    gothic: { label: 'ゴシック', families: ['Yu Gothic', 'YuGothic', 'Hiragino Sans', 'Meiryo', 'Noto Sans CJK JP', 'Noto Sans JP'], weight: 'bold', thick: 0.02 },
  };
  const glyph = font => { const c = newCanvas(48, 48), x = c.getContext('2d'); x.font = font; x.fillText('永', 4, 40); return c.toDataURL(); };
  let fontCache = null;
  // 使える書体：{ kaisho: CSS の font-family か null, mincho: …, gothic: … }
  function sealFonts() {
    if (fontCache) return fontCache;
    const serif = glyph('40px serif'), sans = glyph('40px sans-serif');
    // その名前の書体があれば、ないとき（serif / sans-serif で代わりに描かれる）と形が変わる
    const has = f => glyph(`40px "${f}", serif`) !== serif || glyph(`40px "${f}", sans-serif`) !== sans;
    fontCache = {};
    for (const [k, v] of Object.entries(SEAL_FONTS)) {
      const f = v.families.find(has);
      fontCache[k] = f ? `"${f}"` : k === 'gothic' ? 'sans-serif' : k === 'mincho' && serif !== sans ? 'serif' : null;
    }
    return fontCache;
  }

  /* ---------- 印鑑 ---------- */
  // name：名字（1〜4文字）、font：'kaisho' | 'mincho' | 'gothic'、rough：かすれ（本物っぽく）
  // 1〜3文字は縦に1列、4文字は2列（右の列から読む）
  function seal({ name, font = 'kaisho', rough = false }) {
    const chars = [...String(name).replace(/\s/g, '')].slice(0, 4), n = chars.length;
    const fonts = sealFonts(), key = fonts[font] ? font : 'gothic', spec = SEAL_FONTS[key];
    const c = newCanvas(D, D), x = c.getContext('2d');
    const lw = D * 0.05, R = D / 2 - lw / 2 - 2, Ri = R - lw / 2 - D * 0.025;
    x.fillStyle = x.strokeStyle = SEAL_COLOR;
    x.lineWidth = lw; x.beginPath(); x.arc(D / 2, D / 2, R, 0, Math.PI * 2); x.stroke();
    // 文字の枠：[中心 x, 中心 y, 幅, 高さ]（内側の半径に対する割合）
    const boxes = [];
    if (n === 1) boxes.push([0, 0, 1.25, 1.25]);
    else if (n <= 3) {
      const bw = n === 2 ? 1.05 : 0.9, bh = n === 2 ? 0.8 : 0.58;
      for (let i = 0; i < n; i++) boxes.push([0, (i - (n - 1) / 2) * bh, bw, bh]);
    } else { const s = 0.72; boxes.push([s / 2, -s / 2, s, s], [s / 2, s / 2, s, s], [-s / 2, -s / 2, s, s], [-s / 2, s / 2, s, s]); }
    // 字と字の間を少しあける（太らせた線どうしがくっつかないように）
    const [fw, fh] = n === 1 ? [0.94, 0.94] : n === 2 ? [0.94, 0.8] : n === 3 ? [0.94, 0.86] : [0.88, 0.88];
    chars.forEach((ch, i) => {
      const [bx, by, bw, bh] = boxes[i].map(v => v * Ri);
      drawFit(x, ch, `${spec.weight} 200px ${fonts[key]}`, spec.thick, D / 2 + bx, D / 2 + by, bw * fw, bh * fh);
    });
    if (rough) roughen(x, hash(chars.join('') + key));
    return c;
  }
  // 1文字を、枠いっぱいに（縦横それぞれ伸び縮みさせて）描く。インクのある所の大きさで合わせる
  function drawFit(x, ch, font, thick, cx, cy, bw, bh) {
    x.font = font;
    const m = x.measureText(ch);
    const l = m.actualBoundingBoxLeft || 0, r = m.actualBoundingBoxRight || m.width;
    const a = m.actualBoundingBoxAscent || 170, d = m.actualBoundingBoxDescent || 20;
    let sx = bw / Math.max(1, l + r), sy = bh / Math.max(1, a + d);
    if (sx > sy * 1.6) sx = sy * 1.6; else if (sy > sx * 1.6) sy = sx * 1.6; // 細長い字（一・川など）を伸ばしすぎない
    x.save();
    x.translate(cx, cy); x.scale(sx, sy);
    x.textBaseline = 'alphabetic'; x.textAlign = 'left';
    const ox = -(r - l) / 2, oy = (a - d) / 2; // インクのある所のまん中が (0, 0) に来るように
    x.lineWidth = 200 * thick; x.lineJoin = 'round';
    x.fillText(ch, ox, oy); x.strokeText(ch, ox, oy); // 少し太らせて、ハンコらしく
    x.restore();
  }
  // かすれ：インクのつかない小さな所と、ところどころ薄い所を作る（名前から決まる乱数なので、毎回同じ形）
  function rnd(seed) {
    return () => {
      seed = (seed + 0x6D2B79F5) | 0;
      let t = Math.imul(seed ^ (seed >>> 15), 1 | seed);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }
  const hash = s => { let h = 2166136261; for (const ch of s) h = Math.imul(h ^ ch.codePointAt(0), 16777619); return h >>> 0; };
  function roughen(x, seed) {
    const r = rnd(seed);
    x.save();
    x.globalCompositeOperation = 'destination-out';
    for (let i = 0; i < 240; i++) {
      const s = D * (0.003 + r() * r() * 0.018);
      x.globalAlpha = 0.45 + r() * 0.55;
      x.beginPath(); x.ellipse(r() * D, r() * D, s * (0.5 + r()), s, r() * Math.PI, 0, Math.PI * 2); x.fill();
    }
    x.globalAlpha = 1;
    for (let i = 0; i < 6; i++) {
      const gx = r() * D, gy = r() * D, gr = D * (0.15 + r() * 0.2), g = x.createRadialGradient(gx, gy, 0, gx, gy, gr);
      g.addColorStop(0, 'rgba(0,0,0,0.35)'); g.addColorStop(1, 'rgba(0,0,0,0)');
      x.fillStyle = g; x.fillRect(0, 0, D, D);
    }
    x.restore();
  }

  /* ---------- 写真・スキャンから ---------- */
  // 紙の色を場所ごとに求める（影や明るさのむらがあっても、白い紙として扱えるように）
  //   格子に分け、明るいほうから 25% の画素の平均を、その格子の紙の色とする。
  //   インクで埋まった格子（まわりよりずっと暗い）は、まわりの紙の色で埋める。
  function paperGrid(p, W, H) {
    const cell = Math.max(16, Math.round(Math.max(W, H) / 10)), gw = Math.ceil(W / cell), gh = Math.ceil(H / cell);
    const vals = new Float32Array(gw * gh * 3), lum = new Float32Array(gw * gh), hist = new Uint32Array(256);
    const L = j => (p[j] * 77 + p[j + 1] * 150 + p[j + 2] * 29) >> 8;
    for (let gy = 0; gy < gh; gy++) for (let gx = 0; gx < gw; gx++) {
      const x0 = gx * cell, y0 = gy * cell, x1 = Math.min(W, x0 + cell), y1 = Math.min(H, y0 + cell);
      hist.fill(0);
      let n = 0;
      for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) { hist[L((y * W + x) * 4)]++; n++; }
      let k = 255, acc = 0;
      while (k > 0 && acc + hist[k] < n * 0.25) acc += hist[k--];
      let sr = 0, sg = 0, sb = 0, m = 0;
      for (let y = y0; y < y1; y++) for (let x = x0; x < x1; x++) {
        const j = (y * W + x) * 4;
        if (L(j) >= k) { sr += p[j]; sg += p[j + 1]; sb += p[j + 2]; m++; }
      }
      const g = gy * gw + gx;
      vals[g * 3] = sr / m; vals[g * 3 + 1] = sg / m; vals[g * 3 + 2] = sb / m; lum[g] = k;
    }
    let top = 0;
    for (const v of lum) top = Math.max(top, v);
    let ok = Uint8Array.from(lum, v => (v >= top * 0.75 ? 1 : 0));
    for (let pass = 0; pass < gw + gh && ok.includes(0); pass++) {
      const next = ok.slice();
      for (let gy = 0; gy < gh; gy++) for (let gx = 0; gx < gw; gx++) {
        const g = gy * gw + gx;
        if (ok[g]) continue;
        let s0 = 0, s1 = 0, s2 = 0, m = 0;
        for (const [dx, dy] of [[1, 0], [-1, 0], [0, 1], [0, -1], [1, 1], [-1, -1], [1, -1], [-1, 1]]) {
          const nx = gx + dx, ny = gy + dy;
          if (nx < 0 || ny < 0 || nx >= gw || ny >= gh || !ok[ny * gw + nx]) continue;
          const q = (ny * gw + nx) * 3; s0 += vals[q]; s1 += vals[q + 1]; s2 += vals[q + 2]; m++;
        }
        if (m) { vals[g * 3] = s0 / m; vals[g * 3 + 1] = s1 / m; vals[g * 3 + 2] = s2 / m; next[g] = 1; }
      }
      ok = next;
    }
    for (let i = 0; i < vals.length; i++) vals[i] = Math.max(40, vals[i]);
    // (x, y) の紙の色を out に入れる（格子の中心どうしの間をなめらかにつなぐ）
    return (x, y, out) => {
      const fx = Math.min(gw - 1, Math.max(0, (x + 0.5) / cell - 0.5)), fy = Math.min(gh - 1, Math.max(0, (y + 0.5) / cell - 0.5));
      const ix = Math.min(gw - 2, Math.floor(fx)), iy = Math.min(gh - 2, Math.floor(fy));
      if (gw === 1 || gh === 1) { const g = (gw === 1 ? Math.round(fy) * gw : Math.round(fx)) * 3; out[0] = vals[g]; out[1] = vals[g + 1]; out[2] = vals[g + 2]; return; }
      const tx = fx - ix, ty = fy - iy;
      for (let c = 0; c < 3; c++) {
        const a = vals[(iy * gw + ix) * 3 + c], b = vals[(iy * gw + ix + 1) * 3 + c];
        const d = vals[((iy + 1) * gw + ix) * 3 + c], e = vals[((iy + 1) * gw + ix + 1) * 3 + c];
        out[c] = (a * (1 - tx) + b * tx) * (1 - ty) + (d * (1 - tx) + e * tx) * ty;
      }
    };
  }
  // src：canvas か ImageBitmap、rect：使う範囲（画素。null なら全体）、strength：1〜10（大きいほど白くする）
  // color：'orig'（元の色）| 'red' | 'black' | 'blue'
  function fromImage(src, { rect = null, strength = 5, color = 'orig', maxSide = 1200 } = {}) {
    const r = rect || { x: 0, y: 0, w: src.width, h: src.height };
    const s = Math.min(1, maxSide / Math.max(r.w, r.h));
    const W = Math.max(1, Math.round(r.w * s)), H = Math.max(1, Math.round(r.h * s));
    const c = newCanvas(W, H), x = c.getContext('2d', { willReadFrequently: true });
    x.imageSmoothingQuality = 'high';
    x.drawImage(src, r.x, r.y, r.w, r.h, 0, 0, W, H);
    const img = x.getImageData(0, 0, W, H), p = img.data, paper = paperGrid(p, W, H), bg = [0, 0, 0];
    const t0 = 0.04 + strength * 0.025, t1 = t0 + 0.25, ink = INKS[color] ? hexRgb(INKS[color]) : null;
    const clamp01 = v => (v < 0 ? 0 : v > 1 ? 1 : v);
    // 1回目：紙の色で割って「紙 = 1」にそろえ、インクの濃さ a（いちばん減った色で見る。赤いインクなら緑と青が減る）を求める
    const o = new Float32Array(W * H * 3), hist = new Uint32Array(256);
    let strong = 0;
    for (let y = 0, j = 0, q = 0; y < H; y++) for (let xx = 0; xx < W; xx++, j += 4, q += 3) {
      paper(xx, y, bg);
      o[q] = Math.min(1, p[j] / bg[0]); o[q + 1] = Math.min(1, p[j + 1] / bg[1]); o[q + 2] = Math.min(1, p[j + 2] / bg[2]);
      const a = 1 - Math.min(o[q], o[q + 1], o[q + 2]);
      if (a > t1) { hist[Math.round(a * 255)]++; strong++; }
    }
    // しっかりインクのある所の濃さ（まん中の値）。線のふちの薄い所は、この濃さのインクが少しだけ乗っているとみなす
    let aInk = t1;
    for (let k = 0, acc = 0; k < 256 && strong; k++) { acc += hist[k]; if (acc >= strong / 2) { aInk = Math.max(t1, k / 255); break; } }
    // 2回目：透明度と色
    for (let i = 0, j = 0, q = 0; i < W * H; i++, j += 4, q += 3) {
      const a = 1 - Math.min(o[q], o[q + 1], o[q + 2]), al = clamp01((a - t0) / (t1 - t0));
      p[j + 3] = Math.round(al * 255);
      if (ink) { p[j] = ink[0]; p[j + 1] = ink[1]; p[j + 2] = ink[2]; }
      else if (al > 0) { // インクそのものの色にもどす（紙の白さを差し引く）
        const s = Math.min(1, a / aInk);
        p[j] = Math.round(clamp01(1 - (1 - o[q]) / s) * 255);
        p[j + 1] = Math.round(clamp01(1 - (1 - o[q + 1]) / s) * 255);
        p[j + 2] = Math.round(clamp01(1 - (1 - o[q + 2]) / s) * 255);
      }
    }
    x.putImageData(img, 0, 0);
    return trim(c, Math.round(Math.max(W, H) * 0.02));
  }

  /* ---------- 手書きの線 ---------- */
  // strokes：[[x, y, x, y, …], …]（書いた枠の座標）。点と点の間はなめらかな曲線でつなぐ
  function drawStrokes(x, strokes, color, lw) {
    x.strokeStyle = x.fillStyle = color; x.lineWidth = lw; x.lineCap = x.lineJoin = 'round';
    for (const s of strokes) {
      const n = s.length / 2;
      if (!n) continue;
      if (n === 1) { x.beginPath(); x.arc(s[0], s[1], lw / 2, 0, Math.PI * 2); x.fill(); continue; }
      x.beginPath(); x.moveTo(s[0], s[1]);
      for (let i = 1; i < n - 1; i++) x.quadraticCurveTo(s[2 * i], s[2 * i + 1], (s[2 * i] + s[2 * i + 2]) / 2, (s[2 * i + 1] + s[2 * i + 3]) / 2);
      x.lineTo(s[2 * n - 2], s[2 * n - 1]); x.stroke();
    }
  }
  function strokesToCanvas(strokes, { width, height, scale = 1, color = INKS.black, lw = 3 }) {
    const c = newCanvas(Math.max(1, Math.round(width * scale)), Math.max(1, Math.round(height * scale))), x = c.getContext('2d');
    x.scale(scale, scale);
    drawStrokes(x, strokes, color, lw);
    return c;
  }

  /* ---------- 余白を切る ---------- */
  // 透明でない所を囲む大きさに切る（ぽつんとある1〜2画素の点は無視）。何もなければ null
  function trim(c, pad = 0) {
    const W = c.width, H = c.height, d = c.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, W, H).data;
    const cols = new Uint32Array(W), rows = new Uint32Array(H);
    for (let y = 0, j = 3; y < H; y++) for (let x = 0; x < W; x++, j += 4) if (d[j] > 24) { cols[x]++; rows[y]++; }
    const need = a => Math.max(1, Math.round(a.length * 0.002));
    const first = (a, k) => { for (let i = 0; i < a.length; i++) if (a[i] > k) return i; return -1; };
    const last = (a, k) => { for (let i = a.length - 1; i >= 0; i--) if (a[i] > k) return i; return -1; };
    const kc = need(rows) - 1, kr = need(cols) - 1;
    let x0 = first(cols, kc), x1 = last(cols, kc), y0 = first(rows, kr), y1 = last(rows, kr);
    if (x0 < 0 || y0 < 0) return null;
    x0 = Math.max(0, x0 - pad); y0 = Math.max(0, y0 - pad); x1 = Math.min(W - 1, x1 + pad); y1 = Math.min(H - 1, y1 + pad);
    const o = newCanvas(x1 - x0 + 1, y1 - y0 + 1);
    o.getContext('2d').drawImage(c, -x0, -y0);
    return o;
  }

  // 保存用の PNG と、保存した PNG から canvas に戻す
  const toPng = c => new Blob([encodeCanvas(c, 'image/png')], { type: 'image/png' });
  async function fromBlob(blob) {
    const bmp = await createImageBitmap(blob), c = newCanvas(bmp.width, bmp.height);
    c.getContext('2d').drawImage(bmp, 0, 0); bmp.close();
    return c;
  }

  return { SEAL_COLOR, INKS, SEAL_FONTS, sealFonts, seal, fromImage, drawStrokes, strokesToCanvas, trim, toPng, fromBlob };
})();

// 作ったハンコ・サインを、この端末のこのブラウザの中にしまっておく（IndexedDB）
//   { id, kind: 'seal' | 'sign' | 'image', label, png: Blob, w, h（置くときの大きさ mm）, created }
const StampStore = (() => {
  let dbp = null;
  const db = () => dbp || (dbp = new Promise((res, rej) => {
    const r = indexedDB.open('file-tools-stamps', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('stamps', { keyPath: 'id', autoIncrement: true });
    r.onsuccess = () => res(r.result);
    r.onerror = () => { dbp = null; rej(r.error); };
  }));
  async function tx(mode, fn) {
    const d = await db();
    return new Promise((res, rej) => {
      const t = d.transaction('stamps', mode), req = fn(t.objectStore('stamps'));
      t.oncomplete = () => res(req.result);
      t.onerror = t.onabort = () => rej(t.error);
    });
  }
  return {
    list: () => tx('readonly', s => s.getAll()),
    add: rec => tx('readwrite', s => s.add(rec)),
    remove: id => tx('readwrite', s => s.delete(id)),
  };
})();
