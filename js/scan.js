'use strict';
// =========================================================
// スキャン補正の計算部分（画面とは関係ない部分。Node でもテストできる）
//   detectDocument … 写真の中から書類の四隅を見つける
//   paperAspect    … 四隅の形から、元の紙の縦横比を推定する
//   warpQuad       … 四隅を長方形に引き伸ばす（台形補正）
//   enhance        … 影を消す・線をくっきり・白黒にする
// 画像は { width, height, data: Uint8ClampedArray（RGBA） } の形でやりとりする。
// 四隅は [左上, 右上, 右下, 左下] の順で、それぞれ [x, y]（ピクセル）。
// =========================================================
const Scan = (() => {
  const DEG = Math.PI / 180;
  const clamp = (v, a, b) => v < a ? a : v > b ? b : v;

  /* ---------- ぼかし ---------- */
  function gaussBlur(src, w, h, sigma) {
    const r = Math.max(1, Math.ceil(sigma * 2.5)), k = new Float32Array(2 * r + 1);
    let s = 0;
    for (let i = -r; i <= r; i++) s += k[i + r] = Math.exp(-i * i / (2 * sigma * sigma));
    for (let i = 0; i < k.length; i++) k[i] /= s;
    const tmp = new Float32Array(w * h), out = new Float32Array(w * h);
    for (let y = 0; y < h; y++) {
      const o = y * w;
      for (let x = 0; x < w; x++) {
        let a = 0;
        for (let i = -r; i <= r; i++) a += src[o + clamp(x + i, 0, w - 1)] * k[i + r];
        tmp[o + x] = a;
      }
    }
    for (let y = 0; y < h; y++) {
      for (let x = 0; x < w; x++) {
        let a = 0;
        for (let i = -r; i <= r; i++) a += tmp[clamp(y + i, 0, h - 1) * w + x] * k[i + r];
        out[y * w + x] = a;
      }
    }
    return out;
  }
  // 小さな格子の最大値・最小値フィルタ（半径 r マス）
  function rankFilter(src, w, h, r, isMax) {
    const tmp = new Float32Array(w * h), out = new Float32Array(w * h);
    const pick = isMax ? Math.max : Math.min;
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      let v = src[y * w + x];
      for (let i = Math.max(0, x - r); i <= Math.min(w - 1, x + r); i++) v = pick(v, src[y * w + i]);
      tmp[y * w + x] = v;
    }
    for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
      let v = tmp[y * w + x];
      for (let j = Math.max(0, y - r); j <= Math.min(h - 1, y + r); j++) v = pick(v, tmp[j * w + x]);
      out[y * w + x] = v;
    }
    return out;
  }

  /* ---------- 四隅の並べ替え・面積 ---------- */
  function orderQuad(pts) {
    const cx = pts.reduce((s, p) => s + p[0], 0) / 4, cy = pts.reduce((s, p) => s + p[1], 0) / 4;
    const q = [...pts].sort((a, b) => Math.atan2(a[1] - cy, a[0] - cx) - Math.atan2(b[1] - cy, b[0] - cx));
    let k = 0;
    for (let i = 1; i < 4; i++) if (q[i][0] + q[i][1] < q[k][0] + q[k][1]) k = i;
    return [0, 1, 2, 3].map(i => q[(i + k) % 4]);
  }
  function quadArea(q) {
    let a = 0;
    for (let i = 0; i < 4; i++) { const p = q[i], n = q[(i + 1) % 4]; a += p[0] * n[1] - n[0] * p[1]; }
    return Math.abs(a) / 2;
  }

  /* ---------- 書類の四隅を見つける ---------- */
  // 小さめ（長辺 500px 前後）の画像を渡す。見つからなければ found: false で画像全体を返す
  function detectDocument(img, debug) {
    const W = img.width, H = img.height, d = img.data, N = W * H;
    const whole = { quad: [[0, 0], [W, 0], [W, H], [0, H]], found: false, confidence: 0 };
    if (W < 16 || H < 16) return whole;

    // 1) ぼかしてから、色の境目（勾配）を R・G・B のうち一番強いもので求める
    const sigma = Math.max(1, Math.max(W, H) / 320);
    const chans = [0, 1, 2].map(c => {
      const a = new Float32Array(N);
      for (let i = 0; i < N; i++) a[i] = d[i * 4 + c];
      return gaussBlur(a, W, H, sigma);
    });
    const gx = new Float32Array(N), gy = new Float32Array(N), mag = new Float32Array(N);
    for (let y = 1; y < H - 1; y++) for (let x = 1; x < W - 1; x++) {
      const i = y * W + x;
      let best = -1, bx = 0, by = 0;
      for (let c = 0; c < 3; c++) {
        const a = chans[c];
        const sx = (a[i - W + 1] + 2 * a[i + 1] + a[i + W + 1]) - (a[i - W - 1] + 2 * a[i - 1] + a[i + W - 1]);
        const sy = (a[i + W - 1] + 2 * a[i + W] + a[i + W + 1]) - (a[i - W - 1] + 2 * a[i - W] + a[i - W + 1]);
        const m = sx * sx + sy * sy;
        if (m > best) { best = m; bx = sx; by = sy; }
      }
      gx[i] = bx / 4; gy[i] = by / 4; mag[i] = Math.sqrt(best) / 4;
    }

    // 2) 境目を細い線にする（勾配の向きで一番強い点だけ残す）
    const T_EDGE = 4, ex = [], ey = [];
    for (let y = 2; y < H - 2; y++) for (let x = 2; x < W - 2; x++) {
      const i = y * W + x, m = mag[i];
      if (m < T_EDGE) continue;
      const ax = Math.abs(gx[i]), ay = Math.abs(gy[i]);
      let n1, n2;
      if (ax > 2.414 * ay) { n1 = i - 1; n2 = i + 1; }
      else if (ay > 2.414 * ax) { n1 = i - W; n2 = i + W; }
      else if (gx[i] * gy[i] > 0) { n1 = i - W - 1; n2 = i + W + 1; }
      else { n1 = i - W + 1; n2 = i + W - 1; }
      if (m < mag[n1] || m <= mag[n2]) continue;
      ex.push(x); ey.push(y);
    }

    // 3) ハフ変換：まっすぐな線の候補を投票で探す（x cosθ + y sinθ = ρ、画像の中心が原点）。
    //    θ は 0〜360° で「どちら側が明るいか」も区別する。紙のふちは明るい側がそろっているので票が集まり、
    //    市松模様のように明暗が入れかわる線は票が割れる
    const NT = 360, cx = W / 2, cy = H / 2, R = Math.ceil(Math.hypot(W, H) / 2) + 3, NR = 2 * R + 1;
    const cosT = new Float32Array(NT), sinT = new Float32Array(NT);
    for (let t = 0; t < NT; t++) { cosT[t] = Math.cos(t * DEG); sinT[t] = Math.sin(t * DEG); }
    const acc = new Float32Array(NT * NR);
    for (let k = 0; k < ex.length; k++) {
      const x = ex[k] - cx, y = ey[k] - cy, i = ey[k] * W + ex[k];
      let t0 = Math.round(Math.atan2(gy[i], gx[i]) / DEG);
      t0 = ((t0 % NT) + NT) % NT;
      const w = Math.min(mag[i], 25);
      for (let dt = -3; dt <= 3; dt++) {
        const t = (t0 + dt + NT) % NT;
        acc[t * NR + Math.round(x * cosT[t] + y * sinT[t]) + R] += w * (1 - Math.abs(dt) * 0.2);
      }
    }
    const at = (t, r) => r < 0 || r >= NR ? 0 : acc[((t + NT) % NT) * NR + r];
    const minVotes = Math.min(W, H) * 0.06 * 5;
    const peaks = [];
    for (let t = 0; t < NT; t++) for (let r = 0; r < NR; r++) {
      const v = acc[t * NR + r];
      if (v < minVotes) continue;
      let isMax = true;
      for (let dt = -2; dt <= 2 && isMax; dt++) for (let dr = -3; dr <= 3; dr++) {
        if (!dt && !dr) continue;
        const u = at(t + dt, r + dr);
        if (u > v || (u === v && (dt < 0 || (dt === 0 && dr < 0)))) { isMax = false; break; }
      }
      if (isMax) peaks.push({ t: t * DEG, rho: r - R, v });
    }
    peaks.sort((a, b) => b.v - a.v);
    const same = (a, b) => { // 同じ向き・ほぼ同じ位置の線（細かい模様の線に、近くの紙のふちが消されないよう狭めに）
      const dt = Math.abs(a.t - b.t) % (2 * Math.PI);
      return Math.min(dt, 2 * Math.PI - dt) < 3 * DEG && Math.abs(a.rho - b.rho) < 3.5;
    };
    const cand = [];
    for (const p of peaks) { if (cand.length >= 120) break; if (!cand.some(q => same(p, q))) cand.push(p); }

    // 4) 線の位置を、近くの境目の点から細かく合わせ直す
    function refineLine(p) {
      const nx = Math.cos(p.t), ny = Math.sin(p.t), cosLim = Math.cos(12 * DEG);
      let sw = 0, sx = 0, sy = 0;
      const px = [], py = [], pw = [];
      for (let k = 0; k < ex.length; k++) {
        const x = ex[k] - cx, y = ey[k] - cy;
        if (Math.abs(x * nx + y * ny - p.rho) > 2.5) continue;
        const i = ey[k] * W + ex[k];
        if (gx[i] * nx + gy[i] * ny < cosLim * mag[i]) continue;
        const w = Math.min(mag[i], 25);
        px.push(x); py.push(y); pw.push(w); sw += w; sx += w * x; sy += w * y;
      }
      if (px.length < 10) return mkLine(p.t, p.rho);
      const mx = sx / sw, my = sy / sw;
      let cxx = 0, cxy = 0, cyy = 0;
      for (let k = 0; k < px.length; k++) {
        const dx = px[k] - mx, dy = py[k] - my;
        cxx += pw[k] * dx * dx; cxy += pw[k] * dx * dy; cyy += pw[k] * dy * dy;
      }
      // 散らばりが一番小さい向き＝線の法線（明るい側を向くように元の向きに合わせる）
      let t = 0.5 * Math.atan2(2 * cxy, cxx - cyy) + Math.PI / 2;
      if (Math.cos(t - p.t) < 0) t += Math.PI;
      if (Math.cos(t - p.t) < Math.cos(8 * DEG)) return mkLine(p.t, p.rho); // 大きくずれたら元のまま
      return mkLine(t, mx * Math.cos(t) + my * Math.sin(t));
    }
    function mkLine(t, rho, border = false) {
      const nx = Math.cos(t), ny = Math.sin(t);
      return { t, rho, nx, ny, border, ...(border ? {} : supportProfile(nx, ny, rho)) };
    }
    // 線の両側の明るさの差。+ なら法線の向きの側が明るい。
    // 紙と机の境目は「段差」になるが、表の罫線や文字のような細い線は両側とも紙なので差が出ない。
    // 線の位置が細い線のふちに少しずれていても間違えないよう、近く（±K）と遠く（±2K）の両方で段差を確かめる
    const K = Math.max(3, Math.round(Math.max(W, H) / 120)), T_STEP = 10;
    function diffAt(x, y, nx, ny, k) {
      const ax = Math.round(x + k * nx), ay = Math.round(y + k * ny), bx = Math.round(x - k * nx), by = Math.round(y - k * ny);
      if (ax < 0 || ay < 0 || bx < 0 || by < 0 || ax >= W || bx >= W || ay >= H || by >= H) return 0;
      let best = 0;
      for (let c = 0; c < 3; c++) { const v = chans[c][ay * W + ax] - chans[c][by * W + bx]; if (Math.abs(v) > Math.abs(best)) best = v; }
      return best;
    }
    function stepAt(x, y, nx, ny) {
      const a = diffAt(x, y, nx, ny, K), b = diffAt(x, y, nx, ny, 2 * K);
      return a > T_STEP && b > T_STEP ? 1 : a < -T_STEP && b < -T_STEP ? -1 : 0;
    }
    // 線にそって1ピクセルごとに「本当に境目（段差）があるか」を数えた累積和（四角形の辺の評価を速くする）。
    // 明るい側が法線の＋側か−側かで分けて数える（紙の4辺は、どれも「内側が明るい」などそろっているはず）
    function supportProfile(nx, ny, rho) {
      const len = 2 * R + 1, pos = new Float32Array(len + 1), neg = new Float32Array(len + 1), cosLim = Math.cos(18 * DEG);
      for (let k = 0; k < len; k++) {
        const s = k - R, bx = nx * rho - ny * s + cx, by = ny * rho + nx * s + cy;
        let sup = 0;
        if (bx >= 1 && by >= 1 && bx < W - 1 && by < H - 1) {
          for (let o = -2; o <= 2 && !sup; o++) {
            const qx = Math.round(bx + o * nx), qy = Math.round(by + o * ny);
            if (qx < 1 || qy < 1 || qx >= W - 1 || qy >= H - 1) continue;
            const i = qy * W + qx, m = mag[i];
            if (m >= T_EDGE * 1.5 && Math.abs(gx[i] * nx + gy[i] * ny) >= cosLim * m) sup = 1;
          }
        }
        let p = 0, q = 0;
        if (sup) { const st = stepAt(bx, by, nx, ny); if (st > 0) p = 1; else if (st < 0) q = 1; }
        pos[k + 1] = pos[k] + p; neg[k + 1] = neg[k] + q;
      }
      return { pos, neg };
    }
    // 「段差」になっている長さが長い線から順に使う（文字や罫線のような細い線はここで落ちる）
    const lines = [];
    const strength = L => Math.max(L.pos[L.pos.length - 1], L.neg[L.neg.length - 1]);
    for (const L of cand.map(refineLine).sort((a, b) => strength(b) - strength(a))) {
      if (lines.length >= 30 || strength(L) < Math.min(W, H) * 0.08) break;
      if (!lines.some(q => same(L, q))) lines.push(L);
    }
    // 画像の端も「辺の候補」に入れる（書類がはみ出して写っているとき用。点数は低め）
    // （ここまでの座標は「ピクセルの番号」。番号 i のピクセルは i〜i+1 の範囲なので、画像の端は -0.5 と W-0.5）
    lines.push(mkLine(90 * DEG, -0.5 - cy, true), mkLine(90 * DEG, H - 0.5 - cy, true), mkLine(0, -0.5 - cx, true), mkLine(0, W - 0.5 - cx, true));

    // 5) 4本の線の組み合わせで四角形を作り、辺にそった境目の多さで点数をつける
    const angDiff = (a, b) => { const x = Math.abs(a - b) % Math.PI; return Math.min(x, Math.PI - x); };
    const cross = (A, B) => {
      const det = A.nx * B.ny - A.ny * B.nx;
      if (Math.abs(det) < 1e-6) return null;
      return [(A.rho * B.ny - A.ny * B.rho) / det, (A.nx * B.rho - A.rho * B.nx) / det];
    };
    // 辺 P→Q の長さと、そのうち「内側が明るい段差」「内側が暗い段差」になっている長さ。C は四角形の中心
    const sideSupport = (L, P, Q, C) => {
      const s1 = (P[0] - L.nx * L.rho) * -L.ny + (P[1] - L.ny * L.rho) * L.nx;
      const s2 = (Q[0] - L.nx * L.rho) * -L.ny + (Q[1] - L.ny * L.rho) * L.nx;
      const len = Math.abs(s2 - s1);
      if (L.border) return { bright: len * 0.3, dark: len * 0.3, len };
      const k1 = clamp(Math.round(Math.min(s1, s2)) + R, 0, 2 * R + 1), k2 = clamp(Math.round(Math.max(s1, s2)) + R, 0, 2 * R + 1);
      const p = L.pos[k2] - L.pos[k1], n = L.neg[k2] - L.neg[k1];
      const insidePlus = C[0] * L.nx + C[1] * L.ny > L.rho;
      return insidePlus ? { bright: p, dark: n, len } : { bright: n, dark: p, len };
    };
    const minSep = Math.min(W, H) * 0.12, pairs = [];
    for (let i = 0; i < lines.length; i++) for (let j = i + 1; j < lines.length; j++) {
      const A = lines[i], B = lines[j];
      // 向かい合う2辺の候補。強く斜めから撮ると向かいの辺どうしも50°近く傾くので、ゆるめに
      if (angDiff(A.t, B.t) > 55 * DEG) continue;
      const sgn = Math.cos(A.t - B.t) > 0 ? 1 : -1;
      if (Math.abs(A.rho - sgn * B.rho) < minSep) continue;
      const dir = 0.5 * Math.atan2(Math.sin(2 * A.t) + Math.sin(2 * B.t), Math.cos(2 * A.t) + Math.cos(2 * B.t)); // 2本の平均の向き
      pairs.push([A, B, dir]);
    }
    const minArea = W * H * 0.1, lim = 1.06;
    let best = null;
    for (let a = 0; a < pairs.length; a++) for (let b = a + 1; b < pairs.length; b++) {
      const [A1, A2, da] = pairs[a], [B1, B2, db] = pairs[b];
      if (A1 === B1 || A1 === B2 || A2 === B1 || A2 === B2) continue;
      if (angDiff(da, db) < 40 * DEG) continue;
      const nb = A1.border + A2.border + B1.border + B2.border;
      if (nb >= 3) continue;
      const P = [cross(A1, B1), cross(A1, B2), cross(A2, B2), cross(A2, B1)];
      if (P.some(p => !p || Math.abs(p[0]) > cx * lim || Math.abs(p[1]) > cy * lim)) continue;
      // 凸の四角形で、角度がつぶれすぎていないこと
      let sign = 0, okShape = true;
      for (let k = 0; k < 4 && okShape; k++) {
        const p0 = P[(k + 3) % 4], p1 = P[k], p2 = P[(k + 1) % 4];
        const ux = p0[0] - p1[0], uy = p0[1] - p1[1], vx = p2[0] - p1[0], vy = p2[1] - p1[1];
        const cr = ux * vy - uy * vx, ang = Math.acos(clamp((ux * vx + uy * vy) / (Math.hypot(ux, uy) * Math.hypot(vx, vy) || 1), -1, 1));
        if (!cr || (sign && Math.sign(cr) !== sign) || ang < 30 * DEG || ang > 150 * DEG) okShape = false;
        sign = Math.sign(cr);
      }
      if (!okShape || quadArea(P) < minArea) continue;
      const C = [(P[0][0] + P[1][0] + P[2][0] + P[3][0]) / 4, (P[0][1] + P[1][1] + P[2][1] + P[3][1]) / 4];
      const sides = [[A1, P[0], P[1]], [B2, P[1], P[2]], [A2, P[2], P[3]], [B1, P[3], P[0]]].map(([L, p, q]) => ({ L, ...sideSupport(L, p, q, C) }));
      // 4辺とも同じ向きの段差（ふつうは「紙のほうが明るい」）になっているほうで数える
      for (const key of ['bright', 'dark']) {
        let score = 0, fsum = 0, nReal = 0, bad = false;
        for (const sd of sides) {
          if (!sd.L.border) { const f = sd[key] / (sd.len || 1); if (f < 0.3) { bad = true; break; } fsum += f; nReal++; }
          score += sd[key];
        }
        if (bad) continue;
        if (!best || score > best.score) best = { score, P, confidence: fsum / nReal, nb, key, fr: sides.map(sd => +(sd[key] / sd.len).toFixed(2)) };
      }
    }
    if (debug) debug.lines = lines.map(L => ({ t: +(L.t / DEG).toFixed(1), rho: +L.rho.toFixed(1), s: L.border ? -1 : strength(L), a: [+(-L.ny * 0 + L.nx * L.rho + cx).toFixed(0), +(L.ny * L.rho + cy).toFixed(0)] }));
    if (!best || best.confidence < 0.6) return whole; // 自信がないときは画像全体（手で合わせてもらう）
    const quad = orderQuad(best.P.map(p => [clamp(p[0] + cx + 0.5, 0, W), clamp(p[1] + cy + 0.5, 0, H)]));
    return { quad, found: true, confidence: best.confidence, borders: best.nb, detail: best.key + ' ' + best.fr.join(',') };
  }

  /* ---------- 紙の縦横比の推定 ---------- */
  // 斜めから撮った長方形の四隅から、カメラの焦点距離を推定して本来の縦横比（幅÷高さ）を求める
  // （Zhang & He「Whiteboard scanning and image enhancement」の方法）。うまくいかなければ辺の長さの比
  function paperAspect(q, W, H) {
    const [tl, tr, br, bl] = q, dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
    const simple = (dist(tl, tr) + dist(bl, br)) / ((dist(tl, bl) + dist(tr, br)) || 1);
    const m1 = [tl[0], tl[1], 1], m2 = [tr[0], tr[1], 1], m3 = [bl[0], bl[1], 1], m4 = [br[0], br[1], 1];
    const cr = (a, b) => [a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0]];
    const dot = (a, b) => a[0] * b[0] + a[1] * b[1] + a[2] * b[2];
    const k2 = dot(cr(m1, m4), m3) / dot(cr(m2, m4), m3), k3 = dot(cr(m1, m4), m2) / dot(cr(m3, m4), m2);
    const n2 = m2.map((v, i) => k2 * v - m1[i]), n3 = m3.map((v, i) => k3 * v - m1[i]);
    const u0 = W / 2, v0 = H / 2, den = n2[2] * n3[2];
    if (!isFinite(k2) || !isFinite(k3)) return simple;
    if (Math.abs(den) < 1e-9) return Math.sqrt((n2[0] ** 2 + n2[1] ** 2) / (n3[0] ** 2 + n3[1] ** 2)) || simple;
    const f2 = -((n2[0] * n3[0] - (n2[0] * n3[2] + n2[2] * n3[0]) * u0 + den * u0 * u0) +
                 (n2[1] * n3[1] - (n2[1] * n3[2] + n2[2] * n3[1]) * v0 + den * v0 * v0)) / den;
    const L = Math.max(W, H);
    if (!(f2 > 0)) return simple;
    const f = Math.sqrt(f2);
    if (f < 0.35 * L || f > 6 * L) return simple; // ほぼ真上から撮ったときは推定がぶれるので辺の比を使う
    const a = v => [(v[0] - u0 * v[2]) / f, (v[1] - v0 * v[2]) / f, v[2]];
    const p2 = a(n2), p3 = a(n3), r = Math.sqrt(dot(p2, p2) / dot(p3, p3));
    return isFinite(r) && r > 0.1 && r < 10 ? r : simple;
  }
  // 出来上がりの大きさ。写真の解像度を落とさない大きさにする
  function outputSize(q, aspect) {
    const [tl, tr, br, bl] = q, dist = (a, b) => Math.hypot(a[0] - b[0], a[1] - b[1]);
    const w = Math.max(dist(tl, tr), dist(bl, br)), h = Math.max(dist(tl, bl), dist(tr, br));
    const H = Math.max(h, w / aspect);
    return [Math.max(1, Math.round(H * aspect)), Math.max(1, Math.round(H))];
  }

  /* ---------- 台形補正（四隅 → 長方形） ---------- */
  // inset: 端を少しだけ内側から取る（割合）。書類のふちの外の机などが細く残るのを防ぐ
  function warpQuad(src, q, ow, oh, inset = 0) {
    const [[x0, y0], [x1, y1], [x2, y2], [x3, y3]] = q;
    let a, b, c, d, e, f, g, h;
    const sx = x0 - x1 + x2 - x3, sy = y0 - y1 + y2 - y3;
    if (Math.abs(sx) < 1e-9 && Math.abs(sy) < 1e-9) {
      a = x1 - x0; b = x3 - x0; c = x0; d = y1 - y0; e = y3 - y0; f = y0; g = 0; h = 0;
    } else {
      const dx1 = x1 - x2, dx2 = x3 - x2, dy1 = y1 - y2, dy2 = y3 - y2, det = dx1 * dy2 - dx2 * dy1;
      g = (sx * dy2 - dx2 * sy) / det; h = (dx1 * sy - sx * dy1) / det;
      a = x1 - x0 + g * x1; b = x3 - x0 + h * x3; c = x0;
      d = y1 - y0 + g * y1; e = y3 - y0 + h * y3; f = y0;
    }
    const sw = src.width, sh = src.height, s = src.data, out = new Uint8ClampedArray(ow * oh * 4), span = 1 - 2 * inset;
    for (let oy = 0; oy < oh; oy++) {
      const v = inset + span * (oy + 0.5) / oh;
      let o = oy * ow * 4;
      for (let ox = 0; ox < ow; ox++, o += 4) {
        const u = inset + span * (ox + 0.5) / ow, den = g * u + h * v + 1;
        let x = (a * u + b * v + c) / den - 0.5, y = (d * u + e * v + f) / den - 0.5;
        x = x < 0 ? 0 : x > sw - 1 ? sw - 1 : x;
        y = y < 0 ? 0 : y > sh - 1 ? sh - 1 : y;
        const xi = x | 0, yi = y | 0, fx = x - xi, fy = y - yi;
        const xj = xi + 1 < sw ? 1 : 0, yj = yi + 1 < sh ? sw : 0;
        const p = (yi * sw + xi) * 4, p10 = p + xj * 4, p01 = p + yj * 4, p11 = p01 + xj * 4;
        const w00 = (1 - fx) * (1 - fy), w10 = fx * (1 - fy), w01 = (1 - fx) * fy, w11 = fx * fy;
        out[o] = s[p] * w00 + s[p10] * w10 + s[p01] * w01 + s[p11] * w11;
        out[o + 1] = s[p + 1] * w00 + s[p10 + 1] * w10 + s[p01 + 1] * w01 + s[p11 + 1] * w11;
        out[o + 2] = s[p + 2] * w00 + s[p10 + 2] * w10 + s[p01 + 2] * w01 + s[p11 + 2] * w11;
        out[o + 3] = 255;
      }
    }
    return { width: ow, height: oh, data: out };
  }

  /* ---------- 影消し・線くっきり ---------- */
  // 紙の明るさ（背景）を粗いマス目で推定して割り算し、紙が 1.0 になるようにそろえる（影や色かぶりが消える）
  function normalizeByBackground(plane, W, H) {
    const cell = Math.max(4, Math.round(Math.max(W, H) / 120));
    const gw = Math.ceil(W / cell), gh = Math.ceil(H / cell), G = gw * gh;
    const colCell = new Int32Array(W);
    for (let x = 0; x < W; x++) colCell[x] = (x / cell) | 0;
    const sum = new Float64Array(G), cnt = new Float64Array(G);
    for (let y = 0; y < H; y++) {
      const row = ((y / cell) | 0) * gw, o = y * W;
      for (let x = 0; x < W; x++) { const g = row + colCell[x]; sum[g] += plane[o + x]; cnt[g]++; }
    }
    const mean = new Float32Array(G);
    for (let g = 0; g < G; g++) mean[g] = sum[g] / cnt[g];
    // マスの中の「平均より明るい部分」の平均＝文字を除いた紙の明るさ
    sum.fill(0); cnt.fill(0);
    for (let y = 0; y < H; y++) {
      const row = ((y / cell) | 0) * gw, o = y * W;
      for (let x = 0; x < W; x++) { const g = row + colCell[x], v = plane[o + x]; if (v >= mean[g]) { sum[g] += v; cnt[g]++; } }
    }
    let bg = new Float32Array(G);
    for (let g = 0; g < G; g++) bg[g] = cnt[g] ? sum[g] / cnt[g] : mean[g];
    // 太い文字や線も消えるように、明るいほうへ広げてから戻す（クロージング）。最後になめらかに
    bg = rankFilter(rankFilter(bg, gw, gh, 2, true), gw, gh, 2, false);
    bg = gaussBlur(bg, gw, gh, 1);
    for (let g = 0; g < G; g++) bg[g] = Math.max(bg[g], 24);
    // マスの中心どうしの間をなめらかにつないで（双線形補間）割り算
    const x0 = new Int32Array(W), x1 = new Int32Array(W), wx = new Float32Array(W);
    for (let x = 0; x < W; x++) {
      const fx = clamp((x + 0.5) / cell - 0.5, 0, gw - 1), i = Math.min(gw - 2, fx | 0);
      x0[x] = Math.max(0, i); x1[x] = Math.min(gw - 1, i + 1); wx[x] = gw > 1 ? fx - x0[x] : 0;
    }
    for (let y = 0; y < H; y++) {
      const fy = clamp((y + 0.5) / cell - 0.5, 0, gh - 1), j = Math.max(0, Math.min(gh - 2, fy | 0));
      const j1 = Math.min(gh - 1, j + 1), wy = gh > 1 ? fy - j : 0, r0 = j * gw, r1 = j1 * gw, o = y * W;
      for (let x = 0; x < W; x++) {
        const top = bg[r0 + x0[x]] * (1 - wx[x]) + bg[r0 + x1[x]] * wx[x];
        const bot = bg[r1 + x0[x]] * (1 - wx[x]) + bg[r1 + x1[x]] * wx[x];
        plane[o + x] /= top * (1 - wy) + bot * wy;
      }
    }
  }
  // アンシャープマスク（輪郭を強調）。ぼかした画像との差を足す
  function sharpen(plane, W, H, r, amount) {
    const n = 2 * r + 1, tmp = new Float32Array(W * H);
    for (let y = 0; y < H; y++) {
      const o = y * W;
      let a = 0;
      for (let i = -r; i <= r; i++) a += plane[o + clamp(i, 0, W - 1)];
      for (let x = 0; x < W; x++) {
        tmp[o + x] = a / n;
        a += plane[o + Math.min(W - 1, x + r + 1)] - plane[o + Math.max(0, x - r)];
      }
    }
    const col = new Float32Array(W);
    for (let i = -r; i <= r; i++) { const o = clamp(i, 0, H - 1) * W; for (let x = 0; x < W; x++) col[x] += tmp[o + x]; }
    for (let y = 0; y < H; y++) {
      const o = y * W, add = Math.min(H - 1, y + r + 1) * W, sub = Math.max(0, y - r) * W;
      for (let x = 0; x < W; x++) {
        const v = plane[o + x];
        plane[o + x] = v + amount * (v - col[x] / n);
        col[x] += tmp[add + x] - tmp[sub + x];
      }
    }
  }
  // mode: 'color' | 'gray' | 'bw' | 'none'、strength: 0〜1（線の濃さ）。img を書き換えて返す
  function enhance(img, mode, strength = 0.5) {
    if (mode === 'none') return img;
    const W = img.width, H = img.height, N = W * H, d = img.data, s = clamp(strength, 0, 1);
    const plane = new Float32Array(N), sharpR = Math.max(1, Math.round(Math.max(W, H) / 1600));
    // 明るさ（紙=1.0）→ 0〜255 の対応表
    const LUT_N = 1536, lut = new Uint8ClampedArray(LUT_N + 1);
    if (mode === 'bw') {
      const T = 0.66 + 0.2 * s;
      for (let k = 0; k <= LUT_N; k++) lut[k] = clamp((k / 1024 - T) / 0.08 + 0.5, 0, 1) * 255;
    } else {
      const wp = 0.9 - 0.06 * s, bp = 0.12 + 0.24 * s, gamma = 1 + s;
      for (let k = 0; k <= LUT_N; k++) lut[k] = Math.pow(clamp((k / 1024 - bp) / (wp - bp), 0, 1), gamma) * 255;
    }
    const toLut = v => { const k = v * 1024; return lut[k <= 0 ? 0 : k >= LUT_N ? LUT_N : k | 0]; };
    if (mode !== 'color') {
      for (let i = 0, j = 0; i < N; i++, j += 4) plane[i] = 0.299 * d[j] + 0.587 * d[j + 1] + 0.114 * d[j + 2];
      normalizeByBackground(plane, W, H);
      sharpen(plane, W, H, sharpR, 0.5 + s);
      for (let i = 0, j = 0; i < N; i++, j += 4) d[j] = d[j + 1] = d[j + 2] = toLut(plane[i]);
      return img;
    }
    // カラー：①R・G・B をそれぞれ紙の明るさで割る（影と色かぶりが消えて、紙が真っ白に）
    for (let c = 0; c < 3; c++) {
      for (let i = 0, j = c; i < N; i++, j += 4) plane[i] = d[j];
      normalizeByBackground(plane, W, H);
      for (let i = 0, j = c; i < N; i++, j += 4) d[j] = plane[i] * 255;
    }
    // ②明るさだけ輪郭強調とトーンカーブで濃くし、色み（明るさとの差）は少し強めて足し戻す。
    //   色ごとに濃くすると青や赤のペンが黒くなってしまうため
    for (let i = 0, j = 0; i < N; i++, j += 4) plane[i] = (0.299 * d[j] + 0.587 * d[j + 1] + 0.114 * d[j + 2]) / 255;
    sharpen(plane, W, H, sharpR, 0.5 + s);
    const SAT = 1.3;
    for (let i = 0, j = 0; i < N; i++, j += 4) {
      const r = d[j], g = d[j + 1], b = d[j + 2], L = 0.299 * r + 0.587 * g + 0.114 * b, Lp = toLut(plane[i]);
      d[j] = Lp + (r - L) * SAT; d[j + 1] = Lp + (g - L) * SAT; d[j + 2] = Lp + (b - L) * SAT;
    }
    return img;
  }

  return { detectDocument, orderQuad, quadArea, paperAspect, outputSize, warpQuad, enhance };
})();

if (typeof module !== 'undefined') module.exports = Scan;
