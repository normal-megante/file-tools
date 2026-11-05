'use strict';
// =========================================================
// スキャン補正の「ページの中のカメラ」（自動シャッター）
//   書類のふちを見つけて枠で見せる。スマホが止まって、ピントが合ったら自動で撮る（ボタンでも撮れる）。
//   撮った写真は File にして onShot に渡す。続けて何枚でも撮れる（single なら1枚で閉じる）。
//   使うもの：js/common.js（$, toast, newCanvas, encodeCanvas）、js/scan.js（Scan）
// =========================================================
const ScanCam = (() => {
  const AUTO_KEY = 'scan-cam-auto';
  const TICK_MS = 100;        // 映像を調べる間かく
  const STEADY_MS = 700;      // これだけ止まっていたら、自動で撮る
  const MOTION_T = 2.5;       // 小さく縮めた映像の、前のコマとの明るさの差（平均）。これ未満なら「止まっている」
  const GYRO_T = 10;          // スマホの回転の速さ（度/秒）。これ未満なら「止まっている」（取れる機種だけ）
  const CHANGE_T = 9;         // 前に撮ったときの映像との差。これを超えたら「次のページになった」とみなす
  const BLUR_MAX = 2.6;       // 映像のぼやけ（Scan.measureBlur の blur）。これより大きいと「ピントが合っていない」
  const el = id => document.getElementById(id);
  const st = {
    stream: null, track: null, caps: {}, ic: null, photoSettings: null, timer: 0, onShot: null, single: false,
    small: null, prevSig: null, steadySince: 0, gyro: 0, gyroAt: 0, lastDetect: 0, dets: [], quad: null, quadSeen: 0,
    widths: [], lastShotSig: null, armed: true, busy: false, count: 0, auto: true, focus: null, pushed: false, gen: 0,
  };
  try { st.auto = localStorage.getItem(AUTO_KEY) !== 'off'; } catch (e) { /* 既定はオン */ }

  const setStatus = t => { if (el('shStatus').textContent !== t) el('shStatus').textContent = t; };
  const now = () => performance.now();

  /* ---------- 開く・閉じる ---------- */
  async function open({ onShot, single = false } = {}) {
    if (st.stream) return;
    if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) { toast('このブラウザではカメラを使えません'); return false; }
    const gen = ++st.gen;
    Object.assign(st, { onShot, single, count: 0, armed: true, lastShotSig: null, prevSig: null, steadySince: 0, dets: [], quad: null, widths: [], focus: null, busy: false });
    el('shDone').textContent = single ? 'やめる' : '完了';
    el('shLast').hidden = true;
    el('shTorch').hidden = true; el('shTorch').setAttribute('aria-pressed', 'false');
    el('shAuto').setAttribute('aria-pressed', String(st.auto));
    el('shShutter').disabled = true;
    setStatus('カメラを準備しています…');
    el('shoot').hidden = false;
    document.body.style.overflow = 'hidden';
    // スマホの「戻る」でページから出てしまわないよう、履歴を1つ足しておく（戻る＝カメラを閉じる）
    try { history.pushState({ scanCam: 1 }, ''); st.pushed = true; } catch (e) { st.pushed = false; }
    // ImageCapture（高画質の写真）が使えるなら映像はふつうの大きさ、使えないなら映像の1コマを写真にするので大きめに
    const big = !('ImageCapture' in window);
    const size = big ? { width: { ideal: 3840 }, height: { ideal: 2160 } } : { width: { ideal: 1920 }, height: { ideal: 1080 } };
    let stream;
    try {
      stream = await navigator.mediaDevices.getUserMedia({ audio: false, video: { facingMode: { ideal: 'environment' }, ...size } });
    } catch (e) {
      close();
      toast(e.name === 'NotAllowedError' ? 'カメラの使用が許可されていません。アドレスバーの左のマークから許可できます'
        : e.name === 'NotFoundError' || e.name === 'OverconstrainedError' ? 'カメラが見つかりませんでした'
        : 'カメラを開けませんでした（ほかのアプリが使っているかもしれません）');
      return false;
    }
    if (gen !== st.gen || el('shoot').hidden) { stream.getTracks().forEach(t => t.stop()); return false; } // 準備中に閉じられた
    st.stream = stream;
    st.track = stream.getVideoTracks()[0];
    st.caps = (st.track && st.track.getCapabilities) ? st.track.getCapabilities() : {};
    const v = el('shVideo');
    v.srcObject = stream;
    try { await v.play(); } catch (e) { /* 自動再生できなくても、映像は出る */ }
    // ピントは「ずっと合わせ続ける」に（できる機種だけ）
    if (st.caps.focusMode && st.caps.focusMode.includes('continuous')) st.track.applyConstraints({ advanced: [{ focusMode: 'continuous' }] }).catch(() => {});
    el('shTorch').hidden = !st.caps.torch;
    st.ic = null; st.photoSettings = null;
    if ('ImageCapture' in window && st.track) {
      try { st.ic = new ImageCapture(st.track); } catch (e) { st.ic = null; }
    }
    el('shShutter').disabled = false;
    window.addEventListener('devicemotion', onMotion);
    layout();
    tick();
    return true;
  }
  function close(fromHistory) {
    st.gen++;
    clearTimeout(st.timer); st.timer = 0;
    window.removeEventListener('devicemotion', onMotion);
    if (st.stream) st.stream.getTracks().forEach(t => t.stop());
    st.stream = st.track = st.ic = null;
    el('shVideo').srcObject = null;
    el('shoot').hidden = true;
    document.body.style.overflow = '';
    if (st.pushed) { st.pushed = false; if (!fromHistory) history.back(); }
  }
  window.addEventListener('popstate', () => { if (!el('shoot').hidden) { st.pushed = false; close(true); } });
  // ほかのアプリに切り替えたらカメラを止める（電池とプライバシーのため）。撮った写真はそのまま
  document.addEventListener('visibilitychange', () => { if (document.hidden && !el('shoot').hidden) close(); });
  document.addEventListener('keydown', e => { if (e.key === 'Escape' && !el('shoot').hidden) close(); });

  function onMotion(e) {
    const r = e.rotationRate; if (!r) return;
    const m = Math.hypot(r.alpha || 0, r.beta || 0, r.gamma || 0);
    if (Number.isFinite(m)) { st.gyro = m; st.gyroAt = now(); }
  }

  /* ---------- 画面の大きさ合わせ ---------- */
  // 映像の縦横比のまま、画面に収まる大きさにする（枠を描く canvas も同じ大きさ）
  function layout() {
    const v = el('shVideo'), view = el('shView'), box = el('shBox');
    const vw = v.videoWidth || 3, vh = v.videoHeight || 4, W = view.clientWidth, H = view.clientHeight;
    if (!W || !H) return;
    const ar = vw / vh, bw = Math.min(W, H * ar), bh = bw / ar;
    box.style.width = Math.round(bw) + 'px'; box.style.height = Math.round(bh) + 'px';
    const dpr = Math.min(2, window.devicePixelRatio || 1), c = el('shOverlay');
    c.width = Math.round(bw * dpr); c.height = Math.round(bh * dpr);
  }
  el('shVideo').addEventListener('loadedmetadata', layout);
  el('shVideo').addEventListener('resize', layout); // スマホを横にしたときなど
  window.addEventListener('resize', () => { if (!el('shoot').hidden) layout(); });

  /* ---------- 映像を調べる ---------- */
  // 小さな絵（だいたい 40×30）の明るさ。動いたか・ページが替わったかを見るのに使う
  function signature(img) {
    const W = img.width, H = img.height, d = img.data, gw = 40, gh = Math.max(1, Math.round(40 * H / W));
    const s = new Float32Array(gw * gh), n = new Float32Array(gw * gh);
    for (let y = 0; y < H; y++) {
      const gy = Math.min(gh - 1, (y * gh / H) | 0);
      for (let x = 0; x < W; x++) {
        const g = gy * gw + Math.min(gw - 1, (x * gw / W) | 0), j = (y * W + x) * 4;
        s[g] += 0.299 * d[j] + 0.587 * d[j + 1] + 0.114 * d[j + 2]; n[g]++;
      }
    }
    for (let i = 0; i < s.length; i++) s[i] /= n[i] || 1;
    return s;
  }
  const sigDiff = (a, b) => {
    if (!a || !b || a.length !== b.length) return 99;
    let t = 0; for (let i = 0; i < a.length; i++) t += Math.abs(a[i] - b[i]);
    return t / a.length;
  };
  // 書類の四隅が、続けて同じあたりに見えているか
  function docStable() {
    const d = st.dets;
    if (d.length < 2 || !d[d.length - 1] || !d[d.length - 2]) return false;
    const a = d[d.length - 1], b = d[d.length - 2];
    return a.every((p, i) => Math.hypot(p[0] - b[i][0], p[1] - b[i][1]) < 0.03);
  }
  // 止まっている間に測ったぼやけのうち、いちばん小さいものに近く、それが続いていれば「ピントが合った」
  function focusState() {
    const w = st.widths;
    if (w.length < 3) return 'wait';
    const best = Math.min(...w.map(x => x.blur)), last3 = w.slice(-3).map(x => x.blur);
    if (best > BLUR_MAX) return 'blurry';
    return last3.every(b => b <= best * 1.15 + 0.3) ? 'sharp' : 'wait';
  }

  function tick() {
    if (!st.stream) return;
    const v = el('shVideo');
    if (v.readyState >= 2 && v.videoWidth && !st.busy) {
      try { analyze(v); } catch (e) { console.error(e); }
    }
    if (st.stream) st.timer = setTimeout(tick, TICK_MS);
  }
  function analyze(v) {
    const t = now(), vw = v.videoWidth, vh = v.videoHeight;
    const s = 320 / Math.max(vw, vh), sw = Math.max(1, Math.round(vw * s)), sh = Math.max(1, Math.round(vh * s));
    if (!st.small || st.small.width !== sw || st.small.height !== sh) { st.small = newCanvas(sw, sh); layout(); }
    const sx = st.small.getContext('2d', { willReadFrequently: true });
    sx.drawImage(v, 0, 0, sw, sh);
    const img = sx.getImageData(0, 0, sw, sh), sig = signature(img);
    // ① 止まっているか（映像の変化と、取れればスマホの回転の速さ）
    const diff = sigDiff(sig, st.prevSig); st.prevSig = sig;
    const gyroOk = !(st.gyroAt && t - st.gyroAt < 600) || st.gyro < GYRO_T;
    const still = diff < MOTION_T && gyroOk;
    if (!still) { st.steadySince = 0; st.widths = []; } else if (!st.steadySince) st.steadySince = t;
    // 前に撮ってから、ページが替わった（または大きく動いた）ら、次を撮れるようにする
    if (!st.armed && sigDiff(sig, st.lastShotSig) > CHANGE_T) st.armed = true;
    // ② 書類の四隅（0.35 秒ごと）
    if (t - st.lastDetect > 350) {
      st.lastDetect = t;
      const r = Scan.detectDocument(img);
      const q = r.found ? r.quad.map(([x, y]) => [x / sw, y / sh]) : null;
      st.dets.push(q); if (st.dets.length > 4) st.dets.shift();
      if (q) { st.quadSeen = t; st.quad = st.quad ? st.quad.map((p, i) => [p[0] + (q[i][0] - p[0]) * 0.6, p[1] + (q[i][1] - p[1]) * 0.6]) : q; }
    }
    if (st.quad && t - st.quadSeen > 900) st.quad = null;
    const found = !!st.quad && docStable();
    // ③ ピント：止まっている間、書類のまん中あたりを元の細かさのまま切り出して、ぼやけを測る
    if (still) {
      const T = Math.max(64, Math.min(256, Math.round(Math.min(vw, vh) / 4)));
      const c = st.quad ? st.quad.reduce((a, p) => [a[0] + p[0] / 4, a[1] + p[1] / 4], [0, 0]) : [0.5, 0.5];
      const x0 = Math.min(vw - T, Math.max(0, Math.round(c[0] * vw - T / 2))), y0 = Math.min(vh - T, Math.max(0, Math.round(c[1] * vh - T / 2)));
      if (!st.tile) st.tile = newCanvas(T, T);
      if (st.tile.width !== T) st.tile.width = st.tile.height = T;
      const tx = st.tile.getContext('2d', { willReadFrequently: true });
      tx.drawImage(v, x0, y0, T, T, 0, 0, T, T);
      const m = Scan.measureBlur([tx.getImageData(0, 0, T, T)]);
      if (m.ok) st.widths.push({ t, blur: m.blur });
      if (st.widths.length > 12) st.widths.shift();
    }
    const fs = focusState(), stillFor = st.steadySince ? t - st.steadySince : 0;
    // ④ 自動シャッター
    let progress = 0;
    if (!st.auto) setStatus('シャッターボタンで撮ってください');
    else if (!st.armed) setStatus(st.count ? '撮りました。次のページにしてください' : 'そのまま');
    else if (!found) setStatus('書類全体が入るように、少し離れてください');
    else if (!still) setStatus('そのまま止めてください');
    else if (fs === 'blurry') setStatus('ピントが合いません。少し離れるか、画面をタップしてください');
    else {
      progress = Math.min(1, stillFor / STEADY_MS) * (fs === 'sharp' ? 1 : 0.85);
      setStatus(fs === 'sharp' ? 'そのまま…' : 'ピントを合わせています…');
      if (fs === 'sharp' && stillFor >= STEADY_MS) { draw(found, 1); shoot(); return; }
    }
    draw(found, progress);
  }

  /* ---------- 枠・ピントの輪・シャッターの輪 ---------- */
  function draw(found, progress) {
    const c = el('shOverlay'), x = c.getContext('2d'), W = c.width, H = c.height, dpr = W / (c.clientWidth || W);
    x.clearRect(0, 0, W, H);
    if (st.quad) {
      const P = st.quad.map(([u, v]) => [u * W, v * H]);
      x.beginPath(); x.moveTo(P[0][0], P[0][1]); for (let k = 1; k < 4; k++) x.lineTo(P[k][0], P[k][1]); x.closePath();
      const ready = found && progress > 0;
      x.fillStyle = ready ? 'rgba(47,111,237,.22)' : 'rgba(255,255,255,.12)'; x.fill();
      x.lineWidth = 3 * dpr; x.strokeStyle = ready ? '#4c8dff' : 'rgba(255,255,255,.9)'; x.lineJoin = 'round'; x.stroke();
    }
    if (st.focus && now() - st.focus.t < 1200) {
      const r = 34 * dpr, fx = st.focus.x * W, fy = st.focus.y * H;
      x.lineWidth = 2 * dpr; x.strokeStyle = '#ffd400'; x.strokeRect(fx - r, fy - r, 2 * r, 2 * r);
    }
    el('shRing').setAttribute('stroke-dashoffset', String(226.2 * (1 - progress)));
  }

  /* ---------- 撮る ---------- */
  const withTimeout = (p, ms) => Promise.race([p, new Promise((_, rej) => setTimeout(() => rej(new Error('timeout')), ms))]);
  async function takePhoto() {
    if (!st.photoSettings) {
      st.photoSettings = {};
      try {
        const pc = await withTimeout(st.ic.getPhotoCapabilities(), 3000);
        if (pc.fillLightMode && pc.fillLightMode.includes('off')) st.photoSettings.fillLightMode = 'off'; // フラッシュはテカるので使わない
      } catch (e) { /* 分からなければ、そのまま撮る */ }
    }
    try { return await withTimeout(st.ic.takePhoto(st.photoSettings), 6000); }
    catch (e) { return await withTimeout(st.ic.takePhoto(), 6000); }
  }
  // 写真の向きが、画面の映像と違う機種があるので、映像と見比べて向きをそろえる
  function similarity(src, rot, ref) {
    const S = 48, rw = S, rh = Math.max(8, Math.round(S * ref.height / ref.width));
    const a = newCanvas(rw, rh), ax = a.getContext('2d', { willReadFrequently: true });
    ax.drawImage(ref, 0, 0, rw, rh);
    const b = newCanvas(rw, rh), bx = b.getContext('2d', { willReadFrequently: true });
    const swap = rot % 180 !== 0, w = swap ? src.height : src.width, h = swap ? src.width : src.height;
    const k = Math.max(rw / w, rh / h); // 映像の範囲（真ん中）に合わせて切り取る
    bx.translate(rw / 2, rh / 2); bx.rotate(rot * Math.PI / 180); bx.scale(k, k);
    bx.drawImage(src, -src.width / 2, -src.height / 2);
    const g = c => { const d = c.getContext('2d').getImageData(0, 0, rw, rh).data, o = new Float32Array(rw * rh); for (let i = 0; i < o.length; i++) o[i] = d[i * 4] + d[i * 4 + 1] + d[i * 4 + 2]; return o; };
    const p = g(a), q = g(b);
    let mp = 0, mq = 0; for (let i = 0; i < p.length; i++) { mp += p[i]; mq += q[i]; }
    mp /= p.length; mq /= q.length;
    let sxy = 0, sxx = 0, syy = 0;
    for (let i = 0; i < p.length; i++) { const u = p[i] - mp, v = q[i] - mq; sxy += u * v; sxx += u * u; syy += v * v; }
    return sxy / Math.sqrt(sxx * syy + 1e-9);
  }
  async function orient(blob, ref, name) {
    const bmp = await createImageBitmap(blob, { imageOrientation: 'from-image' });
    try {
      const cands = (bmp.height > bmp.width) === (ref.height > ref.width) ? [0, 180] : [90, 270];
      let best = cands[0], bestScore = -Infinity;
      for (const r of cands) { const sc = similarity(bmp, r, ref); if (sc > bestScore) { bestScore = sc; best = r; } }
      if (best === 0) return new File([blob], name, { type: blob.type || 'image/jpeg' });
      const swap = best % 180 !== 0, c = newCanvas(swap ? bmp.height : bmp.width, swap ? bmp.width : bmp.height), x = c.getContext('2d');
      x.translate(c.width / 2, c.height / 2); x.rotate(best * Math.PI / 180); x.drawImage(bmp, -bmp.width / 2, -bmp.height / 2);
      const data = encodeCanvas(c, 'image/jpeg', 0.92);
      c.width = c.height = 0;
      return new File([data], name, { type: 'image/jpeg' });
    } finally { bmp.close(); }
  }
  let lastBase = '', dup = 0;
  function fileName() {
    const d = new Date(), p = n => String(n).padStart(2, '0');
    const base = `カメラ_${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
    if (base === lastBase) dup++; else { lastBase = base; dup = 0; } // 同じ秒に撮ったら _2, _3 …
    return base + (dup ? '_' + (dup + 1) : '') + '.jpg';
  }
  async function shoot() {
    if (st.busy || !st.stream) return;
    const gen = st.gen, v = el('shVideo'), vw = v.videoWidth, vh = v.videoHeight;
    if (!vw) return;
    st.busy = true; el('shShutter').disabled = true;
    const fl = el('shFlash'); fl.classList.add('on'); setTimeout(() => fl.classList.remove('on'), 60);
    if (navigator.vibrate && (!navigator.userActivation || navigator.userActivation.hasBeenActive)) navigator.vibrate(40);
    setStatus('撮っています…');
    // 今の映像（向きの見比べ・サムネイル・次のページの見分けに使う）
    const s = 320 / Math.max(vw, vh), ref = newCanvas(Math.round(vw * s), Math.round(vh * s));
    ref.getContext('2d').drawImage(v, 0, 0, ref.width, ref.height);
    const sig = signature(ref.getContext('2d').getImageData(0, 0, ref.width, ref.height));
    const name = fileName();
    let file = null;
    if (st.ic) {
      try { file = await orient(await takePhoto(), ref, name); }
      catch (e) { console.warn('takePhoto に失敗。映像の1コマを使います', e); st.ic = null; }
    }
    if (gen !== st.gen) return; // 撮っている間に閉じられた
    if (!file) {
      const c = newCanvas(vw, vh);
      c.getContext('2d').drawImage(v, 0, 0);
      file = new File([encodeCanvas(c, 'image/jpeg', 0.92)], name, { type: 'image/jpeg' });
      c.width = c.height = 0;
    }
    st.count++; st.lastShotSig = sig; st.armed = false; st.steadySince = 0; st.widths = [];
    const last = el('shLast'), lx = last.getContext('2d'), m = Math.min(ref.width, ref.height);
    lx.drawImage(ref, (ref.width - m) / 2, (ref.height - m) / 2, m, m, 0, 0, last.width, last.height);
    last.hidden = false;
    if (!st.single) el('shDone').textContent = `完了（${st.count}枚）`;
    st.busy = false; el('shShutter').disabled = false;
    try { st.onShot && st.onShot(file); } catch (e) { console.error(e); }
    if (st.single) close();
  }

  /* ---------- ボタン・タップ ---------- */
  el('shShutter').addEventListener('click', () => { st.armed = true; shoot(); });
  el('shDone').addEventListener('click', () => close());
  el('shAuto').addEventListener('click', () => {
    st.auto = !st.auto;
    el('shAuto').setAttribute('aria-pressed', String(st.auto));
    try { localStorage.setItem(AUTO_KEY, st.auto ? 'on' : 'off'); } catch (e) { /* 覚えられなくても動く */ }
  });
  el('shTorch').addEventListener('click', () => {
    if (!st.track) return;
    const on = el('shTorch').getAttribute('aria-pressed') !== 'true';
    st.track.applyConstraints({ advanced: [{ torch: on }] })
      .then(() => el('shTorch').setAttribute('aria-pressed', String(on)))
      .catch(() => toast('ライトをつけられませんでした'));
  });
  // 画面をタップしたところにピントを合わせる（できる機種だけ）
  el('shOverlay').addEventListener('pointerdown', e => {
    if (!st.track || !st.caps.focusMode) return;
    const r = el('shOverlay').getBoundingClientRect();
    const fx = Math.min(1, Math.max(0, (e.clientX - r.left) / r.width)), fy = Math.min(1, Math.max(0, (e.clientY - r.top) / r.height));
    st.focus = { x: fx, y: fy, t: now() };
    const set = { pointsOfInterest: [{ x: fx, y: fy }] };
    if (st.caps.focusMode.includes('continuous')) set.focusMode = 'continuous';
    else if (st.caps.focusMode.includes('single-shot')) set.focusMode = 'single-shot';
    st.track.applyConstraints({ advanced: [set] }).catch(() => {});
    st.steadySince = 0; st.widths = []; // 新しいピントを待ってから撮る
  });

  return { open, close, isOpen: () => !el('shoot').hidden };
})();
