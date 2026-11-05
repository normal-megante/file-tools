'use strict';
// =========================================================
// どのツールでも使う共通の部品
//   ・小さな便利関数（サイズの表示、ファイル名の整理など）
//   ・保存（ダウンロード）と共有、ZIP 作り
//   ・画像を入れる場所（ドラッグ＆ドロップ・選択・貼り付け）
//   ・画像の一覧（ImageList）… サムネイル、回転、削除、並べ替え
// =========================================================

const $ = (s, root = document) => root.querySelector(s);
const collator = new Intl.Collator('ja', { numeric: true, sensitivity: 'base' });
const IMAGE_EXT = /\.(jpe?g|jfif|png|webp|gif|bmp|avif|heic|heif|tiff?)$/i;
const isImageFile = f => (f.type && f.type.startsWith('image/')) || IMAGE_EXT.test(f.name);
const stem = name => name.replace(/\.[^.]+$/, '');
const sanitize = s => s.replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_').trim().replace(/\.(pdf|zip|jpe?g|png|webp)$/i, '').trim();

function fmtSize(n) {
  if (n < 1024) return n + ' B';
  if (n < 1048576) return (n / 1024).toFixed(n < 10240 ? 1 : 0) + ' KB';
  return (n / 1048576).toFixed(n < 10485760 ? 2 : 1) + ' MB';
}

let toastTimer;
function toast(msg) {
  let t = $('#toast');
  if (!t) { t = document.createElement('div'); t.id = 'toast'; t.className = 'toast'; t.setAttribute('role', 'status'); document.body.appendChild(t); }
  t.textContent = msg; t.classList.add('show');
  clearTimeout(toastTimer); toastTimer = setTimeout(() => t.classList.remove('show'), 3500);
}

function latin1(str) {
  const u = new Uint8Array(str.length);
  for (let i = 0; i < str.length; i++) u[i] = str.charCodeAt(i) & 255;
  return u;
}
function concatBytes(parts) {
  let n = 0; for (const p of parts) n += p.length;
  const out = new Uint8Array(n); let o = 0;
  for (const p of parts) { out.set(p, o); o += p.length; }
  return out;
}
const wait = ms => new Promise(r => setTimeout(r, ms));

/* ---------- 画像の書き出し ---------- */
function newCanvas(w, h) { const c = document.createElement('canvas'); c.width = w; c.height = h; return c; }

// toBlob はブラウザの空き時間待ちで1枚1秒近くかかることがあるので、同期の toDataURL を使う
function encodeCanvas(c, type, q) {
  const url = c.toDataURL(type, q);
  if (!url.startsWith('data:' + type)) throw new Error('画像の書き出しに失敗しました');
  const bin = atob(url.slice(url.indexOf(',') + 1)), data = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) data[i] = bin.charCodeAt(i);
  return data;
}
// そのブラウザがその形式で書き出せるか（WebP は古い Safari だと書き出せない）
const canEncode = type => newCanvas(1, 1).toDataURL(type).startsWith('data:' + type);

// 大きく縮小するときは半分ずつ縮めてジャギーを防ぐ。background を渡すと透明部分をその色で塗る
function downscale(src, tw, th, background) {
  let cur = src, sw = src.width, sh = src.height, tmp = null;
  while (sw / 2 >= tw && sh / 2 >= th) {
    const nw = Math.round(sw / 2), nh = Math.round(sh / 2), c = newCanvas(nw, nh), x = c.getContext('2d');
    x.imageSmoothingQuality = 'high'; x.drawImage(cur, 0, 0, nw, nh);
    if (tmp) tmp.width = tmp.height = 0;
    tmp = cur = c; sw = nw; sh = nh;
  }
  const out = newCanvas(tw, th), ctx = out.getContext('2d', { willReadFrequently: true });
  if (background) { ctx.fillStyle = background; ctx.fillRect(0, 0, tw, th); }
  ctx.imageSmoothingQuality = 'high'; ctx.drawImage(cur, 0, 0, tw, th);
  if (tmp) tmp.width = tmp.height = 0;
  return out;
}

// 回転（右回り rot 度）→ 左右反転（flip）→ 切り抜き（crop: 0〜1 の割合）の順に当てて target に描く
function drawEdited(target, src, { rot = 0, flip = false, crop = null } = {}, scale = 1) {
  const sw = src.width, sh = src.height, swap = rot % 180 !== 0;
  const RW = swap ? sh : sw, RH = swap ? sw : sh;
  const c = crop || { x: 0, y: 0, w: 1, h: 1 };
  const cx = c.x * RW, cy = c.y * RH, cw = c.w * RW, ch = c.h * RH;
  target.width = Math.max(1, Math.round(cw * scale)); target.height = Math.max(1, Math.round(ch * scale));
  const x = target.getContext('2d');
  x.imageSmoothingQuality = 'high';
  x.scale(target.width / cw, target.height / ch);
  x.translate(-cx, -cy);
  if (flip) { x.translate(RW, 0); x.scale(-1, 1); }
  x.translate(RW / 2, RH / 2); x.rotate(rot * Math.PI / 180); x.drawImage(src, -sw / 2, -sh / 2);
  return target;
}

/* ---------- 保存と共有 ---------- */
function downloadUrl(url, name) {
  const a = document.createElement('a');
  a.href = url; a.download = name;
  document.body.appendChild(a); a.click(); a.remove();
}
function saveBlob(blob, name) {
  const url = URL.createObjectURL(blob);
  downloadUrl(url, name);
  setTimeout(() => URL.revokeObjectURL(url), 60000);
}
// スマホの「共有」でギャラリーやドライブに送れるか
function canShareFiles(files) {
  try { return !!(navigator.canShare && navigator.canShare({ files })); } catch (e) { return false; }
}
async function shareFiles(files) {
  try { await navigator.share({ files }); return true; }
  catch (e) { if (e.name !== 'AbortError') toast('共有できませんでした'); return false; }
}

/* ---------- ZIP（無圧縮で格納。中身はすでに圧縮済み） ---------- */
const CRC_TABLE = (() => {
  const t = new Uint32Array(256);
  for (let n = 0; n < 256; n++) { let c = n; for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; }
  return t;
})();
function crc32(u8) { let c = 0xFFFFFFFF; for (let i = 0; i < u8.length; i++) c = CRC_TABLE[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
async function makeZip(files) {
  const enc = new TextEncoder(), d = new Date(), parts = [], central = [];
  const dosTime = (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1);
  const dosDate = ((d.getFullYear() - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate();
  let offset = 0;
  for (const f of files) {
    const data = new Uint8Array(await f.blob.arrayBuffer()), name = enc.encode(f.name), crc = crc32(data);
    const lh = new DataView(new ArrayBuffer(30));
    lh.setUint32(0, 0x04034b50, true); lh.setUint16(4, 20, true); lh.setUint16(6, 0x0800, true); // UTF-8 のファイル名
    lh.setUint16(10, dosTime, true); lh.setUint16(12, dosDate, true); lh.setUint32(14, crc, true);
    lh.setUint32(18, data.length, true); lh.setUint32(22, data.length, true); lh.setUint16(26, name.length, true);
    parts.push(new Blob([lh.buffer, name, data]));
    const ch = new DataView(new ArrayBuffer(46));
    ch.setUint32(0, 0x02014b50, true); ch.setUint16(4, 20, true); ch.setUint16(6, 20, true); ch.setUint16(8, 0x0800, true);
    ch.setUint16(12, dosTime, true); ch.setUint16(14, dosDate, true); ch.setUint32(16, crc, true);
    ch.setUint32(20, data.length, true); ch.setUint32(24, data.length, true); ch.setUint16(28, name.length, true);
    ch.setUint32(42, offset, true);
    central.push(ch.buffer, name);
    offset += 30 + name.length + data.length;
  }
  const cdSize = central.reduce((s, p) => s + p.byteLength, 0), end = new DataView(new ArrayBuffer(22));
  end.setUint32(0, 0x06054b50, true); end.setUint16(8, files.length, true); end.setUint16(10, files.length, true);
  end.setUint32(12, cdSize, true); end.setUint32(16, offset, true);
  return new Blob([...parts, ...central, end.buffer], { type: 'application/zip' });
}
// 同じ名前が重なったら「名前 (2).拡張子」にする
function uniqueNamer() {
  const used = new Set();
  return (base, ext) => {
    let name = base + ext;
    for (let k = 2; used.has(name.toLowerCase()); k++) name = `${base} (${k})${ext}`;
    used.add(name.toLowerCase());
    return name;
  };
}

/* ---------- アイコン ---------- */
const SVG = (d, size = 17) => `<svg viewBox="0 0 24 24" width="${size}" height="${size}" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">${d}</svg>`;
const ICON = {
  rotl: SVG('<path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/><path d="M3 3v5h5"/>'),
  rotr: SVG('<path d="M21 12a9 9 0 1 1-9-9c2.52 0 4.93 1 6.74 2.74L21 8"/><path d="M21 3v5h-5"/>'),
  del:  SVG('<path d="M3 6h18"/><path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/><path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>'),
  crop: SVG('<path d="M6 2v14a2 2 0 0 0 2 2h14"/><path d="M18 22V8a2 2 0 0 0-2-2H2"/>'),
  flip: SVG('<path d="M12 3v18"/><path d="M8 7 3 12l5 5z"/><path d="m16 7 5 5-5 5z"/>'),
  left: SVG('<path d="m15 18-6-6 6-6"/>'),
  right: SVG('<path d="m9 18 6-6-6-6"/>'),
};
const ACTION_LABEL = { rotl: '左に90°回転', rotr: '右に90°回転', crop: '回転・切り抜き', del: '削除', left: '前へ移動', right: '後ろへ移動' };

/* ---------- 画像を入れる場所 ---------- */
// フォルダがドロップされたら中の画像を全部拾う（entries はイベント中に同期で取得する必要がある）
async function filesFromDataTransfer(dt) {
  const entries = [...(dt.items || [])].filter(i => i.kind === 'file').map(i => i.webkitGetAsEntry && i.webkitGetAsEntry());
  if (!entries.some(en => en && en.isDirectory)) return [...dt.files];
  const out = [];
  async function walk(en) {
    if (!en) return;
    if (en.isFile) {
      const f = await new Promise((res, rej) => en.file(res, rej));
      f._path = en.fullPath; out.push(f);
    } else if (en.isDirectory) {
      const rd = en.createReader(); let batch;
      do { batch = await new Promise((res, rej) => rd.readEntries(res, rej)); for (const b of batch) await walk(b); } while (batch.length);
    }
  }
  for (const en of entries) await walk(en);
  return out;
}

// drop: 枠の要素、input: <input type=file>、onFiles: 受け取ったファイルの処理、
// isBusy: 処理中は受け付けない、isInternalDrag: 一覧の並べ替え中はファイル扱いしない
function setupFileInput({ drop, input, onFiles, isBusy = () => false, isInternalDrag = () => false }) {
  const hasFiles = e => e.dataTransfer && [...e.dataTransfer.types].includes('Files');
  for (const ev of ['dragenter', 'dragover']) window.addEventListener(ev, e => {
    if (isInternalDrag() || !hasFiles(e)) return;
    e.preventDefault(); e.dataTransfer.dropEffect = 'copy'; drop.classList.add('over');
  });
  window.addEventListener('dragleave', e => { if (!e.relatedTarget) drop.classList.remove('over'); });
  window.addEventListener('drop', async e => {
    if (isInternalDrag() || !hasFiles(e)) return;
    e.preventDefault(); drop.classList.remove('over');
    if (isBusy()) return toast('処理中は追加できません');
    onFiles(await filesFromDataTransfer(e.dataTransfer));
  });
  drop.addEventListener('click', () => { if (!isBusy()) input.click(); });
  drop.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); drop.click(); } });
  input.addEventListener('change', () => { onFiles([...input.files]); input.value = ''; });
  let pasteNo = 0;
  document.addEventListener('paste', e => {
    const fs = [...((e.clipboardData && e.clipboardData.files) || [])].filter(isImageFile);
    if (!fs.length || isBusy()) return;
    e.preventDefault();
    onFiles(fs.map(f => new File([f], `貼り付け_${++pasteNo}.${(f.type.split('/')[1] || 'png').replace('jpeg', 'jpg')}`, { type: f.type })));
  });
}

/* ---------- 画像の一覧 ---------- */
// 1枚ごとに { file, name, size, w, h, isJpeg, rot, flip, crop, status, thumb, el } を持つ
class ImageList {
  constructor(grid, { actions = ['rotl', 'rotr', 'del'], onChange = () => {}, onAction = () => {}, numbered = true } = {}) {
    this.grid = grid; this.actions = actions; this.onChange = onChange; this.onAction = onAction; this.numbered = numbered;
    this.items = []; this.seq = 0; this.dragEl = null; this.busy = false;
    grid.addEventListener('dragstart', e => {
      const card = e.target.closest('.card'); if (!card) return;
      this.dragEl = card; card.classList.add('dragging');
      e.dataTransfer.effectAllowed = 'move'; e.dataTransfer.setData('text/plain', '');
    });
    grid.addEventListener('dragover', e => {
      if (!this.dragEl) return;
      e.preventDefault();
      const t = e.target.closest('.card'); if (!t || t === this.dragEl) return;
      const r = t.getBoundingClientRect();
      grid.insertBefore(this.dragEl, e.clientX > r.left + r.width / 2 ? t.nextSibling : t);
    });
    grid.addEventListener('drop', e => { if (this.dragEl) e.preventDefault(); });
    grid.addEventListener('dragend', () => {
      if (!this.dragEl) return;
      this.dragEl.classList.remove('dragging'); this.dragEl = null; this.syncFromDom();
    });
    grid.addEventListener('click', e => {
      const b = e.target.closest('button[data-act]'); if (!b) return;
      const it = this.items.find(i => String(i.id) === b.closest('.card').dataset.id); if (!it) return;
      const act = b.dataset.act;
      if (act === 'del') return this.remove(it);
      if (act === 'left' || act === 'right') return this.move(it, act === 'left' ? -1 : 1);
      if (it.status !== 'ok') return;
      if (act === 'rotl' || act === 'rotr') {
        it.rot = (it.rot + (act === 'rotr' ? (it.flip ? 270 : 90) : (it.flip ? 90 : 270))) % 360;
        if (it.crop) it.crop = rotateCrop(it.crop, act === 'rotr');
        this.drawThumb(it); this.onChange();
      } else this.onAction(it, act);
    });
  }
  get ok() { return this.items.filter(i => i.status === 'ok'); }
  get loading() { return this.items.some(i => i.status === 'loading'); }
  get errors() { return this.items.filter(i => i.status === 'error').length; }
  get isDragging() { return !!this.dragEl; }

  async add(files) {
    const all = [...files];
    const imgs = all.filter(isImageFile).sort((a, b) => collator.compare(a._path || a.name, b._path || b.name));
    if (all.length > imgs.length) toast(`画像ではないファイルを ${all.length - imgs.length} 件スキップしました`);
    if (!imgs.length) return [];
    const added = imgs.map(f => ({ id: ++this.seq, file: f, name: f.name, size: f.size, rot: 0, flip: false, crop: null, status: 'loading' }));
    for (const it of added) { this.items.push(it); this.grid.appendChild(this.makeCard(it)); }
    this.refresh();
    for (const it of added) await this.load(it);
    return added;
  }
  makeCard(it) {
    const el = document.createElement('div');
    el.className = 'card loading'; el.draggable = true; el.dataset.id = it.id;
    const btns = ['left', 'right', ...this.actions].map(a =>
      `<button type="button" class="icon${a === 'del' ? ' del' : ''}${a === 'left' || a === 'right' ? ' move' : ''}" data-act="${a}" title="${ACTION_LABEL[a]}" aria-label="${ACTION_LABEL[a]}">${ICON[a]}</button>`).join('');
    el.innerHTML = `<div class="thumb"><canvas width="1" height="1"></canvas><span class="num"${this.numbered ? '' : ' hidden'}></span><div class="spin"></div></div>
      <div class="meta"><div class="name"></div><div class="sub">読み込み中…</div><div class="note" hidden></div></div>
      <div class="actions">${btns}</div>`;
    const nm = el.querySelector('.name'); nm.textContent = it.name; nm.title = it.name;
    it.el = el;
    return el;
  }
  async load(it) {
    try {
      const head = new Uint8Array(await it.file.slice(0, 3).arrayBuffer());
      it.isJpeg = head[0] === 0xFF && head[1] === 0xD8 && head[2] === 0xFF;
      const bmp = await createImageBitmap(it.file, { imageOrientation: 'from-image' });
      it.w = bmp.width; it.h = bmp.height;
      const s = Math.min(1, 300 / Math.max(it.w, it.h));
      const c = newCanvas(Math.max(1, Math.round(it.w * s)), Math.max(1, Math.round(it.h * s)));
      const x = c.getContext('2d'); x.imageSmoothingQuality = 'high'; x.drawImage(bmp, 0, 0, c.width, c.height);
      bmp.close();
      it.thumb = c; it.status = 'ok';
      this.drawThumb(it);
      it.el.querySelector('.sub').textContent = `${it.w}×${it.h} ・ ${fmtSize(it.size)}`;
    } catch (e) {
      it.status = 'error';
      it.el.classList.add('error');
      it.el.querySelector('.sub').textContent = '読み込めない形式です（使われません）';
    }
    it.el.classList.remove('loading');
    this.refresh();
  }
  drawThumb(it) { drawEdited(it.el.querySelector('.thumb canvas'), it.thumb, it, 1); }
  setNote(it, text) { const n = it.el.querySelector('.note'); n.textContent = text || ''; n.hidden = !text; }
  refresh() {
    let n = 0;
    for (const it of this.items) it.el.querySelector('.num').textContent = it.status === 'ok' ? ++n : '—';
    this.grid.classList.toggle('busy', this.busy);
    this.onChange();
  }
  setBusy(b) { this.busy = b; this.refresh(); }
  syncFromDom() {
    const byId = new Map(this.items.map(i => [String(i.id), i]));
    this.items = [...this.grid.children].map(el => byId.get(el.dataset.id));
    this.refresh();
  }
  reorder(items) {
    this.items = items;
    for (const it of items) this.grid.appendChild(it.el);
    this.refresh();
  }
  move(it, d) {
    const i = this.items.indexOf(it), j = i + d;
    if (j < 0 || j >= this.items.length) return;
    const items = [...this.items]; [items[i], items[j]] = [items[j], items[i]];
    this.reorder(items);
  }
  sortByName() { this.reorder([...this.items].sort((a, b) => collator.compare(a.file._path || a.name, b.file._path || b.name))); }
  reverse() { this.reorder([...this.items].reverse()); }
  remove(it) {
    this.items = this.items.filter(i => i !== it);
    it.el.remove();
    if (it.thumb) it.thumb.width = it.thumb.height = 0;
    this.refresh();
  }
  clear() { for (const it of [...this.items]) this.remove(it); }
}

// 切り抜き枠を、画像といっしょに90°回したときの位置にする（cw: 右回り）
function rotateCrop(c, cw) {
  return cw ? { x: 1 - (c.y + c.h), y: c.x, w: c.h, h: c.w } : { x: c.y, y: 1 - (c.x + c.w), w: c.h, h: c.w };
}

/* ---------- アプリとして動かすしくみ（ネットなしでも開ける） ---------- */
if ('serviceWorker' in navigator && /^https?:$/.test(location.protocol)) {
  navigator.serviceWorker.register('./sw.js').catch(e => console.warn('サービスワーカーを登録できませんでした', e));
}
