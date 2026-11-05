'use strict';
// =========================================================
// PDF を作る部品（画像を1ページずつ入れる）
//   ・JPEG をそのまま入れるための解析（向き・ICC プロファイル）
//   ・画像の圧縮（JPEG / 可逆圧縮 / 白黒2値）
//   ・PDF の組み立て（buildPdf）
// common.js のあとに読み込む
// =========================================================

const HAS_DEFLATE = typeof CompressionStream !== 'undefined';
// 用紙サイズ（縦向き, pt）
const PAPER = { A4: [595.28, 841.89], A3: [841.89, 1190.55], A5: [419.53, 595.28], B5: [515.91, 728.5], B4: [728.5, 1031.81], Letter: [612, 792] };

/* ---------- JPEG 解析（再圧縮せずにそのまま入れるため） ---------- */
function exifOrientation(b, t, end) {
  if (t + 8 > end) return 1;
  const le = b[t] === 0x49 && b[t + 1] === 0x49, be = b[t] === 0x4D && b[t + 1] === 0x4D;
  if (!le && !be) return 1;
  const u16 = o => le ? b[o] | (b[o + 1] << 8) : (b[o] << 8) | b[o + 1];
  const u32 = o => (le ? (b[o] | (b[o + 1] << 8) | (b[o + 2] << 16) | (b[o + 3] << 24)) : ((b[o] << 24) | (b[o + 1] << 16) | (b[o + 2] << 8) | b[o + 3])) >>> 0;
  const ifd = t + u32(t + 4);
  if (ifd + 2 > end) return 1;
  const n = u16(ifd);
  for (let k = 0; k < n; k++) {
    const e = ifd + 2 + k * 12;
    if (e + 12 > end) break;
    if (u16(e) === 0x0112) { const v = u16(e + 8); return v >= 1 && v <= 8 ? v : 1; }
  }
  return 1;
}
const bytesAscii = (b, s, n) => String.fromCharCode(...b.subarray(s, s + n));

// SOF・EXIFの向き・ICCプロファイルを読み取り、表示に不要なメタデータ（EXIF・サムネイル・付属画像）を除いたJPEGを返す
function analyzeJpeg(b) {
  if (b.length < 4 || b[0] !== 0xFF || b[1] !== 0xD8) return null;
  const parts = [b.subarray(0, 2)], icc = [];
  let i = 2, info = null, orient = 1, sawScan = false;
  while (i + 3 < b.length) {
    if (b[i] !== 0xFF) return null;
    const m = b[i + 1];
    if (m === 0xFF) { i++; continue; }
    if (m === 0x01 || (m >= 0xD0 && m <= 0xD8)) { i += 2; continue; }
    if (m === 0xD9) break;
    const len = (b[i + 2] << 8) | b[i + 3], s = i + 4, e = i + 2 + len;
    if (len < 2 || e > b.length) return null;
    if (m === 0xDA) {
      // エントロピー符号化データを走査して、メイン画像の EOI を探す（後ろに付いたゲインマップ等は捨てる）
      let j = e, end = -1;
      while (j + 1 < b.length) {
        if (b[j] !== 0xFF) { j++; continue; }
        const n = b[j + 1];
        if (n === 0x00 || (n >= 0xD0 && n <= 0xD7)) { j += 2; continue; }
        if (n === 0xFF) { j++; continue; }
        if (n === 0xD9) { end = j + 2; break; }
        j += 2 + ((b[j + 2] << 8) | b[j + 3]);
      }
      parts.push(b.subarray(i, end > 0 ? end : b.length));
      sawScan = true;
      break;
    }
    if (m >= 0xC0 && m <= 0xCF && m !== 0xC4 && m !== 0xC8 && m !== 0xCC)
      info = { sof: m, prec: b[s], h: (b[s + 1] << 8) | b[s + 2], w: (b[s + 3] << 8) | b[s + 4], comps: b[s + 5] };
    if (m === 0xE1 && len >= 16 && bytesAscii(b, s, 6) === 'Exif\0\0') orient = exifOrientation(b, s + 6, e);
    if (m === 0xE2 && len >= 16 && bytesAscii(b, s, 12) === 'ICC_PROFILE\0') icc.push([b[s + 12], b.subarray(s + 14, e)]);
    const drop = (m >= 0xE1 && m <= 0xEF && m !== 0xEE) || m === 0xFE; // APP1〜15（Adobe APP14 以外）とコメント
    if (!drop) parts.push(b.subarray(i, e));
    i = e;
  }
  if (!info || !sawScan) return null;
  info.orient = orient;
  info.ok = (info.sof === 0xC0 || info.sof === 0xC1 || info.sof === 0xC2) && info.prec === 8 &&
            (info.comps === 1 || info.comps === 3) && info.w > 0 && info.h > 0;
  info.data = concatBytes(parts);
  info.icc = null;
  if (icc.length) {
    const prof = concatBytes(icc.sort((a, b) => a[0] - b[0]).map(c => c[1]));
    const sig = prof.length >= 128 ? bytesAscii(prof, 16, 4) : '';
    if ((info.comps === 3 && sig === 'RGB ') || (info.comps === 1 && sig === 'GRAY')) info.icc = prof;
  }
  return info;
}

/* ---------- 画像の圧縮（PDF に入れる形） ---------- */
// どれも { w, h, comps, bits?, filter, parms?, data, icc?, orient, rot? } を返す
const jpegFromCanvas = (c, q) => ({ w: c.width, h: c.height, comps: 3, filter: 'DCTDecode', data: encodeCanvas(c, 'image/jpeg', q), orient: 1 });

// PNG と同じ行ごとのフィルタ（PDF の /Predictor 15）。スクショや図は JPEG より小さく・くっきりになることが多い
function pngFilter(raw, w, h, bpp) {
  const rl = w * bpp, out = new Uint8Array(h * (rl + 1)), zero = new Uint8Array(rl);
  const f = [0, 1, 2, 3, 4].map(() => new Uint8Array(rl));
  for (let y = 0; y < h; y++) {
    const o = y * rl, cur = raw.subarray(o, o + rl), up = y ? raw.subarray(o - rl, o) : zero;
    const f0 = f[0], f1 = f[1], f2 = f[2], f3 = f[3], f4 = f[4];
    for (let i = 0; i < rl; i++) {
      const x = cur[i], a = i >= bpp ? cur[i - bpp] : 0, b = up[i], c = i >= bpp ? up[i - bpp] : 0;
      f0[i] = x; f1[i] = x - a; f2[i] = x - b; f3[i] = x - ((a + b) >> 1);
      const p = a + b - c, pa = p > a ? p - a : a - p, pb = p > b ? p - b : b - p, pc = p > c ? p - c : c - p;
      f4[i] = x - (pa <= pb && pa <= pc ? a : pb <= pc ? b : c);
    }
    let best = 0, bestSum = Infinity;
    for (let k = 0; k < 5; k++) {
      const t = f[k]; let s = 0;
      for (let i = 0; i < rl; i++) { const v = t[i]; s += v < 128 ? v : 256 - v; }
      if (s < bestSum) { bestSum = s; best = k; }
    }
    out[y * (rl + 1)] = best; out.set(f[best], y * (rl + 1) + 1);
  }
  return out;
}
async function deflate(u8) {
  const stream = new Blob([u8]).stream().pipeThrough(new CompressionStream('deflate'));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
async function flateFromCanvas(c, gray) {
  const w = c.width, h = c.height, px = c.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, w, h).data;
  const bpp = gray ? 1 : 3, raw = new Uint8Array(w * h * bpp);
  if (gray) for (let i = 0, j = 0; j < raw.length; i += 4, j++) raw[j] = px[i];
  else for (let i = 0, j = 0; j < raw.length; i += 4, j += 3) { raw[j] = px[i]; raw[j + 1] = px[i + 1]; raw[j + 2] = px[i + 2]; }
  const data = await deflate(pngFilter(raw, w, h, bpp));
  return { w, h, comps: bpp, filter: 'FlateDecode', parms: `<< /Predictor 15 /Colors ${bpp} /BitsPerComponent 8 /Columns ${w} >>`, data, orient: 1 };
}
// 白黒2値（1ピクセル1ビット）。白黒の書類はこれがいちばん小さい
async function bitonalFromCanvas(c, threshold = 128) {
  const w = c.width, h = c.height, px = c.getContext('2d', { willReadFrequently: true }).getImageData(0, 0, w, h).data;
  const rb = (w + 7) >> 3, raw = new Uint8Array(rb * h); // 1 = 白
  for (let y = 0; y < h; y++) {
    let i = y * w * 4;
    for (let x = 0; x < w; x++, i += 4) if (px[i] >= threshold) raw[y * rb + (x >> 3)] |= 0x80 >> (x & 7);
  }
  const data = await deflate(pngFilter(raw, rb, h, 1));
  return { w, h, comps: 1, bits: 1, filter: 'FlateDecode', parms: `<< /Predictor 15 /Colors 1 /BitsPerComponent 1 /Columns ${w} >>`, data, orient: 1 };
}

/* ---------- PDF の組み立て ---------- */
const pdfNum = n => String(Math.round(n * 1000) / 1000 || 0);
function pdfText(s) { // UTF-16BE の16進文字列（日本語タイトル用）
  let h = '<FEFF';
  for (let i = 0; i < s.length; i++) h += s.charCodeAt(i).toString(16).padStart(4, '0').toUpperCase();
  return h + '>';
}
function pdfDate(d = new Date()) {
  const p = n => String(n).padStart(2, '0'), tz = -d.getTimezoneOffset(), a = Math.abs(tz);
  return `D:${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}${tz >= 0 ? '+' : '-'}${p(a / 60 | 0)}'${p(a % 60)}'`;
}
function hashBytes(u8) { let h = 0x811c9dc5; for (let i = 0; i < u8.length; i++) h = Math.imul(h ^ u8[i], 16777619); return u8.length + ':' + (h >>> 0); }

// 向き（EXIF）と回転を反映した表示上の縦横
function displaySize(img, rot) {
  let W = img.orient >= 5 ? img.h : img.w, H = img.orient >= 5 ? img.w : img.h;
  return rot % 180 ? [H, W] : [W, H];
}
// o: { paper: 'fit' | 'A4' など, orient: 'auto' | 'portrait' | 'landscape', margin: mm }
function pageLayout(W, H, o) {
  const m = o.margin * 72 / 25.4;
  if (o.paper === 'fit') { // 長辺を A4 の長辺に揃える
    const s = (841.89 - 2 * m) / Math.max(W, H), iw = W * s, ih = H * s;
    return { pw: iw + 2 * m, ph: ih + 2 * m, box: { x: m, y: m, w: iw, h: ih } };
  }
  const [a, b] = PAPER[o.paper];
  const land = o.orient === 'landscape' || (o.orient === 'auto' && W > H);
  const pw = land ? b : a, ph = land ? a : b;
  const s = Math.min((pw - 2 * m) / W, (ph - 2 * m) / H), iw = W * s, ih = H * s;
  return { pw, ph, box: { x: (pw - iw) / 2, y: (ph - ih) / 2, w: iw, h: ih } };
}
// 画像空間 (u,v)∈[0,1]² → ページ座標 の変換行列。EXIFの向きとユーザーの回転をまとめて表現する
function imageMatrix(img, rot, box) {
  const { w, h, orient } = img, [W] = displaySize(img, rot), s = box.w / W;
  const map = (u, v) => {
    const x = u * w, y = (1 - v) * h; let X, Y;
    switch (orient) {
      case 2: X = w - x; Y = y; break;
      case 3: X = w - x; Y = h - y; break;
      case 4: X = x; Y = h - y; break;
      case 5: X = y; Y = x; break;
      case 6: X = h - y; Y = x; break;
      case 7: X = h - y; Y = w - x; break;
      case 8: X = y; Y = w - x; break;
      default: X = x; Y = y;
    }
    let cw = orient >= 5 ? h : w, ch = orient >= 5 ? w : h;
    for (let k = 0; k < rot / 90; k++) { [X, Y] = [ch - Y, X]; [cw, ch] = [ch, cw]; } // 右に90°ずつ
    return [box.x + X * s, box.y + (ch - Y) * s];
  };
  const p0 = map(0, 0), p1 = map(1, 0), p2 = map(0, 1);
  return [p1[0] - p0[0], p1[1] - p0[1], p2[0] - p0[0], p2[1] - p0[1], p0[0], p0[1]];
}

class PdfWriter {
  constructor() { this.blobs = []; this.pending = []; this.pos = 0; this.offsets = []; this.nextId = 1; }
  alloc() { return this.nextId++; }
  raw(x) { const u = typeof x === 'string' ? latin1(x) : x; this.pending.push(u); this.pos += u.length; }
  obj(id, dict, stream) {
    this.offsets[id] = this.pos;
    if (stream) { this.raw(`${id} 0 obj\n${dict}\nstream\n`); this.raw(stream); this.raw('\nendstream\nendobj\n'); }
    else this.raw(`${id} 0 obj\n${dict}\nendobj\n`);
  }
  flush() { if (this.pending.length) { this.blobs.push(new Blob(this.pending)); this.pending = []; } } // メモリを溜めすぎない
}

// items を1つずつ encode(item, o) で画像にして、1ページずつ入れる。
// 回転は画像の rot（あれば）か item.rot を使う
async function buildPdf(items, o, title, { encode, onPage } = {}) {
  const pw = new PdfWriter();
  pw.raw('%PDF-1.5\n%\xE2\xE3\xCF\xD3\n');
  const catalogId = pw.alloc(), pagesId = pw.alloc(), infoId = pw.alloc(), kids = [], iccIds = new Map();
  for (let i = 0; i < items.length; i++) {
    if (onPage) onPage(i);
    const it = items[i], img = await encode(it, o), rot = 'rot' in img ? img.rot : (it.rot || 0);
    let cs = img.comps === 1 ? '/DeviceGray' : '/DeviceRGB';
    if (img.icc) {
      const key = hashBytes(img.icc);
      if (!iccIds.has(key)) {
        const id = pw.alloc(); iccIds.set(key, id);
        pw.obj(id, `<< /N ${img.comps} /Alternate ${cs} /Length ${img.icc.length} >>`, img.icc);
      }
      cs = `[/ICCBased ${iccIds.get(key)} 0 R]`;
    }
    const [W, H] = displaySize(img, rot), lay = pageLayout(W, H, o);
    const content = `q ${imageMatrix(img, rot, lay.box).map(pdfNum).join(' ')} cm /Im0 Do Q`;
    const pageId = pw.alloc(), contId = pw.alloc(), imId = pw.alloc();
    pw.obj(imId, `<< /Type /XObject /Subtype /Image /Width ${img.w} /Height ${img.h} /ColorSpace ${cs} /BitsPerComponent ${img.bits || 8} /Filter /${img.filter}${img.parms ? ' /DecodeParms ' + img.parms : ''} /Length ${img.data.length} >>`, img.data);
    pw.obj(contId, `<< /Length ${content.length} >>`, latin1(content));
    pw.obj(pageId, `<< /Type /Page /Parent ${pagesId} 0 R /MediaBox [0 0 ${pdfNum(lay.pw)} ${pdfNum(lay.ph)}] /Resources << /XObject << /Im0 ${imId} 0 R >> >> /Contents ${contId} 0 R >>`);
    kids.push(pageId);
    pw.flush();
  }
  pw.obj(pagesId, `<< /Type /Pages /Kids [${kids.map(k => k + ' 0 R').join(' ')}] /Count ${kids.length} >>`);
  pw.obj(catalogId, `<< /Type /Catalog /Pages ${pagesId} 0 R >>`);
  pw.obj(infoId, `<< /Title ${pdfText(title)} /Producer (file-tools) /CreationDate (${pdfDate()}) >>`);
  const xref = pw.pos, n = pw.nextId;
  let x = `xref\n0 ${n}\n0000000000 65535 f \n`;
  for (let id = 1; id < n; id++) x += String(pw.offsets[id]).padStart(10, '0') + ' 00000 n \n';
  x += `trailer\n<< /Size ${n} /Root ${catalogId} 0 R /Info ${infoId} 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  pw.raw(x); pw.flush();
  return new Blob(pw.blobs, { type: 'application/pdf' });
}
