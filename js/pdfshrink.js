'use strict';
// =========================================================
// PDF を小さくする部品（PDF の圧縮）
//   ・ふつう（compress）：文字や図・しおり・リンクはそのまま。中の写真だけを
//     「ページに表示される大きさ × dpi」まで縮めて、JPEG にし直す
//   ・強力（rasterize）：ページを丸ごと画像にする（pdf.js で描いて JPEG）。文字は選べなくなる
// common.js・pdf.js（js/pdf.js）・pdfdoc.js・pdfview.js のあとに読み込む
// =========================================================
const PdfShrink = (() => {
  const { PName, PStr, PStream } = PdfDoc.types;
  const nameOf = v => (v instanceof PName ? v.n : null);

  // Filter・DecodeParms が別オブジェクトのときに、中身に置き換えたストリーム
  async function resolved(doc, st) {
    const d = new Map(st.dict);
    for (const k of ['Filter', 'DecodeParms']) {
      let v = await doc.get(d.get(k));
      if (Array.isArray(v)) v = await Promise.all(v.map(x => doc.get(x)));
      if (v !== undefined) d.set(k, v);
    }
    return new PStream(d, st.data);
  }

  // 画像の色の種類 → { n: 1（グレー）か 3（RGB）, indexed? }。CMYK・特色・Lab などは null（さわらない）
  async function colorSpace(doc, v, depth = 0) {
    v = await doc.get(v);
    const n = nameOf(v);
    if (n) return ['DeviceGray', 'G', 'CalGray'].includes(n) ? { n: 1 } : ['DeviceRGB', 'RGB', 'CalRGB'].includes(n) ? { n: 3 } : null;
    if (!Array.isArray(v) || !v.length || depth > 2) return null;
    const fam = nameOf(await doc.get(v[0]));
    if (fam === 'CalGray') return { n: 1 };
    if (fam === 'CalRGB') return { n: 3 };
    if (fam === 'ICCBased') {
      const s = await doc.get(v[1]), N = s instanceof PStream ? await doc.get(s.dict.get('N')) : 0;
      return N === 1 || N === 3 ? { n: N } : null;
    }
    if ((fam === 'Indexed' || fam === 'I') && depth === 0) { // 色の表（パレット）を使う画像
      const base = await colorSpace(doc, v[1], depth + 1), hival = await doc.get(v[2]);
      let lut = await doc.get(v[3]);
      if (!base || base.indexed || typeof hival !== 'number') return null;
      if (lut instanceof PStream) lut = await PdfDoc.decodeStream(await resolved(doc, lut));
      else if (lut instanceof PStr) lut = lut.b;
      else return null;
      if (lut.length < (hival + 1) * base.n) return null;
      return { n: 1, indexed: { base: base.n, lut, hival } };
    }
    return null;
  }

  // 画像を1つ縮める。縮めた PStream か、さわらないなら null
  //   drawn：ページに描かれている大きさ（pt）
  async function shrinkImage(doc, st, drawn, { dpi, quality, maxPixels }) {
    const g = k => doc.get(st.dict.get(k));
    if (nameOf(await g('Subtype')) !== 'Image' || (await g('ImageMask')) === true) return null;
    const w = await g('Width'), h = await g('Height');
    if (!(w > 0 && h > 0) || (await g('BitsPerComponent')) !== 8) return null;
    if (st.dict.has('Decode') || Array.isArray(await g('Mask')) || st.data.length < 8192) return null; // 小さい画像は縮めても得がない
    let f = await g('Filter');
    f = Array.isArray(f) ? f : f ? [f] : [];
    const filters = [];
    for (const x of f) filters.push(nameOf(await doc.get(x)));
    const cs = await colorSpace(doc, st.dict.get('ColorSpace'));
    if (!cs) return null;

    // 目標の大きさ：表示される大きさ × dpi。少し大きいくらいなら縮めない（画質を落とさない）
    let s = Math.max(drawn.w / 72 * dpi / w, drawn.h / 72 * dpi / h);
    if (s >= 0.75) s = 1;
    if (w * h * s * s > maxPixels) s = Math.sqrt(maxPixels / (w * h));
    const nw = Math.max(1, Math.round(w * s)), nh = Math.max(1, Math.round(h * s));

    let canvas;
    if (filters.length === 1 && (filters[0] === 'DCTDecode' || filters[0] === 'DCT')) {
      // JPEG：ブラウザで開く（EXIF の向きなどは外してから。PDF では使わないため）
      if (cs.indexed) return null;
      const info = analyzeJpeg(st.data);
      if (!info || !info.ok || info.comps !== cs.n || info.w !== w || info.h !== h) return null;
      const blob = new Blob([info.data], { type: 'image/jpeg' });
      const bmp = await createImageBitmap(blob, s < 1 ? { resizeWidth: nw, resizeHeight: nh, resizeQuality: 'high' } : {});
      canvas = newCanvas(nw, nh);
      const ctx = canvas.getContext('2d');
      ctx.imageSmoothingQuality = 'high'; ctx.drawImage(bmp, 0, 0, nw, nh); bmp.close();
    } else if (filters.every(x => x === 'FlateDecode' || x === 'Fl')) {
      // 可逆圧縮（PNG と同じしくみ）の画像：色の値を取り出して canvas に描く
      if (w * h > 50e6) return null;
      const data = await PdfDoc.decodeStream(await resolved(doc, st)), n = cs.n, cnt = w * h;
      if (data.length < cnt * n) return null;
      const full = newCanvas(w, h), ctx = full.getContext('2d'), id = ctx.createImageData(w, h), px = id.data;
      if (cs.indexed) {
        const { base, lut, hival } = cs.indexed;
        for (let j = 0, i = 0; j < cnt; j++, i += 4) {
          const k = Math.min(data[j], hival) * base;
          if (base === 3) { px[i] = lut[k]; px[i + 1] = lut[k + 1]; px[i + 2] = lut[k + 2]; } else px[i] = px[i + 1] = px[i + 2] = lut[k];
          px[i + 3] = 255;
        }
      } else if (n === 3) {
        for (let j = 0, i = 0; j < cnt * 3; j += 3, i += 4) { px[i] = data[j]; px[i + 1] = data[j + 1]; px[i + 2] = data[j + 2]; px[i + 3] = 255; }
      } else {
        for (let j = 0, i = 0; j < cnt; j++, i += 4) { px[i] = px[i + 1] = px[i + 2] = data[j]; px[i + 3] = 255; }
      }
      ctx.putImageData(id, 0, 0);
      canvas = s < 1 ? downscale(full, nw, nh) : full;
      if (canvas !== full) full.width = full.height = 0;
    } else return null; // JPEG2000・JBIG2 などはさわらない

    const jpg = encodeCanvas(canvas, 'image/jpeg', quality);
    canvas.width = canvas.height = 0;
    // 小さくならないなら元のまま（縮めないときは、はっきり小さくなるときだけ JPEG にし直す）
    if (jpg.length > st.data.length * (s < 1 ? 0.9 : 0.7)) return null;
    const d = new Map(st.dict);
    for (const k of ['Filter', 'DecodeParms', 'Length', 'DL', 'Decode', 'Interpolate']) d.delete(k);
    d.set('Width', nw); d.set('Height', nh); d.set('BitsPerComponent', 8);
    d.set('ColorSpace', new PName('DeviceRGB')); d.set('Filter', new PName('DCTDecode'));
    return new PStream(d, jpg);
  }

  // ふつう：中の写真だけを縮める。doc は PdfDoc.open で開いたもの
  //   onProgress(step, done, total)  step は 'image'（写真を縮める）か 'write'（書き出す）
  async function compress(doc, { dpi = 150, quality = 0.75, maxPixels = 25e6, onProgress } = {}) {
    const usage = await PdfDoc.imageUsage(doc), replace = new Map();
    let k = 0;
    for (const [num, drawn] of usage) {
      if (onProgress) onProgress('image', k++, usage.size);
      const st = await doc.getObj(num);
      if (!(st instanceof PStream)) continue;
      try {
        const r = await shrinkImage(doc, st, drawn, { dpi, quality, maxPixels });
        if (r) replace.set(num, r);
      } catch (e) { console.warn('画像を縮められませんでした', num, e); }
      await wait(0); // 画面の表示を止めない
    }
    if (onProgress) onProgress('write', 0, 1);
    const blob = await PdfDoc.rewrite(doc, { replace });
    return { blob, images: usage.size, shrunk: replace.size };
  }

  // 強力：ページを丸ごと画像にする。pdf は PdfView.open で開いたもの（ページの大きさはそのまま）
  async function rasterize(pdf, { dpi = 150, quality = 0.7, maxPixels = 25e6, title = '', onProgress } = {}) {
    const pages = Array.from({ length: pdf.numPages }, (_, i) => i);
    return buildPdf(pages, { paper: 'fit', margin: 0 }, title, {
      onPage: i => onProgress && onProgress('page', i, pages.length),
      encode: async i => {
        const r = await PdfView.renderPage(pdf, i, { scale: dpi / 72, maxPixels });
        const img = jpegFromCanvas(r.canvas, quality);
        r.canvas.width = r.canvas.height = 0;
        img.page = PdfView.pageSize(r.view, r.rotation);
        return img;
      },
    });
  }

  return { compress, rasterize, shrinkImage };
})();
