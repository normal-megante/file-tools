'use strict';
// =========================================================
// PDF のページを絵にする部品（pdf.js を使う）
//   ・pdf.js は必要になってから読み込み、裏で読む係（ワーカー）は1つを使い回す
//   ・ページを canvas に描く（renderPage）、カード用の小さな見本（thumb）
//   ・見えているカードだけ、順番に見本を描く（ThumbQueue）
// common.js のあとに読み込む
// =========================================================
const PdfView = (() => {
  const DIR = new URL('vendor/pdfjs/', document.baseURI).href;
  let libPromise = null, worker = null;

  function lib() {
    if (!libPromise) libPromise = import(DIR + 'pdf.min.mjs').then(l => {
      l.GlobalWorkerOptions.workerSrc = DIR + 'pdf.worker.min.mjs';
      worker = new l.PDFWorker(); // ファイルが何個あっても1つで済ませる（スマホのメモリ節約）
      return l;
    });
    return libPromise;
  }

  // PDF を開く。pdf.js は渡したデータを持っていってしまうので、写しを渡す。
  // onPassword(again) はパスワードを返す（やめるなら null）
  async function open(bytes, { onPassword } = {}) {
    const l = await lib();
    const task = l.getDocument({
      worker, data: bytes.slice(),
      cMapUrl: DIR + 'cmaps/', cMapPacked: true,
      standardFontDataUrl: DIR + 'standard_fonts/', wasmUrl: DIR + 'wasm/', iccUrl: DIR + 'iccs/',
      enableXfa: false,
    });
    if (onPassword) {
      task.onPassword = (update, reason) => {
        const pw = onPassword(reason === l.PasswordResponses.INCORRECT_PASSWORD);
        if (pw == null) task.destroy(); else update(pw);
      };
    }
    return task.promise;
  }
  // 開いた PDF を閉じてメモリを空ける（この版の pdf.js では doc.destroy() はなく、読み込み係を止める）
  const close = doc => doc.loadingTask.destroy().catch(() => {});

  // ページを描く。rotate はページ本来の向きに足す角度。
  // 大きさは scale（1 = 72dpi）か maxSide（長い辺のピクセル数）。maxPixels を超えるときは縮める
  async function renderPage(doc, index, { rotate = 0, scale, maxSide, maxPixels = Infinity } = {}) {
    const page = await doc.getPage(index + 1);
    try {
      const r = (((page.rotate + rotate) % 360) + 360) % 360, v1 = page.getViewport({ scale: 1, rotation: r });
      let s = scale || maxSide / Math.max(v1.width, v1.height);
      const px = v1.width * v1.height * s * s;
      if (px > maxPixels) s *= Math.sqrt(maxPixels / px);
      s = Math.min(s, 16000 / Math.max(v1.width, v1.height)); // ブラウザの canvas の限界
      const vp = page.getViewport({ scale: s, rotation: r });
      const canvas = newCanvas(Math.max(1, Math.floor(vp.width)), Math.max(1, Math.floor(vp.height)));
      // 注釈（フォームに入力された文字なども）を絵として一緒に描く。既定ではフォームは描かれない
      await page.render({ canvas, viewport: vp, annotationMode: (await lib()).AnnotationMode.ENABLE }).promise;
      return { canvas, view: page.view, rotation: r, scale: s, shrunk: px > maxPixels };
    } finally { page.cleanup(); }
  }

  // カード用の小さな見本（JPEG の data URL。canvas のまま持つよりメモリが少なくて済む）
  async function thumb(doc, index, rotate = 0) {
    const r = await renderPage(doc, index, { rotate, maxSide: Math.round(190 * Math.min(2, window.devicePixelRatio || 1)) });
    const url = r.canvas.toDataURL('image/jpeg', 0.85);
    r.canvas.width = r.canvas.height = 0;
    return { url, view: r.view, rotation: r.rotation };
  }

  // ページの大きさ（pt）。rotation で縦横を入れ替える
  function pageSize(view, rotation) {
    const w = view[2] - view[0], h = view[3] - view[1];
    return rotation % 180 ? [h, w] : [w, h];
  }
  // 紙の大きさの名前（A4 など）
  const PAPERS = [['A3', 297, 420], ['A4', 210, 297], ['A5', 148, 210], ['B4', 257, 364], ['B5', 182, 257], ['レター', 216, 279], ['リーガル', 216, 356]];
  function paperName(view, rotation) {
    const [w, h] = pageSize(view, rotation).map(v => v * 25.4 / 72);
    const p = PAPERS.find(([, a, b]) => (Math.abs(w - a) < 3 && Math.abs(h - b) < 3) || (Math.abs(w - b) < 3 && Math.abs(h - a) < 3));
    return p ? p[0] + (w > h ? ' 横' : '') : `${Math.round(w)}×${Math.round(h)} mm`;
  }

  // 見えているもの（とその少し先）だけ、1つずつ順番に描く
  //   draw(it): 描く処理、stale(it): 描き直しが必要か。it.visible はここで更新する
  class ThumbQueue {
    constructor({ draw, stale, rootMargin = '500px 0px' }) {
      this.draw = draw; this.stale = stale; this.want = new Set(); this.busy = false; this.items = new WeakMap();
      this.io = new IntersectionObserver(entries => {
        for (const e of entries) {
          const it = this.items.get(e.target); if (!it) continue;
          it.visible = e.isIntersecting;
          if (it.visible && this.stale(it)) this.want.add(it);
        }
        this.pump();
      }, { rootMargin });
    }
    observe(el, it) { this.items.set(el, it); this.io.observe(el); }
    unobserve(el) { this.io.unobserve(el); }
    queue(it) { this.want.add(it); this.pump(); }
    async pump() {
      if (this.busy) return;
      this.busy = true;
      try {
        while (this.want.size) {
          let it = null;
          for (const x of this.want) if (x.visible) { it = x; break; }
          if (!it) break;
          this.want.delete(it);
          if (this.stale(it)) await this.draw(it);
        }
      } finally { this.busy = false; }
    }
  }

  return { lib, open, close, renderPage, thumb, pageSize, paperName, ThumbQueue };
})();
