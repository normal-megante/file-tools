'use strict';
// =========================================================
// PDF に書き込む部品（文字・チェックなどの記号・白塗り・手書き）
//   書き込んだもの（items）から「ページに書き足す命令」を作り、PdfDoc.rewrite で保存する
//   ・座標は「画面で見えている向きのページ」の左上が原点、下向きが +y、単位は pt（1/72 インチ）
//   ・文字は、見た目を画像（透明度つき）で貼り、その上に検索・コピー用の見えない文字を重ねる
//     （どの端末・どの PDF アプリでも同じに見えて、日本語のフォントを PDF に入れなくて済む）
//   ・記号・手書き・白塗りは線や面（図形）で描く（拡大してもきれい）
// common.js・pdf.js（js/pdf.js）・pdfdoc.js のあとに読み込む
// =========================================================
const PdfWrite = (() => {
  const { PName, PStr, PStream, PNew } = PdfDoc.types;
  const N = n => new PName(n);
  const ascii = s => new TextEncoder().encode(s);
  const fmt = v => String(Math.round(v * 1000) / 1000 || 0);

  const FONTS = {
    gothic: '"Hiragino Sans", "Hiragino Kaku Gothic ProN", "Yu Gothic Medium", "Yu Gothic", YuGothic, Meiryo, "Noto Sans JP", "Noto Sans CJK JP", sans-serif',
    mincho: '"Hiragino Mincho ProN", "Yu Mincho", YuMincho, "Noto Serif JP", "Noto Serif CJK JP", serif',
  };
  const COLORS = { black: '#000000', red: '#d32f2f', blue: '#1f4fd1' };
  const LINE = 1.3;  // 行の高さ（文字の大きさの何倍か）
  const PAD = 0.15;  // 文字の左右の余白（文字の大きさの何倍か）
  const K = 4;       // 文字の画像の細かさ（1pt あたり何画素か。4 なら 288 dpi）
  const rgb = hex => [1, 3, 5].map(i => parseInt(hex.slice(i, i + 2), 16) / 255);
  const colorOf = it => COLORS[it.color] || COLORS.black;

  /* ---------- 文字 ---------- */
  let mctx = null;
  // 文字の並び方（pt）。w・h は余白こみの枠の大きさ、base は行の上から基準線までの距離
  function textLayout(it) {
    if (!mctx) mctx = newCanvas(1, 1).getContext('2d');
    mctx.font = `100px ${FONTS[it.font] || FONTS.gothic}`;
    const lines = String(it.text).split('\n'), size = it.size, lineH = size * LINE, pad = size * PAD;
    const widths = lines.map(l => mctx.measureText(l).width / 100 * size);
    const m = mctx.measureText('あAg');
    const asc = (m.fontBoundingBoxAscent || 88) / 100 * size, desc = (m.fontBoundingBoxDescent || 12) / 100 * size;
    return { lines, widths, lineH, pad, base: (lineH - asc - desc) / 2 + asc, w: Math.max(size * 0.6, ...widths) + 2 * pad, h: lines.length * lineH };
  }
  // 文字を ctx に描く（s：1pt あたり何画素か。左上が枠の左上）
  function drawText(ctx, it, s, fill) {
    const L = textLayout(it);
    ctx.font = `${it.size * s}px ${FONTS[it.font] || FONTS.gothic}`;
    ctx.fillStyle = fill || colorOf(it);
    ctx.textBaseline = 'alphabetic';
    L.lines.forEach((line, i) => ctx.fillText(line, L.pad * s, (i * L.lineH + L.base) * s));
    return L;
  }
  // 文字の画像：色は一色で、形は透明度（SMask）で表す
  async function textImage(it) {
    const L = textLayout(it), k = Math.min(K, 12000 / Math.max(L.w, L.h, 1)); // canvas の大きさの限界をこえないように
    const W = Math.max(1, Math.round(L.w * k)), H = Math.max(1, Math.round(L.h * k));
    const c = newCanvas(W, H), ctx = c.getContext('2d', { willReadFrequently: true });
    drawText(ctx, it, k, '#000');
    const px = ctx.getImageData(0, 0, W, H).data, alpha = new Uint8Array(W * H);
    for (let i = 0, j = 3; i < alpha.length; i++, j += 4) alpha[i] = px[j];
    c.width = c.height = 0;
    const mask = new PStream(new Map([['Type', N('XObject')], ['Subtype', N('Image')], ['Width', W], ['Height', H],
      ['ColorSpace', N('DeviceGray')], ['BitsPerComponent', 8], ['Filter', N('FlateDecode')],
      ['DecodeParms', new Map([['Predictor', 15], ['Colors', 1], ['BitsPerComponent', 8], ['Columns', W]])]]),
      await deflate(pngFilter(alpha, W, H, 1)));
    const [r, g, b] = rgb(colorOf(it)).map(v => Math.round(v * 255)), solid = new Uint8Array(W * H * 3);
    for (let i = 0; i < solid.length; i += 3) { solid[i] = r; solid[i + 1] = g; solid[i + 2] = b; }
    const img = new PStream(new Map([['Type', N('XObject')], ['Subtype', N('Image')], ['Width', W], ['Height', H],
      ['ColorSpace', N('DeviceRGB')], ['BitsPerComponent', 8], ['Filter', N('FlateDecode')], ['SMask', mask]]), await deflate(solid));
    return { img, L };
  }
  const hex4 = c => c.toString(16).toUpperCase().padStart(4, '0');
  const hexText = s => { let h = ''; for (let i = 0; i < s.length; i++) h += hex4(s.charCodeAt(i)); return h; };
  // 検索・コピー用の見えない文字のフォント（フォント本体は入れない。日本語の標準の名前だけ）
  function hiddenFont(texts) {
    const codes = new Set();
    for (const it of texts) for (let i = 0; i < it.text.length; i++) { const c = it.text.charCodeAt(i); if (c !== 10) codes.add(c); }
    const list = [...codes].sort((a, b) => a - b);
    let body = '';
    for (let i = 0; i < list.length; i += 100) {
      const chunk = list.slice(i, i + 100);
      body += `${chunk.length} beginbfchar\n${chunk.map(c => `<${hex4(c)}> <${hex4(c)}>`).join('\n')}\nendbfchar\n`;
    }
    const toUni = '/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n' +
      '/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n' +
      `1 begincodespacerange\n<0000> <FFFF>\nendcodespacerange\n${body}endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n`;
    const S = s => new PStr(ascii(s), false);
    const desc = new Map([['Type', N('FontDescriptor')], ['FontName', N('HeiseiKakuGo-W5')], ['Flags', 4], ['FontBBox', [-92, -250, 1010, 922]],
      ['ItalicAngle', 0], ['Ascent', 880], ['Descent', -120], ['CapHeight', 737], ['StemV', 114]]);
    const cid = new Map([['Type', N('Font')], ['Subtype', N('CIDFontType0')], ['BaseFont', N('HeiseiKakuGo-W5')],
      ['CIDSystemInfo', new Map([['Registry', S('Adobe')], ['Ordering', S('Japan1')], ['Supplement', 2]])], ['FontDescriptor', new PNew(desc)], ['DW', 1000]]);
    return new PNew(new Map([['Type', N('Font')], ['Subtype', N('Type0')], ['BaseFont', N('HeiseiKakuGo-W5-UniJIS-UCS2-H')],
      ['Encoding', N('UniJIS-UCS2-H')], ['DescendantFonts', [new PNew(cid)]], ['ToUnicode', new PStream(new Map(), ascii(toUni))]]));
  }

  /* ---------- 記号・手書き（線） ---------- */
  // 線の形：[['M', x, y], ['L', x, y], ['C', x1, y1, x2, y2, x, y], ['Z']]（枠の左上からの pt）
  const markWidth = it => Math.max(0.8, it.w * 0.09);
  function markPath(it) {
    const w = it.w, h = it.h, P = (x, y) => [x * w, y * h];
    if (it.type === 'check') return [['M', ...P(0.12, 0.55)], ['L', ...P(0.4, 0.84)], ['L', ...P(0.9, 0.16)]];
    if (it.type === 'cross') return [['M', ...P(0.18, 0.18)], ['L', ...P(0.82, 0.82)], ['M', ...P(0.82, 0.18)], ['L', ...P(0.18, 0.82)]];
    const lw = markWidth(it), rx = w / 2 - lw / 2, ry = h / 2 - lw / 2, cx = w / 2, cy = h / 2, k = 0.5523; // 丸（4本の曲線）
    return [['M', cx + rx, cy], ['C', cx + rx, cy + k * ry, cx + k * rx, cy + ry, cx, cy + ry], ['C', cx - k * rx, cy + ry, cx - rx, cy + k * ry, cx - rx, cy],
      ['C', cx - rx, cy - k * ry, cx - k * rx, cy - ry, cx, cy - ry], ['C', cx + k * rx, cy - ry, cx + rx, cy - k * ry, cx + rx, cy], ['Z']];
  }
  // 手書き：点と点の間をなめらかな曲線でつなぐ（点は「ページの座標」で持っている）
  function inkPath(it) {
    const cmds = [];
    for (const s of it.strokes) {
      const n = s.length / 2, X = i => s[2 * i] - it.x, Y = i => s[2 * i + 1] - it.y;
      if (!n) continue;
      cmds.push(['M', X(0), Y(0)]);
      if (n === 1) { cmds.push(['L', X(0) + 0.01, Y(0)]); continue; } // 点
      let px = X(0), py = Y(0);
      for (let i = 1; i < n - 1; i++) {
        const qx = X(i), qy = Y(i), ex = (qx + X(i + 1)) / 2, ey = (qy + Y(i + 1)) / 2;
        cmds.push(['C', px + 2 / 3 * (qx - px), py + 2 / 3 * (qy - py), ex + 2 / 3 * (qx - ex), ey + 2 / 3 * (qy - ey), ex, ey]);
        px = ex; py = ey;
      }
      cmds.push(['L', X(n - 1), Y(n - 1)]);
    }
    return cmds;
  }
  const shape = it => (it.type === 'ink' ? { cmds: inkPath(it), lw: it.lw } : { cmds: markPath(it), lw: markWidth(it) });
  const r2 = v => Math.round(v * 100) / 100;
  const svgPath = cmds => cmds.map(c => c[0] + c.slice(1).map(r2).join(' ')).join('');
  const pdfPath = (cmds, ox, oy) => cmds.map(c => {
    const p = (i) => `${fmt(c[i] + ox)} ${fmt(c[i + 1] + oy)}`;
    return c[0] === 'M' ? p(1) + ' m' : c[0] === 'L' ? p(1) + ' l' : c[0] === 'C' ? `${p(1)} ${p(3)} ${p(5)} c` : 'h';
  }).join('\n');

  /* ---------- 保存 ---------- */
  // doc：PdfDoc.open で開いたもの、pageInfo[i].inv：画面の座標 → PDF の座標 の変換（pdf.js の viewport の逆）
  async function build(doc, pageInfo, items, { onProgress } = {}) {
    const byPage = new Map();
    for (const it of items) { if (!byPage.has(it.page)) byPage.set(it.page, []); byPage.get(it.page).push(it); }
    const tag = Math.random().toString(36).slice(2, 6); // 元のページの部品の名前とぶつからないように
    const texts = items.filter(it => it.type === 'text' && it.text.trim());
    const font = texts.length ? hiddenFont(texts) : null;
    const edits = new Map();
    let done = 0, nImg = 0;
    for (const [pi, list] of [...byPage].sort((a, b) => a[0] - b[0])) {
      const ops = ['q', pageInfo[pi].inv.map(fmt).join(' ') + ' cm'], xobj = new Map(), fonts = new Map();
      for (const it of list) {
        if (onProgress) onProgress(done++, items.length);
        if (it.type === 'white') ops.push(`q 1 1 1 rg ${fmt(it.x)} ${fmt(it.y)} ${fmt(it.w)} ${fmt(it.h)} re f Q`);
        else if (it.type === 'text') {
          if (!it.text.trim()) continue;
          const name = `W${tag}i${nImg++}`, { img, L } = await textImage(it);
          xobj.set(name, img);
          ops.push(`q ${fmt(L.w)} 0 0 ${fmt(-L.h)} ${fmt(it.x)} ${fmt(it.y + L.h)} cm /${name} Do Q`);
          // 見えない文字（3 Tr）。行の幅が画像の文字と合うように横の伸び縮み（Tz）で合わせる
          const fn = `W${tag}f`;
          fonts.set(fn, font);
          ops.push(`BT 3 Tr /${fn} ${fmt(it.size)} Tf`);
          L.lines.forEach((line, i) => {
            if (!line.length) return;
            ops.push(`${fmt(100 * L.widths[i] / (line.length * it.size))} Tz 1 0 0 -1 ${fmt(it.x + L.pad)} ${fmt(it.y + i * L.lineH + L.base)} Tm <${hexText(line)}> Tj`);
          });
          ops.push('ET');
        } else {
          const { cmds, lw } = shape(it), [r, g, b] = rgb(colorOf(it));
          ops.push(`q ${fmt(r)} ${fmt(g)} ${fmt(b)} RG ${fmt(lw)} w 1 J 1 j`, pdfPath(cmds, it.x, it.y), 'S Q');
        }
        await wait(0);
      }
      ops.push('Q');
      edits.set(pi, { content: ascii(ops.join('\n') + '\n'), resources: { XObject: xobj, Font: fonts } });
    }
    return PdfDoc.rewrite(doc, { pageEdits: edits });
  }

  return { FONTS, COLORS, textLayout, drawText, shape, svgPath, markWidth, build };
})();
