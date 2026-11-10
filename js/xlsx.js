'use strict';
// =========================================================
// Excel（.xlsx）を読む（自作）。js/zip.js・js/drawingml.js・js/numfmt.js を先に読む
//   const book = await Xlsx.open(zip)
//     book.sheets：[{ name, state（visible / hidden / veryHidden）, kind（sheet / chart / other）, tab（色） }]
//     book.active：開いたときに出ていたシート、book.sheet(i) → シートの中身（下）、book.dispose()
//     book.style(s)：セルの書式（文字・塗り・罫線・配置・表示形式）、book.text(cell, style, fit)：セルに出す文字
//   const sh = await book.sheet(i)
//     sh.cell(r, c) → { v, t（n 数・s 文字・r 書式つき文字・b TRUE/FALSE・e エラー・z 空）, s（書式の番号）, runs }
//     sh.colW(c)・sh.rowH(r)：100% のときの幅・高さ（px）。隠した列・行は 0
//     sh.nR・sh.nC：表示する行・列の数。sh.merges・sh.mergeAt(r, c)、sh.freeze、sh.grid（枠線を出すか）
//     sh.cf（条件付き書式の結果）・sh.notes（メモ）・sh.links（リンク）・sh.anchors（図）・sh.formula(r, c)（数式の文字）
//   表に描くのは js/xlgrid.js
// =========================================================
const Xlsx = (() => {
  const { parseXml, kid, kids, pathOf, at, atNum, truthy, rAttr, hex, rgb2hsl, hsl2rgb, SERIF } = DML;

  /* ---------- セルの番地 ---------- */
  const MAXR = 1048575, MAXC = 16383;
  function colName(c) { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + (n - 1) % 26) + s; return s; }
  const addr = (r, c) => colName(c) + (r + 1);
  function colNum(s) { let n = 0; for (let i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) & 31); return n - 1; }
  // "A1:C3"・"B2"・"A:A"・"2:5" → { r0, c0, r1, c1 }
  function range(s) {
    const [a, b = a] = String(s || '').replace(/\$/g, '').split(':');
    const one = x => {
      let m = /^([A-Za-z]{1,3})(\d+)$/.exec(x);
      if (m) return { r: +m[2] - 1, c: colNum(m[1]) };
      if ((m = /^([A-Za-z]{1,3})$/.exec(x))) return { c: colNum(m[1]) };
      if ((m = /^(\d+)$/.exec(x))) return { r: +m[1] - 1 };
      return null;
    };
    const p = one(a), q = one(b);
    if (!p || !q) return null;
    const r0 = p.r != null ? p.r : 0, c0 = p.c != null ? p.c : 0, r1 = q.r != null ? q.r : MAXR, c1 = q.c != null ? q.c : MAXC;
    return { r0: Math.min(r0, r1), c0: Math.min(c0, c1), r1: Math.max(r0, r1), c1: Math.max(c0, c1) };
  }
  const rangeList = s => String(s || '').trim().split(/\s+/).map(range).filter(Boolean);
  const key = (r, c) => r * 16384 + c;

  /* ---------- 文字の小道具 ---------- */
  const ENT = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" };
  const unxml = s => (s.indexOf('&') < 0 ? s : s.replace(/&(#[xX][0-9a-fA-F]+|#\d+|\w+);/g, (m, e) => (e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : +e.slice(1)) : ENT[e] != null ? ENT[e] : m)));
  // Excel の文字の中の _x000D_ など（改行などの字を書きかえたもの）
  const unesc = s => (s.indexOf('_x') < 0 ? s : s.replace(/_x([0-9A-Fa-f]{4})_/g, (m, x) => String.fromCharCode(parseInt(x, 16))));
  function attrs(s) {
    const o = {}, re = /([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    let m;
    while ((m = re.exec(s))) o[m[1].replace(/^[\w.-]+:/, '')] = m[2] != null ? m[2] : m[3];
    return o;
  }
  // 文字（t・r）を取り出す。ふりがな（rPh）は使わない
  function richOf(el) {
    const runs = [];
    let fmt = false;
    for (const c of kids(el)) {
      if (c.localName === 't') runs.push({ t: unesc(c.textContent) });
      else if (c.localName === 'r') {
        const t = kid(c, 't'), f = fontOf(kid(c, 'rPr'));
        runs.push({ t: unesc(t ? t.textContent : ''), f });
        if (f) fmt = true;
      }
    }
    const text = runs.map(r => r.t).join('');
    return fmt ? { text, runs } : text;
  }

  /* ---------- 色 ---------- */
  // 番号の色（[Color8] = 黒 …）。0〜7 は 8〜15 と同じ
  const PALETTE = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', ...NumFmt.PALETTE];
  const THEME_ORDER = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
  function colorSpec(el) {
    if (!el) return null;
    const o = {};
    if (truthy(at(el, 'auto'))) o.auto = true;
    const rgb = at(el, 'rgb');
    if (rgb) o.rgb = rgb.slice(-6);
    if (at(el, 'theme') != null) o.theme = +at(el, 'theme');
    if (at(el, 'indexed') != null) o.indexed = +at(el, 'indexed');
    if (+at(el, 'tint')) o.tint = +at(el, 'tint');
    return o;
  }
  const toHex = c => [c.r, c.g, c.b].map(v => Math.round(Math.max(0, Math.min(255, v))).toString(16).padStart(2, '0')).join('');
  // 明るさの調整（Excel は HLS の明るさを変える）
  function tintHex(h, t) {
    const [hh, s, l] = rgb2hsl(hex(h));
    return toHex(hsl2rgb(hh, s, t < 0 ? l * (1 + t) : l * (1 - t) + t));
  }

  /* ---------- 書式（styles.xml） ---------- */
  function fontOf(el) {
    if (!el) return null;
    const o = {};
    for (const c of kids(el)) {
      const n = c.localName, v = at(c, 'val');
      switch (n) {
        case 'name': case 'rFont': o.name = v || ''; break;
        case 'sz': o.sz = +v || 11; break;
        case 'b': case 'i': case 'strike': o[n] = v == null || !(v === '0' || v === 'false'); break;
        case 'u': o.u = v || 'single'; break;
        case 'color': o.color = colorSpec(c); break;
        case 'vertAlign': o.vert = v; break;
      }
    }
    return o;
  }
  // dxf：条件付き書式の塗り（塗りつぶしの色は bgColor に入っている）
  function fillOf(el, dxf) {
    if (!el) return null;
    const pf = kid(el, 'patternFill'), gf = kid(el, 'gradientFill');
    if (gf) {
      return { grad: { type: at(gf, 'type') || 'linear', deg: atNum(gf, 'degree', 0), l: atNum(gf, 'left', 0), r: atNum(gf, 'right', 0), t: atNum(gf, 'top', 0), b: atNum(gf, 'bottom', 0),
        stops: kids(gf, 'stop').map(s => ({ pos: atNum(s, 'position', 0), color: colorSpec(kid(s, 'color')) })) } };
    }
    if (!pf) return null;
    const type = at(pf, 'patternType'), fg = colorSpec(kid(pf, 'fgColor')), bg = colorSpec(kid(pf, 'bgColor'));
    if (dxf && (!type || type === 'solid')) return bg || fg ? { solid: bg || fg } : null;
    if (!type || type === 'none') return null;
    if (type === 'solid') return { solid: fg || { indexed: 64 } };
    return { pattern: type, fg: fg || { indexed: 64 }, bg: bg || { indexed: 65 } };
  }
  function borderOf(el) {
    if (!el) return null;
    const o = {};
    for (const [k, names] of [['l', ['left', 'start']], ['r', ['right', 'end']], ['t', ['top']], ['b', ['bottom']], ['d', ['diagonal']]]) {
      const e = names.map(n => kid(el, n)).find(Boolean), st = at(e, 'style');
      if (st && st !== 'none') o[k] = { style: st, color: colorSpec(kid(e, 'color')) };
    }
    if (o.d) { o.du = truthy(at(el, 'diagonalUp')); o.dd = truthy(at(el, 'diagonalDown')); if (!o.du && !o.dd) delete o.d; }
    return o;
  }
  function xfOf(el) {
    const a = kid(el, 'alignment'), al = {};
    if (a) {
      const g = n => at(a, n);
      if (g('horizontal')) al.h = g('horizontal');
      if (g('vertical')) al.v = g('vertical');
      if (truthy(g('wrapText'))) al.wrap = true;
      if (truthy(g('shrinkToFit'))) al.shrink = true;
      if (+g('indent')) al.indent = +g('indent');
      if (+g('textRotation')) al.rot = +g('textRotation');
    }
    return { fmt: atNum(el, 'numFmtId', 0), font: atNum(el, 'fontId', 0), fill: atNum(el, 'fillId', 0), border: atNum(el, 'borderId', 0), al, quote: truthy(at(el, 'quotePrefix')) };
  }
  // 罫線の種類 → 太さ（px）・点線の形・強さ（重なったときに強いほうを描く）
  const LINE = {
    thin: { w: 1, k: 1 }, hair: { w: 1, dash: [1, 1], k: 0.5 }, dotted: { w: 1, dash: [1, 2], k: 0.6 }, dashed: { w: 1, dash: [3, 1], k: 0.7 },
    dashDot: { w: 1, dash: [3, 1, 1, 1], k: 0.7 }, dashDotDot: { w: 1, dash: [3, 1, 1, 1, 1, 1], k: 0.7 },
    medium: { w: 2, k: 2 }, mediumDashed: { w: 2, dash: [6, 2], k: 1.8 }, mediumDashDot: { w: 2, dash: [6, 2, 2, 2], k: 1.8 },
    mediumDashDotDot: { w: 2, dash: [6, 2, 2, 2, 2, 2], k: 1.8 }, slantDashDot: { w: 2, dash: [8, 1, 3, 1], k: 1.9 },
    thick: { w: 3, k: 3 }, double: { w: 3, dbl: true, k: 2.9 },
  };

  /* ---------- 書体 ---------- */
  const ALIAS = [[/^(ＭＳ|MS)\s*(Ｐ|P)?\s*(明朝|Mincho)/i, ['MS Mincho', 'Hiragino Mincho ProN']], [/^(ＭＳ|MS)\s*(Ｐ|P)\s*(ゴシック|Gothic)/i, ['MS PGothic', 'Hiragino Sans']],
    [/^(ＭＳ|MS)\s*(ゴシック|Gothic)/i, ['MS Gothic', 'Hiragino Sans']], [/^(游明朝|Yu Mincho)/i, ['Yu Mincho', 'YuMincho', 'Hiragino Mincho ProN']],
    [/^(游ゴシック|Yu Gothic)/i, ['Yu Gothic', 'YuGothic', 'Hiragino Sans']], [/^(メイリオ|Meiryo)$/i, ['Meiryo']], [/^Meiryo UI$/i, ['Meiryo UI', 'Meiryo']]];
  function family(name) {
    const list = [];
    if (name) { list.push(name); for (const [re, al] of ALIAS) if (re.test(name)) for (const a of al) if (!list.includes(a)) list.push(a); }
    const serif = /明朝|mincho|教科書|楷書|行書/i.test(name || '') || (SERIF.test(name || '') && !/sans|gothic|ゴシック/i.test(name || ''));
    return list.map(n => `"${n.replace(/["\\]/g, '')}"`).concat(serif ? 'serif' : 'sans-serif').join(',');
  }
  // 標準の書体の「数字の幅」（列の幅の単位。Excel で測った）。表にない書体は、この端末で測る
  const MDW = [[/^(游ゴシック|Yu Gothic)$/i, { 11: 8 }], [/^Calibri$/i, { 11: 7, 10: 7 }], [/^(ＭＳ\s*Ｐゴシック|MS PGothic)$/i, { 11: 7 }], [/^(メイリオ|Meiryo)$/i, { 10: 8 }], [/^Arial$/i, { 10: 7, 11: 8 }]];
  let measureCtx = null;
  function mdwOf(font) {
    const name = font.name || '', sz = font.sz || 11;
    for (const [re, t] of MDW) if (re.test(name) && t[sz]) return t[sz];
    if (typeof document === 'undefined') return Math.max(1, Math.round(sz * 0.68));
    // 大きな字で測って縮める（Windows の Chrome は ＭＳ ゴシックなどを小さい字で広めに並べるため）
    measureCtx = measureCtx || document.createElement('canvas').getContext('2d');
    measureCtx.font = `100px ${family(name)}`;
    let w = 0;
    for (let d = 0; d < 10; d++) w = Math.max(w, measureCtx.measureText(String(d)).width);
    return Math.max(1, Math.round(w / 100 * sz * 4 / 3));
  }
  // 書体の「1 行」の高さ（字の大きさの何倍か）。Excel の標準の行の高さ・折り返した行の高さ（Excel で測った）
  const ROWH = [[/^(游ゴシック|Yu Gothic|游明朝|Yu Mincho)/i, 1.66], [/^Calibri$/i, 1.33], [/^(ＭＳ|MS)\s*(Ｐ|P)?\s*(ゴシック|Gothic|明朝|Mincho)$/i, 1.2],
    [/^(メイリオ|Meiryo)$/i, 1.6], [/^Meiryo UI$/i, 1.33], [/^Arial$/i, 1.28]];
  const lineRatio = name => (ROWH.find(([re]) => re.test(name || '')) || [0, 1.33])[1];
  const defaultRowPt = font => Math.round((font.sz || 11) * lineRatio(font.name) / 0.75) * 0.75;

  /* ---------- ブックを開く ---------- */
  async function open(zip) {
    const P = DML.pkg(zip);
    const rootRels = await P.rels('');
    let wbPath = 'xl/workbook.xml';
    for (const r of rootRels.values()) if (r.type === 'officeDocument') wbPath = r.target;
    const wb = await P.xml(wbPath);
    if (!wb) throw new Error('workbook.xml を読めません');
    const W = wb.documentElement, rels = await P.rels(wbPath);
    const date1904 = truthy(at(kid(W, 'workbookPr'), 'date1904'));
    const active = atNum(pathOf(W, 'bookViews', 'workbookView'), 'activeTab', 0);
    const sheets = kids(kid(W, 'sheets'), 'sheet').map(s => {
      const rel = rels.get(rAttr(s, 'id'));
      const type = rel ? rel.type : '';
      return { name: at(s, 'name') || '', state: at(s, 'state') || 'visible', path: rel && rel.target, kind: type === 'worksheet' ? 'sheet' : type === 'chartsheet' ? 'chart' : 'other' };
    });
    const themeRel = P.relOfType(rels, 'theme');
    const theme = themeRel ? await P.theme(themeRel.target) : { colors: {}, major: {}, minor: {} };
    // 共有の文字
    let sst = [];
    const sr = P.relOfType(rels, 'sharedStrings');
    if (sr) { const d = await P.xml(sr.target); if (d) sst = kids(d.documentElement, 'si').map(richOf); }
    // 書式
    const stRel = P.relOfType(rels, 'styles'), sd = stRel ? await P.xml(stRel.target) : null, S = sd && sd.documentElement;
    const palette = PALETTE.slice();
    const ic = pathOf(S, 'colors', 'indexedColors');
    if (ic) kids(ic, 'rgbColor').forEach((e, i) => { const v = (at(e, 'rgb') || '').slice(-6); if (/^[0-9a-f]{6}$/i.test(v)) palette[i] = v; });
    const numFmts = {};
    for (const nf of kids(kid(S, 'numFmts'), 'numFmt')) { const f16 = [...nf.attributes].find(a => a.localName === 'formatCode16'); numFmts[at(nf, 'numFmtId')] = (f16 && f16.value) || at(nf, 'formatCode') || 'General'; }
    const fonts = kids(kid(S, 'fonts'), 'font').map(fontOf);
    const fills = kids(kid(S, 'fills'), 'fill').map(f => fillOf(f, false));
    const borders = kids(kid(S, 'borders'), 'border').map(borderOf);
    const xfs = kids(kid(S, 'cellXfs'), 'xf').map(xfOf);
    const styleXfs = kids(kid(S, 'cellStyleXfs'), 'xf').map(xfOf);
    const dxfEls = kids(kid(S, 'dxfs'), 'dxf');
    if (!fonts.length) fonts.push({ name: 'Calibri', sz: 11 });
    if (!xfs.length) xfs.push({ fmt: 0, font: 0, fill: 0, border: 0, al: {} });
    const normalFont = { name: 'Calibri', sz: 11, ...fonts[(styleXfs[0] || xfs[0]).font] || fonts[0] };
    const mdw = mdwOf(normalFont);

    // 色の指定 → "#RRGGBB"（auto：自動の色）
    function color(spec, auto = '#000000') {
      if (!spec) return auto;
      let h = null;
      if (spec.rgb && /^[0-9a-f]{6}$/i.test(spec.rgb)) h = spec.rgb;
      else if (spec.theme != null) h = theme.colors[THEME_ORDER[spec.theme]] || (spec.theme === 0 ? 'FFFFFF' : spec.theme === 1 ? '000000' : null);
      else if (spec.indexed != null) { if (spec.indexed === 64) return auto || '#000000'; h = spec.indexed === 65 ? 'FFFFFF' : palette[spec.indexed]; }
      if (!h) return auto;
      if (spec.tint) h = tintHex(h, spec.tint);
      return '#' + h.toUpperCase();
    }
    const fontCss = f => ({ ...f, fam: family(f.name), color: color(f.color, '#000000') });
    function resolveBorder(b) {
      if (!b) return null;
      const o = {};
      for (const k of ['l', 'r', 't', 'b', 'd']) if (b[k]) { const L = LINE[b[k].style] || LINE.thin; o[k] = { ...L, color: color(b[k].color, '#000000') }; }
      if (o.d) { o.du = b.du; o.dd = b.dd; }
      return Object.keys(o).length ? o : null;
    }
    function resolveFill(f) {
      if (!f) return null;
      if (f.solid) return { color: color(f.solid, null) };
      if (f.pattern) return { pattern: f.pattern, fg: color(f.fg, '#000000'), bg: color(f.bg, '#FFFFFF') };
      if (f.grad) return { grad: { ...f.grad, stops: f.grad.stops.map(s => ({ pos: s.pos, color: color(s.color, '#FFFFFF') })) } };
      return null;
    }
    const styleCache = [];
    function style(i) {
      i = i | 0;
      if (styleCache[i]) return styleCache[i];
      const x = xfs[i] || xfs[0];
      const st = {
        font: fontCss({ name: normalFont.name, sz: 11, ...fonts[x.font] || normalFont }),
        fill: resolveFill(fills[x.fill]),
        border: resolveBorder(borders[x.border]),
        al: x.al, quote: x.quote,
        fmt: numFmts[x.fmt] != null ? numFmts[x.fmt] : NumFmt.builtin(x.fmt),
      };
      return (styleCache[i] = st);
    }
    const dxfCache = [];
    function dxf(i) {
      if (dxfCache[i]) return dxfCache[i];
      const el = dxfEls[i], o = {};
      if (el) {
        const f = fontOf(kid(el, 'font'));
        if (f) { o.font = { ...f }; if (f.color) o.font.color = color(f.color, '#000000'); else delete o.font.color; }
        const fl = resolveFill(fillOf(kid(el, 'fill'), true));
        if (fl) o.fill = fl;
        const b = resolveBorder(borderOf(kid(el, 'border')));
        if (b) o.border = b;
        const nf = kid(el, 'numFmt');
        if (nf) o.fmt = at(nf, 'formatCode');
      }
      return (dxfCache[i] = o);
    }

    // セルに出す文字。fit(text) → その文字がセルに入るか（数だけに使う。入らなければ ### や小数を減らす）
    function text(cell, st, fit, fmt) {
      if (!cell) return null;
      const code = fmt || st.fmt;
      switch (cell.t) {
        case 'n': { const r = NumFmt.format(cell.v, code, { date1904, fit }); r.align = 'r'; r.num = true; return r; }
        case 's': case 'r': { const r = NumFmt.format(cell.v, code, {}); r.align = 'l'; return r; }
        case 'b': { const t = cell.v ? 'TRUE' : 'FALSE'; return { text: t, parts: [{ s: t }], color: null, align: 'c', kind: 'bool' }; }
        case 'e': return { text: cell.v, parts: [{ s: cell.v }], color: null, align: 'c', kind: 'error' };
      }
      return null;
    }

    const cache = new Map();
    const book = {
      sheets, active: Math.min(active, Math.max(0, sheets.length - 1)), date1904, theme, mdw, normalFont, P, sst,
      color, style, dxf, text, family,
      sheet(i) {
        if (!cache.has(i)) cache.set(i, loadSheet(book, sheets[i]).catch(e => { cache.delete(i); throw e; }));
        return cache.get(i);
      },
      dispose() { P.dispose(); cache.clear(); },
    };
    // 人の名前（スレッド形式のコメントを書いた人）
    const pr = P.relOfType(rels, 'person');
    book.persons = new Map();
    if (pr) { const d = await P.xml(pr.target); if (d) for (const p of kids(d.documentElement, 'person')) book.persons.set(at(p, 'id'), at(p, 'displayName') || ''); }
    return book;
  }

  /* ---------- シートを読む ---------- */
  async function loadSheet(book, info) {
    if (!info || !info.path) throw new Error('シートが見つかりません');
    const P = book.P;
    if (info.kind !== 'sheet') {
      // グラフだけのシートなど：グラフの場所と大きさ（ext は EMU）
      let title = '', chart = null, w = 0, hh = 0;
      try {
        const d = await P.xml(info.path), rels = await P.rels(info.path);
        const dr = d && kid(d.documentElement, 'drawing'), rel = dr && rels.get(rAttr(dr, 'id'));
        if (rel) {
          const dd = await P.xml(rel.target), drels = await P.rels(rel.target);
          const ch = dd && [...dd.getElementsByTagNameNS('*', 'chart')].find(e => rAttr(e, 'id'));
          const cr = ch && drels.get(rAttr(ch, 'id')), cd = cr && await P.xml(cr.target);
          const t = cd && pathOf(cd.documentElement, 'chart', 'title');
          if (t) title = [...t.getElementsByTagNameNS('*', 't')].map(x => x.textContent).join('');
          if (cd) chart = cr.target;
          let an = ch;
          while (an && !/Anchor$/.test(an.localName)) an = an.parentNode;
          const ext = an && kid(an, 'ext');
          if (ext) { w = atNum(ext, 'cx', 0) / 9525; hh = atNum(ext, 'cy', 0) / 9525; }
        }
      } catch (e) { /* 題名なし */ }
      return { kind: info.kind, name: info.name, title, chart, w: w || 868, h: hh || 614 };
    }
    const entry = P.zip.get(info.path);
    if (!entry) throw new Error('シートが見つかりません');
    const xml = await P.zip.text(entry);
    const sh = { kind: 'sheet', name: info.name, rows: [], merges: [], notes: new Map(), links: new Map(), cf: new Map(), anchors: [], shared: new Map(), maxR: -1, maxC: -1 };

    // セル（sheetData）は大きいので、正規表現で読む。ほかは DOM で
    const sdm = /<((?:[\w.-]+:)?)sheetData\b[^>]*?(\/?)>/.exec(xml);
    let rest = xml;
    if (sdm) {
      const start = sdm.index + sdm[0].length;
      let end = start, close = start;
      if (!sdm[2]) {
        close = xml.indexOf(`</${sdm[1]}sheetData>`, start);
        if (close < 0) close = xml.length;
        end = close + `</${sdm[1]}sheetData>`.length;
        readCells(sh, xml, start, close, book);
      }
      rest = xml.slice(0, sdm.index) + xml.slice(sdm[2] ? sdm.index + sdm[0].length : Math.min(end, xml.length));
    }
    const doc = parseXml(rest), R = doc && doc.documentElement;
    if (!R) throw new Error('シートを読めません');
    const rels = await P.rels(info.path);

    // 見た目の設定
    const sv = pathOf(R, 'sheetViews', 'sheetView');
    sh.grid = at(sv, 'showGridLines') !== '0' && at(sv, 'showGridLines') !== 'false';
    sh.headers = at(sv, 'showRowColHeaders') !== '0' && at(sv, 'showRowColHeaders') !== 'false';
    sh.zoom = atNum(sv, 'zoomScale', 100) / 100;
    sh.rtl = truthy(at(sv, 'rightToLeft'));
    const pane = kid(sv, 'pane'), tl = range(at(sv, 'topLeftCell') || 'A1');
    sh.freeze = { rows: 0, cols: 0, top: tl ? tl.r0 : 0, left: tl ? tl.c0 : 0 };
    if (pane && /^frozen/.test(at(pane, 'state') || '')) {
      sh.freeze.rows = Math.max(0, Math.round(atNum(pane, 'ySplit', 0)));
      sh.freeze.cols = Math.max(0, Math.round(atNum(pane, 'xSplit', 0)));
      const p = range(at(pane, 'topLeftCell') || '');
      sh.freeze.top = p ? Math.max(p.r0, sh.freeze.rows) : sh.freeze.rows;
      sh.freeze.left = p ? Math.max(p.c0, sh.freeze.cols) : sh.freeze.cols;
    }
    // 選んであったセル（作業中のペインの選択）
    const sels = kids(sv, 'selection'), act = at(pane, 'activePane');
    const sel = sels.find(x => (at(x, 'pane') || 'topLeft') === (act || 'topLeft')) || sels[sels.length - 1];
    const ac = sel && range(at(sel, 'activeCell') || '');
    sh.active = ac ? { r: ac.r0, c: ac.c0 } : null;
    const tab = pathOf(R, 'sheetPr', 'tabColor');
    if (tab) info.tab = book.color(colorSpec(tab), null);

    // 列の幅・行の高さ（100% のときの px）
    const fp = kid(R, 'sheetFormatPr'), mdw = book.mdw;
    const colPx = w => (w <= 0 ? 0 : Math.trunc(((256 * w + Math.trunc(128 / mdw)) / 256) * mdw));
    const defW = at(fp, 'defaultColWidth') != null ? colPx(atNum(fp, 'defaultColWidth', 8)) : Math.ceil((atNum(fp, 'baseColWidth', 8) * mdw + 5) / 8) * 8;
    const defH = Math.round((at(fp, 'defaultRowHeight') != null ? atNum(fp, 'defaultRowHeight', 15) : defaultRowPt(book.normalFont)) / 0.75);
    const zeroH = truthy(at(fp, 'zeroHeight'));
    sh.defW = defW; sh.defH = defH;
    const cols = [];
    for (const c of kids(kid(R, 'cols'), 'col')) {
      const mn = atNum(c, 'min', 1) - 1, mx = Math.min(MAXC, atNum(c, 'max', mn + 1) - 1);
      cols.push({ mn, mx, w: truthy(at(c, 'hidden')) ? 0 : at(c, 'width') != null ? colPx(atNum(c, 'width', 0)) : defW, s: at(c, 's') != null ? atNum(c, 's', 0) : null });
    }
    cols.sort((a, b) => a.mn - b.mn);
    const colOf = c => { for (const x of cols) { if (c < x.mn) return null; if (c <= x.mx) return x; } return null; };
    sh.colW = c => { const x = colOf(c); return x ? x.w : defW; };
    sh.colStyle = c => { const x = colOf(c); return x && x.s != null ? x.s : null; };
    sh.rowH = r => {
      const row = sh.rows[r];
      if (!row) return zeroH ? 0 : defH;
      if (row.hidden) return 0;
      return row.ht != null ? Math.round(row.ht / 0.75) : row.auto || (zeroH ? 0 : defH);
    };

    // 結合したセル
    sh.mergeMap = new Map();
    for (const m of kids(kid(R, 'mergeCells'), 'mergeCell')) {
      const g = range(at(m, 'ref'));
      if (!g || (g.r0 === g.r1 && g.c0 === g.c1)) continue;
      g.r1 = Math.min(g.r1, g.r0 + 20000); g.c1 = Math.min(g.c1, g.c0 + 2000);
      sh.merges.push(g);
      for (let r = g.r0; r <= g.r1; r++) for (let c = g.c0; c <= g.c1; c++) sh.mergeMap.set(key(r, c), g);
    }
    sh.mergeAt = (r, c) => sh.mergeMap.get(key(r, c)) || null;

    // リンク
    for (const hl of kids(kid(R, 'hyperlinks'), 'hyperlink')) {
      const g = range(at(hl, 'ref'));
      if (!g) continue;
      const rel = rels.get(rAttr(hl, 'id'));
      const link = { url: rel && rel.external ? rel.target : null, loc: at(hl, 'location') || null, tip: at(hl, 'tooltip') || '' };
      if (!link.url && !link.loc) continue;
      for (let r = g.r0; r <= Math.min(g.r1, g.r0 + 1000); r++) for (let c = g.c0; c <= Math.min(g.c1, g.c0 + 200); c++) sh.links.set(key(r, c), link);
    }

    // メモ（コメント）と、スレッド形式のコメント
    const cr = P.relOfType(rels, 'comments');
    if (cr) {
      const d = await P.xml(cr.target), C = d && d.documentElement;
      const authors = kids(kid(C, 'authors'), 'author').map(a => a.textContent);
      for (const cm of kids(kid(C, 'commentList'), 'comment')) {
        const g = range(at(cm, 'ref')), t = kid(cm, 'text');
        if (!g || !t) continue;
        const rt = richOf(t);
        sh.notes.set(key(g.r0, g.c0), { text: typeof rt === 'string' ? rt : rt.text, author: authors[atNum(cm, 'authorId', -1)] || '' });
      }
    }
    for (const r of rels.values()) {
      if (r.type !== 'threadedComment') continue;
      const d = await P.xml(r.target), C = d && d.documentElement;
      for (const tc of kids(C, 'threadedComment')) {
        const g = range(at(tc, 'ref'));
        if (!g) continue;
        const k = key(g.r0, g.c0), t = kid(tc, 'text'), who = book.persons.get(at(tc, 'personId')) || '';
        const prev = sh.notes.get(k);
        const line = `${who ? who + '：' : ''}${t ? t.textContent : ''}`;
        if (prev && prev.thread) prev.text += '\n' + line;
        else sh.notes.set(k, { text: line, author: '', thread: true });
      }
    }

    // 図（画像・図形・グラフ）
    const dr = kid(R, 'drawing'), drel = dr && rels.get(rAttr(dr, 'id'));
    if (drel && !drel.external) {
      const d = await P.xml(drel.target);
      sh.drawingPath = drel.target;
      sh.drawingRels = await P.rels(drel.target);
      const pos = el => el && { c: +((kid(el, 'col') || {}).textContent || 0), dx: +((kid(el, 'colOff') || {}).textContent || 0) / 9525, r: +((kid(el, 'row') || {}).textContent || 0), dy: +((kid(el, 'rowOff') || {}).textContent || 0) / 9525 };
      if (d) for (const a of kids(d.documentElement)) {
        if (!/Anchor$/.test(a.localName)) continue;
        let el = kids(a).find(k => /^(sp|pic|graphicFrame|grpSp|cxnSp|AlternateContent|contentPart)$/.test(k.localName));
        if (el && el.localName === 'AlternateContent') el = kids(kid(el, 'Fallback'))[0] || kids(kid(el, 'Choice'))[0];
        if (!el || el.localName === 'contentPart') continue;
        const ext = kid(a, 'ext'), p = kid(a, 'pos');
        sh.anchors.push({
          type: a.localName, el, from: pos(kid(a, 'from')), to: pos(kid(a, 'to')),
          ext: ext && { w: atNum(ext, 'cx', 0) / 9525, h: atNum(ext, 'cy', 0) / 9525 }, pos: p && { x: atNum(p, 'x', 0) / 9525, y: atNum(p, 'y', 0) / 9525 },
        });
      }
    }

    // 表示する広さ：使っている所＋少し
    let nR = sh.maxR + 1, nC = sh.maxC + 1;
    for (const g of sh.merges) { nR = Math.max(nR, g.r1 + 1); nC = Math.max(nC, g.c1 + 1); }
    for (const a of sh.anchors) { const e = a.to || a.from; if (e) { nR = Math.max(nR, e.r + 1); nC = Math.max(nC, e.c + 1); } }
    sh.usedR = nR; sh.usedC = nC;
    sh.nR = Math.min(MAXR + 1, Math.max(nR + 30, sh.freeze.top + 30, 60));
    sh.nC = Math.min(MAXC + 1, Math.max(nC + 6, sh.freeze.left + 6, 20));

    // 条件付き書式
    try { conditional(sh, R, book); } catch (e) { console.warn('条件付き書式を読めませんでした', e); }
    sh.cell = (r, c) => { const row = sh.rows[r]; return row ? row.cells[c] : undefined; };
    sh.formula = (r, c) => formulaOf(sh, r, c);
    return sh;
  }

  // sheetData を読む（行 <row>・セル <c>）
  function readCells(sh, xml, start, end, book) {
    const sst = book.sst;
    const rowRe = /<((?:[\w.-]+:)?)row\b([^>]*?)(\/?)>/g;
    const cRe = /<((?:[\w.-]+:)?)c\b([^>]*?)(\/?)>/g;
    let rIdx = -1, maxR = -1, maxC = -1;
    rowRe.lastIndex = start;
    for (;;) {
      const m = rowRe.exec(xml);
      if (!m || m.index >= end) break;
      const a = attrs(m[2]);
      rIdx = a.r ? +a.r - 1 : rIdx + 1;
      const row = { cells: [] };
      if (a.ht != null && (truthy(a.customHeight) || true)) row.ht = +a.ht;
      if (truthy(a.hidden)) row.hidden = true;
      if (truthy(a.customFormat) && a.s != null) row.s = +a.s;
      sh.rows[rIdx] = row;
      if (row.ht != null || row.hidden || row.s != null) maxR = Math.max(maxR, rIdx);
      if (m[3]) continue;
      const close = xml.indexOf(`</${m[1]}row>`, rowRe.lastIndex);
      const rEnd = close < 0 || close > end ? end : close;
      let cIdx = -1;
      cRe.lastIndex = rowRe.lastIndex;
      for (;;) {
        const cm = cRe.exec(xml);
        if (!cm || cm.index >= rEnd) break;
        const ca = attrs(cm[2]);
        if (ca.r) { let n = 0, i = 0; const s = ca.r; for (; i < s.length; i++) { const k = s.charCodeAt(i); if (k >= 48 && k <= 57) break; n = n * 26 + (k & 31); } cIdx = n - 1; }
        else cIdx++;
        const cell = { s: ca.s != null ? +ca.s : 0, t: 'z', v: undefined };
        if (!cm[3]) {
          const cClose = xml.indexOf(`</${cm[1]}c>`, cRe.lastIndex);
          const inner = xml.slice(cRe.lastIndex, cClose < 0 ? rEnd : cClose);
          cRe.lastIndex = cClose < 0 ? rEnd : cClose;
          readCell(cell, ca.t, inner, sst, sh, rIdx, cIdx);
        }
        row.cells[cIdx] = cell;
        if (cell.t !== 'z' || cell.s) { maxR = Math.max(maxR, rIdx); maxC = Math.max(maxC, cIdx); }
      }
      rowRe.lastIndex = rEnd;
    }
    sh.maxR = maxR; sh.maxC = maxC;
  }
  const vRe = /<(?:[\w.-]+:)?v>([\s\S]*?)<\/(?:[\w.-]+:)?v>/, fRe = /<(?:[\w.-]+:)?f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[\w.-]+:)?f>)/, isRe = /<(?:[\w.-]+:)?is>([\s\S]*?)<\/(?:[\w.-]+:)?is>/;
  function readCell(cell, t, inner, sst, sh, r, c) {
    const fm = fRe.exec(inner);
    if (fm) {
      const fa = attrs(fm[1]), ft = fm[2] != null ? unxml(fm[2]) : '';
      if (fa.t === 'shared' && fa.si != null) {
        if (ft) { sh.shared.set(fa.si, { r, c, f: ft }); cell.f = ft; } else cell.fsi = fa.si;
      } else if (ft) { cell.f = ft; if (fa.t === 'array' && !/\bcm=/.test(inner)) cell.arr = true; }
    }
    const vm = vRe.exec(inner);
    const v = vm ? unxml(vm[1]) : null;
    switch (t) {
      case 's': { const e = sst[+v]; if (e == null) break; if (typeof e === 'string') { cell.t = 's'; cell.v = e; } else { cell.t = 'r'; cell.v = e.text; cell.runs = e.runs; } break; }
      case 'str': if (v != null) { cell.t = 's'; cell.v = unesc(v); } break;
      case 'inlineStr': {
        const im = isRe.exec(inner);
        if (im) {
          const d = parseXml('<is>' + im[1].replace(/<(\/?)[\w.-]+:/g, '<$1') + '</is>');
          const e = d ? richOf(d.documentElement) : unesc(unxml(im[1].replace(/<[^>]+>/g, '')));
          if (typeof e === 'string') { cell.t = 's'; cell.v = e; } else { cell.t = 'r'; cell.v = e.text; cell.runs = e.runs; }
        }
        break;
      }
      case 'b': if (v != null) { cell.t = 'b'; cell.v = v === '1' || v === 'true'; } break;
      case 'e': if (v != null) { cell.t = 'e'; cell.v = v; } break;
      case 'd': if (v) { const d = Date.parse(v.length <= 10 ? v + 'T00:00:00Z' : v.replace(/Z?$/, 'Z')); if (isFinite(d)) { cell.t = 'n'; cell.v = d / 86400000 + 25569; } } break;
      default: if (v != null && v !== '') { const n = +v; if (isFinite(n)) { cell.t = 'n'; cell.v = n; } }
    }
  }

  /* ---------- 数式（セルの上の欄に出す） ---------- */
  function formulaOf(sh, r, c) {
    const cell = sh.cell(r, c);
    if (!cell) return null;
    let f = cell.f;
    if (f == null && cell.fsi != null) {
      const m = sh.shared.get(cell.fsi);
      if (m) f = shiftFormula(m.f, r - m.r, c - m.c);
    }
    if (f == null) return null;
    f = '=' + f.replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\.|_xludf\./g, '');
    return cell.arr ? `{${f}}` : f;
  }
  // 相対参照（$ のない所）を dr 行・dc 列ずらす。文字列（"…"）とシート名（'…'）の中は変えない
  function shiftFormula(f, dr, dc) {
    let out = '', i = 0;
    while (i < f.length) {
      const ch = f[i];
      if (ch === '"' || ch === "'") {
        let j = i + 1;
        for (; j < f.length; j++) if (f[j] === ch) { if (f[j + 1] === ch) j++; else break; }
        out += f.slice(i, j + 1); i = j + 1; continue;
      }
      const prev = i ? f[i - 1] : '';
      const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)(?![\w(.])/.exec(f.slice(i, i + 16));
      if (m && !/[\w.$]/.test(prev)) {
        let c = colNum(m[2]), r = +m[4] - 1;
        if (!m[1]) c += dc;
        if (!m[3]) r += dr;
        out += c < 0 || r < 0 || c > MAXC || r > MAXR ? '#REF!' : `${m[1]}${colName(c)}${m[3]}${r + 1}`;
        i += m[0].length; continue;
      }
      // 名前・関数名はそのまま（途中の文字を番地と見ないように）
      const w = /^[A-Za-z_\\][\w.]*/.exec(f.slice(i));
      if (w && !/[\d]/.test(w[0][0])) { out += w[0]; i += w[0].length; continue; }
      out += ch; i++;
    }
    return out;
  }

  /* ---------- 条件付き書式（よく使うものだけ。結果を sh.cf に入れる） ---------- */
  /* ---------- 条件付き書式の式（小さな計算機）----------
     同じシートの中の番地・範囲、数・文字・TRUE/FALSE、+ - * / ^ & % と比べる記号、よく使う関数だけ。
     それ以外（ほかのシート・名前・知らない関数）は parse が null を返すか、run が投げる → その条件は出さない */
  const Formula = (() => {
    const ERR = e => ({ err: e });
    const isErr = v => v && typeof v === 'object' && v.err;
    function tokens(src) {
      const out = [];
      let i = 0;
      const s = src.replace(/^=/, '');
      while (i < s.length) {
        const ch = s[i];
        if (/\s/.test(ch)) { i++; continue; }
        let m;
        const rest = s.slice(i);
        if ((m = /^"((?:[^"]|"")*)"/.exec(rest))) { out.push({ k: 's', v: m[1].replace(/""/g, '"') }); i += m[0].length; continue; }
        if ((m = /^(\d+\.?\d*|\.\d+)(E[-+]?\d+)?/i.exec(rest))) {
          // 「1:3」（行の範囲）は数ではない
          if (s[i + m[0].length] === ':' && /^\d+$/.test(m[0])) {
            const r = /^(\$?\d+):(\$?\d+)/.exec(rest);
            if (r) { out.push({ k: 'rows', a: r[1], b: r[2] }); i += r[0].length; continue; }
          }
          out.push({ k: 'n', v: +m[0] }); i += m[0].length; continue;
        }
        if ((m = /^(\$?[A-Za-z]{1,3}\$?\d+)(?::(\$?[A-Za-z]{1,3}\$?\d+))?(?![\w(.])/.exec(rest))) {
          out.push(m[2] ? { k: 'rng', a: m[1], b: m[2] } : { k: 'ref', a: m[1] }); i += m[0].length; continue;
        }
        if ((m = /^(\$?[A-Za-z]{1,3}):(\$?[A-Za-z]{1,3})(?![\w(])/.exec(rest))) { out.push({ k: 'cols', a: m[1], b: m[2] }); i += m[0].length; continue; }
        if ((m = /^(TRUE|FALSE)(?![\w(.])/i.exec(rest))) { out.push({ k: 'b', v: /^TRUE$/i.test(m[1]) }); i += m[0].length; continue; }
        if ((m = /^([A-Za-z_][\w.]*)\s*\(/.exec(rest))) { out.push({ k: 'fn', v: m[1].toUpperCase().replace(/^_XLFN\./, '') }); i += m[0].length; continue; }
        if ((m = /^(<>|<=|>=|[-+*/^&=<>%(),])/.exec(rest))) { out.push({ k: 'op', v: m[1] }); i += m[0].length; continue; }
        return null; // ほかのシート（!）・名前・配列 {…} など
      }
      return out;
    }
    // 優先順位：比べる < & < + - < * / < ^ < % < 符号
    function parse(src) {
      const tk = tokens(src);
      if (!tk) return null;
      let p = 0;
      const peek = () => tk[p], isOp = v => tk[p] && tk[p].k === 'op' && tk[p].v === v;
      const bin = (next, ops) => () => {
        let a = next();
        while (tk[p] && tk[p].k === 'op' && ops.includes(tk[p].v)) { const op = tk[p++].v; a = { t: 'op', op, a, b: next() }; }
        return a;
      };
      const unary = () => {
        if (isOp('-')) { p++; return { t: 'neg', a: unary() }; }
        if (isOp('+')) { p++; return unary(); }
        let a = atom();
        while (isOp('%')) { p++; a = { t: 'op', op: '/', a, b: { t: 'n', v: 100 } }; }
        return a;
      };
      const pow = bin(unary, ['^']), mul = bin(pow, ['*', '/']), add = bin(mul, ['+', '-']), cat = bin(add, ['&']), cmp = bin(cat, ['=', '<>', '<', '>', '<=', '>=']);
      function atom() {
        const t = tk[p++];
        if (!t) throw new Error('end');
        switch (t.k) {
          case 'n': return { t: 'n', v: t.v };
          case 's': return { t: 's', v: t.v };
          case 'b': return { t: 'b', v: t.v };
          case 'ref': return { t: 'ref', ...cellRef(t.a) };
          case 'rng': return { t: 'rng', a: cellRef(t.a), b: cellRef(t.b) };
          case 'cols': { const a = colRef(t.a), b = colRef(t.b); return { t: 'rng', a: { r: 0, ra: true, ...a }, b: { r: MAXR, ra: true, ...b } }; }
          case 'rows': { const a = rowRef(t.a), b = rowRef(t.b); return { t: 'rng', a: { c: 0, ca: true, ...a }, b: { c: MAXC, ca: true, ...b } }; }
          case 'fn': {
            const args = [];
            if (!isOp(')')) { for (;;) { args.push(isOp(',') || isOp(')') ? { t: 'blank' } : cmp()); if (isOp(',')) { p++; continue; } break; } }
            if (!isOp(')')) throw new Error(')');
            p++;
            return { t: 'fn', name: t.v, args };
          }
          case 'op':
            if (t.v === '(') { const e = cmp(); if (!isOp(')')) throw new Error(')'); p++; return e; }
        }
        throw new Error('?');
      }
      try {
        const e = cmp();
        return p === tk.length && peek() === undefined ? e : null;
      } catch (e) { return null; }
    }
    function cellRef(s) { const m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(s); return { c: colNum(m[2]), ca: !!m[1], r: +m[4] - 1, ra: !!m[3] }; }
    function colRef(s) { const m = /^(\$?)([A-Za-z]{1,3})$/.exec(s); return { c: colNum(m[2]), ca: !!m[1] }; }
    function rowRef(s) { const m = /^(\$?)(\d+)$/.exec(s); return { r: +m[2] - 1, ra: !!m[1] }; }

    // セルの値：数・文字・true/false・null（空）・{ err }
    const cellVal = (sh, r, c) => {
      const cell = sh.rows[r] && sh.rows[r].cells[c];
      if (!cell || cell.t === 'z') return null;
      if (cell.t === 'e') return ERR(cell.v);
      return cell.v;
    };
    const pos = (x, ctx) => ({ r: x.ra ? x.r : x.r + ctx.dr, c: x.ca ? x.c : x.c + ctx.dc });
    // 範囲の値（使っている所まで）
    function* cells(n, ctx) {
      const a = pos(n.a, ctx), b = pos(n.b, ctx);
      const r1 = Math.min(Math.max(a.r, b.r), ctx.sh.maxR), c1 = Math.min(Math.max(a.c, b.c), ctx.sh.maxC);
      for (let r = Math.min(a.r, b.r); r <= r1; r++) for (let c = Math.min(a.c, b.c); c <= c1; c++) yield cellVal(ctx.sh, r, c);
    }
    const num = v => {
      if (isErr(v)) throw v;
      if (v == null) return 0;
      if (typeof v === 'boolean') return v ? 1 : 0;
      if (typeof v === 'number') return v;
      const s = String(v).trim();
      if (/^[-+]?(\d+\.?\d*|\.\d+)(E[-+]?\d+)?%?$/i.test(s)) return parseFloat(s) / (s.endsWith('%') ? 100 : 1);
      throw ERR('#VALUE!');
    };
    const str = v => {
      if (isErr(v)) throw v;
      if (v == null) return '';
      if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
      if (typeof v === 'number') { const t = String(+v.toPrecision(15)); return /e/.test(t) ? NumFmt.format(v, 'General').text : t; }
      return String(v);
    };
    // Excel の比べ方：数 < 文字 < TRUE/FALSE。空は相手の型の「空」。文字は大文字小文字を区別しない
    function compare(a, b) {
      if (isErr(a)) throw a;
      if (isErr(b)) throw b;
      const rank = v => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
      if (a == null && b == null) return 0;
      if (a == null) a = typeof b === 'number' ? 0 : typeof b === 'string' ? '' : false;
      if (b == null) b = typeof a === 'number' ? 0 : typeof a === 'string' ? '' : false;
      if (rank(a) !== rank(b)) return rank(a) - rank(b);
      if (typeof a === 'string') { const x = a.toLowerCase(), y = b.toLowerCase(); return x < y ? -1 : x > y ? 1 : 0; }
      return a === b ? 0 : a < b ? -1 : 1;
    }
    // COUNTIF の条件（">5"・"<>x"・"り*"）
    function criteria(cv) {
      let op = '=', v = cv;
      if (typeof cv === 'string') {
        const m = /^(<=|>=|<>|<|>|=)?(.*)$/s.exec(cv);
        op = m[1] || '='; v = m[2];
        if (/^[-+]?(\d+\.?\d*|\.\d+)(E[-+]?\d+)?$/i.test(v.trim())) v = +v;
        else if (/^(TRUE|FALSE)$/i.test(v)) v = /^TRUE$/i.test(v);
      }
      const wild = typeof v === 'string' && (op === '=' || op === '<>') && /[*?]/.test(v)
        ? new RegExp('^' + v.replace(/~([*?])|([*?])|[.+^${}()|[\]\\]/g, (m, esc, w) => (esc ? '\\' + esc : w ? (w === '*' ? '[\\s\\S]*' : '[\\s\\S]') : '\\' + m)) + '$', 'i') : null;
      return x => {
        if (isErr(x)) return false;
        if (wild) { const ok = x != null && typeof x === 'string' && wild.test(x); return op === '=' ? ok : !ok; }
        if (v === '' && op === '=') return x == null || x === '';
        if (x == null) return op === '<>';
        if (typeof x !== typeof v) return op === '<>';
        const k = compare(x, v);
        return op === '=' ? k === 0 : op === '<>' ? k !== 0 : op === '<' ? k < 0 : op === '>' ? k > 0 : op === '<=' ? k <= 0 : k >= 0;
      };
    }
    function flat(args, ctx) {
      const out = [];
      for (const a of args) { if (a.t === 'rng') for (const v of cells(a, ctx)) out.push({ v, ref: true }); else out.push({ v: run(a, ctx), ref: a.t === 'ref' }); }
      return out;
    }
    const FN = {
      MOD: (a, b) => { const x = num(a), y = num(b); if (!y) throw ERR('#DIV/0!'); return x - y * Math.floor(x / y); },
      ABS: a => Math.abs(num(a)),
      INT: a => Math.floor(num(a)),
      ROUND: (a, b) => { const k = Math.pow(10, Math.trunc(num(b == null ? 0 : b))); const x = num(a); return Math.sign(x) * Math.round(Math.abs(x) * k + 1e-9) / k; },
      NOT: a => !truth(a),
      LEN: a => str(a).length,
      LEFT: (a, n) => str(a).slice(0, n == null ? 1 : num(n)),
      RIGHT: (a, n) => { const s = str(a), k = n == null ? 1 : num(n); return k ? s.slice(-k) : ''; },
      MID: (a, s0, n) => str(a).substr(num(s0) - 1, num(n)),
      TRIM: a => str(a).trim().replace(/ +/g, ' '),
      UPPER: a => str(a).toUpperCase(),
      LOWER: a => str(a).toLowerCase(),
      EXACT: (a, b) => str(a) === str(b),
      FIND: (a, b, s0) => { const i = str(b).indexOf(str(a), s0 == null ? 0 : num(s0) - 1); if (i < 0) throw ERR('#VALUE!'); return i + 1; },
      SEARCH: (a, b, s0) => { const i = str(b).toLowerCase().indexOf(str(a).toLowerCase(), s0 == null ? 0 : num(s0) - 1); if (i < 0) throw ERR('#VALUE!'); return i + 1; },
      YEAR: a => NumDate(Math.floor(num(a))).y,
      MONTH: a => NumDate(Math.floor(num(a))).m + 1,
      DAY: a => NumDate(Math.floor(num(a))).d,
      WEEKDAY: (a, t) => { const w = NumDate(Math.floor(num(a))).w, k = t == null ? 1 : num(t); return k === 2 ? (w + 6) % 7 + 1 : k === 3 ? (w + 6) % 7 : w + 1; },
      DATE: (y, m, d) => Math.round(Date.UTC(num(y), num(m) - 1, num(d)) / 86400000) + 25569,
      ISODD: a => Math.floor(Math.abs(num(a))) % 2 === 1,
      ISEVEN: a => Math.floor(Math.abs(num(a))) % 2 === 0,
    };
    function run(n, ctx) {
      switch (n.t) {
        case 'n': case 's': case 'b': return n.v;
        case 'blank': return null;
        case 'ref': { const q = pos(n, ctx); return cellVal(ctx.sh, q.r, q.c); }
        case 'rng': { const q = pos(n.a, ctx); return cellVal(ctx.sh, q.r, q.c); } // 1 つの値として使ったときは左上
        case 'neg': return -num(run(n.a, ctx));
        case 'op': {
          const a = run(n.a, ctx), b = run(n.b, ctx);
          switch (n.op) {
            case '+': return num(a) + num(b);
            case '-': return num(a) - num(b);
            case '*': return num(a) * num(b);
            case '/': { const y = num(b); if (!y) throw ERR('#DIV/0!'); return num(a) / y; }
            case '^': return Math.pow(num(a), num(b));
            case '&': return str(a) + str(b);
            case '=': return compare(a, b) === 0;
            case '<>': return compare(a, b) !== 0;
            case '<': return compare(a, b) < 0;
            case '>': return compare(a, b) > 0;
            case '<=': return compare(a, b) <= 0;
            case '>=': return compare(a, b) >= 0;
          }
          break;
        }
        case 'fn': {
          const A = n.args, name = n.name;
          switch (name) {
            case 'ROW': return A.length && A[0].t !== 'blank' ? pos(A[0].t === 'rng' ? A[0].a : A[0], ctx).r + 1 : ctx.r + 1;
            case 'COLUMN': return A.length && A[0].t !== 'blank' ? pos(A[0].t === 'rng' ? A[0].a : A[0], ctx).c + 1 : ctx.c + 1;
            case 'TODAY': return ctx.today;
            case 'NOW': { const d = new Date(); return ctx.today + (d.getHours() * 3600 + d.getMinutes() * 60 + d.getSeconds()) / 86400; }
            case 'IF': { const c = truth(run(A[0], ctx)); return c ? (A[1] ? run(A[1], ctx) : true) : A[2] ? run(A[2], ctx) : false; }
            case 'IFERROR': try { const v = run(A[0], ctx); if (isErr(v)) throw v; return v; } catch (e) { if (isErr(e)) return run(A[1], ctx); throw e; }
            case 'AND': case 'OR': {
              const vs = flat(A, ctx).filter(x => !(x.ref && (x.v == null || typeof x.v === 'string')));
              if (!vs.length) throw ERR('#VALUE!');
              return name === 'AND' ? vs.every(x => truth(x.v, true)) : vs.some(x => truth(x.v, true));
            }
            case 'ISBLANK': return A[0].t === 'ref' && run(A[0], ctx) == null;
            case 'ISNUMBER': case 'ISTEXT': case 'ISLOGICAL': case 'ISERROR': case 'ISERR': case 'ISNA': case 'ISNONTEXT': {
              let v;
              try { v = run(A[0], ctx); } catch (e) { if (isErr(e)) v = e; else throw e; }
              if (name === 'ISERROR') return !!isErr(v);
              if (name === 'ISERR') return !!isErr(v) && v.err !== '#N/A';
              if (name === 'ISNA') return !!isErr(v) && v.err === '#N/A';
              if (name === 'ISNONTEXT') return typeof v !== 'string';
              return name === 'ISNUMBER' ? typeof v === 'number' : name === 'ISTEXT' ? typeof v === 'string' : typeof v === 'boolean';
            }
            case 'SUM': case 'COUNT': case 'COUNTA': case 'AVERAGE': case 'MAX': case 'MIN': {
              const vs = flat(A, ctx);
              if (name === 'COUNTA') return vs.filter(x => x.v != null).length;
              const ns = [];
              for (const x of vs) { if (isErr(x.v)) { if (name === 'COUNT') continue; throw x.v; } if (typeof x.v === 'number') ns.push(x.v); else if (!x.ref && x.v != null && name !== 'COUNT') ns.push(num(x.v)); }
              if (name === 'COUNT') return ns.length;
              if (name === 'SUM') return ns.reduce((s, x) => s + x, 0);
              if (name === 'AVERAGE') { if (!ns.length) throw ERR('#DIV/0!'); return ns.reduce((s, x) => s + x, 0) / ns.length; }
              return ns.length ? (name === 'MAX' ? Math.max(...ns) : Math.min(...ns)) : 0;
            }
            case 'COUNTIF': case 'SUMIF': {
              if (A[0].t !== 'rng' && A[0].t !== 'ref') throw ERR('#VALUE!');
              const ok = criteria(run(A[1], ctx));
              const rg = A[0].t === 'ref' ? { a: A[0], b: A[0] } : A[0];
              if (name === 'COUNTIF') { let k = 0; for (const v of cells(rg, ctx)) if (ok(v)) k++; return k; }
              const vals = [...cells(rg, ctx)], sums = A[2] ? [...cells(A[2].t === 'ref' ? { a: A[2], b: A[2] } : A[2], ctx)] : vals;
              let s = 0;
              vals.forEach((v, i) => { if (ok(v) && typeof sums[i] === 'number') s += sums[i]; });
              return s;
            }
          }
          const f = FN[name];
          if (!f) throw new Error('知らない関数 ' + name);
          return f(...A.map(a => (a.t === 'blank' ? null : run(a, ctx))));
        }
      }
      throw new Error('?');
    }
    // 条件として正しいか（0 以外の数・TRUE）
    function truth(v, strict) {
      if (isErr(v)) { if (strict) throw v; return false; }
      if (typeof v === 'boolean') return v;
      if (typeof v === 'number') return v !== 0;
      if (strict && typeof v === 'string') throw ERR('#VALUE!');
      return false;
    }
    return { parse, run, truth };
  })();

  function conditional(sh, R, book) {
    // Excel 2010 からのデータバーの追加の設定（extLst の中）
    const x14 = new Map();
    for (const e of R.getElementsByTagNameNS('*', 'cfRule')) {
      if (e.parentNode && e.parentNode.parentNode && e.parentNode.parentNode.localName === 'conditionalFormattings') { const id = at(e, 'id'); if (id) x14.set(id, e); }
    }
    const rules = [];
    for (const cf of kids(R, 'conditionalFormatting')) {
      const rs = rangeList(at(cf, 'sqref'));
      if (!rs.length) continue;
      for (const rule of kids(cf, 'cfRule')) {
        const idEl = [...rule.getElementsByTagNameNS('*', 'id')][0];
        rules.push({ rs, rule, pri: atNum(rule, 'priority', 1e9), ext: idEl ? x14.get(idEl.textContent.trim()) : null });
      }
    }
    if (!rules.length) return;
    rules.sort((a, b) => a.pri - b.pri);
    const stopped = new Set();
    // 式の条件は、空いたセル（表の外の少し先まで）にもつける（1 行おきの色など）
    const lim = (g, wide) => ({ r0: g.r0, c0: g.c0, r1: Math.min(g.r1, sh.maxR + (wide ? 60 : 0)), c1: Math.min(g.c1, sh.maxC + (wide ? 10 : 0)) });
    const valOf = cell => (!cell || cell.t === 'z' ? null : cell.v);
    const today = Math.floor((Date.now() - new Date().getTimezoneOffset() * 60000) / 86400000) + 25569;
    for (const { rs, rule, ext } of rules) {
      const type = at(rule, 'type');
      // 範囲の中のセル
      const list = [];
      const blanks = type === 'cellIs' || type === 'containsBlanks' || type === 'notContainsBlanks' || type === 'expression';
      for (const g0 of rs) {
        const g = lim(g0, type === 'expression');
        if ((g.r1 - g.r0 + 1) * (g.c1 - g.c0 + 1) > 300000) continue;
        for (let r = g.r0; r <= g.r1; r++) {
          const row = sh.rows[r];
          if (!row && !blanks) continue;
          for (let c = g.c0; c <= g.c1; c++) {
            const cell = row && row.cells[c];
            if (!blanks && (!cell || cell.t === 'z')) continue;
            list.push({ r, c, cell, v: valOf(cell) });
          }
        }
      }
      if (!list.length) continue;
      const nums = list.filter(x => typeof x.v === 'number').map(x => x.v);
      const top = rs[0];
      const operand = (txt, x) => evalOperand(sh, txt, x.r - top.r0, x.c - top.c0);
      let test = null, special = null;
      switch (type) {
        case 'cellIs': {
          const fs = [...rule.getElementsByTagNameNS('*', 'formula')].map(e => e.textContent);
          const op = at(rule, 'operator');
          test = x => {
            const v = x.v == null ? 0 : x.v, a = operand(fs[0], x), b = fs[1] != null ? operand(fs[1], x) : undefined;
            if (a === undefined) return false;
            const cmp = (p, q) => (typeof p === 'number' && typeof q === 'number' ? p - q : typeof p === 'string' && typeof q === 'string' ? p.toLowerCase().localeCompare(q.toLowerCase()) : NaN);
            const k = cmp(v, a);
            switch (op) {
              case 'equal': return k === 0;
              case 'notEqual': return k !== 0;
              case 'greaterThan': return k > 0;
              case 'lessThan': return k < 0;
              case 'greaterThanOrEqual': return k >= 0;
              case 'lessThanOrEqual': return k <= 0;
              case 'between': case 'notBetween': {
                if (b === undefined) return false;
                const lo = cmp(a, b) <= 0 ? a : b, hi = lo === a ? b : a;
                const inside = cmp(v, lo) >= 0 && cmp(v, hi) <= 0;
                return op === 'between' ? inside : !inside && !isNaN(cmp(v, lo));
              }
            }
            return false;
          };
          break;
        }
        case 'containsText': case 'notContainsText': case 'beginsWith': case 'endsWith': {
          const t = String(at(rule, 'text') || '').toLowerCase();
          test = x => {
            if (x.v == null) return type === 'notContainsText';
            const s = String(typeof x.v === 'boolean' ? (x.v ? 'TRUE' : 'FALSE') : x.v).toLowerCase();
            return type === 'containsText' ? s.includes(t) : type === 'notContainsText' ? !s.includes(t) : type === 'beginsWith' ? s.startsWith(t) : s.endsWith(t);
          };
          break;
        }
        case 'containsBlanks': test = x => x.v == null || (typeof x.v === 'string' && !x.v.trim()); break;
        case 'notContainsBlanks': test = x => !(x.v == null || (typeof x.v === 'string' && !x.v.trim())); break;
        case 'containsErrors': test = x => x.cell && x.cell.t === 'e'; break;
        case 'notContainsErrors': test = x => !(x.cell && x.cell.t === 'e'); break;
        case 'top10': {
          if (!nums.length) break;
          const s = nums.slice().sort((a, b) => b - a), bottom = truthy(at(rule, 'bottom'));
          if (bottom) s.reverse();
          let n = atNum(rule, 'rank', 10);
          if (truthy(at(rule, 'percent'))) n = Math.max(1, Math.floor(s.length * n / 100));
          const th = s[Math.min(s.length, Math.max(1, n)) - 1];
          test = x => typeof x.v === 'number' && (bottom ? x.v <= th : x.v >= th);
          break;
        }
        case 'aboveAverage': {
          if (!nums.length) break;
          const avg = nums.reduce((a, b) => a + b, 0) / nums.length;
          const sd = nums.length > 1 ? Math.sqrt(nums.reduce((a, b) => a + (b - avg) * (b - avg), 0) / (nums.length - 1)) : 0;
          const above = at(rule, 'aboveAverage') !== '0', eq = truthy(at(rule, 'equalAverage')), k = atNum(rule, 'stdDev', 0);
          const th = avg + (above ? 1 : -1) * k * sd;
          test = x => typeof x.v === 'number' && (above ? (eq ? x.v >= th : x.v > th) : (eq ? x.v <= th : x.v < th));
          break;
        }
        case 'duplicateValues': case 'uniqueValues': {
          const cnt = new Map(), k = v => (typeof v === 'string' ? 's' + v.toLowerCase() : typeof v + v);
          for (const x of list) if (x.v != null && x.v !== '') cnt.set(k(x.v), (cnt.get(k(x.v)) || 0) + 1);
          test = x => x.v != null && x.v !== '' && (type === 'duplicateValues' ? cnt.get(k(x.v)) > 1 : cnt.get(k(x.v)) === 1);
          break;
        }
        case 'timePeriod': {
          const p = at(rule, 'timePeriod'), dt = NumDate(today);
          const inMonth = (v, k) => { const d = NumDate(Math.floor(v)); const m = dt.y * 12 + dt.m + k; return d.y * 12 + d.m === m; };
          const wk = today - dt.w;
          test = x => {
            if (typeof x.v !== 'number') return false;
            const d = Math.floor(x.v);
            switch (p) {
              case 'today': return d === today;
              case 'yesterday': return d === today - 1;
              case 'tomorrow': return d === today + 1;
              case 'last7Days': return d > today - 7 && d <= today;
              case 'thisWeek': return d >= wk && d < wk + 7;
              case 'lastWeek': return d >= wk - 7 && d < wk;
              case 'nextWeek': return d >= wk + 7 && d < wk + 14;
              case 'thisMonth': return inMonth(d, 0);
              case 'lastMonth': return inMonth(d, -1);
              case 'nextMonth': return inMonth(d, 1);
            }
            return false;
          };
          break;
        }
        case 'colorScale': special = colorScale(rule, nums, book, sh); break;
        case 'dataBar': special = dataBar(rule, ext, nums, book, sh); break;
        case 'expression': {
          // 式（同じシートの中だけ。知らない関数・ほかのシートを使う式は出さない）
          const f = [...rule.getElementsByTagNameNS('*', 'formula')].map(e => e.textContent)[0];
          const ast = f ? Formula.parse(f) : null;
          if (!ast) break;
          test = x => { try { return Formula.truth(Formula.run(ast, { sh, dr: x.r - top.r0, dc: x.c - top.c0, r: x.r, c: x.c, today })); } catch (e) { return false; } };
          break;
        }
        case 'iconSet': special = iconSet(rule, nums, sh); break;
      }
      if (!test && !special) continue;
      const d = test ? book.dxf(atNum(rule, 'dxfId', -1)) : null;
      const stop = truthy(at(rule, 'stopIfTrue'));
      for (const x of list) {
        const k = key(x.r, x.c);
        if (stopped.has(k)) continue;
        let o = null;
        if (special) {
          if (typeof x.v !== 'number') continue;
          o = special(x.v);
          if (!o) continue;
        } else if (!test(x)) continue;
        let cur = sh.cf.get(k);
        if (!cur) sh.cf.set(k, (cur = {}));
        if (special) { for (const p in o) if (cur[p] === undefined) cur[p] = o[p]; }
        else if (d) {
          if (d.fill && cur.fill === undefined) cur.fill = d.fill;
          if (d.font) { cur.font = cur.font || {}; for (const p in d.font) if (cur.font[p] === undefined) cur.font[p] = d.font[p]; }
          if (d.border && cur.border === undefined) cur.border = d.border;
          if (d.fmt && cur.fmt === undefined) cur.fmt = d.fmt;
        }
        if (stop) stopped.add(k);
      }
    }
  }
  function NumDate(serial) {
    const dt = new Date((serial - 25569) * 86400000);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth(), d: dt.getUTCDate(), w: dt.getUTCDay() };
  }
  // 条件の値：数・"文字"・セルの番地（相対ならずらす）。それ以外（式）は分からない（undefined）
  function evalOperand(sh, txt, dr, dc) {
    if (txt == null) return undefined;
    let s = txt.trim().replace(/^=/, '');
    if (/^[-+]?(\d+\.?\d*|\.\d+)(E[-+]?\d+)?$/i.test(s)) return +s;
    let m = /^"((?:[^"]|"")*)"$/.exec(s);
    if (m) return m[1].replace(/""/g, '"');
    if (/^(TRUE|FALSE)$/i.test(s)) return /^TRUE$/i.test(s);
    m = /^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$/.exec(s);
    if (m) {
      const r = +m[4] - 1 + (m[3] ? 0 : dr), c = colNum(m[2]) + (m[1] ? 0 : dc);
      const cell = sh.rows[r] && sh.rows[r].cells[c];
      return !cell || cell.t === 'z' ? 0 : cell.v;
    }
    return undefined;
  }
  function percentile(sorted, p) {
    if (!sorted.length) return 0;
    const h = (sorted.length - 1) * Math.max(0, Math.min(1, p)), lo = Math.floor(h);
    return sorted[lo] + (h - lo) * ((sorted[lo + 1] != null ? sorted[lo + 1] : sorted[lo]) - sorted[lo]);
  }
  // 色の段階・データバーのしきい値
  function cfvoVal(el, sorted, sh) {
    const t = at(el, 'type'), v = at(el, 'val');
    const mn = sorted[0], mx = sorted[sorted.length - 1];
    switch (t) {
      case 'min': return mn;
      case 'max': return mx;
      case 'autoMin': return Math.min(0, mn);
      case 'autoMax': return Math.max(0, mx);
      case 'num': { const x = evalOperand(sh, v, 0, 0); return typeof x === 'number' ? x : NaN; }
      case 'percent': return mn + (mx - mn) * (+v || 0) / 100;
      case 'percentile': return percentile(sorted, (+v || 0) / 100);
      case 'formula': { const x = evalOperand(sh, v, 0, 0); return typeof x === 'number' ? x : NaN; }
    }
    return NaN;
  }
  const rgbOf = h => [parseInt(h.slice(1, 3), 16), parseInt(h.slice(3, 5), 16), parseInt(h.slice(5, 7), 16)];
  function colorScale(rule, nums, book, sh) {
    const cs = kid(rule, 'colorScale');
    if (!cs || !nums.length) return null;
    const sorted = nums.slice().sort((a, b) => a - b);
    const ths = kids(cs, 'cfvo').map(e => cfvoVal(e, sorted, sh));
    const cols = kids(cs, 'color').map(e => rgbOf(book.color(colorSpec(e), '#000000')));
    if (ths.length < 2 || cols.length < ths.length || ths.some(isNaN)) return null;
    return v => {
      let k = 0;
      if (v <= ths[0]) k = 0;
      else if (v >= ths[ths.length - 1]) return { fill: { color: hexOf(cols[ths.length - 1]) } };
      else { while (k < ths.length - 2 && v > ths[k + 1]) k++; }
      const a = ths[k], b = ths[k + 1], t = b > a ? Math.max(0, Math.min(1, (v - a) / (b - a))) : 0;
      return { fill: { color: hexOf(cols[k].map((x, i) => x + (cols[k + 1][i] - x) * t)) } };
    };
  }
  // アイコン（3 つ〜5 つに分ける。idx 0 がいちばん小さい値）
  const ICON_N = { '3Arrows': 3, '3ArrowsGray': 3, '3Flags': 3, '3TrafficLights1': 3, '3TrafficLights2': 3, '3Signs': 3, '3Symbols': 3, '3Symbols2': 3, '3Stars': 3, '3Triangles': 3,
    '4Arrows': 4, '4ArrowsGray': 4, '4RedToBlack': 4, '4Rating': 4, '4TrafficLights': 4, '5Arrows': 5, '5ArrowsGray': 5, '5Rating': 5, '5Quarters': 5, '5Boxes': 5 };
  function iconSet(rule, nums, sh) {
    const el = kid(rule, 'iconSet');
    if (!el || !nums.length) return null;
    const set = at(el, 'iconSet') || '3TrafficLights1', n = ICON_N[set];
    const vos = kids(el, 'cfvo');
    if (!n || vos.length < n) return null;
    const sorted = nums.slice().sort((a, b) => a - b);
    const ths = vos.map(e => ({ v: cfvoVal(e, sorted, sh), gte: at(e, 'gte') !== '0' }));
    if (ths.some(t => isNaN(t.v))) return null;
    const rev = truthy(at(el, 'reverse')), show = at(el, 'showValue') !== '0';
    return v => {
      let k = 0;
      for (let i = 1; i < n; i++) if (ths[i].gte ? v >= ths[i].v : v > ths[i].v) k = i;
      return { icon: { set, idx: rev ? n - 1 - k : k, n, show } };
    };
  }
  const hexOf = a => '#' + a.map(v => Math.round(v).toString(16).padStart(2, '0')).join('').toUpperCase();
  function dataBar(rule, ext, nums, book, sh) {
    const db = kid(rule, 'dataBar');
    if (!db || !nums.length) return null;
    const sorted = nums.slice().sort((a, b) => a - b);
    const xb = ext && kid(ext, 'dataBar');
    const vos = kids(xb || db, 'cfvo');
    if (vos.length < 2) return null;
    let lo = cfvoVal(vos[0], sorted, sh), hi = cfvoVal(vos[1], sorted, sh);
    if (isNaN(lo) || isNaN(hi)) return null;
    const color = book.color(colorSpec(kid(db, 'color')), '#638EC6');
    const o = {
      color, show: at(db, 'showValue') !== '0',
      min: (xb ? atNum(xb, 'minLength', 10) : atNum(db, 'minLength', 10)) / 100,
      max: (xb ? atNum(xb, 'maxLength', 90) : atNum(db, 'maxLength', 90)) / 100,
      grad: xb ? at(xb, 'gradient') !== '0' : true,
      border: xb && truthy(at(xb, 'border')) ? book.color(colorSpec(kid(xb, 'borderColor')), color) : null,
      neg: xb ? book.color(colorSpec(kid(xb, 'negativeFillColor')), '#FF0000') : '#FF0000',
    };
    const axisMode = xb ? at(xb, 'axisPosition') || 'automatic' : 'none';
    // 軸（0 の位置）。負の数があるときだけ
    let axis = 0;
    if (axisMode !== 'none' && lo < 0 && hi > 0) axis = axisMode === 'middle' ? 0.5 : -lo / (hi - lo);
    else if (axisMode === 'middle') axis = 0.5;
    return v => {
      if (axis > 0) {
        if (v >= 0) return { bar: { ...o, from: axis, to: axis + (1 - axis) * Math.max(0, Math.min(1, hi > 0 ? v / hi : 0)) } };
        return { bar: { ...o, color: o.neg, from: axis - axis * Math.max(0, Math.min(1, lo < 0 ? v / lo : 0)), to: axis } };
      }
      const t = hi > lo ? Math.max(0, Math.min(1, (v - lo) / (hi - lo))) : 1;
      return { bar: { ...o, from: 0, to: o.min + (o.max - o.min) * t } };
    };
  }

  return { open, colName, addr, range, key, shiftFormula, unesc, family, lineRatio, LINE, digitW: mdwOf, Formula };
})();
if (typeof module !== 'undefined') module.exports = Xlsx;
