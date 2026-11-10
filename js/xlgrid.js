'use strict';
// =========================================================
// Excel のシートを表として描く（自作）。js/drawingml.js・js/numfmt.js・js/xlsx.js を先に読む
//   const g = XlGrid.create(host, book, sh, { zoom, onSelect(info), onZoom(z) })
//     host の中に、スクロールできる表を作る（列の記号・行の番号・固定した行と列）。
//     見えている所だけを 512px 四方の canvas に描き、遠くの canvas は片付ける（スマホのメモリのため）
//     g.setZoom(z)・g.zoom・g.select(r, c)・g.goto(r, c)・g.dispose()
//   図（画像・図形・グラフの枠）は DOM（js/drawingml.js）で作って、表の上に置く
//   座標：表全体の左上（行の番号の左・列の記号の上）を 0 とした px（拡大縮小をかけたもの）
// =========================================================
const XlGrid = (() => {
  const { kid, kids, pathOf, at, atNum, truthy, rAttr, desc, h, xfrmOf, placeBox, fillOf, fillEl, lineOf, mergeLine, color: dmlColor, colorEl,
    geometry, preset, textRect, drawText, shapeSvg, blipImage } = DML;
  const TILE = 512, MARGIN = 300;
  const GRID = '#D9D9D9', HBG = '#F3F3F3', HFG = '#505050', HLINE = '#C9C9C9';

  /* ---------- 字の幅・高さ（測ったものは覚えておく） ---------- */
  const mctx = typeof document !== 'undefined' ? document.createElement('canvas').getContext('2d') : null;
  const widths = new Map();
  function tw(font, t) {
    let m = widths.get(font);
    if (!m) { if (widths.size > 400) widths.clear(); m = new Map(); widths.set(font, m); }
    let w = m.get(t);
    if (w === undefined) { mctx.font = font; w = mctx.measureText(t).width; if (m.size < 30000) m.set(t, w); }
    return w;
  }
  const vms = new Map();
  function vm(font) {
    let v = vms.get(font);
    if (!v) {
      mctx.font = font;
      const m = mctx.measureText('国Mg'), px = parseFloat(/([\d.]+)px/.exec(font)[1]);
      v = { asc: m.fontBoundingBoxAscent || px * 0.88, desc: m.fontBoundingBoxDescent || px * 0.12, px };
      vms.set(font, v);
    }
    return v;
  }
  const fontCss = (f, z, k = 1) => `${f.i ? 'italic ' : ''}${f.b ? 'bold ' : ''}${+((f.sz || 11) * 4 / 3 * z * k).toFixed(2)}px ${f.fam}`;

  /* ---------- 塗りの模様（Excel の 18 種類） ---------- */
  const PAT = {
    darkGray: ['1110', '1011'], mediumGray: ['10', '01'], lightGray: ['1000', '0010'], gray125: ['1000', '0000', '0010', '0000'], gray0625: ['10000000', '00000000', '00001000', '00000000'],
    darkHorizontal: ['1111', '1111', '0000', '0000'], lightHorizontal: ['1111', '0000', '0000', '0000'], darkVertical: ['1100'], lightVertical: ['1000'],
    darkDown: ['1100', '0110', '0011', '1001'], lightDown: ['1000', '0100', '0010', '0001'], darkUp: ['0011', '0110', '1100', '1001'], lightUp: ['0001', '0010', '0100', '1000'],
    darkGrid: ['1100', '1100', '0011', '0011'], lightGrid: ['1111', '1000', '1000', '1000'], darkTrellis: ['1001', '1111', '0110', '1111'], lightTrellis: ['1010', '0100', '1010', '0001'],
  };
  const patCache = new Map();
  // 模様は 1 点を css の 1px にする（Excel も画面の拡大率に合わせて広げる）
  function pattern(ctx, f, dpr) {
    const p = ctx.createPattern(patCanvas(f), 'repeat');
    if (dpr !== 1 && p.setTransform) p.setTransform(new DOMMatrix([dpr, 0, 0, dpr, 0, 0]));
    return p;
  }
  function patCanvas(f) {
    const k = f.pattern + f.fg + f.bg;
    let c = patCache.get(k);
    if (!c) {
      const rows = PAT[f.pattern] || PAT.mediumGray;
      c = document.createElement('canvas');
      c.width = rows[0].length; c.height = rows.length;
      const x = c.getContext('2d');
      x.fillStyle = f.bg; x.fillRect(0, 0, c.width, c.height);
      x.fillStyle = f.fg;
      rows.forEach((row, y) => { for (let i = 0; i < row.length; i++) if (row[i] === '1') x.fillRect(i, y, 1, 1); });
      patCache.set(k, c);
    }
    return c;
  }

  /* ---------- 条件付き書式のアイコン（Excel の形をまねた簡単な図） ---------- */
  const IR = '#D65532', IY = '#EAC282', IG = '#68A490', IK = '#505050', IGR = '#A6A6A6', IP = '#F2A39A', IGOLD = '#F2B600', IBL = '#4472C4', IAG = '#7F7F7F';
  const ICONS = {
    '3TrafficLights1': ['c' + IR, 'c' + IY, 'c' + IG], '3TrafficLights2': ['c' + IR, 'c' + IY, 'c' + IG],
    '4TrafficLights': ['c' + IK, 'c' + IR, 'c' + IY, 'c' + IG], '4RedToBlack': ['c' + IK, 'c' + IGR, 'c' + IP, 'c' + IR],
    '3Signs': ['d' + IR, 't' + IY, 'c' + IG], '3Flags': ['f' + IR, 'f' + IY, 'f' + IG],
    '3Arrows': ['a270' + IR, 'a0' + IY, 'a90' + IG], '3ArrowsGray': ['a270' + IAG, 'a0' + IAG, 'a90' + IAG],
    '4Arrows': ['a270' + IR, 'a315' + IY, 'a45' + IY, 'a90' + IG], '4ArrowsGray': ['a270' + IAG, 'a315' + IAG, 'a45' + IAG, 'a90' + IAG],
    '5Arrows': ['a270' + IR, 'a315' + IY, 'a0' + IY, 'a45' + IY, 'a90' + IG], '5ArrowsGray': ['a270' + IAG, 'a315' + IAG, 'a0' + IAG, 'a45' + IAG, 'a90' + IAG],
    '3Triangles': ['v' + IR, '-' + IY, '^' + IG],
    '3Symbols': ['X' + IR, '!' + IY, 'V' + IG], '3Symbols2': ['x' + IR, 'i' + IY, 'k' + IG],
    '3Stars': ['s0', 's1', 's2'], '4Rating': ['r1', 'r2', 'r3', 'r4'], '5Rating': ['r0', 'r1', 'r2', 'r3', 'r4'],
    '5Quarters': ['q0', 'q1', 'q2', 'q3', 'q4'], '5Boxes': ['b0', 'b1', 'b2', 'b3', 'b4'],
  };
  function drawIcon(ctx, icon, x, y, s) {
    const spec = (ICONS[icon.set] || [])[icon.idx];
    if (!spec) return;
    const k = spec[0], col = (/#\w{6}$/.exec(spec) || [IK])[0], lv = +spec.slice(1, 2);
    const cx = x + s / 2, cy = y + s / 2, R = s * 0.46;
    ctx.save();
    ctx.lineJoin = ctx.lineCap = 'round';
    const circle = (fill, r = R) => { ctx.beginPath(); ctx.arc(cx, cy, r, 0, Math.PI * 2); ctx.fillStyle = fill; ctx.fill(); };
    const poly = (pts, fill) => { ctx.beginPath(); pts.forEach(([px, py], i) => (i ? ctx.lineTo(x + px * s, y + py * s) : ctx.moveTo(x + px * s, y + py * s))); ctx.closePath(); ctx.fillStyle = fill; ctx.fill(); };
    const stroke = (pts, color, w) => { ctx.beginPath(); pts.forEach(([px, py], i) => (i ? ctx.lineTo(x + px * s, y + py * s) : ctx.moveTo(x + px * s, y + py * s))); ctx.strokeStyle = color; ctx.lineWidth = w * s; ctx.stroke(); };
    switch (k) {
      case 'c': circle(col); break;
      case 'd': poly([[0.5, 0.06], [0.94, 0.5], [0.5, 0.94], [0.06, 0.5]], col); break;
      case 't': poly([[0.5, 0.08], [0.95, 0.9], [0.05, 0.9]], col); break;
      case 'f': poly([[0.25, 0.12], [0.85, 0.3], [0.25, 0.5]], col); stroke([[0.25, 0.1], [0.25, 0.92]], IK, 0.09); break;
      case 'a': {
        const deg = +/^a(\d+)/.exec(spec)[1], a = -deg * Math.PI / 180;
        ctx.translate(cx, cy); ctx.rotate(a); ctx.translate(-cx, -cy);
        poly([[0.08, 0.38], [0.5, 0.38], [0.5, 0.12], [0.94, 0.5], [0.5, 0.88], [0.5, 0.62], [0.08, 0.62]], col);
        break;
      }
      case 'v': poly([[0.1, 0.25], [0.9, 0.25], [0.5, 0.8]], col); break;
      case '^': poly([[0.1, 0.75], [0.9, 0.75], [0.5, 0.2]], col); break;
      case '-': poly([[0.15, 0.4], [0.85, 0.4], [0.85, 0.6], [0.15, 0.6]], col); break;
      case 'X': circle(col); stroke([[0.33, 0.33], [0.67, 0.67]], '#fff', 0.1); stroke([[0.67, 0.33], [0.33, 0.67]], '#fff', 0.1); break;
      case '!': circle(col); stroke([[0.5, 0.26], [0.5, 0.56]], '#fff', 0.11); stroke([[0.5, 0.73], [0.5, 0.74]], '#fff', 0.12); break;
      case 'V': circle(col); stroke([[0.3, 0.52], [0.45, 0.67], [0.71, 0.36]], '#fff', 0.1); break;
      case 'x': stroke([[0.22, 0.22], [0.78, 0.78]], col, 0.16); stroke([[0.78, 0.22], [0.22, 0.78]], col, 0.16); break;
      case 'i': stroke([[0.5, 0.12], [0.5, 0.6]], col, 0.18); stroke([[0.5, 0.85], [0.5, 0.86]], col, 0.18); break;
      case 'k': stroke([[0.15, 0.52], [0.4, 0.78], [0.86, 0.22]], col, 0.16); break;
      case 's': {
        const pts = [];
        for (let i = 0; i < 10; i++) { const a = -Math.PI / 2 + i * Math.PI / 5, rr = i % 2 ? 0.2 : 0.47; pts.push([0.5 + rr * Math.cos(a), 0.53 + rr * Math.sin(a)]); }
        poly(pts, '#fff'); stroke([...pts, pts[0]], IGOLD, 0.06);
        if (lv) { ctx.save(); ctx.beginPath(); ctx.rect(x, y, lv === 1 ? s / 2 : s, s); ctx.clip(); poly(pts, IGOLD); ctx.restore(); }
        break;
      }
      case 'r': for (let i = 0; i < 4; i++) { const hh = 0.25 + i * 0.22; ctx.fillStyle = i < lv ? IBL : '#D9D9D9'; ctx.fillRect(x + (0.08 + i * 0.23) * s, y + (0.95 - hh) * s, 0.17 * s, hh * s); } break;
      case 'q': {
        circle(IK); circle('#fff', R - Math.max(1, s * 0.07));
        if (lv) { ctx.beginPath(); ctx.moveTo(cx, cy); ctx.arc(cx, cy, R, -Math.PI / 2, -Math.PI / 2 + lv * Math.PI / 2); ctx.closePath(); ctx.fillStyle = IK; ctx.fill(); }
        break;
      }
      case 'b': for (let i = 0; i < 4; i++) { ctx.fillStyle = i < lv ? IBL : '#D9D9D9'; ctx.fillRect(x + (0.1 + (i % 2) * 0.42) * s, y + (0.1 + (i >> 1) * 0.42) * s, 0.38 * s, 0.38 * s); } break;
    }
    ctx.restore();
  }

  const VERT = { '「': '﹁', '」': '﹂', '『': '﹃', '』': '﹄', '（': '︵', '）': '︶', '(': '︵', ')': '︶', '【': '︻', '】': '︼', '〔': '︹', '〕': '︺',
    '｛': '︷', '｝': '︸', '〈': '︿', '〉': '﹀', '《': '︽', '》': '︾', 'ー': '丨', '－': '丨', '―': '丨', '…': '︙', '‥': '︰', '～': '≀', '〜': '≀' };

  /* ---------- 文字を行に分ける（折り返し） ---------- */
  const NO_START = /[、。，．,.)）」』】〕〉》｝\]!！?？:：;；ゝゞー〜…‥・ぁぃぅぇぉっゃゅょゎァィゥェォッャュョヮヵヶ%％]/;
  const NO_END = /[(（「『【〔〈《｛\[]/;
  const CJK = /[　-ヿ㐀-鿿豈-﫿！-｠￠-￯]/;
  // segs：[{ t, font, f }] → [{ segs, w }]
  function breakLines(segs, maxW) {
    const chars = [];
    segs.forEach((s, i) => { for (const ch of s.t) chars.push({ ch, i, w: ch === '\n' ? 0 : tw(s.font, ch) }); });
    const lines = [];
    let cur = [], curW = 0;
    const push = arr => {
      while (arr.length && arr[arr.length - 1].ch === ' ') arr.pop();
      lines.push(arr);
    };
    for (let k = 0; k < chars.length; k++) {
      const c = chars[k];
      if (c.ch === '\n') { push(cur); cur = []; curW = 0; continue; }
      if (curW + c.w > maxW && cur.length) {
        // 折り返す所：空白のあと・和文の前後。なければ、その字の前
        let bp = -1;
        for (let j = cur.length - 1; j >= 0; j--) {
          const a = cur[j].ch, b = j + 1 < cur.length ? cur[j + 1].ch : c.ch;
          if (a === ' ' || ((CJK.test(a) || CJK.test(b)) && !NO_START.test(b) && !NO_END.test(a))) { bp = j; break; }
        }
        let line = bp >= 0 ? cur.slice(0, bp + 1) : cur.slice();
        let rest = bp >= 0 ? cur.slice(bp + 1) : [];
        // 行の頭に来てはいけない字（、。」など）は、前の字と一緒に送る
        const first = rest.length ? rest[0].ch : c.ch;
        if (NO_START.test(first) && line.length > 1 && bp < 0) { rest.unshift(line.pop()); }
        push(line);
        while (rest.length && rest[0].ch === ' ') rest.shift();
        cur = rest;
        curW = cur.reduce((s, x) => s + x.w, 0);
        if (c.ch === ' ' && !cur.length) continue;
      }
      cur.push(c);
      curW += c.w;
    }
    if (cur.length || !lines.length) push(cur);
    // 字の並び → 書式ごとのかたまり
    return lines.map(arr => {
      const out = [];
      for (const c of arr) {
        const last = out[out.length - 1];
        if (last && last.i === c.i) last.t += c.ch; else out.push({ i: c.i, t: c.ch });
      }
      const ss = out.map(o => ({ ...segs[o.i], t: o.t }));
      return { segs: ss, w: ss.reduce((s, x) => s + tw(x.font, x.t), 0) };
    });
  }

  /* ---------- 図（DrawingML。pptx.js と同じ描き方で、プレースホルダーなし） ---------- */
  const CLRMAP = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
  let defLstEl = null;
  // Excel の図形の文字の標準の大きさは 11pt
  const defLst = () => defLstEl || (defLstEl = DML.parseXml('<a:lstStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:defPPr><a:defRPr sz="1100"/></a:defPPr><a:lvl1pPr><a:defRPr sz="1100"/></a:lvl1pPr></a:lstStyle>').documentElement);
  function styleOf(sp, ctx) {
    const st = kid(sp, 'style'), out = {};
    if (!st) return out;
    const T = ctx.theme, lr = kid(st, 'lnRef'), fr = kid(st, 'fillRef'), fo = kid(st, 'fontRef');
    if (lr) { const i = atNum(lr, 'idx', 0); if (i > 0 && T.lines[i - 1]) out.line = lineOf(T.lines[i - 1], { ...ctx, ph: dmlColor(lr, ctx) }); }
    if (fr) {
      const i = atNum(fr, 'idx', 0), f = i >= 1001 ? T.bgFills[i - 1001] : T.fills[i - 1];
      if (i > 0) out.fill = f ? fillEl(f, { ...ctx, ph: dmlColor(fr, ctx) }) : { color: dmlColor(fr, ctx) };
      if (out.fill && out.fill.blip) out.fill = { color: dmlColor(fr, ctx) };
    }
    if (fo) out.font = { font: at(fo, 'idx') === 'major' ? 'major' : 'minor', color: colorEl(fo) ? fo : null };
    return out;
  }
  function missing(text) { const d = h('div', 'pp-missing'); d.textContent = text; return d; }
  async function drawTree(tree, ctx, parent) {
    for (let el of kids(tree)) {
      if (el.localName === 'AlternateContent') el = kids(kid(el, 'Fallback'))[0] || kids(kid(el, 'Choice'))[0];
      if (!el) continue;
      try { await drawObj(el, ctx, parent, null); } catch (e) { console.warn('図形を描けませんでした', e); }
    }
  }
  async function drawObj(el, ctx, parent, xf) {
    const n = el.localName;
    if (n === 'sp' || n === 'cxnSp') return drawSp(el, ctx, parent, xf);
    if (n === 'pic') return drawPic(el, ctx, parent, xf);
    if (n === 'grpSp') return drawGroup(el, ctx, parent, xf);
    if (n === 'graphicFrame') return drawFrame(el, ctx, parent, xf);
  }
  async function drawSp(sp, ctx, parent, xf0) {
    const nv = kids(sp).find(c => /^nv/.test(c.localName));
    if (truthy(at(kid(nv, 'cNvPr'), 'hidden'))) return;
    const spPr = kid(sp, 'spPr'), own = xfrmOf(spPr);
    const xf = xf0 ? { ...(own || {}), ...xf0 } : own;
    if (!xf) return;
    const st = styleOf(sp, ctx), box = placeBox(xf);
    const fill = fillOf(spPr, ctx) || st.fill || null;
    const line = mergeLine(st.line, lineOf(kid(spPr, 'ln'), ctx));
    const paths = geometry(spPr, xf.w, xf.h) || (sp.localName === 'cxnSp' ? preset('line', xf.w, xf.h, {}) : preset('rect', xf.w, xf.h, {}));
    if ((fill && !fill.none) || (line && !line.none && line.color)) box.appendChild(await shapeSvg(paths, fill, line, xf.w, xf.h, { h: xf.flipH, v: xf.flipV }, ctx));
    const tx = kid(sp, 'txBody');
    if (tx) {
      const chain = [defLst(), st.font ? { font: st.font.font, color: st.font.color } : null, kid(tx, 'lstStyle')];
      const t = drawText(tx, chain, [], ctx, textRect(spPr, xf) || { x: 0, y: 0, w: xf.w, h: xf.h });
      if (t) box.appendChild(t);
    }
    if (box.childNodes.length) parent.appendChild(box);
  }
  async function drawPic(pic, ctx, parent, xf0) {
    if (truthy(at(pathOf(pic, 'nvPicPr', 'cNvPr'), 'hidden'))) return;
    const spPr = kid(pic, 'spPr'), own = xfrmOf(spPr);
    const xf = xf0 ? { ...(own || {}), ...xf0 } : own;
    if (!xf) return;
    const box = placeBox(xf);
    const bf = kid(pic, 'blipFill'), m = await blipImage(bf, ctx.src, ctx);
    const paths = geometry(spPr, xf.w, xf.h), prst = at(kid(spPr, 'prstGeom'), 'prst');
    if (m && m.url) {
      const img = h('img');
      img.alt = ''; img.draggable = false; img.decoding = 'async'; img.src = m.url;
      if (m.opacity != null) img.style.opacity = m.opacity;
      const sr = kid(bf, 'srcRect'), c = k => atNum(sr, k, 0) / 100000;
      const l = c('l'), t = c('t'), r = c('r'), b = c('b');
      const iw = xf.w / Math.max(0.01, 1 - l - r), ih = xf.h / Math.max(0.01, 1 - t - b);
      Object.assign(img.style, { position: 'absolute', left: -l * iw + 'px', top: -t * ih + 'px', width: iw + 'px', height: ih + 'px', maxWidth: 'none' });
      const fl = [xf.flipH ? 'scaleX(-1)' : '', xf.flipV ? 'scaleY(-1)' : ''].join(' ').trim();
      if (fl) box.style.transform = (box.style.transform + ' ' + fl).trim();
      const clipBox = h('div', null, { position: 'absolute', inset: '0', overflow: 'hidden' });
      if (paths && prst && prst !== 'rect') clipBox.style.clipPath = `path('${paths.filter(p => p.fill).map(p => p.d).join('')}')`;
      clipBox.appendChild(img);
      box.appendChild(clipBox);
    } else box.appendChild(missing(m && m.unsupported ? `この図は表示できません（${m.unsupported} 形式）` : '図を読めませんでした'));
    const line = mergeLine(styleOf(pic, ctx).line, lineOf(kid(spPr, 'ln'), ctx));
    if (line && !line.none && line.color) box.appendChild(await shapeSvg(paths || preset('rect', xf.w, xf.h, {}), { none: true }, line, xf.w, xf.h, {}, ctx));
    parent.appendChild(box);
  }
  async function drawGroup(g, ctx, parent, xf0) {
    if (truthy(at(pathOf(g, 'nvGrpSpPr', 'cNvPr'), 'hidden'))) return;
    const gp = kid(g, 'grpSpPr'), own = xfrmOf(gp);
    const xf = xf0 ? { ...(own || {}), ...xf0 } : own;
    const c2 = { ...ctx, grpFill: fillOf(gp, ctx) || ctx.grpFill };
    if (!xf) { await drawTree(g, c2, parent); return; }
    const box = placeBox(xf, 'pp-grp');
    const fl = [xf.flipH ? 'scaleX(-1)' : '', xf.flipV ? 'scaleY(-1)' : ''].join(' ').trim();
    if (fl) box.style.transform = (box.style.transform + ' ' + fl).trim();
    const ow = own ? own.w : xf.w, oh = own ? own.h : xf.h;
    const cw = (own && own.cw) || ow, ch = (own && own.ch) || oh;
    const inner = h('div', 'pp-grp', { left: '0', top: '0', width: cw + 'px', height: ch + 'px', transformOrigin: '0 0',
      transform: `scale(${cw ? xf.w / cw : 1},${ch ? xf.h / ch : 1}) translate(${-((own && own.cx) || 0)}px,${-((own && own.cy) || 0)}px)` });
    box.appendChild(inner);
    await drawTree(g, c2, inner);
    parent.appendChild(box);
  }
  async function drawFrame(gf, ctx, parent, xf0) {
    if (truthy(at(pathOf(gf, 'nvGraphicFramePr', 'cNvPr'), 'hidden'))) return;
    const own = xfrmOf(gf), xf = xf0 ? { ...(own || {}), ...xf0 } : own;
    if (!xf) return;
    const gd = pathOf(gf, 'graphic', 'graphicData'), uri = at(gd, 'uri') || '';
    const box = placeBox(xf);
    if (/\/chart$/.test(uri)) {
      let title = '';
      const id = rAttr(kid(gd, 'chart'), 'id'), rel = id && ctx.src.rels.get(id);
      if (rel) {
        const doc = await ctx.xml(rel.target), t = doc && pathOf(doc.documentElement, 'chart', 'title');
        if (t) title = [...desc(t, 't')].map(x => x.textContent).join('');
      }
      box.appendChild(missing(`グラフ${title ? '「' + title + '」' : ''}（この画面ではまだ表示できません）`));
    } else {
      const pic = desc(gd, 'pic')[0];
      if (pic) return drawPic(pic, ctx, parent, xf);
      box.appendChild(missing('この部分は表示できません（埋め込まれたもの）'));
    }
    parent.appendChild(box);
  }

  /* ---------- 表 ---------- */
  function create(host, book, sh, opts = {}) {
    const G = { zoom: opts.zoom || 1, sel: null, dead: false };
    const wrap = h('div', 'xg-wrap'), sc = h('div', 'xg');
    sc.tabIndex = 0;
    const top = h('div', 'xg-band xg-top'), bot = h('div', 'xg-band xg-bot');
    const mkPane = key => {
      const el = h('div', 'xg-pane xg-' + key), tl = h('div', 'xg-layer'), dl = h('div', 'xg-layer xg-dl'), ol = h('div', 'xg-layer');
      el.append(tl, dl, ol);
      return { key, el, tl, dl, ol, tiles: new Map(), gx: 0, gy: 0, w: 0, h: 0, left: key === 'tl' || key === 'bl', upper: key === 'tl' || key === 'tr' };
    };
    const P = { tl: mkPane('tl'), tr: mkPane('tr'), bl: mkPane('bl'), br: mkPane('br') };
    const panes = [P.tl, P.tr, P.bl, P.br];
    top.append(P.tl.el, P.tr.el);
    bot.append(P.bl.el, P.br.el);
    sc.append(top, bot);
    wrap.appendChild(sc);
    host.appendChild(wrap);

    const nR = sh.nR, nC = sh.nC;
    const fr = Math.min(sh.freeze.rows, nR - 1), fc = Math.min(sh.freeze.cols, nC - 1);
    let z = 1, RW = 0, CH = 0, FW = 0, FH = 0, TW = 0, TH = 0, colX, rowY, gw = 1, dpr = 1;
    const st0 = book.style(0);
    const normal = book.style(0).font;
    // 図の位置（100% のときの左上からの位置）
    const colX1 = new Float64Array(nC + 1), rowY1 = new Float64Array(nR + 1);
    for (let c = 0; c < nC; c++) colX1[c + 1] = colX1[c] + sh.colW(c);
    for (let r = 0; r < nR; r++) rowY1[r + 1] = rowY1[r] + sh.rowH(r);

    function layout() {
      z = G.zoom;
      dpr = Math.min(3, window.devicePixelRatio || 1);
      gw = Math.max(1, Math.floor(dpr * 0.75)) / dpr; // 枠線の太さ（css px。画面の 1 点）
      RW = sh.headers ? Math.round((Math.max(3, String(nR).length) * 7.5 + 10) * z) : 0;
      CH = sh.headers ? Math.round(Math.max(18, Math.min(sh.defH, 30)) * z) : 0;
      colX = new Float64Array(nC + 1); rowY = new Float64Array(nR + 1);
      colX[0] = RW; rowY[0] = CH;
      for (let c = 0; c < nC; c++) colX[c + 1] = colX[c] + Math.round(sh.colW(c) * z);
      for (let r = 0; r < nR; r++) rowY[r + 1] = rowY[r] + Math.round(sh.rowH(r) * z);
      FW = colX[fc] - RW; FH = rowY[fr] - CH; TW = colX[nC] - RW; TH = rowY[nR] - CH;
      const geo = { tl: [0, 0, RW + FW, CH + FH], tr: [RW + FW, 0, TW - FW, CH + FH], bl: [0, CH + FH, RW + FW, TH - FH], br: [RW + FW, CH + FH, TW - FW, TH - FH] };
      for (const p of panes) {
        [p.gx, p.gy, p.w, p.h] = geo[p.key];
        Object.assign(p.el.style, { width: p.w + 'px', height: p.h + 'px', display: p.w > 0 && p.h > 0 ? '' : 'none' });
        for (const t of p.tiles.values()) { t.width = t.height = 0; t.remove(); }
        p.tiles.clear();
      }
      Object.assign(top.style, { width: RW + TW + 'px', height: CH + FH + 'px' });
      Object.assign(bot.style, { width: RW + TW + 'px', height: TH - FH + 'px' });
      rowCache.clear();
      placeDrawings();
      if (G.sel) showSel();
    }
    // 位置 → 列・行（二分探索）
    const colAt = x => { let lo = 0, hi = nC - 1; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (colX[m] <= x) lo = m; else hi = m - 1; } return lo; };
    const rowAt = y => { let lo = 0, hi = nR - 1; while (lo < hi) { const m = (lo + hi + 1) >> 1; if (rowY[m] <= y) lo = m; else hi = m - 1; } return lo; };

    /* ----- セルの書式（条件付き書式を重ねる） ----- */
    function styleIdx(r, c, cell) {
      if (cell) return cell.s;
      const row = sh.rows[r];
      if (row && row.s != null) return row.s;
      const cs = sh.colStyle(c);
      return cs != null ? cs : 0;
    }
    const cfStyleCache = new Map();
    function cellStyle(r, c, cell) {
      const st = book.style(styleIdx(r, c, cell));
      const cf = sh.cf.size ? sh.cf.get(Xlsx.key(r, c)) : null;
      if (!cf || (!cf.fill && !cf.font && !cf.border && !cf.fmt)) return st;
      let o = cfStyleCache.get(cf);
      if (!o || o.base !== st) {
        o = { ...st, base: st };
        if (cf.fill) o.fill = cf.fill;
        if (cf.font) { o.font = { ...st.font, ...cf.font }; if (cf.font.color) o.font.color = cf.font.color; }
        if (cf.border) o.border = { ...(st.border || {}), ...cf.border };
        if (cf.fmt) o.fmt = cf.fmt;
        cfStyleCache.set(cf, o);
      }
      return o;
    }
    const isEmpty = (r, c) => { const cell = sh.cell(r, c); return (!cell || cell.t === 'z') && !sh.mergeAt(r, c); };

    /* ----- 1 行の文字の並べ方（はみ出しも）。行ごとに覚えておく ----- */
    const rowCache = new Map();
    const padX = () => 2 * z;
    function rowLayout(r, left) {
      const k = r * 2 + (left ? 1 : 0);
      let L = rowCache.get(k);
      if (L) return L;
      L = [];
      const row = sh.rows[r];
      if (row && rowY[r + 1] > rowY[r]) {
        const cMin = left ? 0 : fc, cMax = left ? fc - 1 : nC - 1, cs = row.cells;
        for (let c = cMin; c <= cMax && c < cs.length; c++) {
          const cell = cs[c];
          if (!cell || cell.t === 'z' || sh.mergeAt(r, c)) continue;
          const it = cellItem(r, c, cell, colX[c], colX[c + 1], rowY[r], rowY[r + 1]);
          if (!it) continue;
          if (it.cc) ccSpan(it, r, c, cMax); else spill(it, r, c, cMin, cMax);
          L.push(it);
        }
      }
      if (rowCache.size > 4000) rowCache.clear();
      rowCache.set(k, L);
      return L;
    }
    // セルの文字の情報
    function cellItem(r, c, cell, x0, x1, y0, y1) {
      if (x1 - x0 <= 0 || y1 - y0 <= 0) return null;
      const st = cellStyle(r, c, cell), al = st.al || {};
      const font = fontCss(st.font, z, st.font.vert && st.font.vert !== 'baseline' ? 0.7 : 1);
      // 字下げ 1 つ = 標準の書体の空白 3 つ分
      const indent = (al.indent || 0) * 3 * tw(fontCss(normal, z), ' ');
      const cf = sh.cf.size ? sh.cf.get(Xlsx.key(r, c)) : null;
      const icon = cf && cf.icon && cell.t === 'n' ? cf.icon : null;
      const iconW = icon ? iconSize(st, y1 - y0) + padX() : 0;
      const avail = x1 - x0 - gw - 2 * padX() - indent - iconW;
      const tx = book.text(cell, st, cell.t === 'n' ? fitter(st, font, avail) : null);
      if (!tx || (!tx.text && !tx.parts.some(p => p.fill))) return null;
      const rot = al.rot || 0;
      let ha = al.h || 'general';
      // 「標準」の配置：回した文字は数・文字に関係なく、上向き（1〜89°）と −90° は左、90° と下向きは右、縦書きは中央（Excel で調べた）
      if (ha === 'general') ha = rot ? (rot === 255 ? 'center' : rot < 90 || rot === 180 ? 'left' : 'right') : tx.align === 'r' ? 'right' : tx.align === 'c' ? 'center' : 'left';
      const cc = ha === 'centerContinuous';
      if (cc) ha = 'center';
      const wrap = !tx.num && (al.wrap || ha === 'justify' || ha === 'distributed' || al.v === 'justify' || al.v === 'distributed');
      const color = (cf && cf.font && cf.font.color) || tx.color || st.font.color || '#000';
      let segs;
      if (cell.t === 'r' && cell.runs && !tx.num) {
        segs = cell.runs.map(run => {
          const f = run.f ? { ...st.font, ...run.f, fam: run.f.name ? Xlsx.family(run.f.name) : st.font.fam } : st.font;
          const fcol = run.f && run.f.color ? book.color(run.f.color, color) : color;
          return { t: run.t, font: fontCss(f, z, f.vert && f.vert !== 'baseline' ? 0.7 : 1), f, color: fcol };
        });
      } else segs = null;
      return { r, c, st, tx, ha, cc, va: al.v || 'bottom', wrap, shrink: !!al.shrink && !wrap, rot, font, color, segs, indent, x0, x1, y0, y1, clip0: x0, clip1: x1, hide: [], bar: cf && cf.bar, icon, iconW };
    }
    // 数がセルに入るか。Excel は字の幅を整数の点で数えるので、それに合わせる。
    // 「標準」の形式は、字の数（数字の幅を 1、小数点を 0.5）で決める（細い列で 1E+05 か 100000 かなど）
    function fitter(st, font, avail) {
      const d0 = tw(font, '0'), d1 = Xlsx.digitW(st.font) * z;
      const gen = /^(general|g\/標準)?$/i.test(String(st.fmt == null ? '' : st.fmt).trim());
      return t => {
        if (/^#+$/.test(t)) return tw(font, t) <= avail + 0.01;
        if (gen) { let u = 0; for (const ch of t) u += ch === '.' ? 0.5 : 1; return u * d1 <= avail + 0.01; }
        return tw(font, t) * d1 / d0 <= avail + 0.01;
      };
    }
    const iconSize = (st, rowH) => Math.max(4, Math.min(rowH - 3 * z, Math.round((st.font.sz || 11) * 4 / 3 * 1.2) * z));
    // 文字が長いとき、となりの空いたセルへはみ出す（数・折り返し・回転・縮小・結合のときは、はみ出さない）
    function spill(it, r, c, cMin, cMax) {
      if (it.tx.num || it.wrap || it.shrink || it.rot || it.ha === 'fill' || it.tx.kind === 'bool' || it.tx.kind === 'error' || it.tx.parts.some(p => p.fill)) return;
      const w = textWidth(it) + it.indent + 2 * padX() + gw;
      if (w <= it.x1 - it.x0) return;
      let a = c, b = c;
      if (it.ha === 'left' || it.ha === 'justify') {
        while (colX[b + 1] - it.x0 < w && b + 1 <= cMax && isEmpty(r, b + 1)) b++;
      } else if (it.ha === 'right') {
        while (it.x1 - colX[a] < w && a - 1 >= cMin && isEmpty(r, a - 1)) a--;
      } else {
        const need = (w - (it.x1 - it.x0)) / 2;
        while (it.x0 - colX[a] < need && a - 1 >= cMin && isEmpty(r, a - 1)) a--;
        while (colX[b + 1] - it.x1 < need && b + 1 <= cMax && isEmpty(r, b + 1)) b++;
      }
      it.clip0 = colX[a]; it.clip1 = colX[b + 1];
      for (let k = a + 1; k <= b; k++) it.hide.push(k); // この列の左の枠線は消す
    }
    // 「選択範囲内で中央」：右に続く空いたセル（同じ配置）までを 1 つの幅として中央に
    function ccSpan(it, r, c, cMax) {
      let b = c;
      while (b + 1 <= cMax && isEmpty(r, b + 1) && (cellStyle(r, b + 1, sh.cell(r, b + 1)).al || {}).h === 'centerContinuous') b++;
      if (b === c) return;
      it.x1 = it.clip1 = colX[b + 1];
      for (let k = c + 1; k <= b; k++) it.hide.push(k);
    }
    const segsOf = it => it.segs || [{ t: it.tx.text, font: it.font, f: it.st.font, color: it.color }];
    function textWidth(it) {
      if (it.segs) return it.segs.reduce((s, x) => s + tw(x.font, x.t.replace(/\n/g, '')), 0);
      return partsWidth(it.tx.parts, it.font);
    }
    function partsWidth(parts, font) {
      let w = 0;
      for (const p of parts) w += p.s != null ? tw(font, p.s) : p.pad != null ? tw(font, p.pad) : 0;
      return w;
    }

    /* ----- 描く ----- */
    function paintTile(p, i, j) {
      const lx = i * TILE, ly = j * TILE, w = Math.min(TILE, p.w - lx), hh = Math.min(TILE, p.h - ly);
      if (w <= 0 || hh <= 0) return;
      const cv = document.createElement('canvas');
      cv.width = Math.ceil(w * dpr); cv.height = Math.ceil(hh * dpr);
      Object.assign(cv.style, { left: lx + 'px', top: ly + 'px', width: w + 'px', height: hh + 'px' });
      try { paint(cv.getContext('2d'), p, p.gx + lx, p.gy + ly, w, hh); } catch (e) { console.warn('表を描けませんでした', e); }
      p.tl.appendChild(cv);
      p.tiles.set(i + ',' + j, cv);
    }
    function paint(ctx, p, gx, gy, w, hh) {
      const d = dpr;
      const X = x => Math.round((x - gx) * d), Y = y => Math.round((y - gy) * d);
      const gwd = Math.round(gw * d);
      ctx.fillStyle = '#fff';
      ctx.fillRect(0, 0, ctx.canvas.width, ctx.canvas.height);
      const cMin = p.left ? 0 : fc, cMax = p.left ? fc - 1 : nC - 1, rMin = p.upper ? 0 : fr, rMax = p.upper ? fr - 1 : nR - 1;
      const c0 = Math.max(cMin, colAt(Math.max(gx, RW))), c1 = Math.min(cMax, colAt(gx + w - 0.01));
      const r0 = Math.max(rMin, rowAt(Math.max(gy, CH))), r1 = Math.min(rMax, rowAt(gy + hh - 0.01));
      const hasCells = gx + w > RW && gy + hh > CH && c0 <= c1 && r0 <= r1 && cMin <= cMax && rMin <= rMax;
      if (hasCells) {
        const cellRect = (x0, y0, x1, y1, style) => { ctx.fillStyle = style; ctx.fillRect(X(x0), Y(y0), X(x1) - X(x0), Y(y1) - Y(y0)); };
        // 塗り（枠線の上までかぶせる）
        const fillRect = (fill, x0, y0, x1, y1) => {
          const dx0 = X(x0) - gwd, dy0 = Y(y0) - gwd, dx1 = X(x1), dy1 = Y(y1);
          if (fill.color) ctx.fillStyle = fill.color;
          else if (fill.pattern) ctx.fillStyle = pattern(ctx, fill, d);
          else if (fill.grad) {
            const g = fill.grad, W = dx1 - dx0, H = dy1 - dy0;
            let gr;
            if (g.type === 'path') gr = ctx.createRadialGradient(dx0 + W * (g.l + g.r) / 2, dy0 + H * (g.t + g.b) / 2, 0, dx0 + W * (g.l + g.r) / 2, dy0 + H * (g.t + g.b) / 2, Math.hypot(W, H) / 2);
            else { const a = g.deg * Math.PI / 180, cx = dx0 + W / 2, cy = dy0 + H / 2, L = (Math.abs(W * Math.cos(a)) + Math.abs(H * Math.sin(a))) / 2; gr = ctx.createLinearGradient(cx - Math.cos(a) * L, cy - Math.sin(a) * L, cx + Math.cos(a) * L, cy + Math.sin(a) * L); }
            for (const s of g.stops) gr.addColorStop(Math.max(0, Math.min(1, s.pos)), s.color);
            ctx.fillStyle = gr;
          } else return;
          ctx.fillRect(dx0, dy0, dx1 - dx0, dy1 - dy0);
        };
        // 1. 枠線
        if (sh.grid) {
          ctx.fillStyle = GRID;
          const ya = Y(rowY[r0]) - gwd, yb = Y(rowY[r1 + 1]), xa = X(colX[c0]) - gwd, xb = X(colX[c1 + 1]);
          for (let c = c0; c <= c1; c++) ctx.fillRect(X(colX[c + 1]) - gwd, ya, gwd, yb - ya);
          for (let r = r0; r <= r1; r++) ctx.fillRect(xa, Y(rowY[r + 1]) - gwd, xb - xa, gwd);
        }
        // 2. 塗り
        const anyFill = st0.fill || sh.cf.size || true;
        if (anyFill) for (let r = r0; r <= r1; r++) {
          if (rowY[r + 1] <= rowY[r]) continue;
          const row = sh.rows[r];
          const rowStyled = row && row.s != null;
          for (let c = c0; c <= c1; c++) {
            const cell = row && row.cells[c];
            if (!cell && !rowStyled && sh.colStyle(c) == null && !st0.fill && !sh.cf.size) continue;
            if (sh.mergeAt(r, c)) continue;
            const st = cellStyle(r, c, cell);
            if (st.fill) fillRect(st.fill, colX[c], rowY[r], colX[c + 1], rowY[r + 1]);
          }
        }
        // 3. 結合したセル（中の枠線を消して、左上のセルの塗り）
        const merges = [];
        for (const g of sh.merges) {
          if (g.r1 < r0 || g.r0 > r1 || g.c1 < c0 || g.c0 > c1) continue;
          merges.push(g);
          const st = cellStyle(g.r0, g.c0, sh.cell(g.r0, g.c0));
          const x0 = colX[g.c0], x1 = colX[Math.min(nC, g.c1 + 1)], y0 = rowY[g.r0], y1 = rowY[Math.min(nR, g.r1 + 1)];
          if (st.fill) fillRect(st.fill, x0, y0, x1, y1);
          else { ctx.fillStyle = '#fff'; ctx.fillRect(X(x0), Y(y0), X(x1) - gwd - X(x0), Y(y1) - gwd - Y(y0)); }
        }
        // 4. はみ出した文字の下の枠線を消す
        const items = [];
        for (let r = r0; r <= r1; r++) {
          for (const it of rowLayout(r, p.left)) {
            if (it.clip1 <= gx || it.clip0 >= gx + w) continue;
            items.push(it);
            for (const k of it.hide) {
              const lc = cellStyle(r, k - 1, sh.cell(r, k - 1)), rc = cellStyle(r, k, sh.cell(r, k));
              if (lc.fill || rc.fill) continue;
              ctx.fillStyle = '#fff';
              ctx.fillRect(X(colX[k]) - gwd, Y(rowY[r]), gwd, Y(rowY[r + 1]) - gwd - Y(rowY[r]));
            }
          }
        }
        // 5. データバー
        for (const it of items) if (it.bar) drawBar(ctx, it, X, Y);
        // 6. 文字
        ctx.save();
        ctx.setTransform(d, 0, 0, d, -gx * d, -gy * d);
        for (const it of items) drawItem(ctx, it);
        for (const g of merges) {
          const cell = sh.cell(g.r0, g.c0);
          if (!cell || cell.t === 'z') continue;
          const x0 = colX[g.c0], x1 = colX[Math.min(nC, g.c1 + 1)], y0 = rowY[g.r0], y1 = rowY[Math.min(nR, g.r1 + 1)];
          const it = cellItem(g.r0, g.c0, cell, x0, x1, y0, y1);
          if (it) { if (it.bar) { ctx.restore(); drawBar(ctx, it, X, Y); ctx.save(); ctx.setTransform(d, 0, 0, d, -gx * d, -gy * d); } drawItem(ctx, it); }
        }
        ctx.restore();
        // 7. 罫線（強いほうを描く）
        drawBorders(ctx, X, Y, gwd, Math.max(rMin, r0 - 1), Math.min(rMax, r1 + 1), Math.max(cMin, c0 - 1), Math.min(cMax, c1 + 1));
        // 8. メモの印（右上の赤い三角）
        if (sh.notes.size) {
          const s = Math.round(Math.max(4, 6 * Math.min(z, 1.5)) * d);
          for (const [k, n] of sh.notes) {
            const r = Math.floor(k / 16384), c = k % 16384;
            if (r < r0 || r > r1 || c < c0 || c > c1) continue;
            const g = sh.mergeAt(r, c), cc = g ? g.c1 : c;
            const xr = X(colX[Math.min(nC, cc + 1)]) - gwd, yt = Y(rowY[g ? g.r0 : r]);
            ctx.fillStyle = n.thread ? '#7D3CBD' : '#E0301E';
            ctx.beginPath(); ctx.moveTo(xr - s, yt); ctx.lineTo(xr, yt); ctx.lineTo(xr, yt + s); ctx.closePath(); ctx.fill();
          }
        }
      }
      // 列の記号・行の番号
      if (sh.headers) paintHeaders(ctx, p, gx, gy, w, hh, X, Y, gwd, cMin, cMax, rMin, rMax);
    }
    function paintHeaders(ctx, p, gx, gy, w, hh, X, Y, gwd, cMin, cMax, rMin, rMax) {
      const font = fontCss({ ...normal, b: false, i: false, sz: Math.min(11, normal.sz || 11) }, z, 0.9), v = vm(font);
      const dTxt = (t, cx, cy) => {
        ctx.save(); ctx.setTransform(dpr, 0, 0, dpr, -gx * dpr, -gy * dpr);
        ctx.font = font; ctx.fillStyle = HFG; ctx.textAlign = 'center';
        ctx.fillText(t, cx, cy + (v.asc - v.desc) / 2);
        ctx.restore();
      };
      if (gy < CH && cMin <= cMax) {
        const a = Math.max(cMin, colAt(Math.max(gx, RW))), b = Math.min(cMax, colAt(gx + w - 0.01));
        ctx.fillStyle = HBG; ctx.fillRect(X(colX[a]), Y(0), X(colX[b + 1]) - X(colX[a]), Y(CH) - Y(0));
        ctx.fillStyle = HLINE;
        ctx.fillRect(X(colX[a]), Y(CH) - gwd, X(colX[b + 1]) - X(colX[a]), gwd);
        for (let c = a; c <= b; c++) {
          if (colX[c + 1] <= colX[c]) continue;
          ctx.fillStyle = HLINE; ctx.fillRect(X(colX[c + 1]) - gwd, Y(0), gwd, Y(CH) - Y(0));
          if (colX[c + 1] - colX[c] > 8) dTxt(Xlsx.colName(c), (colX[c] + colX[c + 1]) / 2, CH / 2);
        }
      }
      if (gx < RW && rMin <= rMax) {
        const a = Math.max(rMin, rowAt(Math.max(gy, CH))), b = Math.min(rMax, rowAt(gy + hh - 0.01));
        ctx.fillStyle = HBG; ctx.fillRect(X(0), Y(rowY[a]), X(RW) - X(0), Y(rowY[b + 1]) - Y(rowY[a]));
        ctx.fillStyle = HLINE;
        ctx.fillRect(X(RW) - gwd, Y(rowY[a]), gwd, Y(rowY[b + 1]) - Y(rowY[a]));
        for (let r = a; r <= b; r++) {
          if (rowY[r + 1] <= rowY[r]) continue;
          ctx.fillStyle = HLINE; ctx.fillRect(X(0), Y(rowY[r + 1]) - gwd, X(RW) - X(0), gwd);
          if (rowY[r + 1] - rowY[r] > 6) dTxt(String(r + 1), RW / 2, (rowY[r] + rowY[r + 1]) / 2);
        }
      }
      if (gx < RW && gy < CH) {
        ctx.fillStyle = HBG; ctx.fillRect(X(0), Y(0), X(RW) - X(0), Y(CH) - Y(0));
        ctx.fillStyle = HLINE; ctx.fillRect(X(RW) - gwd, Y(0), gwd, Y(CH) - Y(0)); ctx.fillRect(X(0), Y(CH) - gwd, X(RW) - X(0), gwd);
        ctx.fillStyle = '#B8B8B8';
        const s = Math.min(RW, CH) * 0.45;
        ctx.beginPath(); ctx.moveTo(X(RW - 3 * z), Y(CH - 3 * z - s)); ctx.lineTo(X(RW - 3 * z), Y(CH - 3 * z)); ctx.lineTo(X(RW - 3 * z - s), Y(CH - 3 * z)); ctx.closePath(); ctx.fill();
      }
    }
    // データバー
    function drawBar(ctx, it, X, Y) {
      const b = it.bar, x0 = it.x0 + 1, W = it.x1 - it.x0 - gw - 2, y0 = it.y0 + 2, y1 = it.y1 - gw - 2;
      if (W <= 0 || y1 <= y0) return;
      const a = X(x0 + W * Math.min(b.from, b.to)), e = X(x0 + W * Math.max(b.from, b.to)), t = Y(y0), u = Y(y1);
      if (e - a < 1) return;
      if (b.grad) {
        const gr = ctx.createLinearGradient(b.to >= b.from ? a : e, 0, b.to >= b.from ? e : a, 0);
        gr.addColorStop(0, b.color); gr.addColorStop(1, mixWhite(b.color, 0.9));
        ctx.fillStyle = gr;
      } else ctx.fillStyle = b.color;
      ctx.fillRect(a, t, e - a, u - t);
      if (b.border) { ctx.strokeStyle = b.border; ctx.lineWidth = Math.max(1, Math.round(dpr)); ctx.strokeRect(a + 0.5, t + 0.5, e - a - 1, u - t - 1); }
    }
    const mixWhite = (hex, k) => { const n = parseInt(hex.slice(1), 16); const f = v => Math.round(v + (255 - v) * k); return `rgb(${f(n >> 16)},${f((n >> 8) & 255)},${f(n & 255)})`; };

    // セルの文字を描く（ctx は css px の座標）
    function drawItem(ctx, it) {
      const x0 = it.x0 + it.iconW, x1 = it.x1 - gw, y0 = it.y0, y1 = it.y1 - gw;
      if (it.bar && !it.bar.show) return;
      ctx.save();
      ctx.beginPath(); ctx.rect(it.clip0, y0, it.clip1 - gw - it.clip0, y1 - y0); ctx.clip();
      const pad = padX(), lr = Xlsx.lineRatio(it.st.font.name);
      // アイコン（条件付き書式）はセルの左に
      if (it.icon) {
        const s = it.iconW - pad, va = it.va;
        const iy = va === 'top' ? y0 + 1.5 * z : va === 'center' || va === 'justify' || va === 'distributed' ? (y0 + y1 - s) / 2 : y1 - 1.5 * z - s;
        drawIcon(ctx, it.icon, it.x0 + pad, iy, s);
        if (!it.icon.show) { ctx.restore(); return; }
      }
      let segs = segsOf(it), font = it.font;
      const lineH = Math.round((it.st.font.sz || 11) * lr * 4 / 3) * z;
      const avail = x1 - x0 - 2 * pad - it.indent;
      // 縮小して全体を表示
      if (it.shrink) {
        const w0 = textWidth(it);
        if (w0 > avail && w0 > 0) {
          const k = Math.max(0.1, avail / w0);
          segs = segs.map(s => ({ ...s, font: s.font.replace(/([\d.]+)px/, (m, v) => (+v * k).toFixed(2) + 'px') }));
          font = font.replace(/([\d.]+)px/, (m, v) => (+v * k).toFixed(2) + 'px');
        }
      }
      if (it.rot) { drawRotated(ctx, it, segs, lineH, x0, x1, y0, y1, pad); ctx.restore(); return; }
      let lines;
      if (it.wrap) lines = breakLines(segs, Math.max(1, avail));
      else if (it.segs || !it.tx.parts.some(p => p.pad != null || p.fill != null)) {
        const ss = segs.map(s => ({ ...s, t: s.t.replace(/\r?\n/g, '') }));
        lines = [{ segs: ss, w: ss.reduce((a, s) => a + tw(s.font, s.t), 0) }];
      } else lines = null; // 数の形式の空白・くり返し（会計の形式など）
      const v = vm(font);
      const asc = v.asc, dsc = v.desc;
      const n = lines ? lines.length : 1, block = n * lineH;
      const va = it.va;
      let top = va === 'top' || va === 'justify' || va === 'distributed' ? y0 + 1 : va === 'center' ? (y0 + y1 - block) / 2 : y1 - 1 - block;
      if (it.wrap && block > y1 - y0 && va !== 'top') top = va === 'center' ? (y0 + y1 - block) / 2 : y1 - 1 - block;
      const base = i => top + i * lineH + (lineH - (asc + dsc)) / 2 + asc;
      ctx.textBaseline = 'alphabetic';
      if (!lines) {
        // 数の形式の部品（_x は空白、*x はくり返して幅いっぱい）
        const parts = it.tx.parts, fixed = partsWidth(parts, font);
        const nf = parts.filter(p => p.fill != null).length;
        const fillW = nf ? Math.max(0, avail - fixed) / nf : 0;
        // Excel では入る数が、この端末の書体では少し広いとき（ＭＳ Ｐゴシックなど）は、字の間を少しつめる
        const nCh = parts.reduce((a, p) => a + (p.s != null ? [...p.s].length : 0), 0);
        const ls = !nf && nCh && fixed > avail + 0.5 ? squeeze(ctx, fixed, avail, nCh) : 0, fw = fixed + ls * nCh;
        let x = nf ? x0 + pad + (it.ha === 'right' ? 0 : it.indent) : it.ha === 'right' ? x1 - pad - it.indent - fw : it.ha === 'center' ? (x0 + x1 - fw) / 2 : x0 + pad + it.indent;
        ctx.font = font; ctx.fillStyle = it.color;
        const y = base(0);
        for (const p of parts) {
          if (p.s != null) { const pw = tw(font, p.s) + ls * [...p.s].length; ctx.fillText(p.s, x, y); deco(ctx, it.st.font, x, y, pw, v); x += pw; }
          else if (p.pad != null) x += tw(font, p.pad);
          else if (p.fill != null) {
            const cw = tw(font, p.fill);
            if (p.fill !== ' ' && cw > 0) { const k = Math.floor(fillW / cw); for (let j = 0; j < k; j++) ctx.fillText(p.fill, x + j * cw, y); }
            x += fillW;
          }
        }
        ctx.restore();
        return;
      }
      // 数（1 行）が、この端末の書体で少し広いときは、字の間を少しつめる（上と同じ）
      if (it.tx.num && lines.length === 1 && lines[0].w > avail + 0.5 && it.ha !== 'fill') {
        const ln = lines[0], n = [...ln.segs.map(q => q.t).join('')].length;
        const w = ln.w + squeeze(ctx, ln.w, avail, n) * n;
        const x = it.ha === 'right' ? x1 - pad - it.indent - w : it.ha === 'center' ? (x0 + x1 - w) / 2 : x0 + pad + it.indent;
        drawSegs(ctx, ln.segs, x, base(0));
        ctx.restore();
        return;
      }
      lines.forEach((ln, i) => {
        const y = base(i);
        let x;
        if (it.ha === 'right') x = x1 - pad - it.indent - ln.w;
        else if (it.ha === 'center' || (it.ha === 'distributed' && !it.wrap)) x = (x0 + x1 - ln.w) / 2;
        else if (it.ha === 'fill') x = x0 + pad;
        else x = x0 + pad + it.indent;
        if (it.ha === 'fill' && ln.w > 0) {
          // 文字をくり返してセルを埋める
          const k = Math.max(1, Math.floor(avail / ln.w));
          for (let j = 0; j < k; j++) drawSegs(ctx, ln.segs, x + j * ln.w, y);
        } else if (it.ha === 'distributed' || (it.ha === 'justify' && i < lines.length - 1)) {
          // 均等割り付け（どの行も両端まで）・両端揃え（最後の行のほかは両端まで）
          drawSpread(ctx, ln, x0 + pad + it.indent, avail - (it.ha === 'distributed' ? it.indent : 0), y);
        } else drawSegs(ctx, ln.segs, x, y);
      });
      ctx.restore();
    }
    // 字の間をつめる量（1 字あたり。広さの 1/4 まで）。字の間を変えられない古いブラウザでは、つめない
    function squeeze(ctx, w, avail, n) {
      if (!('letterSpacing' in ctx)) return 0;
      const ls = Math.max(-w * 0.25 / n, (avail - w) / n);
      ctx.letterSpacing = ls + 'px';
      return ls;
    }
    function drawSpread(ctx, ln, x, width, y) {
      const chars = [];
      for (const s of ln.segs) for (const ch of s.t) chars.push({ ch, s });
      const n = chars.length, cjk = chars.some(c => CJK.test(c.ch));
      const gaps = cjk ? n - 1 : chars.filter((c, i) => c.ch === ' ' && i < n - 1).length;
      if (n < 2 || !gaps || ln.w >= width) { drawSegs(ctx, ln.segs, n < 2 ? x + (width - ln.w) / 2 : x, y); return; }
      const extra = (width - ln.w) / gaps;
      chars.forEach((c, i) => {
        drawSegs(ctx, [{ ...c.s, t: c.ch }], x, y);
        x += tw(c.s.font, c.ch) + (i < n - 1 && (cjk || c.ch === ' ') ? extra : 0);
      });
    }
    function drawSegs(ctx, segs, x, y) {
      for (const s of segs) {
        if (!s.t) continue;
        ctx.font = s.font; ctx.fillStyle = s.color;
        const v = vm(s.font), w = tw(s.font, s.t);
        let yy = y;
        if (s.f && s.f.vert === 'superscript') yy -= v.px * 0.45; else if (s.f && s.f.vert === 'subscript') yy += v.px * 0.2;
        ctx.fillText(s.t, x, yy);
        deco(ctx, s.f, x, yy, w, v);
        x += w;
      }
    }
    // 下線・取り消し線
    function deco(ctx, f, x, y, w, v) {
      if (!f || (!f.u && !f.strike) || w <= 0) return;
      const t = Math.max(1 / dpr, v.px / 16);
      if (f.u && f.u !== 'none') {
        const uy = y + Math.max(1, v.px * 0.1);
        ctx.fillRect(x, uy, w, t);
        if (/double/i.test(f.u)) ctx.fillRect(x, uy + t * 2, w, t);
      }
      if (f.strike) ctx.fillRect(x, y - v.px * 0.3, w, t);
    }
    // 回転した文字・縦に並べた文字
    function drawRotated(ctx, it, segs, lineH, x0, x1, y0, y1, pad) {
      const text = segs.map(s => s.t).join('').replace(/\n/g, '');
      const font = segs[0].font, v = vm(font);
      ctx.font = font; ctx.fillStyle = segs[0].color;
      if (it.rot === 255) {
        // かっこ・長音は縦書き用の形に（Excel も 「 を ﹁ で出す）
        const chars = [...text].map(ch => VERT[ch] || ch), n = chars.length, H = n * lineH;
        const cx = it.ha === 'left' ? x0 + pad + v.px / 2 : it.ha === 'right' ? x1 - pad - v.px / 2 : (x0 + x1) / 2;
        let ty = it.va === 'top' ? y0 + 1 : it.va === 'center' ? (y0 + y1 - H) / 2 : y1 - 1 - H;
        ctx.textAlign = 'center';
        for (const ch of chars) { ctx.fillText(ch, cx, ty + (lineH - (v.asc + v.desc)) / 2 + v.asc); ty += lineH; }
        ctx.textAlign = 'start';
        return;
      }
      const deg = it.rot <= 90 ? it.rot : -(it.rot - 90), a = deg * Math.PI / 180;
      const w = tw(font, text), W = Math.abs(w * Math.cos(a)) + Math.abs(lineH * Math.sin(a)), H = Math.abs(w * Math.sin(a)) + Math.abs(lineH * Math.cos(a));
      const cx = it.ha === 'left' ? x0 + pad + W / 2 : it.ha === 'right' ? x1 - pad - W / 2 : (x0 + x1) / 2;
      const cy = it.va === 'top' ? y0 + 1 + H / 2 : it.va === 'center' ? (y0 + y1) / 2 : y1 - 1 - H / 2;
      ctx.translate(cx, cy);
      ctx.rotate(-a);
      ctx.fillText(text, -w / 2, (v.asc - v.desc) / 2);
    }
    // 罫線：同じ線に 2 つのセルの指定があれば、強いほう
    function drawBorders(ctx, X, Y, gwd, r0, r1, c0, c1) {
      const E = new Map();
      const put = (k, L) => { const o = E.get(k); if (!o || L.k > o.k) E.set(k, L); };
      const diag = [];
      for (let r = r0; r <= r1; r++) {
        const row = sh.rows[r];
        for (let c = c0; c <= c1; c++) {
          const cell = row && row.cells[c];
          if (!cell && !(row && row.s != null) && sh.colStyle(c) == null && !sh.cf.size) continue;
          const st = cellStyle(r, c, cell), b = st.border;
          if (!b) continue;
          const g = sh.mergeAt(r, c);
          if (b.t && !(g && r > g.r0)) put(`h${r}:${c}`, b.t);
          if (b.b && !(g && r < g.r1)) put(`h${r + 1}:${c}`, b.b);
          if (b.l && !(g && c > g.c0)) put(`v${r}:${c}`, b.l);
          if (b.r && !(g && c < g.c1)) put(`v${r}:${c + 1}`, b.r);
          if (b.d && (!g || (r === g.r0 && c === g.c0))) diag.push({ r, c, g, b });
        }
      }
      for (const [k, L] of E) {
        const hz = k[0] === 'h', [a, b] = k.slice(1).split(':').map(Number);
        // 太さは画面の点の数（Excel は拡大しても細い線は 1 点のまま。細い 1・中 2・太い 3）
        const t = L.w * gwd;
        const e = Math.floor((t - gwd) / 2);
        ctx.fillStyle = L.color;
        if (hz) {
          if (a > nR) continue;
          const yb = Y(rowY[Math.min(a, nR)]) - gwd - e, xa = X(colX[b]) - gwd - e, xb = X(colX[b + 1]) + Math.max(0, t - gwd - e);
          if (colX[b + 1] <= colX[b]) continue;
          line(ctx, L, true, xa, xb, yb, t);
        } else {
          if (b > nC) continue;
          const xb = X(colX[Math.min(b, nC)]) - gwd - e, ya = Y(rowY[a]) - gwd - e, yc = Y(rowY[a + 1]) + Math.max(0, t - gwd - e);
          if (rowY[a + 1] <= rowY[a]) continue;
          line(ctx, L, false, ya, yc, xb, t);
        }
      }
      for (const { r, c, g, b } of diag) {
        const x0 = X(colX[g ? g.c0 : c]), x1 = X(colX[(g ? g.c1 : c) + 1]) - gwd, y0 = Y(rowY[g ? g.r0 : r]), y1 = Y(rowY[(g ? g.r1 : r) + 1]) - gwd;
        ctx.strokeStyle = b.d.color; ctx.lineWidth = b.d.w * gwd;
        ctx.setLineDash(b.d.dash ? b.d.dash.map(v => v * gwd) : []);
        ctx.beginPath();
        if (b.dd) { ctx.moveTo(x0, y0); ctx.lineTo(x1, y1); }
        if (b.du) { ctx.moveTo(x0, y1); ctx.lineTo(x1, y0); }
        ctx.stroke();
        ctx.setLineDash([]);
      }
    }
    // 罫線 1 本（点線は自分で区切る。二重線は 1 点ずつ 2 本）
    function line(ctx, L, hz, a, b, pos, t) {
      const rect = (s, e, p, th) => (hz ? ctx.fillRect(s, p, e - s, th) : ctx.fillRect(p, s, th, e - s));
      const u = Math.max(1, Math.round(gw * dpr));
      if (L.dbl) { rect(a, b, pos, u); rect(a, b, pos + t - u, u); return; }
      if (!L.dash) { rect(a, b, pos, t); return; }
      const pat = L.dash.map(v => v * u);
      let x = a, k = 0;
      while (x < b) { const len = pat[k % pat.length]; if (k % 2 === 0) rect(x, Math.min(b, x + len), pos, t); x += len; k++; }
    }

    /* ----- 見えている所のタイルを用意する ----- */
    let raf = 0;
    function update() {
      raf = 0;
      if (G.dead) return;
      const sL = sc.scrollLeft, sT = sc.scrollTop, vw = sc.clientWidth, vh = sc.clientHeight;
      const xA = RW + FW + sL, xB = sL + vw, yA = CH + FH + sT, yB = sT + vh;
      vis(P.br, xA, xB, yA, yB); vis(P.tr, xA, xB, 0, CH + FH); vis(P.bl, 0, RW + FW, yA, yB); vis(P.tl, 0, RW + FW, 0, CH + FH);
    }
    function vis(p, x0, x1, y0, y1) {
      if (p.w <= 0 || p.h <= 0) return;
      const lx0 = x0 - p.gx - MARGIN, lx1 = x1 - p.gx + MARGIN, ly0 = y0 - p.gy - MARGIN, ly1 = y1 - p.gy + MARGIN;
      const ni = Math.ceil(p.w / TILE), nj = Math.ceil(p.h / TILE);
      const i0 = Math.max(0, Math.floor(lx0 / TILE)), i1 = Math.min(ni - 1, Math.floor(lx1 / TILE));
      const j0 = Math.max(0, Math.floor(ly0 / TILE)), j1 = Math.min(nj - 1, Math.floor(ly1 / TILE));
      for (const [k, cv] of p.tiles) {
        const [i, j] = k.split(',').map(Number);
        if (i < i0 - 1 || i > i1 + 1 || j < j0 - 1 || j > j1 + 1) { cv.width = cv.height = 0; cv.remove(); p.tiles.delete(k); }
      }
      for (let j = j0; j <= j1; j++) for (let i = i0; i <= i1; i++) if (!p.tiles.has(i + ',' + j)) paintTile(p, i, j);
    }
    const onScroll = () => { if (!raf) raf = requestAnimationFrame(update); };
    sc.addEventListener('scroll', onScroll, { passive: true });
    const ro = new ResizeObserver(() => onScroll());
    ro.observe(sc);

    /* ----- 図 ----- */
    const drawings = [];
    function placeDrawings() {
      for (const d of drawings) {
        const b = boxOf(d.a, false), p = paneOf(d.a.from ? d.a.from.r : 0, d.a.from ? d.a.from.c : 0);
        if (!b) { d.el.style.display = 'none'; continue; }
        if (d.el.parentNode !== p.dl) p.dl.appendChild(d.el);
        Object.assign(d.el.style, { display: '', left: b.x - p.gx + 'px', top: b.y - p.gy + 'px', transform: `scale(${d.w1 ? b.w / d.w1 : 1},${d.h1 ? b.h / d.h1 : 1})` });
      }
    }
    // 図の場所（one：100% のとき、左上を 0 とする）
    function boxOf(a, one) {
      const cx = one ? colX1 : colX, ry = one ? rowY1 : rowY, k = one ? 1 : z;
      const cw = c => cx[Math.min(c + 1, nC)] - cx[Math.min(c, nC)], rh = r => ry[Math.min(r + 1, nR)] - ry[Math.min(r, nR)];
      const pt = q => ({ x: cx[Math.min(q.c, nC)] + Math.min(q.dx * k, cw(q.c)), y: ry[Math.min(q.r, nR)] + Math.min(q.dy * k, rh(q.r)) });
      let x, y, x1, y1;
      if (a.type === 'absoluteAnchor' && a.pos && a.ext) { x = cx[0] + a.pos.x * k; y = ry[0] + a.pos.y * k; x1 = x + a.ext.w * k; y1 = y + a.ext.h * k; }
      else if (a.from && a.to && a.type === 'twoCellAnchor') { const p0 = pt(a.from), p1 = pt(a.to); x = p0.x; y = p0.y; x1 = p1.x; y1 = p1.y; }
      else if (a.from && a.ext) { const p0 = pt(a.from); x = p0.x; y = p0.y; x1 = x + a.ext.w * k; y1 = y + a.ext.h * k; }
      else return null;
      return { x, y, w: Math.max(0, x1 - x), h: Math.max(0, y1 - y) };
    }
    const paneOf = (r, c) => (r < fr ? (c < fc ? P.tl : P.tr) : c < fc ? P.bl : P.br);
    async function buildDrawings() {
      if (!sh.anchors.length) return;
      const ctx = { theme: book.theme, map: CLRMAP, ph: null, src: { path: sh.drawingPath, rels: sh.drawingRels }, media: book.P.media, recolor: book.P.recolor, xml: book.P.xml };
      for (const a of sh.anchors) {
        const b = boxOf(a, true);
        if (!b || G.dead) continue;
        const el = h('div', 'xg-dw', { width: b.w + 'px', height: b.h + 'px' });
        const d = { a, el, w1: b.w, h1: b.h };
        drawings.push(d);
        try {
          const own = a.el.localName === 'grpSp' ? xfrmOf(kid(a.el, 'grpSpPr')) : a.el.localName === 'graphicFrame' ? xfrmOf(a.el) : xfrmOf(kid(a.el, 'spPr'));
          await drawObj(a.el, ctx, el, { ...(own || {}), x: 0, y: 0, w: b.w, h: b.h });
        } catch (e) { console.warn('図を描けませんでした', e); }
        if (G.dead) return;
        placeDrawings();
      }
    }

    /* ----- 選ぶ（タップ） ----- */
    const selEl = h('div', 'xg-sel'), hlC = h('div', 'xg-hl'), hlR = h('div', 'xg-hl');
    function showSel() {
      const s = G.sel;
      if (!s) { selEl.remove(); hlC.remove(); hlR.remove(); return; }
      const g = sh.mergeAt(s.r, s.c) || { r0: s.r, c0: s.c, r1: s.r, c1: s.c };
      const x0 = colX[g.c0], x1 = colX[Math.min(nC, g.c1 + 1)], y0 = rowY[g.r0], y1 = rowY[Math.min(nR, g.r1 + 1)];
      const p = paneOf(g.r0, g.c0);
      Object.assign(selEl.style, { left: x0 - p.gx - 1 + 'px', top: y0 - p.gy - 1 + 'px', width: x1 - x0 + 2 + 'px', height: y1 - y0 + 2 + 'px' });
      p.ol.appendChild(selEl);
      if (sh.headers) {
        const pc = g.c0 < fc ? P.tl : P.tr, pr = g.r0 < fr ? P.tl : P.bl;
        Object.assign(hlC.style, { left: x0 - pc.gx + 'px', top: 0 + 'px', width: x1 - x0 + 'px', height: CH + 'px' });
        Object.assign(hlR.style, { left: 0 + 'px', top: y0 - pr.gy + 'px', width: RW + 'px', height: y1 - y0 + 'px' });
        pc.ol.appendChild(hlC); pr.ol.appendChild(hlR);
      }
    }
    function select(r, c, quiet) {
      r = Math.max(0, Math.min(nR - 1, r)); c = Math.max(0, Math.min(nC - 1, c));
      const g = sh.mergeAt(r, c);
      if (g) { r = g.r0; c = g.c0; }
      G.sel = { r, c };
      showSel();
      if (!quiet && opts.onSelect) opts.onSelect(info(r, c));
    }
    function info(r, c) {
      const cell = sh.cell(r, c), k = Xlsx.key(r, c), g = sh.mergeAt(r, c);
      const st = cellStyle(r, c, cell);
      return { r, c, addr: Xlsx.addr(r, c) + (g ? ':' + Xlsx.addr(g.r1, g.c1) : ''), cell, formula: sh.formula(r, c), value: valueText(cell, st), note: sh.notes.get(k), link: sh.links.get(k) };
    }
    // 上の欄に出す値（Excel の数式バーと同じように、日付は日付・％は％で）
    function valueText(cell, st) {
      if (!cell || cell.t === 'z') return '';
      if (cell.t === 'b') return cell.v ? 'TRUE' : 'FALSE';
      if (cell.t === 'e') return cell.v;
      if (cell.t !== 'n') return (st.quote ? "'" : '') + cell.v;
      const v = cell.v;
      if (NumFmt.isDate(st.fmt) && v >= 0 && v < 2958466) {
        const f = v % 1 ? (v < 1 ? 'h:mm:ss' : 'yyyy/m/d h:mm:ss') : 'yyyy/m/d';
        return NumFmt.format(v, f, { date1904: book.date1904 }).text;
      }
      const plain = x => { let s = String(+x.toPrecision(15)); const m = /^(-?[\d.]+)e([-+])(\d+)$/.exec(s); if (m) s = `${m[1]}E${m[2]}${m[3].padStart(2, '0')}`; return s; };
      if (/%/.test(String(st.fmt).replace(/"[^"]*"/g, ''))) return plain(v * 100) + '%';
      return plain(v);
    }
    function pointToCell(e) {
      const pe = e.target.closest && e.target.closest('.xg-pane');
      const p = panes.find(q => q.el === pe);
      if (!p) return null;
      const rc = p.el.getBoundingClientRect();
      const gx = e.clientX - rc.left + p.gx, gy = e.clientY - rc.top + p.gy;
      if (gx < RW || gy < CH) return null;
      let c = colAt(gx), r = rowAt(gy);
      if (p.left) c = Math.min(c, fc - 1); else c = Math.max(c, fc);
      if (p.upper) r = Math.min(r, fr - 1); else r = Math.max(r, fr);
      return { r, c };
    }
    const onClick = e => { const q = pointToCell(e); if (q) select(q.r, q.c); };
    sc.addEventListener('click', onClick);
    // 矢印キーで動かす
    const onKey = e => {
      if (!G.sel) return;
      const mv = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
      if (!mv) return;
      e.preventDefault();
      let { r, c } = G.sel;
      const g = sh.mergeAt(r, c);
      if (g) { if (mv[0] > 0) r = g.r1; if (mv[1] > 0) c = g.c1; }
      do { r += mv[0]; c += mv[1]; } while (r > 0 && r < nR - 1 && c > 0 && c < nC - 1 && (rowY[r + 1] === rowY[r] || colX[c + 1] === colX[c]));
      select(r, c);
      goto(G.sel.r, G.sel.c, true);
    };
    sc.addEventListener('keydown', onKey);
    // そのセルが見える所までスクロール
    function goto(r, c, gentle) {
      r = Math.max(0, Math.min(nR - 1, r)); c = Math.max(0, Math.min(nC - 1, c));
      const vw = sc.clientWidth - RW - FW, vh = sc.clientHeight - CH - FH;
      if (r >= fr) {
        const y0 = rowY[r] - CH - FH, y1 = rowY[r + 1] - CH - FH;
        if (!gentle || y0 < sc.scrollTop) sc.scrollTop = gentle ? y0 : Math.max(0, y0 - 20 * z);
        else if (y1 > sc.scrollTop + vh) sc.scrollTop = y1 - vh;
      }
      if (c >= fc) {
        const x0 = colX[c] - RW - FW, x1 = colX[c + 1] - RW - FW;
        if (!gentle || x0 < sc.scrollLeft) sc.scrollLeft = gentle ? x0 : Math.max(0, x0 - 20 * z);
        else if (x1 > sc.scrollLeft + vw) sc.scrollLeft = x1 - vw;
      }
      onScroll();
    }

    /* ----- 拡大縮小（ボタン・ピンチ・Ctrl＋ホイール） ----- */
    function setZoom(nz, focal) {
      nz = Math.max(0.25, Math.min(4, Math.round(nz * 100) / 100));
      if (Math.abs(nz - G.zoom) < 0.001) return;
      const fx = focal ? focal.x : RW + FW, fy = focal ? focal.y : CH + FH;
      const ux = (sc.scrollLeft + fx - RW) / z, uy = (sc.scrollTop + fy - CH) / z;
      G.zoom = nz;
      layout();
      sc.scrollLeft = Math.max(0, ux * z + RW - fx);
      sc.scrollTop = Math.max(0, uy * z + CH - fy);
      update();
      if (opts.onZoom) opts.onZoom(nz);
    }
    let pinch = null;
    const dist = e => Math.hypot(e.touches[0].clientX - e.touches[1].clientX, e.touches[0].clientY - e.touches[1].clientY);
    const onTS = e => {
      if (e.touches.length !== 2) return;
      const rc = sc.getBoundingClientRect();
      pinch = { d0: dist(e), z0: G.zoom, k: 1, x: (e.touches[0].clientX + e.touches[1].clientX) / 2 - rc.left, y: (e.touches[0].clientY + e.touches[1].clientY) / 2 - rc.top };
    };
    const onTM = e => {
      if (!pinch || e.touches.length !== 2) return;
      e.preventDefault();
      pinch.k = Math.max(0.25 / pinch.z0, Math.min(4 / pinch.z0, dist(e) / pinch.d0));
      sc.style.transformOrigin = `${pinch.x}px ${pinch.y}px`;
      sc.style.transform = `scale(${pinch.k})`;
    };
    const onTE = e => {
      if (!pinch || e.touches.length >= 2) return;
      const p = pinch;
      pinch = null;
      sc.style.transform = '';
      if (Math.abs(p.k - 1) > 0.03) setZoom(p.z0 * p.k, { x: p.x, y: p.y });
    };
    sc.addEventListener('touchstart', onTS, { passive: true });
    sc.addEventListener('touchmove', onTM, { passive: false });
    sc.addEventListener('touchend', onTE);
    sc.addEventListener('touchcancel', onTE);
    const onWheel = e => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      const rc = sc.getBoundingClientRect();
      setZoom(G.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1), { x: e.clientX - rc.left, y: e.clientY - rc.top });
    };
    sc.addEventListener('wheel', onWheel, { passive: false });

    // はじめの表示：保存したときの場所
    layout();
    sc.scrollTop = Math.max(0, rowY[Math.min(nR - 1, Math.max(fr, sh.freeze.top))] - rowY[fr]);
    sc.scrollLeft = Math.max(0, colX[Math.min(nC - 1, Math.max(fc, sh.freeze.left))] - colX[fc]);
    update();
    buildDrawings();

    Object.assign(G, {
      el: wrap, scroller: sc, setZoom, select, goto, update,
      dispose() {
        G.dead = true;
        if (raf) cancelAnimationFrame(raf);
        ro.disconnect();
        for (const p of panes) { for (const t of p.tiles.values()) t.width = t.height = 0; p.tiles.clear(); }
        wrap.remove();
      },
    });
    return G;
  }

  // 見た目
  if (typeof document !== 'undefined' && !document.getElementById('xg-css')) {
    const st = document.createElement('style');
    st.id = 'xg-css';
    st.textContent = `
.xg-wrap{position:relative;overflow:hidden;background:#fff}
.xg{position:absolute;inset:0;overflow:auto;background:#fff;touch-action:pan-x pan-y;overscroll-behavior:contain;outline:none}
.xg-band{display:flex}
.xg-top{position:sticky;top:0;z-index:3}
.xg-pane{position:relative;overflow:hidden;flex:none;background:#fff}
.xg-tl,.xg-bl{position:sticky;left:0}
.xg-tl{z-index:4}.xg-bl{z-index:2}
.xg-layer{position:absolute;left:0;top:0;width:0;height:0}
.xg-pane canvas{position:absolute;display:block}
.xg-dw{position:absolute;transform-origin:0 0;pointer-events:none}
.xg-sel{position:absolute;border:2px solid #217346;box-sizing:border-box;pointer-events:none;box-shadow:0 0 0 1px rgba(255,255,255,.8)}
.xg-hl{position:absolute;background:rgba(33,115,70,.16);border-bottom:2px solid #217346;box-sizing:border-box;pointer-events:none}`;
    document.head.appendChild(st);
  }

  return { create, breakLines };
})();
