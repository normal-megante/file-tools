'use strict';
// =========================================================
// Office の図（DrawingML）を描く共通の部品（自作）。js/pptx.js・js/docx.js が使う
//   ・XML の小道具、色（テーマの色・明るさの調整）、塗り・線、図形の形（SVG）、画像、
//     図形の中の文字（a:p）、ZIP の中の部品の読み込み（DML.pkg）
//   ・const P = DML.pkg(zip) → { xml(path), rels(path), relOfType(m, type), media(path), recolor(...), theme(path), dispose() }
//   ・色を決めるときの ctx は { theme, map（色の名前の置きかえ）, ph（phClr）, src（画像の場所の基準 { path, rels }）, media, recolor }
// =========================================================
const DML = (() => {
  const EMU = 9525; // 1px（96 dpi）が何 EMU か
  const px = v => (+v || 0) / EMU;
  const PT = 4 / 3; // 1pt が何 px か

  /* ---------- XML の小道具（名前空間の接頭辞は気にせず、名前だけで探す） ---------- */
  function parseXml(s) {
    const d = new DOMParser().parseFromString(s, 'application/xml');
    return d.getElementsByTagName('parsererror').length ? null : d;
  }
  function kid(el, name) {
    if (el) for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (c.localName === name) return c;
    return null;
  }
  function kids(el, name) {
    const out = [];
    if (el) for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (!name || c.localName === name) out.push(c);
    return out;
  }
  const pathOf = (el, ...names) => { for (const n of names) { el = kid(el, n); if (!el) return null; } return el; };
  const at = (el, n) => (el ? el.getAttribute(n) : null);
  const atNum = (el, n, d) => { const v = at(el, n); return v == null || v === '' || isNaN(+v) ? d : +v; };
  const truthy = v => v === '1' || v === 'true' || v === 'on';
  // r:embed・r:id など（接頭辞のついた属性）
  function rAttr(el, local) {
    if (el) for (const a of el.attributes) if (a.localName === local && a.prefix) return a.value;
    return null;
  }
  const desc = (el, name) => (el ? el.getElementsByTagNameNS('*', name) : []);
  const h = (tag, cls, style) => { const e = document.createElement(tag); if (cls) e.className = cls; if (style) Object.assign(e.style, style); return e; };
  const SVGNS = 'http://www.w3.org/2000/svg';
  const sv = (tag, attrs) => { const e = document.createElementNS(SVGNS, tag); if (attrs) for (const k in attrs) if (attrs[k] != null) e.setAttribute(k, attrs[k]); return e; };
  let uidN = 0;
  const uid = () => 'pp' + (++uidN);

  /* ---------- ファイルの場所 ---------- */
  function resolvePath(base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/'); parts.pop();
    for (const seg of target.split('/')) { if (seg === '..') parts.pop(); else if (seg && seg !== '.') parts.push(seg); }
    return parts.join('/');
  }
  const relsPath = p => { const i = p.lastIndexOf('/'); return `${p.slice(0, i + 1)}_rels/${p.slice(i + 1)}.rels`; };

  /* ---------- 色 ---------- */
  const PRESET = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', cyan: '00FFFF', magenta: 'FF00FF', gray: '808080', grey: '808080', ltGray: 'D3D3D3', lightGray: 'D3D3D3', dkGray: 'A9A9A9', darkGray: 'A9A9A9', orange: 'FFA500', purple: '800080', navy: '000080', maroon: '800000', olive: '808000', teal: '008080', silver: 'C0C0C0', lime: '00FF00', brown: 'A52A2A', pink: 'FFC0CB', gold: 'FFD700', darkBlue: '00008B', darkRed: '8B0000', darkGreen: '006400', skyBlue: '87CEEB', violet: 'EE82EE' };
  const COLOR_TAGS = new Set(['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr']);
  const hex = v => { v = /^[0-9a-f]{6}$/i.test(v || '') ? v : '000000'; return { r: parseInt(v.slice(0, 2), 16), g: parseInt(v.slice(2, 4), 16), b: parseInt(v.slice(4, 6), 16), a: 1 }; };
  function rgb2hsl({ r, g, b }) {
    r /= 255; g /= 255; b /= 255;
    const mx = Math.max(r, g, b), mn = Math.min(r, g, b), l = (mx + mn) / 2;
    if (mx === mn) return [0, 0, l];
    const d = mx - mn, s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
    const hh = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4;
    return [hh / 6, s, l];
  }
  function hsl2rgb(hh, s, l) {
    if (!s) { const v = l * 255; return { r: v, g: v, b: v, a: 1 }; }
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const f = t => { t = ((t % 1) + 1) % 1; return 255 * (t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p); };
    return { r: f(hh + 1 / 3), g: f(hh), b: f(hh - 1 / 3), a: 1 };
  }
  const clamp = (v, a, b) => (v < a ? a : v > b ? b : v);
  const s2lin = v => { v = clamp(v, 0, 255) / 255; return v <= 0.04045 ? v / 12.92 : Math.pow((v + 0.055) / 1.055, 2.4); };
  const lin2s = v => { v = clamp(v, 0, 1); return 255 * (v <= 0.0031308 ? v * 12.92 : 1.055 * Math.pow(v, 1 / 2.4) - 0.055); };
  const tintColor = (c, v) => ({ r: lin2s(s2lin(c.r) * v + 1 - v), g: lin2s(s2lin(c.g) * v + 1 - v), b: lin2s(s2lin(c.b) * v + 1 - v), a: c.a });
  function colorEl(el) {
    if (!el) return null;
    if (COLOR_TAGS.has(el.localName)) return el;
    for (let c = el.firstElementChild; c; c = c.nextElementSibling) if (COLOR_TAGS.has(c.localName)) return c;
    return null;
  }
  // 色の要素（または色を中に持つ要素）→ { r, g, b, a }。ctx.ph は「phClr」（スタイルから渡される色）
  function color(el, ctx) {
    const c = colorEl(el);
    if (!c) return null;
    let base;
    switch (c.localName) {
      case 'srgbClr': base = hex(at(c, 'val')); break;
      case 'schemeClr': {
        let v = at(c, 'val');
        if (v === 'phClr') { base = ctx.ph ? { ...ctx.ph } : { r: 0, g: 0, b: 0, a: 1 }; break; }
        v = ctx.map[v] || v;
        base = hex(ctx.theme.colors[v] || (v === 'lt1' || v === 'bg1' ? 'FFFFFF' : '000000'));
        break;
      }
      case 'sysClr': base = hex(at(c, 'lastClr') || (at(c, 'val') === 'window' ? 'FFFFFF' : '000000')); break;
      case 'prstClr': base = hex(PRESET[at(c, 'val')] || '000000'); break;
      case 'scrgbClr': {
        const g = k => 255 * Math.pow(clamp(atNum(c, k, 0) / 100000, 0, 1), 1 / 2.2);
        base = { r: g('r'), g: g('g'), b: g('b'), a: 1 };
        break;
      }
      case 'hslClr': base = hsl2rgb(atNum(c, 'hue', 0) / 21600000, atNum(c, 'sat', 0) / 100000, atNum(c, 'lum', 0) / 100000); break;
    }
    return mods(base, c);
  }
  // 色の調整（明るさ・濃さ・透明度など）
  function mods(c, el) {
    let { r, g, b } = c, a = c.a == null ? 1 : c.a;
    for (const m of kids(el)) {
      const v = atNum(m, 'val', 0) / 100000, n = m.localName;
      if (n === 'alpha') a = v;
      else if (n === 'alphaMod') a *= v;
      else if (n === 'alphaOff') a += v;
      // tint（白に近づける）・shade（黒に近づける）は、PowerPoint と同じく光の量（リニア）で混ぜる
      else if (n === 'tint') { r = lin2s(s2lin(r) * v + 1 - v); g = lin2s(s2lin(g) * v + 1 - v); b = lin2s(s2lin(b) * v + 1 - v); }
      else if (n === 'shade') { r = lin2s(s2lin(r) * v); g = lin2s(s2lin(g) * v); b = lin2s(s2lin(b) * v); }
      else if (n === 'inv') { r = 255 - r; g = 255 - g; b = 255 - b; }
      else if (n === 'gray') { r = g = b = 0.3 * r + 0.59 * g + 0.11 * b; }
      else if (/^(lumMod|lumOff|lum|satMod|satOff|sat|hueMod|hueOff|hue|comp)$/.test(n)) {
        let [hh, s, l] = rgb2hsl({ r, g, b });
        if (n === 'lumMod') l *= v; else if (n === 'lumOff') l += v; else if (n === 'lum') l = v;
        else if (n === 'satMod') s *= v; else if (n === 'satOff') s += v; else if (n === 'sat') s = v;
        else if (n === 'hueMod') hh *= v; else if (n === 'hueOff') hh += atNum(m, 'val', 0) / 21600000; else if (n === 'hue') hh = atNum(m, 'val', 0) / 21600000;
        else if (n === 'comp') hh += 0.5;
        ({ r, g, b } = hsl2rgb(hh, clamp(s, 0, 1), clamp(l, 0, 1)));
      }
    }
    return { r: clamp(r, 0, 255), g: clamp(g, 0, 255), b: clamp(b, 0, 255), a: clamp(a, 0, 1) };
  }
  const css = c => (c ? `rgba(${Math.round(c.r)},${Math.round(c.g)},${Math.round(c.b)},${+c.a.toFixed(3)})` : 'transparent');
  const mix = (x, y, t) => (!x ? y : !y ? x : { r: x.r * (1 - t) + y.r * t, g: x.g * (1 - t) + y.g * t, b: x.b * (1 - t) + y.b * t, a: x.a * (1 - t) + y.a * t });

  /* ---------- 塗り・線 ---------- */
  const FILL_TAGS = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);
  // 塗りの要素 → { none } / { color } / { stops, ang, radial, color } / { blip } / null
  function fillEl(f, ctx) {
    switch (f && f.localName) {
      case 'noFill': return { none: true };
      case 'solidFill': return { color: color(f, ctx) || { r: 0, g: 0, b: 0, a: 1 } };
      case 'gradFill': {
        const stops = kids(kid(f, 'gsLst'), 'gs').map(g => ({ pos: atNum(g, 'pos', 0) / 100000, c: color(g, ctx) })).filter(s => s.c).sort((a, b) => a.pos - b.pos);
        if (!stops.length) return null;
        const lin = kid(f, 'lin');
        return { stops, ang: lin ? atNum(lin, 'ang', 0) / 60000 : 90, radial: !!kid(f, 'path'), color: stops[Math.floor(stops.length / 2)].c };
      }
      case 'blipFill': return { blip: f, src: ctx.src, ph: ctx.ph };
      case 'pattFill': return { color: mix(color(kid(f, 'fgClr'), ctx), color(kid(f, 'bgClr'), ctx), 0.5) || { r: 0, g: 0, b: 0, a: 1 } };
      case 'grpFill': return ctx.grpFill || null;
    }
    return null;
  }
  function fillOf(parent, ctx) {
    if (parent) for (let c = parent.firstElementChild; c; c = c.nextElementSibling) if (FILL_TAGS.has(c.localName)) return fillEl(c, ctx);
    return null;
  }
  // 線 → { none, color, w, dash, head, tail }（指定のない所は入れない）
  function lineOf(ln, ctx) {
    if (!ln) return null;
    const o = {};
    if (at(ln, 'w') != null) o.w = px(at(ln, 'w'));
    const f = fillOf(ln, ctx);
    if (f) { if (f.none) o.none = true; else { o.none = false; o.color = f.color || (f.stops && f.stops[0].c); } }
    const d = kid(ln, 'prstDash');
    if (d) o.dash = at(d, 'val');
    for (const k of ['headEnd', 'tailEnd']) {
      const e = kid(ln, k);
      if (e) o[k === 'headEnd' ? 'head' : 'tail'] = { type: at(e, 'type') || 'none', w: at(e, 'w') || 'med', len: at(e, 'len') || 'med' };
    }
    return o;
  }
  const mergeLine = (...ls) => ls.reduce((a, l) => (l ? { ...a, ...l } : a), {});

  /* ---------- 位置と大きさ ---------- */
  function xfrmOf(el) {
    const x = kid(el, 'xfrm');
    if (!x) return null;
    const off = kid(x, 'off'), ext = kid(x, 'ext');
    if (!off || !ext) return null;
    const o = { x: px(at(off, 'x')), y: px(at(off, 'y')), w: px(at(ext, 'cx')), h: px(at(ext, 'cy')), rot: atNum(x, 'rot', 0) / 60000, flipH: truthy(at(x, 'flipH')), flipV: truthy(at(x, 'flipV')) };
    const co = kid(x, 'chOff'), ce = kid(x, 'chExt');
    if (co && ce) Object.assign(o, { cx: px(at(co, 'x')), cy: px(at(co, 'y')), cw: px(at(ce, 'cx')), ch: px(at(ce, 'cy')) });
    return o;
  }
  function placeBox(xf, cls = 'pp-sp') {
    const d = h('div', cls, { left: xf.x + 'px', top: xf.y + 'px', width: Math.max(0, xf.w) + 'px', height: Math.max(0, xf.h) + 'px' });
    if (xf.rot) d.style.transform = `rotate(${xf.rot}deg)`;
    return d;
  }

  /* ---------- 図形の形（SVG の path。px） ---------- */
  function adjOf(av) {
    const o = {};
    for (const g of kids(av, 'gd')) { const m = /^val\s+(-?[\d.]+)/.exec(at(g, 'fmla') || ''); if (m) o[at(g, 'name')] = +m[1]; }
    return o;
  }
  const poly = pts => 'M' + pts.map(p => `${+p[0].toFixed(2)},${+p[1].toFixed(2)}`).join('L') + 'Z';
  const ellD = (cx, cy, rx, ry) => `M${cx - rx},${cy}A${rx},${ry} 0 1 0 ${cx + rx},${cy}A${rx},${ry} 0 1 0 ${cx - rx},${cy}Z`;
  function rrectD(w, hh, r) {
    r = Math.max(0, Math.min(r, w / 2, hh / 2));
    return `M${r},0H${w - r}A${r},${r} 0 0 1 ${w},${r}V${hh - r}A${r},${r} 0 0 1 ${w - r},${hh}H${r}A${r},${r} 0 0 1 0,${hh - r}V${r}A${r},${r} 0 0 1 ${r},0Z`;
  }
  function starD(n, w, hh, inner) {
    const pts = [];
    for (let i = 0; i < n * 2; i++) {
      const t = -Math.PI / 2 + i * Math.PI / n, k = i % 2 ? inner : 1;
      pts.push([w / 2 + Math.cos(t) * w / 2 * k, hh / 2 + Math.sin(t) * hh / 2 * k]);
    }
    return poly(pts);
  }
  // 決まった形（prstGeom）。[{ d, fill: 塗るか, stroke: 線を引くか, evenodd }]
  function preset(name, w, hh, av) {
    const ss = Math.min(w, hh), A = (k, d) => (av[k] != null ? av[k] : av.adj != null && k === 'adj1' ? av.adj : d) / 100000;
    const one = (d, o) => [{ d, fill: true, stroke: true, ...o }];
    const open = d => [{ d, fill: false, stroke: true }];
    switch (name) {
      case 'roundRect': case 'flowChartAlternateProcess': case 'round1Rect': case 'round2SameRect': case 'round2DiagRect': case 'snipRoundRect':
        return one(rrectD(w, hh, ss * A('adj', name === 'flowChartAlternateProcess' ? 16667 : 16667)));
      case 'flowChartTerminator': return one(rrectD(w, hh, ss / 2));
      case 'ellipse': case 'flowChartConnector': case 'circle': case 'cloud': case 'cloudCallout': case 'teardrop': case 'flowChartOr': case 'flowChartSummingJunction':
        return one(ellD(w / 2, hh / 2, w / 2, hh / 2));
      case 'triangle': case 'flowChartExtract': { const x = w * A('adj', 50000); return one(poly([[x, 0], [w, hh], [0, hh]])); }
      case 'flowChartMerge': return one(poly([[0, 0], [w, 0], [w / 2, hh]]));
      case 'rtTriangle': return one(poly([[0, 0], [w, hh], [0, hh]]));
      case 'diamond': case 'flowChartDecision': return one(poly([[w / 2, 0], [w, hh / 2], [w / 2, hh], [0, hh / 2]]));
      case 'parallelogram': case 'flowChartInputOutput': { const x = name === 'parallelogram' ? ss * A('adj', 25000) : w / 5; return one(poly([[x, 0], [w, 0], [w - x, hh], [0, hh]])); }
      case 'trapezoid': { const x = ss * A('adj', 25000); return one(poly([[x, 0], [w - x, 0], [w, hh], [0, hh]])); }
      case 'flowChartManualOperation': return one(poly([[0, 0], [w, 0], [w * 0.8, hh], [w * 0.2, hh]]));
      case 'flowChartManualInput': return one(poly([[0, hh * 0.2], [w, 0], [w, hh], [0, hh]]));
      case 'pentagon': return one(poly([[w / 2, 0], [w, hh * 0.38], [w * 0.81, hh], [w * 0.19, hh], [0, hh * 0.38]]));
      case 'hexagon': case 'flowChartPreparation': { const x = name === 'hexagon' ? ss * A('adj', 25000) : w * 0.2; return one(poly([[x, 0], [w - x, 0], [w, hh / 2], [w - x, hh], [x, hh], [0, hh / 2]])); }
      case 'octagon': { const x = ss * A('adj', 29289); return one(poly([[x, 0], [w - x, 0], [w, x], [w, hh - x], [w - x, hh], [x, hh], [0, hh - x], [0, x]])); }
      case 'homePlate': { const x = w - ss * A('adj', 50000); return one(poly([[0, 0], [x, 0], [w, hh / 2], [x, hh], [0, hh]])); }
      case 'chevron': { const x = ss * A('adj', 50000); return one(poly([[0, 0], [w - x, 0], [w, hh / 2], [w - x, hh], [0, hh], [x, hh / 2]])); }
      case 'rightArrow': case 'notchedRightArrow': case 'stripedRightArrow': case 'leftArrow': {
        const t = hh * A('adj1', 50000), hx = ss * A('adj2', 50000), y1 = (hh - t) / 2;
        let pts = [[0, y1], [w - hx, y1], [w - hx, 0], [w, hh / 2], [w - hx, hh], [w - hx, y1 + t], [0, y1 + t]];
        if (name === 'leftArrow') pts = pts.map(([x, y]) => [w - x, y]);
        return one(poly(pts));
      }
      case 'upArrow': case 'downArrow': {
        const t = w * A('adj1', 50000), hy = ss * A('adj2', 50000), x1 = (w - t) / 2;
        let pts = [[x1, hh], [x1, hy], [0, hy], [w / 2, 0], [w, hy], [x1 + t, hy], [x1 + t, hh]];
        if (name === 'downArrow') pts = pts.map(([x, y]) => [x, hh - y]);
        return one(poly(pts));
      }
      case 'leftRightArrow': {
        const t = hh * A('adj1', 50000), hx = ss * A('adj2', 50000), y1 = (hh - t) / 2;
        return one(poly([[0, hh / 2], [hx, 0], [hx, y1], [w - hx, y1], [w - hx, 0], [w, hh / 2], [w - hx, hh], [w - hx, y1 + t], [hx, y1 + t], [hx, hh]]));
      }
      case 'upDownArrow': {
        const t = w * A('adj1', 50000), hy = ss * A('adj2', 50000), x1 = (w - t) / 2;
        return one(poly([[w / 2, 0], [w, hy], [x1 + t, hy], [x1 + t, hh - hy], [w, hh - hy], [w / 2, hh], [0, hh - hy], [x1, hh - hy], [x1, hy], [0, hy]]));
      }
      case 'plus': case 'mathPlus': { const x = ss * A('adj', 25000); return one(poly([[x, 0], [w - x, 0], [w - x, x], [w, x], [w, hh - x], [w - x, hh - x], [w - x, hh], [x, hh], [x, hh - x], [0, hh - x], [0, x], [x, x]])); }
      case 'star4': return one(starD(4, w, hh, 0.38));
      case 'star5': return one(starD(5, w, hh, 0.38));
      case 'star6': return one(starD(6, w, hh, 0.58));
      case 'star7': case 'star8': return one(starD(8, w, hh, 0.7));
      case 'star10': case 'star12': case 'star16': case 'star24': case 'star32': case 'irregularSeal1': case 'irregularSeal2': return one(starD(12, w, hh, 0.75));
      case 'donut': { const t = ss * A('adj', 25000); return one(ellD(w / 2, hh / 2, w / 2, hh / 2) + ellD(w / 2, hh / 2, Math.max(0, w / 2 - t), Math.max(0, hh / 2 - t)), { evenodd: true }); }
      case 'frame': { const t = ss * A('adj1', 12500); return one(poly([[0, 0], [w, 0], [w, hh], [0, hh]]) + poly([[t, t], [w - t, t], [w - t, hh - t], [t, hh - t]]), { evenodd: true }); }
      case 'can': case 'flowChartMagneticDisk': {
        const ry = Math.min(hh / 2, ss * A('adj', 25000) / 2);
        return [{ d: `M0,${ry}A${w / 2},${ry} 0 0 1 ${w},${ry}V${hh - ry}A${w / 2},${ry} 0 0 1 0,${hh - ry}Z`, fill: true, stroke: true }, { d: `M0,${ry}A${w / 2},${ry} 0 0 0 ${w},${ry}`, fill: false, stroke: true }];
      }
      case 'cube': {
        const t = ss * A('adj', 25000);
        return [{ d: poly([[0, t], [w - t, t], [w - t, hh], [0, hh]]), fill: true, stroke: true }, { d: poly([[0, t], [t, 0], [w, 0], [w - t, t]]), fill: true, stroke: true, shade: 0.8 }, { d: poly([[w - t, t], [w, 0], [w, hh - t], [w - t, hh]]), fill: true, stroke: true, shade: 0.6 }];
      }
      case 'heart': return one(`M${w / 2},${hh / 4}C${w / 2},0 0,0 0,${hh / 4}C0,${hh / 2} ${w / 2},${hh * 0.75} ${w / 2},${hh}C${w / 2},${hh * 0.75} ${w},${hh / 2} ${w},${hh / 4}C${w},0 ${w / 2},0 ${w / 2},${hh / 4}Z`);
      case 'wave': case 'doubleWave': case 'flowChartPunchedTape':
        return one(`M0,${hh * 0.1}C${w / 3},${-hh * 0.1} ${w * 2 / 3},${hh * 0.3} ${w},${hh * 0.1}V${hh * 0.9}C${w * 2 / 3},${hh * 1.1} ${w / 3},${hh * 0.7} 0,${hh * 0.9}Z`);
      case 'flowChartDocument': case 'flowChartMultidocument':
        return one(`M0,0H${w}V${hh * 0.83}C${w * 0.7},${hh * 0.6} ${w * 0.3},${hh * 1.08} 0,${hh * 0.87}Z`);
      case 'flowChartDelay': return one(`M0,0H${w / 2}A${w / 2},${hh / 2} 0 0 1 ${w / 2},${hh}H0Z`);
      case 'flowChartOffpageConnector': return one(poly([[0, 0], [w, 0], [w, hh * 0.8], [w / 2, hh], [0, hh * 0.8]]));
      case 'flowChartPredefinedProcess': return [...one(poly([[0, 0], [w, 0], [w, hh], [0, hh]])), ...open(`M${w / 8},0V${hh}M${w * 7 / 8},0V${hh}`)];
      case 'snip1Rect': { const x = ss * A('adj', 16667); return one(poly([[0, 0], [w - x, 0], [w, x], [w, hh], [0, hh]])); }
      case 'snip2SameRect': { const x = ss * A('adj1', 16667); return one(poly([[x, 0], [w - x, 0], [w, x], [w, hh], [0, hh], [0, x]])); }
      case 'wedgeRectCallout': case 'wedgeRoundRectCallout': case 'wedgeEllipseCallout': {
        const tx = w / 2 + w * A('adj1', -20833), ty = hh / 2 + hh * A('adj2', 62500);
        const body = name === 'wedgeEllipseCallout' ? ellD(w / 2, hh / 2, w / 2, hh / 2) : name === 'wedgeRoundRectCallout' ? rrectD(w, hh, ss * 0.16667) : poly([[0, 0], [w, 0], [w, hh], [0, hh]]);
        const dx = tx - w / 2, dy = ty - hh / 2, len = Math.hypot(dx, dy) || 1, bw = Math.min(w, hh) * 0.12;
        const tail = poly([[w / 2 - dy / len * bw, hh / 2 + dx / len * bw], [tx, ty], [w / 2 + dy / len * bw, hh / 2 - dx / len * bw]]);
        return [{ d: tail, fill: true, stroke: true }, { d: body, fill: true, stroke: true }];
      }
      case 'leftBracket': return open(`M${w},0Q0,0 0,${hh * 0.1}V${hh * 0.9}Q0,${hh} ${w},${hh}`);
      case 'rightBracket': return open(`M0,0Q${w},0 ${w},${hh * 0.1}V${hh * 0.9}Q${w},${hh} 0,${hh}`);
      case 'leftBrace': return open(`M${w},0Q${w / 2},0 ${w / 2},${hh * 0.1}V${hh * 0.4}Q${w / 2},${hh / 2} 0,${hh / 2}Q${w / 2},${hh / 2} ${w / 2},${hh * 0.6}V${hh * 0.9}Q${w / 2},${hh} ${w},${hh}`);
      case 'rightBrace': return open(`M0,0Q${w / 2},0 ${w / 2},${hh * 0.1}V${hh * 0.4}Q${w / 2},${hh / 2} ${w},${hh / 2}Q${w / 2},${hh / 2} ${w / 2},${hh * 0.6}V${hh * 0.9}Q${w / 2},${hh} 0,${hh}`);
      case 'bracketPair': { const r = ss * 0.1667; return open(`M${r},0Q0,0 0,${r}V${hh - r}Q0,${hh} ${r},${hh}M${w - r},0Q${w},0 ${w},${r}V${hh - r}Q${w},${hh} ${w - r},${hh}`); }
      case 'arc': {
        const a1 = A('adj1', 16200000 / 1) * 100000 / 60000 * Math.PI / 180, a2 = A('adj2', 0) * 100000 / 60000 * Math.PI / 180;
        const rx = w / 2, ry = hh / 2, p = t => [rx + rx * Math.cos(t), ry + ry * Math.sin(t)];
        let sw = a2 - a1; if (sw <= 0) sw += 2 * Math.PI;
        const [x0, y0] = p(a1), [x1, y1] = p(a1 + sw);
        return open(`M${x0},${y0}A${rx},${ry} 0 ${sw > Math.PI ? 1 : 0} 1 ${x1},${y1}`);
      }
      case 'line': case 'straightConnector1': return open(`M0,0L${w},${hh}`);
      case 'bentConnector2': return open(`M0,0L${w},0L${w},${hh}`);
      case 'bentConnector3': case 'bentConnector4': case 'bentConnector5': { const x = w * A('adj1', 50000); return open(`M0,0L${x},0L${x},${hh}L${w},${hh}`); }
      case 'curvedConnector2': case 'curvedConnector3': case 'curvedConnector4': case 'curvedConnector5': { const x = w * A('adj1', 50000); return open(`M0,0C${x},0 ${x},${hh} ${w},${hh}`); }
      default: return one(poly([[0, 0], [w, 0], [w, hh], [0, hh]]));
    }
  }
  // 自由な形（custGeom）
  function custom(cg, w, hh, ew, eh) {
    const guides = adjOf(kid(cg, 'avLst'));
    Object.assign(guides, adjOf(kid(cg, 'gdLst')));
    const V = v => { if (v == null) return 0; const n = +v; return isNaN(n) ? (guides[v] || 0) : n; };
    const out = [];
    for (const p of kids(kid(cg, 'pathLst'), 'path')) {
      const pw = atNum(p, 'w', 0), ph = atNum(p, 'h', 0);
      const sx = pw ? w / pw : (ew ? w / ew : 1 / EMU), sy = ph ? hh / ph : (eh ? hh / eh : 1 / EMU);
      let d = '', cx = 0, cy = 0;
      for (const c of kids(p)) {
        const pts = kids(c, 'pt').map(q => [V(at(q, 'x')) * sx, V(at(q, 'y')) * sy]);
        const n = c.localName;
        if (n === 'moveTo' && pts[0]) { [cx, cy] = pts[0]; d += `M${cx},${cy}`; }
        else if (n === 'lnTo' && pts[0]) { [cx, cy] = pts[0]; d += `L${cx},${cy}`; }
        else if (n === 'cubicBezTo' && pts.length === 3) { d += 'C' + pts.map(q => q.join(',')).join(' '); [cx, cy] = pts[2]; }
        else if (n === 'quadBezTo' && pts.length === 2) { d += 'Q' + pts.map(q => q.join(',')).join(' '); [cx, cy] = pts[1]; }
        else if (n === 'arcTo') {
          const wr = V(at(c, 'wR')) * sx, hr = V(at(c, 'hR')) * sy, st = V(at(c, 'stAng')) / 60000 * Math.PI / 180, sw = V(at(c, 'swAng')) / 60000 * Math.PI / 180;
          const ox = cx - wr * Math.cos(st), oy = cy - hr * Math.sin(st);
          const arc = (s, e) => { const ex = ox + wr * Math.cos(e), ey = oy + hr * Math.sin(e); d += `A${wr},${hr} 0 ${Math.abs(e - s) > Math.PI ? 1 : 0} ${e > s ? 1 : 0} ${ex},${ey}`; cx = ex; cy = ey; };
          if (Math.abs(sw) >= 2 * Math.PI - 1e-6) { arc(st, st + sw / 2); arc(st + sw / 2, st + sw); } else if (sw) arc(st, st + sw);
        } else if (n === 'close') d += 'Z';
      }
      if (d) out.push({ d, fill: at(p, 'fill') !== 'none', stroke: !/^(false|0)$/.test(at(p, 'stroke') || '') });
    }
    return out;
  }
  function geometry(spPr, w, hh) {
    const pg = kid(spPr, 'prstGeom');
    if (pg) return preset(at(pg, 'prst'), w, hh, adjOf(kid(pg, 'avLst')));
    const cg = kid(spPr, 'custGeom');
    if (cg) return custom(cg, w, hh, w * EMU, hh * EMU);
    return null;
  }
  // 図形の中で文字を置く範囲（PowerPoint が形ごとに決めている「文字の四角」に近いもの）。決まりのない形は null（＝図形の四角いっぱい）
  function presetTextRect(name, w, hh, av) {
    const ss = Math.min(w, hh), A = (k, d) => (av[k] != null ? av[k] : av.adj != null && k === 'adj1' ? av.adj : d) / 100000;
    const R = (l, t, r, b) => ({ l, t, r, b });
    switch (name) {
      case 'ellipse': case 'flowChartConnector': case 'donut': case 'wedgeEllipseCallout': case 'flowChartOr': case 'flowChartSummingJunction': case 'circle':
        return R(w * 0.1464, hh * 0.1464, w * 0.8536, hh * 0.8536);
      case 'cloud': case 'cloudCallout': return R(w * 0.15, hh * 0.21, w * 0.85, hh * 0.79);
      case 'roundRect': { const d = ss * A('adj', 16667) * 0.29289; return R(d, d, w - d, hh - d); }
      case 'flowChartTerminator': return R(w * 1018 / 21600, hh * 3163 / 21600, w * 20582 / 21600, hh * 18437 / 21600);
      case 'triangle': case 'flowChartExtract': { const x = w * A('adj', 50000); return R(x / 2, hh / 2, x / 2 + w / 2, hh); }
      case 'flowChartMerge': return R(w / 4, 0, w * 3 / 4, hh / 2);
      case 'rtTriangle': return R(w / 12, hh * 7 / 12, w * 7 / 12, hh * 11 / 12);
      case 'diamond': case 'flowChartDecision': return R(w / 4, hh / 4, w * 3 / 4, hh * 3 / 4);
      case 'parallelogram': { const q = (1 + 5 * A('adj', 25000) * ss / w) / 12; return R(w * q, hh * q, w - w * q, hh - hh * q); }
      case 'flowChartInputOutput': return R(w / 5, 0, w * 4 / 5, hh);
      case 'trapezoid': { const x = ss * A('adj', 25000) * 2 / 3; return R(x, x * hh / w, w - x, hh); }
      case 'pentagon': return R(w * 0.19, hh * 0.31, w * 0.81, hh);
      case 'homePlate': { const x1 = w - ss * A('adj', 50000); return R(0, 0, (x1 + w) / 2, hh); }
      case 'chevron': { const x1 = ss * A('adj', 50000), x2 = w - x1; return x2 > x1 ? R(x1, 0, x2, hh) : null; }
      case 'hexagon': { const x = ss * A('adj', 25000) * 0.75; return R(x, 0, w - x, hh); }
      case 'octagon': { const x = ss * A('adj', 29289) / 2; return R(x, x, w - x, hh - x); }
      case 'plus': case 'mathPlus': { const x = ss * A('adj', 25000); return R(0, x, w, hh - x); }
      case 'frame': { const t = ss * A('adj1', 12500); return R(t, t, w - t, hh - t); }
      case 'star4': return R(w * 0.32, hh * 0.32, w * 0.68, hh * 0.68);
      case 'star5': return R(w * 0.31, hh * 0.38, w * 0.69, hh * 0.73);
      case 'star6': case 'star7': case 'star8': return R(w * 0.2, hh * 0.25, w * 0.8, hh * 0.75);
      case 'star10': case 'star12': case 'star16': case 'star24': case 'star32': case 'irregularSeal1': case 'irregularSeal2': return R(w * 0.18, hh * 0.18, w * 0.82, hh * 0.82);
      case 'heart': return R(w * 0.15, hh * 0.2, w * 0.85, hh * 0.7);
      case 'can': case 'flowChartMagneticDisk': { const ry = Math.min(hh / 2, ss * A('adj', 25000) / 2); return R(0, ry * 2, w, hh - ry); }
      case 'cube': { const t = ss * A('adj', 25000); return R(0, t, w - t, hh); }
      case 'rightArrow': case 'leftArrow': case 'notchedRightArrow': case 'stripedRightArrow': {
        const th = hh * A('adj1', 50000), hx = ss * A('adj2', 50000), y1 = (hh - th) / 2, x2 = w - hx + y1 * hx / (hh / 2);
        return name === 'leftArrow' ? R(w - x2, y1, w, y1 + th) : R(0, y1, x2, y1 + th);
      }
      case 'upArrow': case 'downArrow': {
        const tw = w * A('adj1', 50000), hy = ss * A('adj2', 50000), x1 = (w - tw) / 2, y2 = hh - hy + x1 * hy / (w / 2);
        return name === 'downArrow' ? R(x1, 0, x1 + tw, y2) : R(x1, hh - y2, x1 + tw, hh);
      }
      case 'leftRightArrow': {
        const th = hh * A('adj1', 50000), hx = ss * A('adj2', 50000), y1 = (hh - th) / 2, d = y1 * hx / (hh / 2);
        return R(hx - d, y1, w - hx + d, y1 + th);
      }
      case 'wave': case 'doubleWave': case 'flowChartPunchedTape': return R(0, hh * 0.2, w, hh * 0.8);
      case 'flowChartDocument': case 'flowChartMultidocument': return R(0, 0, w, hh * 0.8);
      case 'flowChartPredefinedProcess': return R(w / 8, 0, w * 7 / 8, hh);
      case 'flowChartDelay': return R(0, hh * 0.146, w * 0.854, hh * 0.854);
      case 'flowChartOffpageConnector': return R(0, 0, w, hh * 0.8);
    }
    return null;
  }
  function textRect(spPr, xf) {
    const pg = kid(spPr, 'prstGeom');
    const r = pg && presetTextRect(at(pg, 'prst'), xf.w, xf.h, adjOf(kid(pg, 'avLst')));
    if (!r) return null;
    let { l, t, r: rr, b } = r;
    if (xf.flipH) [l, rr] = [xf.w - rr, xf.w - l];
    if (xf.flipV) [t, b] = [xf.h - b, xf.h - t];
    return { x: l, y: t, w: Math.max(0, rr - l), h: Math.max(0, b - t) };
  }

  /* ---------- 文字の書式を重ねる（あとから来たものが勝つ） ---------- */
  const EMPTY_R = { a: {}, ch: {} }, EMPTY_P = { a: {}, ch: {}, r: EMPTY_R };
  function mergeRPr(acc, el) {
    if (!el) return acc;
    const o = { a: { ...acc.a }, ch: { ...acc.ch } };
    for (const x of el.attributes) o.a[x.name] = x.value;
    for (const c of kids(el)) {
      const n = c.localName;
      if (FILL_TAGS.has(n)) o.ch.fill = c; else if (n !== 'extLst') o.ch[n] = c;
    }
    return o;
  }
  function mergePPr(acc, el) {
    if (!el) return acc;
    const o = { a: { ...acc.a }, ch: { ...acc.ch }, r: mergeRPr(acc.r, kid(el, 'defRPr')) };
    for (const x of el.attributes) o.a[x.name] = x.value;
    for (const c of kids(el)) {
      const n = c.localName;
      if (n === 'defRPr' || n === 'extLst') continue;
      if (/^bu(None|Char|AutoNum|Blip)$/.test(n)) o.ch.bu = c;
      else if (/^buFont/.test(n)) o.ch.buFont = c;
      else if (/^buClr/.test(n)) o.ch.buClr = c;
      else if (/^buSz/.test(n)) o.ch.buSz = c;
      else o.ch[n] = c;
    }
    return o;
  }
  // chain：低い順に、文字のスタイルの一覧（lstStyle など）か { font, color }（図形のスタイルの文字の色・書体）
  function levelProps(chain, lvl) {
    let pp = EMPTY_P;
    for (const c of chain) {
      if (!c) continue;
      if (c.nodeType) { pp = mergePPr(pp, kid(c, 'defPPr')); pp = mergePPr(pp, kid(c, `lvl${lvl + 1}pPr`)); }
      else {
        const r = { a: { ...pp.r.a }, ch: { ...pp.r.ch } };
        if (c.color) r.ch.fill = c.color;
        if (c.rgb) r.ch.fill = { localName: '__rgb', rgb: c.rgb }; // 計算ずみの色（表のスタイル）
        if (c.font) r.a.__font = c.font;
        if (c.bold != null) r.a.b = c.bold ? '1' : '0';
        pp = { ...pp, r };
      }
    }
    return pp;
  }
  const bodyPrChain = list => list.reduce((acc, b) => {
    if (!b) return acc;
    const o = { a: { ...acc.a }, fit: acc.fit };
    for (const x of b.attributes) o.a[x.name] = x.value;
    for (const n of ['normAutofit', 'spAutoFit', 'noAutofit']) { const e = kid(b, n); if (e) o.fit = e; }
    return o;
  }, { a: {}, fit: null });

  /* ---------- 書体 ---------- */
  const SERIF = /明朝|mincho|times|century|georgia|garamond|cambria|book antiqua|serif/i;
  function fontFamily(r, theme) {
    const face = v => {
      if (!v) return '';
      if (v[0] !== '+') return v;
      const f = v.startsWith('+mj') ? theme.major : theme.minor;
      return /-ea$/.test(v) ? f.ea : f.latin;
    };
    const t = r.a.__font === 'major' ? theme.major : theme.minor;
    const latin = face(at(r.ch.latin, 'typeface')) || t.latin, ea = face(at(r.ch.ea, 'typeface')) || t.ea;
    const list = [...new Set([latin, ea].filter(Boolean))].map(f => `"${f.replace(/["\\]/g, '')}"`);
    list.push(SERIF.test(latin + ea) && !/sans/i.test(latin + ea) ? 'serif' : 'sans-serif');
    return list.join(',');
  }
  // Wingdings などの記号を、ふつうの文字に置きかえる（スマホには Wingdings がない）
  const WINGDINGS = { '§': '■', 'Ø': '➢', 'ü': '✔', 'q': '❑', 'n': '■', 'l': '●', 'v': '❖', 'Ü': '➤', 'è': '➔', 'à': '➔', 'ð': '⇨', 'û': '✘', 'o': '□', 'p': '□', 'w': '◆', 'u': '◆', 'ê': '⇒', 'Ÿ': '•', '': '•' };
  function bulletChar(ch, fontEl) {
    const f = at(fontEl, 'typeface') || '';
    if (/wingdings|symbol|webdings/i.test(f)) return WINGDINGS[ch] || (ch === '·' ? '•' : '•');
    return ch;
  }
  function roman(n) {
    const t = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']];
    let s = ''; for (const [v, c] of t) while (n >= v) { s += c; n -= v; } return s;
  }
  function autoNum(type, n) {
    const m = /^(arabic|romanUc|romanLc|alphaUc|alphaLc|circleNumDb|circleNumWdBlack|circleNumWdWhite|ea1JpnChsDb|ea1JpnKor|ea1ChsPeriod|ea1Chs|hebrew2|thaiAlpha)?(.*)$/.exec(type || 'arabicPeriod');
    const base = m[1] || 'arabic', suf = m[2] || '';
    let s;
    if (base === 'romanUc') s = roman(n); else if (base === 'romanLc') s = roman(n).toLowerCase();
    else if (base === 'alphaUc' || base === 'alphaLc') { let k = n, t = ''; while (k > 0) { k--; t = String.fromCharCode(65 + (k % 26)) + t; k = Math.floor(k / 26); } s = base === 'alphaLc' ? t.toLowerCase() : t; }
    else if (base === 'circleNumDb' || base === 'circleNumWdWhite') s = n <= 20 ? String.fromCodePoint(0x2460 + n - 1) : `(${n})`;
    else if (base === 'circleNumWdBlack') s = n <= 10 ? String.fromCodePoint(0x2776 + n - 1) : `(${n})`;
    else if (base.startsWith('ea1')) s = String(n).replace(/\d/g, d => String.fromCharCode(0xFF10 + +d));
    else s = String(n);
    if (/ParenBoth/.test(suf)) return `(${s})`;
    if (/ParenR/.test(suf)) return s + ')';
    if (/Period/.test(suf)) return s + (base.startsWith('ea1') ? '．' : '.');
    if (/Minus/.test(suf)) return s + '-';
    return s;
  }

  /* ---------- 文字を描く ---------- */
  // txBody → 文字の箱。chain：段落の書式の重なり、bodies：bodyPr の重なり（低い順）、box：{ x, y, w, h }（px）
  function drawText(txBody, chain, bodies, ctx, box) {
    const paras = kids(txBody, 'p');
    if (!paras.some(p => desc(p, 't').length && [...desc(p, 't')].some(t => t.textContent) || desc(p, 'fld').length)) return null;
    const bp = bodyPrChain([...bodies, kid(txBody, 'bodyPr')]);
    const ins = (k, d) => px(bp.a[k] != null ? bp.a[k] : d);
    const L = ins('lIns', 91440), T = ins('tIns', 45720), R = ins('rIns', 91440), B = ins('bIns', 45720);
    const vert = bp.a.vert && bp.a.vert !== 'horz';
    const anchor = bp.a.anchor || 't';
    const d = h('div', 'pp-tx', { left: box.x + L + 'px', top: box.y + T + 'px', width: Math.max(0, box.w - L - R) + 'px', height: Math.max(0, box.h - T - B) + 'px' });
    d.style.justifyContent = anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start';
    if (vert) d.style.writingMode = 'vertical-rl';
    if (bp.a.wrap === 'none') d.style.whiteSpace = 'pre';
    const cols = +bp.a.numCol || 1;
    let fontScale = 1, lnRed = 0;
    if (bp.fit && bp.fit.localName === 'normAutofit') { fontScale = atNum(bp.fit, 'fontScale', 100000) / 100000; lnRed = atNum(bp.fit, 'lnSpcReduction', 0) / 100000; }
    const inner = cols > 1 ? h('div', null, { columnCount: cols, columnGap: px(bp.a.spcCol || 0) + 'px' }) : d;
    if (inner !== d) d.appendChild(inner);
    const counters = [];
    for (const p of paras) {
      const pPr = kid(p, 'pPr'), lvl = clamp(atNum(pPr, 'lvl', 0), 0, 8);
      const pp = mergePPr(levelProps(chain, lvl), pPr);
      inner.appendChild(drawPara(p, pp, ctx, { fontScale, lnRed, counters, lvl }));
    }
    return d;
  }
  function spacing(el, size, isLine) {
    if (!el) return null;
    const pct = kid(el, 'spcPct'), pts = kid(el, 'spcPts');
    if (pct) return isLine ? { mul: atNum(pct, 'val', 100000) / 100000 } : { px: atNum(pct, 'val', 0) / 100000 * size * 1.2 };
    if (pts) return { px: atNum(pts, 'val', 0) / 100 * PT };
    return null;
  }
  function runCss(r, ctx, scale) {
    const a = r.a, s = {};
    let size = (a.sz ? +a.sz : 1800) / 100 * scale;
    const base = +a.baseline || 0;
    if (base) { size *= 0.7; s.verticalAlign = base > 0 ? 'super' : 'sub'; }
    s.fontSize = size * PT + 'px';
    if (truthy(a.b)) s.fontWeight = 'bold';
    if (truthy(a.i)) s.fontStyle = 'italic';
    const deco = [];
    if (a.u && a.u !== 'none') deco.push('underline');
    if (a.strike && a.strike !== 'noStrike') deco.push('line-through');
    if (deco.length) s.textDecoration = deco.join(' ');
    if (a.cap === 'all') s.textTransform = 'uppercase'; else if (a.cap === 'small') s.fontVariant = 'small-caps';
    if (a.spc) s.letterSpacing = (+a.spc / 100) * PT * scale + 'px';
    const f = r.ch.fill;
    if (f) {
      if (f.localName === '__rgb') s.color = css(f.rgb);
      else if (f.localName === 'noFill') s.color = 'transparent';
      else { const c = f.localName === 'gradFill' ? (fillEl(f, ctx) || {}).color : color(f, ctx); if (c) s.color = css(c); }
    }
    if (r.ch.highlight) { const c = color(r.ch.highlight, ctx); if (c) s.backgroundColor = css(c); }
    // 文字のふち（ワードアートなど）
    if (r.ch.ln) {
      const l = lineOf(r.ch.ln, ctx);
      if (l && !l.none && l.color) s.webkitTextStroke = `${Math.max(0.5, Math.min(l.w || 0.75, size * PT * 0.08))}px ${css(l.color)}`;
    }
    s.fontFamily = fontFamily(r, ctx.theme);
    return { s, size };
  }
  function drawPara(p, pp, ctx, st) {
    const d = h('div', 'pp-p');
    const algn = pp.a.algn;
    d.style.textAlign = algn === 'ctr' ? 'center' : algn === 'r' ? 'right' : algn === 'just' || algn === 'dist' ? 'justify' : 'left';
    const marL = px(pp.a.marL || 0), indent = px(pp.a.indent || 0);
    if (marL) d.style.paddingLeft = marL + 'px';
    const items = [];
    let first = null;
    for (const c of kids(p)) {
      if (c.localName === 'r' || c.localName === 'fld') {
        const r = mergeRPr(pp.r, kid(c, 'rPr'));
        let text = (kid(c, 't') || {}).textContent || '';
        if (c.localName === 'fld' && /^slidenum$/i.test(at(c, 'type') || '')) text = String(ctx.slideNo);
        if (!text) continue;
        const { s, size } = runCss(r, ctx, st.fontScale);
        if (!first) first = { r, size };
        items.push({ text, s });
      } else if (c.localName === 'br') {
        const r = mergeRPr(pp.r, kid(c, 'rPr')), { s } = runCss(r, ctx, st.fontScale);
        items.push({ text: '\n', s });
      }
    }
    const end = mergeRPr(pp.r, kid(p, 'endParaRPr'));
    const endCss = runCss(end, ctx, st.fontScale);
    const size = first ? first.size : endCss.size;
    d.style.fontSize = size * PT + 'px';
    d.style.fontFamily = (first ? runCss(first.r, ctx, st.fontScale).s : endCss.s).fontFamily;
    const ln = spacing(pp.ch.lnSpc, size, true);
    if (ln && ln.mul) d.style.lineHeight = String(Math.max(0.5, ln.mul - st.lnRed) * 1.2);
    else if (ln && ln.px) d.style.lineHeight = ln.px + 'px';
    else if (st.lnRed) d.style.lineHeight = String((1 - st.lnRed) * 1.2);
    const sb = spacing(pp.ch.spcBef, size), sa = spacing(pp.ch.spcAft, size);
    if (sb) d.style.paddingTop = sb.px + 'px';
    if (sa) d.style.paddingBottom = sa.px + 'px';
    // 箇条書きの記号・番号
    const bu = pp.ch.bu, lvl = st.lvl, cnt = st.counters;
    for (let k = lvl + 1; k < cnt.length; k++) cnt[k] = 0;
    if (items.length && bu && (bu.localName === 'buChar' || bu.localName === 'buAutoNum' || bu.localName === 'buBlip')) {
      let txt;
      if (bu.localName === 'buAutoNum') { cnt[lvl] = (cnt[lvl] || atNum(bu, 'startAt', 1) - 1) + 1; txt = autoNum(at(bu, 'type'), cnt[lvl]); }
      else { cnt[lvl] = 0; txt = bu.localName === 'buChar' ? bulletChar(at(bu, 'char') || '•', pp.ch.buFont) : '•'; }
      const fr = first ? runCss(first.r, ctx, st.fontScale).s : endCss.s;
      const b = h('span', 'pp-bu', { color: fr.color || '', fontSize: fr.fontSize, fontWeight: fr.fontWeight || '' });
      if (pp.ch.buClr && pp.ch.buClr.localName === 'buClr') { const c = color(pp.ch.buClr, ctx); if (c) b.style.color = css(c); }
      const sz = pp.ch.buSz;
      if (sz && sz.localName === 'buSzPct') b.style.fontSize = size * PT * atNum(sz, 'val', 100000) / 100000 + 'px';
      else if (sz && sz.localName === 'buSzPts') b.style.fontSize = atNum(sz, 'val', 1800) / 100 * PT + 'px';
      const bf = at(pp.ch.buFont, 'typeface');
      if (bf && !/wingdings|symbol|webdings/i.test(bf) && bu.localName === 'buChar') b.style.fontFamily = `"${bf.replace(/["\\]/g, '')}",${fr.fontFamily}`;
      b.textContent = txt;
      if (indent < 0) { d.style.textIndent = indent + 'px'; b.style.width = -indent + 'px'; }
      else { if (indent) d.style.textIndent = indent + 'px'; b.style.paddingRight = '0.4em'; }
      d.appendChild(b);
    } else {
      if (items.length) cnt[lvl] = 0;
      if (indent) d.style.textIndent = indent + 'px';
    }
    if (!items.length) { const e = h('span', null, endCss.s); e.textContent = '\u200b'; d.appendChild(e); return d; }
    for (const it of items) { const e = h('span', null, it.s); e.textContent = it.text; d.appendChild(e); }
    return d;
  }

  /* ---------- 図形を SVG で描く ---------- */
  const DASH = { dash: [4, 3], sysDash: [3, 1], dot: [1, 1], sysDot: [1, 1], dashDot: [4, 3, 1, 3], sysDashDot: [3, 1, 1, 1], sysDashDotDot: [3, 1, 1, 1, 1, 1], lgDash: [8, 3], lgDashDot: [8, 3, 1, 3], lgDashDotDot: [8, 3, 1, 3, 1, 3] };
  async function shapeSvg(paths, fill, line, w, hh, flip, ctx) {
    const s = sv('svg', { width: Math.max(1, w), height: Math.max(1, hh), viewBox: `0 0 ${Math.max(1, w)} ${Math.max(1, hh)}` });
    s.style.overflow = 'visible';
    if (flip.h || flip.v) s.style.transform = `scale(${flip.h ? -1 : 1},${flip.v ? -1 : 1})`;
    const defs = sv('defs');
    s.appendChild(defs);
    let paint = 'none';
    if (fill && !fill.none) {
      if (fill.stops) {
        const id = uid();
        const g = fill.radial ? sv('radialGradient', { id, cx: '50%', cy: '50%', r: '70%' })
          : (() => { const t = fill.ang * Math.PI / 180, cx = Math.cos(t) / 2, cy = Math.sin(t) / 2; return sv('linearGradient', { id, x1: 0.5 - cx, y1: 0.5 - cy, x2: 0.5 + cx, y2: 0.5 + cy }); })();
        for (const st of fill.stops) g.appendChild(sv('stop', { offset: st.pos, 'stop-color': css({ ...st.c, a: 1 }), 'stop-opacity': st.c.a }));
        defs.appendChild(g);
        paint = `url(#${id})`;
      } else if (fill.blip) {
        const m = await blipImage(fill.blip, fill.src || ctx.src, ctx, fill.ph);
        if (m && m.url) {
          const id = uid(), clip = sv('clipPath', { id });
          for (const p of paths) if (p.fill) clip.appendChild(sv('path', { d: p.d }));
          defs.appendChild(clip);
          const img = sv('image', { href: m.url, x: 0, y: 0, width: w, height: hh, preserveAspectRatio: 'none', 'clip-path': `url(#${id})`, opacity: m.opacity });
          s.appendChild(img);
        }
      } else if (fill.color) paint = css(fill.color);
    }
    const stroke = line && !line.none && line.color ? css(line.color) : 'none';
    const sw = line && line.w != null ? Math.max(line.w, 0.75) : 1;
    let mk = {};
    if (stroke !== 'none') for (const end of ['head', 'tail']) {
      const e = line[end];
      if (!e || e.type === 'none') continue;
      const id = uid(), k = { sm: 2, med: 3, lg: 5 }, mw = k[e.w] || 3, ml = k[e.len] || 3;
      const m = sv('marker', { id, viewBox: '0 0 10 10', refX: e.type === 'oval' || e.type === 'diamond' ? 5 : 9, refY: 5, markerWidth: ml, markerHeight: mw, orient: end === 'head' ? 'auto-start-reverse' : 'auto', markerUnits: 'strokeWidth' });
      const shape = e.type === 'oval' ? sv('circle', { cx: 5, cy: 5, r: 4.5 }) : e.type === 'diamond' ? sv('path', { d: 'M5,0L10,5L5,10L0,5Z' })
        : e.type === 'arrow' ? sv('path', { d: 'M0,0L10,5L0,10', fill: 'none', stroke, 'stroke-width': 1.5 }) : sv('path', { d: 'M0,0L10,5L0,10Z' });
      if (e.type !== 'arrow') shape.setAttribute('fill', stroke);
      m.appendChild(shape); defs.appendChild(m);
      mk[end === 'head' ? 'marker-start' : 'marker-end'] = `url(#${id})`;
    }
    for (const p of paths) {
      let pf = p.fill ? paint : 'none';
      if (p.fill && p.shade && fill && fill.color) pf = css({ ...fill.color, r: fill.color.r * p.shade, g: fill.color.g * p.shade, b: fill.color.b * p.shade });
      const e = sv('path', { d: p.d, fill: pf, stroke: p.stroke ? stroke : 'none', 'stroke-width': sw, 'fill-rule': p.evenodd ? 'evenodd' : null, 'stroke-linejoin': 'round' });
      if (line && line.dash && DASH[line.dash] && stroke !== 'none') e.setAttribute('stroke-dasharray', DASH[line.dash].map(v => v * sw).join(' '));
      if (!p.fill) for (const k in mk) e.setAttribute(k, mk[k]);
      s.appendChild(e);
    }
    return s;
  }
  // 画像 → { url, opacity } など。ph：テーマの背景の画像を色づけするときの色
  async function blipImage(blipFill, src, ctx, ph) {
    const blip = kid(blipFill, 'blip'), id = rAttr(blip, 'embed') || rAttr(blip, 'link');
    const rel = id && src && src.rels.get(id);
    if (!rel || rel.external) return null;
    const m = await ctx.media(rel.target);
    if (!m || !m.url) return m;
    const out = { ...m }, amf = kid(blip, 'alphaModFix');
    if (amf) out.opacity = atNum(amf, 'amt', 100000) / 100000;
    // 2色に色づけ（duotone）・白黒（grayscl）。テーマの模様の背景によく使われる
    const duo = kid(blip, 'duotone'), gray = kid(blip, 'grayscl');
    if (duo || gray) {
      const cs = duo ? kids(duo).filter(c => COLOR_TAGS.has(c.localName)).map(c => color(c, { ...ctx, ph: ph || ctx.ph })) : null;
      out.url = await ctx.recolor(`${rel.target}|${cs ? cs.map(css).join('|') : 'gray'}`, m.url, cs && cs.length >= 2 ? cs : null);
    }
    return out;
  }


  /* ---------- ZIP の中の部品を読む（XML・つながり・画像・テーマ）。読んだものは覚えておく ---------- */
  function pkg(zip) {
    const xmlCache = new Map(), relCache = new Map(), mediaCache = new Map(), urls = [];
    const xml = path => {
      if (!xmlCache.has(path)) xmlCache.set(path, (async () => { const e = path && zip.get(path); return e ? parseXml(await zip.text(e)) : null; })());
      return xmlCache.get(path);
    };
    const rels = path => {
      if (!relCache.has(path)) relCache.set(path, (async () => {
        const m = new Map(), d = await xml(relsPath(path));
        if (d) for (const r of kids(d.documentElement, 'Relationship')) {
          const ext = at(r, 'TargetMode') === 'External';
          let t = at(r, 'Target') || '';
          try { t = decodeURIComponent(t); } catch (e) { /* そのまま */ }
          m.set(at(r, 'Id'), { type: (at(r, 'Type') || '').split('/').pop(), target: ext ? t : resolvePath(path, t), external: ext });
        }
        return m;
      })());
      return relCache.get(path);
    };
    const relOfType = (m, type) => { for (const r of m.values()) if (r.type === type) return r; return null; };
    // 画像の色づけ（明るさに合わせて2色の間の色にする。cs が null なら白黒）。結果はファイルの中の画像ごとに1回だけ作る
    const recolorCache = new Map();
    const recolor = (key, url, cs) => {
      if (!recolorCache.has(key)) recolorCache.set(key, (async () => {
        const img = new Image();
        await new Promise((res, rej) => { img.onload = res; img.onerror = rej; setTimeout(() => rej(new Error('timeout')), 8000); img.src = url; });
        const k = Math.min(1, 1600 / Math.max(img.naturalWidth, img.naturalHeight));
        const w = Math.max(1, Math.round(img.naturalWidth * k)), hh = Math.max(1, Math.round(img.naturalHeight * k));
        const c = document.createElement('canvas');
        c.width = w; c.height = hh;
        const x = c.getContext('2d', { willReadFrequently: true });
        x.drawImage(img, 0, 0, w, hh);
        const d = x.getImageData(0, 0, w, hh), p = d.data;
        for (let i = 0; i < p.length; i += 4) {
          const l = (0.299 * p[i] + 0.587 * p[i + 1] + 0.114 * p[i + 2]) / 255;
          if (cs) { p[i] = cs[0].r + (cs[1].r - cs[0].r) * l; p[i + 1] = cs[0].g + (cs[1].g - cs[0].g) * l; p[i + 2] = cs[0].b + (cs[1].b - cs[0].b) * l; }
          else p[i] = p[i + 1] = p[i + 2] = l * 255;
        }
        x.putImageData(d, 0, 0);
        const out = c.toDataURL(/\.jpe?g\|/i.test(key) ? 'image/jpeg' : 'image/png', 0.9);
        c.width = c.height = 0;
        return out;
      })().catch(() => url));
      return recolorCache.get(key);
    };
    const MIME = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', jfif: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp', svg: 'image/svg+xml', avif: 'image/avif' };
    const media = path => {
      if (!mediaCache.has(path)) mediaCache.set(path, (async () => {
        const e = zip.get(path);
        if (!e) return null;
        const ext = (path.split('.').pop() || '').toLowerCase();
        if (/^(emf|wmf|emz|wmz|tiff?|wdp|jxr|pict?)$/.test(ext)) return { unsupported: ext.toUpperCase() };
        const url = URL.createObjectURL(await zip.blob(e, MIME[ext] || ''));
        urls.push(url);
        return { url };
      })());
      return mediaCache.get(path);
    };

    const themeCache = new Map();
    function theme(path) {
      if (!themeCache.has(path)) themeCache.set(path, (async () => {
        const doc = await xml(path), T = { colors: {}, major: { latin: '', ea: '' }, minor: { latin: '', ea: '' }, fills: [], lines: [], bgFills: [] };
        T.src = { path, rels: await rels(path) }; // テーマの中の画像（背景の模様など）の場所の基準
        const te = doc && pathOf(doc.documentElement, 'themeElements');
        if (!te) return T;
        for (const c of kids(kid(te, 'clrScheme'))) {
          const s = kid(c, 'srgbClr'), y = kid(c, 'sysClr');
          T.colors[c.localName] = s ? at(s, 'val') : y ? (at(y, 'lastClr') || (at(y, 'val') === 'window' ? 'FFFFFF' : '000000')) : null;
        }
        for (const k of ['major', 'minor']) {
          const f = pathOf(te, 'fontScheme', k + 'Font'), jp = kids(f, 'font').find(x => at(x, 'script') === 'Jpan');
          T[k] = { latin: at(kid(f, 'latin'), 'typeface') || '', ea: at(kid(f, 'ea'), 'typeface') || (jp ? at(jp, 'typeface') : '') };
        }
        const fs = kid(te, 'fmtScheme');
        T.fills = kids(kid(fs, 'fillStyleLst')); T.lines = kids(kid(fs, 'lnStyleLst')); T.bgFills = kids(kid(fs, 'bgFillStyleLst'));
        return T;
      })());
      return themeCache.get(path);
    }
    return {
      zip, xml, rels, relOfType, recolor, media, theme,
      dispose() { for (const u of urls) URL.revokeObjectURL(u); urls.length = 0; mediaCache.clear(); recolorCache.clear(); },
    };
  }

  // 見た目（図形・文字の箱）
  if (typeof document !== 'undefined' && !document.getElementById('dml-css')) {
    const st = document.createElement('style');
    st.id = 'dml-css';
    st.textContent = `
.pp-sp,.pp-grp{position:absolute}
.pp-sp>svg{position:absolute;left:0;top:0}
.pp-tx{position:absolute;display:flex;flex-direction:column;white-space:pre-wrap;overflow-wrap:break-word;line-break:strict}
.pp-p{margin:0;min-height:0}
.pp-bu{display:inline-block;text-indent:0;white-space:pre}
.pp-missing{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;text-align:center;padding:6px;border:2px dashed #b0b6c0;background:rgba(240,242,245,.85);color:#5b6474;font:14px/1.4 sans-serif;white-space:normal}`;
    document.head.appendChild(st);
  }

  return { EMU, px, PT, parseXml, kid, kids, pathOf, at, atNum, truthy, rAttr, desc, h, SVGNS, sv, uid, resolvePath,
    relsPath, PRESET, COLOR_TAGS, hex, rgb2hsl, hsl2rgb, clamp, s2lin, lin2s, tintColor, colorEl, color, mods, css, mix,
    FILL_TAGS, fillEl, fillOf, lineOf, mergeLine, xfrmOf, placeBox, adjOf, poly, ellD, rrectD, starD, preset, custom,
    geometry, presetTextRect, textRect, EMPTY_R, EMPTY_P, mergeRPr, mergePPr, levelProps, bodyPrChain, SERIF,
    fontFamily, WINGDINGS, bulletChar, roman, autoNum, drawText, spacing, runCss, drawPara, DASH, shapeSvg, blipImage,
    pkg,
  };
})();
