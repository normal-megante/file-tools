'use strict';
// =========================================================
// Office のグラフ（c:chartSpace）を SVG で描く共通の部品（自作）
//   ・Excel（js/xlgrid.js とグラフのシート）、PowerPoint（js/pptx.js）、Word（js/docx.js）が使う
//   ・OChart.render(path, { P, theme, map, w, h, host }) → Promise<SVGSVGElement>
//       P：DML.pkg（xml を使う）／ path：グラフの XML の場所（xl/charts/chart1.xml など）
//       theme・map：色（js/drawingml.js と同じ）／ w・h：グラフの大きさ（px、96 dpi）
//       host：'xl'（Excel）・'pp'（PowerPoint）・'dx'（Word）。何も指定がないときの見た目が少し違う
//   ・数は、ファイルに保存されている値（numCache・strCache）を使う。計算し直さない
//   ・縦棒・横棒・折れ線・面・円・ドーナツ・散布図・バブル・レーダー・株価（積み上げ・100%・2 軸・組み合わせも）
//     3D のグラフは平らに描く。等高線・新しい種類（じょうご・ツリーマップなど）は描けない
//   ・js/drawingml.js を先に読む（表示形式は js/numfmt.js があれば使う）
// =========================================================
const OChart = (() => {
  const { kid, kids, pathOf, at, atNum, truthy, sv, PT } = DML;
  const val = (el, n, d) => { const e = kid(el, n), a = e && at(e, 'val'); return a == null ? d : a; };
  const num = (el, n, d) => { const a = val(el, n, null); return a == null || a === '' || isNaN(+a) ? d : +a; };
  // c: の真偽の要素は、val がなければ「はい」
  const bool = (el, n, d) => { const e = kid(el, n); if (!e) return d; const a = at(e, 'val'); return a == null ? true : truthy(a); };
  const clamp = (v, a, b) => (v < a ? a : v > b ? b : v);
  const r2 = v => Math.round(v * 100) / 100;
  const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

  /* ---------- 保存されている値 ---------- */
  function cacheOf(el) {
    if (el) for (const c of kids(el)) {
      const n = c.localName;
      if (n === 'numRef') { const k = kid(c, 'numCache'); return k ? { num: true, c: k } : null; }
      if (n === 'strRef') { const k = kid(c, 'strCache'); return k ? { num: false, c: k } : null; }
      if (n === 'numLit') return { num: true, c };
      if (n === 'strLit') return { num: false, c };
      if (n === 'multiLvlStrRef') { const k = kid(c, 'multiLvlStrCache'); return k ? { num: false, c: k, multi: true } : null; }
      if (n === 'v') return { text: c.textContent };
    }
    return null;
  }
  function ptMap(c) {
    const m = new Map();
    for (const p of kids(c, 'pt')) { const v = kid(p, 'v'); m.set(atNum(p, 'idx', 0), { v: v ? v.textContent : '', fmt: at(p, 'formatCode') }); }
    return m;
  }
  // → { n, num, fmt, m（番号 → { v, fmt }）, levels（何段もの項目名） }
  function dataOf(el) {
    const k = cacheOf(el);
    if (!k || k.text != null) return null;
    const fc = kid(k.c, 'formatCode');
    let n = num(k.c, 'ptCount', 0);
    if (k.multi) {
      const levels = kids(k.c, 'lvl').map(ptMap);
      return { n, num: false, fmt: null, m: levels[0] || new Map(), levels };
    }
    const m = ptMap(k.c);
    for (const i of m.keys()) if (i + 1 > n) n = i + 1;
    return { n, num: k.num, fmt: fc ? fc.textContent : null, m };
  }
  const numAt = (d, i) => { const p = d && d.m.get(i); if (!p || p.v === '') return null; const v = +p.v; return isFinite(v) ? v : null; };
  function nameOf(tx) {
    const k = cacheOf(tx);
    if (!k) return null;
    if (k.text != null) return k.text;
    const p = ptMap(k.c).get(0);
    return p ? p.v : '';
  }

  /* ---------- 読む ---------- */
  const TYPES = { barChart: 'bar', bar3DChart: 'bar', lineChart: 'line', line3DChart: 'line', areaChart: 'area', area3DChart: 'area', pieChart: 'pie', pie3DChart: 'pie', ofPieChart: 'pie', doughnutChart: 'doughnut', scatterChart: 'scatter', bubbleChart: 'bubble', radarChart: 'radar', stockChart: 'stock', surfaceChart: 'surface', surface3DChart: 'surface' };
  function readSeries(s) {
    const o = { el: s, idx: num(s, 'idx', 0), order: num(s, 'order', 0), name: nameOf(kid(s, 'tx')), spPr: kid(s, 'spPr'), marker: kid(s, 'marker'), dLbls: kid(s, 'dLbls'),
      smooth: bool(s, 'smooth', false), explosion: num(s, 'explosion', 0), trend: kids(s, 'trendline'), dPt: new Map() };
    for (const p of kids(s, 'dPt')) o.dPt.set(num(p, 'idx', 0), p);
    o.cat = dataOf(kid(s, 'cat') || kid(s, 'xVal'));
    o.val = dataOf(kid(s, 'val') || kid(s, 'yVal'));
    o.size = dataOf(kid(s, 'bubbleSize'));
    o.v = [];
    for (let i = 0; i < (o.val ? o.val.n : 0); i++) o.v.push(numAt(o.val, i));
    if (o.name == null) o.name = '系列' + (o.idx + 1);
    return o;
  }
  function readGroup(el) {
    const type = TYPES[el.localName];
    const g = { el, type, name: el.localName, dir: val(el, 'barDir', 'col'), grouping: val(el, 'grouping', type === 'bar' ? 'clustered' : 'standard'),
      vary: bool(el, 'varyColors', true), gap: num(el, 'gapWidth', 150), overlap: num(el, 'overlap', null), hole: num(el, 'holeSize', 50), first: num(el, 'firstSliceAng', 0),
      marker: bool(el, 'marker', true), radarStyle: val(el, 'radarStyle', 'marker'), bubbleScale: num(el, 'bubbleScale', 100), sizeRep: val(el, 'sizeRepresents', 'area'),
      showNeg: bool(el, 'showNegBubbles', false), dLbls: kid(el, 'dLbls'), axIds: kids(el, 'axId').map(a => at(a, 'val')),
      hiLow: kid(el, 'hiLowLines'), upDown: kid(el, 'upDownBars'), drop: kid(el, 'dropLines') };
    g.stacked = /stacked/i.test(g.grouping);
    g.pct = g.grouping === 'percentStacked';
    g.series = kids(el, 'ser').map(readSeries).sort((a, b) => a.order - b.order);
    return g;
  }
  function readAxis(el) {
    const sc = kid(el, 'scaling'), nf = kid(el, 'numFmt');
    const tick = n => { const e = kid(el, n); return e ? at(e, 'val') || 'cross' : n === 'majorTickMark' ? 'out' : 'none'; };
    return { el, kind: el.localName.replace('Ax', ''), id: val(el, 'axId'), del: bool(el, 'delete', false), pos: val(el, 'axPos', 'b'),
      rev: val(sc, 'orientation', 'minMax') === 'maxMin', min: num(sc, 'min', null), max: num(sc, 'max', null), log: num(sc, 'logBase', null),
      cross: val(el, 'crossAx'), crosses: val(el, 'crosses', kid(el, 'crossesAt') ? null : 'autoZero'), crossesAt: num(el, 'crossesAt', null),
      between: val(el, 'crossBetween', null), major: num(el, 'majorUnit', null), minor: num(el, 'minorUnit', null),
      grid: kid(el, 'majorGridlines'), mgrid: kid(el, 'minorGridlines'), title: kid(el, 'title'),
      fmt: nf ? at(nf, 'formatCode') : null, linked: nf ? truthy(at(nf, 'sourceLinked')) : true,
      majorTick: tick('majorTickMark'), minorTick: tick('minorTickMark'), lblPos: val(el, 'tickLblPos', 'nextTo'),
      spPr: kid(el, 'spPr'), txPr: kid(el, 'txPr'), skip: num(el, 'tickLblSkip', null), markSkip: num(el, 'tickMarkSkip', null),
      noMulti: bool(el, 'noMultiLvlLbl', false), dispUnits: kid(el, 'dispUnits'), baseUnit: val(el, 'baseTimeUnit', null) };
  }

  /* ---------- 文字の大きさを測る ---------- */
  let MC = null;
  const fontCss = s => `${s.i ? 'italic ' : ''}${s.b ? 'bold ' : ''}${r2(s.sz * PT)}px ${s.fam}`;
  function tw(s, t) {
    if (!MC) MC = document.createElement('canvas').getContext('2d');
    MC.font = fontCss(s);
    return MC.measureText(t).width;
  }
  // 1 行の高さ：文字の大きさ × 1.2。游ゴシック・メイリオの日本語は × 1.65（Excel で測った）
  const CJK = /[　-鿿＀-￯]/;
  const lineH = (s, t) => s.sz * PT * (s.tall && t && CJK.test(t) ? 1.65 : 1.2);
  const pt = v => v * PT;

  // 折れ線などの点の印
  const MARKS = ['diamond', 'square', 'triangle', 'x', 'star', 'circle', 'plus', 'dot', 'dash'];
  function markD(sym, x, y, r) {
    const f = v => r2(v);
    switch (sym) {
      case 'circle': return `M${f(x - r)},${f(y)}a${f(r)},${f(r)} 0 1 0 ${f(2 * r)},0a${f(r)},${f(r)} 0 1 0 ${f(-2 * r)},0Z`;
      case 'dot': { const q = r / 2; return `M${f(x - q)},${f(y)}a${f(q)},${f(q)} 0 1 0 ${f(2 * q)},0a${f(q)},${f(q)} 0 1 0 ${f(-2 * q)},0Z`; }
      case 'square': return `M${f(x - r)},${f(y - r)}h${f(2 * r)}v${f(2 * r)}h${f(-2 * r)}Z`;
      case 'diamond': return `M${f(x)},${f(y - r)}L${f(x + r)},${f(y)}L${f(x)},${f(y + r)}L${f(x - r)},${f(y)}Z`;
      case 'triangle': return `M${f(x)},${f(y - r)}L${f(x + r)},${f(y + r)}L${f(x - r)},${f(y + r)}Z`;
      case 'dash': return `M${f(x - r)},${f(y - r / 4)}h${f(2 * r)}v${f(r / 2)}h${f(-2 * r)}Z`;
      case 'x': return `M${f(x - r)},${f(y - r)}L${f(x + r)},${f(y + r)}M${f(x + r)},${f(y - r)}L${f(x - r)},${f(y + r)}`;
      case 'plus': return `M${f(x - r)},${f(y)}H${f(x + r)}M${f(x)},${f(y - r)}V${f(y + r)}`;
      case 'star': return `M${f(x - r)},${f(y - r)}L${f(x + r)},${f(y + r)}M${f(x + r)},${f(y - r)}L${f(x - r)},${f(y + r)}M${f(x)},${f(y - r)}V${f(y + r)}`;
    }
    return '';
  }
  const OPEN_MARK = new Set(['x', 'plus', 'star']);

  // 目盛りの間かくの候補（1・2・5 × 10 のべき。from 以上）
  function niceSteps(from) {
    const out = [], e = Math.floor(Math.log10(from));
    for (let k = e - 1; k <= e + 2; k++) for (const m of [1, 2, 5]) { const s = m * Math.pow(10, k); if (s >= from * 0.999) out.push(+s.toPrecision(12)); }
    return out.sort((a, b) => a - b);
  }
  // 目盛りの数を、ずれのない数にそろえる（0.1 + 0.2 など）
  const tidy = v => +(+v).toPrecision(12);

  const DISP = { hundreds: [100, '百'], thousands: [1e3, '千'], tenThousands: [1e4, '万'], hundredThousands: [1e5, '十万'], millions: [1e6, '百万'], tenMillions: [1e7, '千万'], hundredMillions: [1e8, '億'], billions: [1e9, '十億'], trillions: [1e12, '兆'] };
  const TREND = { linear: '線形', exp: '指数', log: '対数', poly: '多項式', power: '累乗', movingAvg: '移動平均' };

  /* ---------- 描く ---------- */
  async function render(path, o) {
    const doc = await o.P.xml(path);
    const root = doc && doc.documentElement;
    if (!root || root.localName !== 'chartSpace') throw new Error('グラフを読めませんでした');
    return draw(root, o);
  }

  let UID = 0;
  function draw(root, o) {
    const W = Math.max(1, o.w), H = Math.max(1, o.h), host = o.host || 'xl';
    const chart = kid(root, 'chart'), plot = kid(chart, 'plotArea');
    const map = { ...(o.map || {}) };
    const ovr = pathOf(root, 'clrMapOvr');
    if (ovr) for (const a of ovr.attributes) map[a.localName] = a.value;
    const cx = { theme: o.theme, map, ph: null, src: { path: '', rels: new Map() } };
    const T = o.theme;
    const date1904 = bool(root, 'date1904', false);
    const style = (() => {
      const alt = [...root.getElementsByTagNameNS('*', 'style')].find(e => e.parentNode && e.parentNode.localName === 'Choice');
      return alt ? atNum(alt, 'val', 102) - 100 : num(root, 'style', 2);
    })();

    const svg = sv('svg', { viewBox: `0 0 ${r2(W)} ${r2(H)}`, width: '100%', height: '100%', preserveAspectRatio: 'none', class: 'och' });
    svg.style.overflow = 'hidden';
    const defs = sv('defs');
    svg.appendChild(defs);
    const uid = () => 'och' + (++UID);
    const add = (parent, tag, attrs) => { const e = sv(tag, attrs); parent.appendChild(e); return e; };

    /* ----- 色 ----- */
    const colorXml = new Map();
    function schemeColor(acc, mods) {
      const key = acc + JSON.stringify(mods);
      if (!colorXml.has(key)) colorXml.set(key, DML.parseXml(`<a:schemeClr xmlns:a="${A_NS}" val="${acc}">${mods.map(([n, v]) => `<a:${n} val="${v}"/>`).join('')}</a:schemeClr>`).documentElement);
      return DML.color(colorXml.get(key), cx);
    }
    const VARS = [[], [['lumMod', 60000]], [['lumMod', 80000], ['lumOff', 20000]], [['lumMod', 80000]], [['lumMod', 60000], ['lumOff', 40000]], [['lumMod', 50000]], [['lumMod', 70000], ['lumOff', 30000]], [['lumMod', 70000]], [['lumMod', 50000], ['lumOff', 50000]]];
    // 何も指定がない系列の色（Excel の「色 1」：アクセント 1〜6、7 本目からは明るさを変えて）
    function autoColor(i, n) {
      const j = ((style - 1) % 8) + 1;
      if (j >= 3 && j <= 8) {
        // 単色のスタイル：アクセントの濃い色から薄い色へ
        const t = n > 1 ? i / (n - 1) : 0.5;
        const v = Math.round(t < 0.5 ? (0.5 + t) * 100000 : (1.5 - t) * 100000);
        return schemeColor('accent' + (j - 2), t < 0.5 ? [['shade', v]] : [['tint', v]]);
      }
      if (j === 1) return schemeColor('tx1', [['lumMod', 50000 + ((i * 15000) % 45000)], ['lumOff', 10000 + ((i * 15000) % 45000)]]);
      return schemeColor('accent' + ((i % 6) + 1), VARS[Math.floor(i / 6) % VARS.length]);
    }
    const css = c => DML.css(c);
    const tx1 = schemeColor('tx1', []);
    const GRAY = { r: 134, g: 134, b: 134, a: 1 };

    /* ----- 塗り・線 ----- */
    function gradient(f) {
      const id = uid();
      let g;
      if (f.radial) g = add(defs, 'radialGradient', { id, cx: '50%', cy: '50%', r: '70%' });
      else {
        const a = (f.ang || 0) * Math.PI / 180, c = Math.cos(a) / 2, s = Math.sin(a) / 2;
        g = add(defs, 'linearGradient', { id, x1: r2(0.5 - c), y1: r2(0.5 - s), x2: r2(0.5 + c), y2: r2(0.5 + s) });
      }
      for (const st of f.stops) add(g, 'stop', { offset: r2(st.pos), 'stop-color': css({ ...st.c, a: 1 }), 'stop-opacity': r2(st.c.a) });
      return id;
    }
    function paint(el, fill, line) {
      if (!fill || fill.none) el.setAttribute('fill', 'none');
      else if (fill.stops) el.setAttribute('fill', `url(#${gradient(fill)})`);
      else el.setAttribute('fill', css(fill.color || GRAY));
      if (!line || line.none || !line.color) { el.setAttribute('stroke', 'none'); return el; }
      const w = line.w != null ? line.w : 1;
      el.setAttribute('stroke', css(line.color));
      el.setAttribute('stroke-width', r2(Math.max(w, 0.3)));
      el.setAttribute('stroke-linejoin', line.join || 'round');
      el.setAttribute('stroke-linecap', line.cap || 'butt');
      const d = line.dash && DML.DASH[line.dash];
      if (d) el.setAttribute('stroke-dasharray', d.map(x => r2(x * Math.max(w, 1))).join(' '));
      return el;
    }
    function lineFrom(spPr, base) {
      const ln = kid(spPr, 'ln');
      const l = DML.lineOf(ln, cx);
      if (!l) return base;
      const o2 = { ...(base || {}), ...l };
      const cap = at(ln, 'cap');
      if (cap) o2.cap = cap === 'rnd' ? 'round' : cap === 'sq' ? 'square' : 'butt';
      if (kid(ln, 'miter')) o2.join = 'miter'; else if (kid(ln, 'bevel')) o2.join = 'bevel'; else if (kid(ln, 'round')) o2.join = 'round';
      return o2;
    }
    const fillFrom = (spPr, base) => { const f = DML.fillOf(spPr, cx); return f && f.blip ? { color: GRAY } : f || base; };

    /* ----- 文字 ----- */
    const face = v => (!v ? '' : v[0] !== '+' ? v : (/^\+mj/.test(v) ? T.major : T.minor)[/-ea$/.test(v) ? 'ea' : 'latin']);
    // 書体の並び。幅のせまい書体（Aptos Narrow など）がない端末では、ほかのせまい書体で代わりに
    function famOf(s) {
      const latin = face(s.latin) || T.minor.latin, ea = face(s.ea) || T.minor.ea;
      const list = [];
      for (const f of [latin, ea]) if (f) { list.push(f); if (/narrow|condensed/i.test(f)) list.push('Arial Narrow', 'sans-serif-condensed'); }
      const out = [...new Set(list)].map(f => `"${f.replace(/["\\]/g, '')}"`);
      out.push(DML.SERIF.test(latin + ea) && !/sans/i.test(latin + ea) ? 'serif' : 'sans-serif');
      s.tall = /游|yu ?(gothic|mincho)|メイリオ|meiryo/i.test(ea);
      return out.join(',');
    }
    function rpr(s, el) {
      if (!el) return s;
      s = { ...s };
      if (at(el, 'sz') != null) { s.sz = atNum(el, 'sz', 1000) / 100; s.szSet = true; }
      if (at(el, 'b') != null) s.b = truthy(at(el, 'b'));
      if (at(el, 'i') != null) s.i = truthy(at(el, 'i'));
      if (at(el, 'u') != null) s.u = at(el, 'u') !== 'none';
      const f = DML.fillOf(el, cx);
      if (f && f.color) s.color = f.color; else if (f && f.stops) s.color = f.stops[0].c;
      const la = kid(el, 'latin'), ea = kid(el, 'ea');
      if (la && at(la, 'typeface')) s.latin = at(la, 'typeface');
      if (ea && at(ea, 'typeface')) s.ea = at(ea, 'typeface');
      s.fam = famOf(s);
      return s;
    }
    // c:txPr・c:rich → 文字の書式（rot は 1/60000 度。-60000000 は「自動」）
    function txs(s, tx) {
      s = { ...s, rot: null, vert: null };
      if (!tx) return s;
      s = rpr(s, pathOf(kid(tx, 'lstStyle'), 'lvl1pPr', 'defRPr'));
      s = rpr(s, pathOf(kid(tx, 'p'), 'pPr', 'defRPr'));
      const bp = kid(tx, 'bodyPr');
      if (bp && at(bp, 'rot') != null) s.rot = atNum(bp, 'rot', 0);
      if (bp && at(bp, 'vert')) s.vert = at(bp, 'vert');
      return s;
    }
    let base = { sz: 10, b: false, i: false, u: false, color: tx1, latin: '', ea: '', szSet: false };
    base.fam = famOf(base);
    base = txs(base, kid(root, 'txPr'));

    // 文字のかたまり：lines = [[{ t, s }]] → { lines, w, h }
    function block(lines) {
      let w = 0, h = 0;
      const out = lines.map(runs => {
        let lw = 0, lh = 0;
        for (const r of runs) { r.w = tw(r.s, r.t); lw += r.w; lh = Math.max(lh, lineH(r.s, r.t)); }
        if (!runs.length) lh = 0;
        w = Math.max(w, lw); h += lh;
        return { runs, w: lw, h: lh };
      });
      return { lines: out, w, h };
    }
    const plain = (t, s) => block(String(t).split('\n').map(x => [{ t: x, s }]));
    // c:rich（題名など、書式の違う文字が混ざるもの）
    function richBlock(rich, s0) {
      const s1 = txs(s0, rich);
      const lines = [];
      for (const p of kids(rich, 'p')) {
        const ps = rpr(s1, pathOf(p, 'pPr', 'defRPr'));
        let runs = [];
        for (const r of kids(p)) {
          if (r.localName === 'r' || r.localName === 'fld') { const t = kid(r, 't'); runs.push({ t: t ? t.textContent : '', s: rpr(ps, kid(r, 'rPr')) }); }
          else if (r.localName === 'br') { lines.push(runs); runs = []; }
        }
        lines.push(runs.length ? runs : [{ t: '', s: ps }]);
      }
      return { ...block(lines), s: s1 };
    }
    // かたまりを (x, y) を中心に置く。align：行の寄せ、rot：度（時計回り）
    function drawBlock(parent, B, x, y, align = 'middle', rot = 0) {
      const g = add(parent, 'g');
      if (rot) g.setAttribute('transform', `rotate(${r2(rot)} ${r2(x)} ${r2(y)})`);
      let top = y - B.h / 2;
      for (const ln of B.lines) {
        const lx = align === 'start' ? x - B.w / 2 : align === 'end' ? x + B.w / 2 : x;
        const t = add(g, 'text', { x: r2(lx), y: r2(top + ln.h / 2), 'text-anchor': align, 'dominant-baseline': 'central' });
        t.style.whiteSpace = 'pre';
        for (const r of ln.runs) {
          const ts = add(t, 'tspan');
          ts.textContent = r.t;
          ts.setAttribute('font-family', r.s.fam);
          ts.setAttribute('font-size', r2(r.s.sz * PT));
          if (r.s.b) ts.setAttribute('font-weight', 'bold');
          if (r.s.i) ts.setAttribute('font-style', 'italic');
          if (r.s.u) ts.setAttribute('text-decoration', 'underline');
          ts.setAttribute('fill', css(r.s.color || tx1));
        }
        top += ln.h;
      }
      return g;
    }
    // 文字を幅に合わせて折り返す（英語は単語の切れ目、日本語はどこでも）
    function wrapText(t, s, maxW) {
      const out = [];
      for (const para of String(t).split('\n')) {
        const toks = para.match(/[\u3000-\u9fff\uff00-\uffef]|[^\s\u3000-\u9fff\uff00-\uffef]+|\s+/g) || [''];
        let cur = '';
        for (const k of toks) {
          const nx = cur + k;
          if (cur && !/^\s+$/.test(k) && tw(s, nx) > maxW) { out.push(cur.trimEnd()); cur = k.trimStart(); }
          else cur = nx;
        }
        out.push(cur.trimEnd());
      }
      return out;
    }

    /* ----- 形式 ----- */
    function fmtNum(v, code) {
      if (v == null) return '';
      if (typeof NumFmt === 'undefined') return String(tidy(v));
      // グラフに保存された「短い日付」は英語の形。日本の Excel と同じ形で出す
      if (code === 'm/d/yyyy' || code === 'm/d/yy') code = 'yyyy/m/d';
      try { return NumFmt.format(typeof v === 'number' ? tidy(v) : v, code || 'General', { date1904 }).text; } catch (e) { return String(v); }
    }

    /* ----- 読み取り ----- */
    const groups = kids(plot).filter(e => TYPES[e.localName]).map(readGroup);
    const axes = new Map();
    for (const el of kids(plot)) if (/^(catAx|valAx|dateAx|serAx)$/.test(el.localName)) { const a = readAxis(el); axes.set(a.id, a); }
    const allSeries = groups.flatMap(g => g.series);
    const nSeries = allSeries.length;

    // 系列・点の見た目
    const LINE_TYPES = new Set(['line', 'scatter', 'stock']);
    const isLineG = g => LINE_TYPES.has(g.type) || (g.type === 'radar' && g.radarStyle !== 'filled');
    const varyPts = g => g.vary && (g.type === 'pie' || g.type === 'doughnut' || (nSeries === 1 && /^(bar|bubble)$/.test(g.type)));
    function look(g, s, i) {
      const ci = varyPts(g) && i >= 0 ? i : s.idx;
      const nc = varyPts(g) ? Math.max(1, ...g.series.map(x => x.v.length)) : nSeries;
      const auto = autoColor(ci, nc);
      const lineT = isLineG(g);
      let fill = lineT ? { none: true } : { color: auto };
      let line = lineT ? { color: auto, w: 28575 / 9525, cap: 'round' } : (g.type === 'pie' || g.type === 'doughnut') && host === 'xl' && style === 2 ? null : null;
      fill = fillFrom(s.spPr, fill); line = lineFrom(s.spPr, line);
      const dp = i >= 0 && s.dPt.get(i);
      if (dp) { const sp = kid(dp, 'spPr'); fill = fillFrom(sp, fill); line = lineFrom(sp, line); }
      // 印
      let mk = null;
      if (g.type === 'line' || g.type === 'scatter' || g.type === 'radar' || g.type === 'stock') {
        const me = (dp && kid(dp, 'marker')) || s.marker;
        let sym = val(me, 'symbol', null) || val(s.marker, 'symbol', null);
        if (!sym) sym = (g.type === 'line' && !g.marker) || (g.type === 'radar' && g.radarStyle !== 'marker') ? 'none' : 'auto';
        if (sym === 'auto') sym = MARKS[s.idx % MARKS.length];
        if (sym !== 'none' && sym !== 'picture') {
          const size = num(me, 'size', num(s.marker, 'size', 7)); // 大きさの指定がなければ 7（PowerPoint で確かめた）
          const msp = kid(me, 'spPr') || kid(s.marker, 'spPr');
          mk = { sym, r: (size * PT) / 2, fill: fillFrom(msp, { color: auto }), line: lineFrom(msp, { color: auto, w: 1 }) };
          if (OPEN_MARK.has(sym)) mk.fill = { none: true };
        }
      }
      return { fill, line, mk, auto };
    }

    /* ----- 題名 ----- */
    // t がない（自動の題名）ときは autoText だけ
    function titleBlock(t, s0, autoText) {
      if (!t && autoText == null) return null;
      const rich = t && pathOf(t, 'tx', 'rich');
      let B;
      if (rich) B = richBlock(rich, s0);
      else {
        const s = txs(s0 === titleBase ? autoBase : s0, t && kid(t, 'txPr'));
        const ref = t ? nameOf(kid(t, 'tx')) : null;
        B = { ...plain(ref != null ? ref : autoText, s), s };
      }
      B.el = t;
      B.overlay = t ? bool(t, 'overlay', false) : false;
      B.spPr = t && kid(t, 'spPr');
      B.layout = t && pathOf(t, 'layout', 'manualLayout');
      return B;
    }
    // 題名の文字：書いた題名は、グラフ全体の大きさが決めてあればそれと同じ。自動の題名（系列の名前）は 1.2 倍の太字
    //   どちらも、何も決めていなければ 18pt の太字（PowerPoint で確かめた）
    const titleBase = base.szSet ? { ...base } : { ...base, sz: 18, b: true };
    const autoBase = { ...base, sz: base.szSet ? base.sz * 1.2 : 18, b: true };
    const titleEl = kid(chart, 'title');
    const autoDel = bool(chart, 'autoTitleDeleted', false);
    // 題名の要素がなくても「自動の題名を消した」でなければ、系列が 1 つのときはその名前を題名にする
    const TB = titleEl ? titleBlock(titleEl, titleBase, nSeries === 1 ? allSeries[0].name : 'グラフ タイトル') : !autoDel && nSeries === 1 ? titleBlock(null, titleBase, allSeries[0].name) : null;

    /* ----- 凡例 ----- */
    const legEl = kid(chart, 'legend');
    const legend = (() => {
      if (!legEl) return null;
      const pos = val(legEl, 'legendPos', 'r');
      const s = txs(base, kid(legEl, 'txPr'));
      const items = [];
      const g1 = groups.find(g => varyPts(g));
      if (g1) {
        const s1 = g1.series[0];
        const n = Math.max(s1.v.length, s1.cat ? s1.cat.n : 0);
        for (let i = 0; i < n; i++) items.push({ text: catText(s1, i, null), g: g1, s: s1, i, kind: 'rect' });
      } else {
        for (const g of groups) {
          if (g.type === 'surface') continue;
          for (const ser of g.series) {
            items.push({ text: ser.name, g, s: ser, i: -1, kind: isLineG(g) ? 'line' : 'rect' });
            for (const tr of ser.trend) {
              const tt = val(tr, 'trendlineType', 'linear');
              const nm = kid(tr, 'name');
              items.push({ text: nm ? nm.textContent : `${tt === 'movingAvg' ? num(tr, 'period', 2) + ' 区間' : ''}${TREND[tt] || ''} (${ser.name})`, g, s: ser, i: -1, kind: 'trend', tr });
            }
          }
        }
      }
      const del = new Set(), ts = new Map();
      for (const e of kids(legEl, 'legendEntry')) { const i = num(e, 'idx', -1); if (bool(e, 'delete', false)) del.add(i); if (kid(e, 'txPr')) ts.set(i, txs(s, kid(e, 'txPr'))); }
      const list = items.map((it, i) => ({ ...it, st: ts.get(i) || s })).filter((it, i) => !del.has(i));
      if (!list.length) return null;
      return { pos, s, items: list, overlay: bool(legEl, 'overlay', false), spPr: kid(legEl, 'spPr'), layout: pathOf(legEl, 'layout', 'manualLayout') };
    })();

    // 項目名
    function catText(s, i, ax) {
      const c = s.cat;
      if (!c) return String(i + 1);
      const p = c.m.get(i);
      if (!p) return '';
      if (c.num) { const v = +p.v; const code = ax && !ax.linked && ax.fmt ? ax.fmt : p.fmt || c.fmt || (ax && ax.fmt) || 'General'; return isFinite(v) && p.v !== '' ? fmtNum(v, code) : p.v; }
      return p.v;
    }

    /* ----- 場所を決める（数は Excel の画面で測った pt） ----- */
    const ML = el => {
      if (!el) return null;
      const g = n => { const e = kid(el, n); return e ? +at(e, 'val') : null; };
      return { target: val(el, 'layoutTarget', 'outer'), xMode: val(el, 'xMode', 'factor'), yMode: val(el, 'yMode', 'factor'), x: g('x'), y: g('y'), w: g('w'), h: g('h') };
    };
    // 題名：上から 6.2pt、高さは 1 行 + 3.3pt、左右のまん中
    let titleBox = null, top0 = 0;
    if (TB) {
      const m = ML(TB.layout);
      const w = TB.w + pt(7), h = TB.h + pt(3.3);
      let x = (W - w) / 2, y = pt(6.2);
      if (m && m.x != null) x = m.xMode === 'edge' ? m.x * W : x + m.x * W;
      if (m && m.y != null) y = m.yMode === 'edge' ? m.y * H : y + m.y * H;
      titleBox = { x, y, w, h };
      if (!TB.overlay && !(m && m.y != null && m.yMode === 'edge')) top0 = y + h;
    }
    // グラフの中の四角（軸の文字もふくむ）：ふちから 11pt ほど、題名・凡例から 11pt
    const F = { l: pt(11.3), t: top0 ? top0 + pt(11) : pt(11.3), r: W - pt(10.7), b: H - pt(10.7) };
    // 凡例：下なら下から 5.8pt、右なら右から 5.7pt。1 行の高さは 文字 × 1.2 + 6.1pt
    let legBox = null;
    const LG = legend && (() => {
      const sz = legend.s.sz;
      const rowH = pt(1.2 * sz + 6.1);
      const mixed = legend.items.some(i => i.kind !== 'rect');
      const keyW = pt(mixed ? 2.33 * sz : 0.49 * sz), gapK = pt(mixed ? 0.17 * sz : 0.25 * sz), sep = pt(0.61 * sz);
      const ws = legend.items.map(it => keyW + gapK + tw(it.st, it.text));
      const p = legend.pos;
      const vertical = p === 'r' || p === 'l' || p === 'tr';
      const m = ML(legend.layout);
      const padX = vertical ? [pt(3.9), pt(5.8)] : [pt(5.75), pt(5.75)];
      let rows;
      if (vertical) rows = legend.items.map((it, i) => [i]);
      else {
        const maxW = (m && m.w != null ? m.w * W : W - pt(11.4)) - padX[0] - padX[1];
        rows = [[]];
        let cw = 0;
        legend.items.forEach((it, i) => {
          const row = rows[rows.length - 1], w = ws[i] + (row.length ? sep : 0);
          if (row.length && cw + w > maxW) { rows.push([i]); cw = ws[i]; } else { row.push(i); cw += w; }
        });
      }
      const rowW = r => r.reduce((a, i, k) => a + ws[i] + (k ? sep : 0), 0);
      let w = Math.max(...rows.map(rowW)) + padX[0] + padX[1], h = rows.length * rowH;
      if (m && m.w != null) w = m.w * W;
      if (m && m.h != null) h = m.h * H;
      w = Math.min(w, W - pt(11.4)); h = Math.min(h, Math.max(rowH, H - top0 - pt(11.6)));
      let x, y;
      const midY = top0 + (H - top0 - h) / 2, topY = top0 ? top0 + pt(6) : pt(6.2);
      if (p === 'r') { x = W - pt(5.7) - w; y = midY; }
      else if (p === 'l') { x = pt(6.3); y = midY; }
      else if (p === 't') { x = (W - w) / 2; y = topY; }
      else if (p === 'tr') { x = W - pt(5.7) - w; y = topY; }
      else { x = (W - w) / 2; y = H - pt(5.8) - h; }
      if (m && m.x != null) x = m.xMode === 'edge' ? m.x * W : x + m.x * W;
      if (m && m.y != null) y = m.yMode === 'edge' ? m.y * H : y + m.y * H;
      legBox = { x, y, w, h };
      if (!legend.overlay && !(m && (m.xMode === 'edge' || m.yMode === 'edge'))) {
        if (p === 'r' || p === 'tr') F.r = Math.min(F.r, x - pt(11));
        else if (p === 'l') F.l = Math.max(F.l, x + w + pt(11));
        else if (p === 't') F.t = Math.max(F.t, y + h + pt(11));
        else F.b = Math.min(F.b, y - pt(11));
      }
      return { rows, ws, keyW, gapK, sep, rowH, padX, vertical, mixed };
    })();
    const A = { x: F.l, y: F.t, w: Math.max(10, F.r - F.l), h: Math.max(10, F.b - F.t) };

    // 描くもの
    const root2 = add(svg, 'g');
    // グラフの外わく
    {
      const sp = kid(root, 'spPr');
      const r = add(root2, 'rect', { x: 0.5, y: 0.5, width: r2(W - 1), height: r2(H - 1) });
      if (bool(root, 'roundedCorners', true)) { r.setAttribute('rx', r2(Math.min(W, H) * 0.04)); }
      const f0 = host === 'xl' ? { color: { r: 255, g: 255, b: 255, a: 1 } } : { none: true };
      const l0 = host === 'xl' ? { color: GRAY, w: 1 } : null;
      paint(r, fillFrom(sp, f0), lineFrom(sp, l0));
    }

    // 縦横の軸があるもの（円・ドーナツ・レーダー・等高線のほか）
    const axisGroups = groups.filter(g => !/^(pie|doughnut|surface|radar)$/.test(g.type));
    const plotML = ML(pathOf(plot, 'layout', 'manualLayout'));
    const plotSp = kid(plot, 'spPr');

    // 軸の組（カテゴリ・値）
    const pairs = [];
    for (const g of axisGroups) {
      let a0 = axes.get(g.axIds[0]), a1 = axes.get(g.axIds[1]);
      if (!a0 || !a1) {
        a0 = a0 || { kind: g.type === 'scatter' || g.type === 'bubble' ? 'val' : 'cat', id: '_c' + pairs.length, del: true, pos: 'b', rev: false, crosses: 'autoZero', lblPos: 'none', linked: true, majorTick: 'none', minorTick: 'none' };
        a1 = a1 || { kind: 'val', id: '_v' + pairs.length, del: true, pos: 'l', rev: false, crosses: 'autoZero', lblPos: 'none', linked: true, majorTick: 'none', minorTick: 'none' };
      }
      // 向き：横棒はカテゴリが縦
      const catV = g.type === 'bar' ? g.dir === 'bar' : false;
      if (catV) { if (!/^[lr]$/.test(a0.pos)) a0.pos = 'l'; if (!/^[tb]$/.test(a1.pos)) a1.pos = 'b'; }
      else { if (!/^[tb]$/.test(a0.pos)) a0.pos = 'b'; if (!/^[lr]$/.test(a1.pos)) a1.pos = 'l'; }
      g.ax = [a0, a1];
      let p = pairs.find(q => q.c === a0 && q.v === a1);
      if (!p) { p = { c: a0, v: a1, groups: [] }; pairs.push(p); }
      p.groups.push(g);
    }

    // 値の範囲
    function extent(p) {
      let lo = Infinity, hi = -Infinity, xlo = Infinity, xhi = -Infinity, n = 0, pct = false, bubble = false;
      const see = v => { if (v != null && isFinite(v)) { if (v < lo) lo = v; if (v > hi) hi = v; } };
      const seeX = v => { if (v != null && isFinite(v)) { if (v < xlo) xlo = v; if (v > xhi) xhi = v; } };
      for (const g of p.groups) {
        for (const s of g.series) n = Math.max(n, s.v.length, s.cat ? s.cat.n : 0);
        if (g.type === 'scatter' || g.type === 'bubble') {
          for (const s of g.series) s.v.forEach((v, i) => { see(v); seeX(s.cat && s.cat.num ? numAt(s.cat, i) : i + 1); });
          if (g.type === 'bubble') bubble = true;
        } else if (g.stacked) {
          const N = Math.max(...g.series.map(s => s.v.length), 0);
          for (let i = 0; i < N; i++) {
            let pos = 0, neg = 0, tot = 0;
            for (const s of g.series) { const v = s.v[i] || 0; tot += Math.abs(v); if (v >= 0) pos += v; else neg += v; }
            if (g.pct) { pct = true; if (tot) { see(pos / tot); see(neg / tot); } }
            else if (g.type === 'line') { let c = 0; for (const s of g.series) { c += s.v[i] || 0; see(c); } }
            else { see(pos); see(neg); }
          }
        } else for (const s of g.series) for (const v of s.v) see(v);
      }
      return { lo, hi, xlo, xhi, n, pct, bubble };
    }

    // 自動の目盛り（Excel と同じ決め方。Excel の 96 の例で確かめた）
    //   ・目盛りの数は 10 まで、かつ「軸の長さ ÷（文字の大きさ × 1.4）」まで。間かくは 1・2・5 × 10 のべき
    //   ・0 以上で、最小が最大の 5/6 より小さければ 0 から。そうでなければ「最小 −（最大 − 最小）÷ 2」から
    //   ・最大は「最大 + 幅の 5%」を切り上げ（レーダーは足さない）。マイナスだけのときは逆向きに同じ
    function scaleOf(a, lo, hi, len, sz, opt = {}) {
      if (a.log) {
        const b = a.log;
        const l0 = lo > 0 ? Math.floor(Math.log(lo) / Math.log(b) + 1e-9) : 0, l1 = hi > 0 ? Math.ceil(Math.log(hi) / Math.log(b) - 1e-9) : 1;
        const min = a.min != null ? a.min : Math.pow(b, l0), max = a.max != null ? a.max : Math.pow(b, Math.max(l1, l0 + 1));
        return { min, max, major: a.major || b, minor: a.minor || b, log: b };
      }
      if (!(lo <= hi)) { lo = 0; hi = 0; }
      if (opt.pct) { lo = Math.min(lo, 0); hi = Math.max(hi, 0); }
      let dmin = a.min != null ? a.min : lo, dmax = a.max != null ? a.max : hi, head = opt.noHead ? 0 : 0.05;
      if (dmin === dmax) { if (dmax === 0) { dmax = 1; head = 0; } else if (dmax > 0) dmin = 0; else dmax = 0; }
      if (dmin > dmax) dmax = dmin + 1;
      const zMin = a.min == null && dmin >= 0 && dmin < dmax * 5 / 6;
      const zMax = a.max == null && dmax <= 0 && dmax > dmin * 5 / 6;
      const rng = dmax - dmin;
      const maxN = Math.max(1, Math.min(10, Math.floor(len / pt(1.4 * sz))));
      const lo1 = a.min != null ? a.min : zMin ? 0 : dmin, hi1 = a.max != null ? a.max : zMax ? 0 : dmax;
      const steps = a.major ? [a.major] : niceSteps(Math.max((hi1 - lo1) / 10, 1e-12));
      let pick = null;
      for (const st of steps) {
        let min, max;
        if (a.min != null) min = a.min;
        else if (zMin) min = 0;
        else if (dmin < 0) min = Math.floor((dmin - head * rng) / st + 1e-9) * st;
        else min = Math.max(0, Math.floor((dmin - rng / 2) / st + 1e-9) * st);
        if (a.max != null) max = a.max;
        else if (zMax) max = 0;
        else if (dmax > 0) max = Math.ceil((dmax + head * rng) / st - 1e-9) * st;
        else max = Math.min(0, Math.ceil((dmax + rng / 2) / st - 1e-9) * st);
        if (opt.pct && a.max == null && max > 1) max = 1;
        if (opt.pct && a.min == null && min < -1) min = -1;
        if (max <= min) max = min + st;
        pick = { min: tidy(min), max: tidy(max), major: st };
        if (Math.round((max - min) / st) <= maxN && (!opt.fits || opt.fits(pick))) break;
      }
      pick.minor = a.minor || pick.major / 5;
      return pick;
    }

    // 軸の文字のかたまり
    function axisStyle(a) { return txs(base, a.txPr); }
    function valFmt(a, p) {
      if (!a.linked && a.fmt) return a.fmt;
      const g = p && p.groups[0];
      if (g && g.pct) return '0%';
      const s = p && p.groups.flatMap(x => x.series).find(x => x.val && (x.val.fmt || (x.val.m.get(0) && x.val.m.get(0).fmt)));
      const f = s && (s.val.fmt || s.val.m.get(0).fmt);
      return f || a.fmt || 'General';
    }
    function xFmt(a, p) {
      if (!a.linked && a.fmt) return a.fmt;
      const s = p && p.groups.flatMap(x => x.series).find(x => x.cat && x.cat.num && x.cat.fmt);
      return (s && s.cat.fmt) || a.fmt || 'General';
    }
    function dispOf(a) {
      if (!a.dispUnits) return null;
      const bu = val(a.dispUnits, 'builtInUnit', null), cu = num(a.dispUnits, 'custUnit', null);
      const d = bu && DISP[bu] ? DISP[bu] : cu ? [cu, ''] : null;
      if (!d) return null;
      const lbl = kid(a.dispUnits, 'dispUnitsLbl');
      return { div: d[0], label: lbl ? d[1] : null, el: lbl };
    }

    // 値の軸の目盛りの文字
    function valTicks(a, sc, code, disp) {
      const out = [];
      if (sc.log) { for (let v = sc.min; v <= sc.max * 1.0000001; v *= sc.log) out.push({ v, t: fmtNum(disp ? v / disp.div : v, code) }); return out; }
      const n = Math.round((sc.max - sc.min) / sc.major);
      for (let k = 0; k <= n && k < 1000; k++) { const v = tidy(sc.min + k * sc.major); out.push({ v, t: fmtNum(disp ? tidy(v / disp.div) : v, code) }); }
      return out;
    }
    // 横の値の軸：目盛りの文字が重ならないか
    const fitsH = (a, len) => sc => {
      const t = valTicks(a, sc, a.code, a.disp);
      const mw = Math.max(0, ...t.map(k => tw(a.st, k.t)));
      return len / Math.max(1, t.length - 1) >= mw + pt(a.st.sz * 0.5);
    };

    // 日付の軸：Excel の通し番号 ⇔ 年月日
    const D0 = date1904 ? 24107 : 25569;
    const ymd = d => { const t = new Date(Math.round((d - D0) * 864e5)); return [t.getUTCFullYear(), t.getUTCMonth(), t.getUTCDate()]; };
    const serial = (y, m, d) => Date.UTC(y, m, d) / 864e5 + D0;
    function dateAxis(c, s0, n0, len) {
      const ds = [];
      for (let i = 0; i < n0; i++) ds.push(numAt(s0.cat, i));
      const ok = ds.filter(d => d != null);
      if (!ok.length) return false;
      const unit = c.baseUnit || (ok.every(d => ymd(d)[2] === 1) ? 'months' : 'days');
      const key = d => { if (unit === 'days') return Math.floor(d); const [y, m] = ymd(d); return unit === 'years' ? y : y * 12 + m; };
      const back = k => (unit === 'days' ? k : unit === 'years' ? serial(k, 0, 1) : serial(Math.floor(k / 12), k % 12, 1));
      const lo = c.min != null ? key(c.min) : Math.min(...ok.map(key)), hi = c.max != null ? key(c.max) : Math.max(...ok.map(key));
      c.n = Math.max(1, hi - lo + 1);
      c.idx = ds.map(d => (d == null ? null : key(d) - lo));
      const code = !c.linked && c.fmt ? c.fmt : s0.cat.fmt || c.fmt || 'yyyy/m/d';
      const txt = k => fmtNum(back(k), code);
      // 目盛りの間かく：決まっていなければ、文字が横か斜めで重ならない、いちばん小さいもの
      let step = c.major;
      if (!step) {
        const per = len / c.n, lh = lineH(c.st, '0') * Math.SQRT2, mw = tw(c.st, txt(hi));
        const cands = unit === 'days' ? [1, 2, 7, 14, 28, 56, 91, 182, 364] : unit === 'months' ? [1, 2, 3, 6, 12, 24, 60, 120] : [1, 2, 5, 10, 20, 50, 100];
        step = cands.find(k => k * per >= Math.min(lh, mw)) || cands[cands.length - 1];
      }
      c.labels = []; c.lpos = [];
      for (let k = lo; k <= hi && c.labels.length < 1000; k += step) { c.labels.push(txt(k)); c.lpos.push(k - lo); }
      c.lstep = step;
      return true;
    }

    // 軸の準備（長さが決まってから、目盛り・文字を求める）
    function prepAxes(P) {
      for (const p of pairs) {
        const ex = extent(p);
        const c = p.c, v = p.v;
        const vVert = /^[lr]$/.test(v.pos);
        v.vert = vVert; c.vert = !vVert;
        const vLen = vVert ? P.h : P.w, cLen = vVert ? P.w : P.h;
        v.st = axisStyle(v); c.st = axisStyle(c);
        v.disp = dispOf(v);
        v.code = valFmt(v, p);
        // Excel は、日本語のカテゴリの文字の行が高くなる分を入れずに目盛りの数を決めているらしい（月別の売上で 0〜200 を 20 ずつ）
        let extra = 0;
        if (vVert && c.labels && !c.lpos) { const k = Math.max(0, ...c.labels.map(t => lineH(c.st, t))) - lineH(c.st, '0'); if (k > 0) extra = k; }
        v.sc = scaleOf(v, ex.lo, ex.hi, vLen + extra, v.st.sz, { pct: ex.pct, horiz: !vVert, fits: vVert ? null : fitsH(v, vLen) });
        v.ticks = valTicks(v, v.sc, v.code, v.disp);
        p.ex = ex;
        if (c.kind === 'val') {
          // 散布図・バブルの横の軸
          let xlo = ex.xlo, xhi = ex.xhi;
          if (!(xlo <= xhi)) { xlo = 0; xhi = 1; }
          c.code = xFmt(c, p);
          c.disp = dispOf(c);
          c.sc = scaleOf(c, xlo, xhi, cLen, c.st.sz, { horiz: !c.vert, fits: c.vert ? null : fitsH(c, cLen) });
          c.ticks = valTicks(c, c.sc, c.code, c.disp);
        } else {
          const n = Math.max(1, ex.n);
          const between = (v.between || (p.groups.some(g => g.type === 'bar' || g.type === 'stock') ? 'between' : p.groups.every(g => g.type === 'area') ? 'midCat' : 'between')) === 'between' || p.groups.some(g => g.type === 'bar');
          const s0 = p.groups.flatMap(g => g.series).find(s => s.cat) || p.groups[0].series[0];
          c.between = between;
          c.idx = null; c.lpos = null; c.lstep = 1;
          if (!(c.kind === 'date' && s0 && s0.cat && s0.cat.num && dateAxis(c, s0, n, cLen))) {
            c.n = n;
            c.labels = [];
            for (let i = 0; i < n; i++) c.labels.push(s0 ? catText(s0, i, c) : String(i + 1));
          }
          if (c.n === 1) c.between = true;
          c.levels = s0 && s0.cat && s0.cat.levels && s0.cat.levels.length > 1 && !c.noMulti ? s0.cat.levels : null;
        }
      }
    }
    // 位置の計算（0〜1 → px）
    const frac = (a, t) => (a.rev ? 1 - t : t);
    function vT(a, v) {
      const s = a.sc;
      if (s.log) { const l = x => Math.log(Math.max(x, 1e-300)); return (l(v) - l(s.min)) / (l(s.max) - l(s.min)); }
      return (v - s.min) / (s.max - s.min || 1);
    }
    function cT(a, i) { return a.between ? (i + 0.5) / a.n : a.n > 1 ? i / (a.n - 1) : 0.5; }
    function toPx(a, t, P) { t = frac(a, t); return a.vert ? P.y + P.h - t * P.h : P.x + t * P.w; }

    function axisTitle(a) {
      if (!a.title) return null;
      if (a._tb !== undefined) return a._tb;
      const s0 = base.szSet ? { ...base } : { ...base, b: true, sz: 10 };
      const B = titleBlock(a.title, s0, '軸ラベル');
      const rot = B.s && B.s.rot != null ? B.s.rot / 60000 : a.vert ? -90 : 0;
      B.rot = rot;
      const side = Math.abs(Math.abs(rot) - 90) < 1;
      B.thick = a.vert ? (side ? B.h : B.w) : side ? B.w : B.h; // 軸と直角の向きの大きさ
      a._tb = B;
      return B;
    }
    // カテゴリの文字：横に並べる・折り返す・斜め（−45°）・間引く
    //   入らなければ、まず折り返し。斜めにすれば間引かずにすむなら斜め、そうでなければ横のまま間引く（Excel と同じ）
    function catLayout(a, P) {
      const s = a.st, n = a.n, len = a.vert ? P.h : P.w;
      const slot = (len / Math.max(1, a.between ? n : n - 1)) * (a.lstep || 1);
      const rot = s.rot != null && s.rot !== -60000000 ? s.rot / 60000 : null;
      const lh = Math.max(lineH(s, '0'), ...a.labels.map(t => lineH(s, t)));
      const maxW = Math.max(0, ...a.labels.map(t => tw(s, t)));
      const one = a.labels.map(t => [t]);
      const hSkip = () => { let k = 1; while (k < a.labels.length && maxW > slot * k * 0.98) k++; return k; };
      if (a.vert) {
        const skip = a.skip || Math.max(1, Math.ceil(lh / Math.max(slot, 1e-6)));
        return { mode: 'h', rot: 0, skip, thick: maxW, lines: one };
      }
      if (rot != null && rot !== 0) {
        const r = Math.abs(rot) * Math.PI / 180;
        const skip = a.skip || Math.max(1, Math.ceil((Math.abs(rot) >= 89 ? lh : lh / Math.sin(r)) / Math.max(slot, 1e-6)));
        return { mode: 'r', rot, skip, thick: maxW * Math.sin(r) + lh * Math.cos(r), lines: one };
      }
      if (maxW <= slot * 0.98 * (a.skip || 1) || rot === 0) return { mode: 'h', rot: 0, skip: a.skip || (maxW <= slot * 0.98 ? 1 : hSkip()), thick: lh, lines: one };
      // 折り返し（3 行まで）
      if (!a.lpos && slot > lh * 1.5) {
        const wr = a.labels.map(t => wrapText(t, s, slot * 0.98));
        if (wr.every(l => l.length <= 3 && l.every(x => tw(s, x) <= slot))) return { mode: 'h', rot: 0, skip: a.skip || 1, thick: lh * Math.max(...wr.map(l => l.length)), lines: wr };
      }
      if (lh * Math.SQRT2 <= slot * (a.skip || 1)) {
        const r = Math.PI / 4;
        return { mode: 'r', rot: -45, skip: a.skip || 1, thick: Math.min(maxW, P.h * 0.6) * Math.sin(r) + lh * Math.cos(r), lines: one };
      }
      return { mode: 'h', rot: 0, skip: a.skip || hSkip(), thick: lh, lines: one };
    }

    // 軸の文字の出る所：side（l・r・t・b）、edge（文字を並べる線の位置）、dir（外向き ±1）、at（プロットのふちか）
    function lblSide(a, o, P) {
      const vert = a.vert;
      const line = vert ? toPx({ ...o, vert: false }, crossPos(a, o), P) : toPx({ ...o, vert: true }, crossPos(a, o), P);
      let dir = a.pos === 'l' || a.pos === 't' ? -1 : 1, edge = line, at;
      if (a.lblPos === 'low' || a.lblPos === 'high') {
        const lowEnd = vert ? (o.rev ? P.x + P.w : P.x) : o.rev ? P.y : P.y + P.h;
        const highEnd = vert ? (o.rev ? P.x : P.x + P.w) : o.rev ? P.y + P.h : P.y;
        edge = a.lblPos === 'low' ? lowEnd : highEnd;
        dir = vert ? (edge <= P.x + P.w / 2 ? -1 : 1) : edge >= P.y + P.h / 2 ? 1 : -1;
        at = true;
      } else at = vert ? (dir < 0 ? line <= P.x + 0.5 : line >= P.x + P.w - 0.5) : dir > 0 ? line >= P.y + P.h - 0.5 : line <= P.y + 0.5;
      return { line, edge, dir, at, side: vert ? (dir < 0 ? 'l' : 'r') : dir < 0 ? 't' : 'b' };
    }
    // 軸の文字と題名が使う幅（縦の軸）・高さ（横の軸）
    //   縦：文字の右はしは中の四角から 文字の大きさ × 0.88、外へ 4.1pt はみ出してよい
    //   横：文字の上は中の四角から 文字の大きさ × 0.55、外へ 3.7pt はみ出してよい
    //   題名があるときは、題名（ふちから）＋ すき間 ＋ 文字
    //   目盛りの印（外向き）があっても、文字の位置は変わらない（Excel の面グラフで確かめた）
    function band(a, o, P) {
      if (a.del) return null;
      const L = lblSide(a, o, P), sz = a.st.sz;
      let lab = 0;
      if (a.lblPos !== 'none' && L.at) {
        if (a.ticks) lab = a.vert ? Math.max(0, ...a.ticks.map(k => tw(a.st, k.t))) : lineH(a.st, '0');
        else if (a.labels) lab = catLayout(a, P).thick;
      }
      const tb = axisTitle(a);
      let size;
      if (a.vert) size = tb ? tb.thick + pt(10) + (lab ? lab + pt(0.88 * sz) : 0) : lab ? Math.max(0, lab + pt(0.88 * sz) - pt(4.1)) : 0;
      else size = tb ? tb.thick + pt(4.95) + (lab ? lab + pt(0.55 * sz) : 0) : lab ? Math.max(0, lab + pt(0.55 * sz) - pt(3.7)) : 0;
      return { side: L.side, size, L };
    }

    // 中の四角（プロット）
    let P, O;
    {
      O = { ...A };
      if (plotML && plotML.w != null && plotML.h != null) {
        const x = plotML.xMode === 'edge' ? plotML.x * W : O.x + (plotML.x || 0) * W;
        const y = plotML.yMode === 'edge' ? plotML.y * H : O.y + (plotML.y || 0) * H;
        O = { x, y, w: plotML.w * W, h: plotML.h * H };
      }
      P = { ...O };
      const inner = plotML && plotML.target === 'inner' && plotML.w != null;
      if (pairs.length) {
        for (let it = 0; it < 4; it++) {
          prepAxes(P);
          if (inner) break;
          const side = { l: 0, r: 0, t: 0, b: 0 };
          for (const p of pairs) for (const [a, o] of [[p.c, p.v], [p.v, p.c]]) {
            const b = band(a, o, P);
            if (!b) continue;
            side[b.side] = Math.max(side[b.side], b.size);
            if (!b.L.at || a.lblPos === 'none') continue;
            // 値の軸のはしの文字が、四角の外にはみ出す分
            if (a.ticks && a.ticks.length) {
              if (a.vert) side.t = Math.max(side.t, lineH(a.st, '0') / 2 - pt(5.9));
              else { const k = a.ticks; side.r = Math.max(side.r, tw(a.st, k[k.length - 1].t) / 2); side.l = Math.max(side.l, tw(a.st, k[0].t) / 2); }
            }
            // カテゴリの文字（横・斜め）は、グラフの中の四角の左右からはみ出さないように（面グラフのはしの文字など）
            if (!a.ticks && a.labels && !a.vert) {
              const Lc = catLayout(a, P);
              const r = Lc.rot * Math.PI / 180, c = Math.abs(Math.cos(r)), sn = Math.abs(Math.sin(r)), lh = lineH(a.st, '0');
              for (let i = 0; i < a.labels.length; i += Lc.skip) {
                const t = frac(a, cT(a, a.lpos ? a.lpos[i] : i));
                let left, right;
                if (Lc.mode === 'r') {
                  const w = Math.min(tw(a.st, a.labels[i]), P.h * 0.6 / Math.max(sn, 0.1));
                  left = Lc.rot < 0 ? w * c : lh * sn; right = Lc.rot < 0 ? lh * sn : w * c;
                } else left = right = Math.max(0, ...Lc.lines[i].map(x => tw(a.st, x))) / 2;
                side.l = Math.max(side.l, left - t * P.w);
                side.r = Math.max(side.r, right - (1 - t) * P.w);
              }
            }
          }
          const np = { x: O.x + side.l, y: O.y + side.t, w: Math.max(10, O.w - side.l - side.r), h: Math.max(10, O.h - side.t - side.b) };
          const same = Math.abs(np.x - P.x) + Math.abs(np.y - P.y) + Math.abs(np.w - P.w) + Math.abs(np.h - P.h) < 0.1;
          P = np;
          if (same) break;
        }
        prepAxes(P);
      }
    }

    /* ----- 描く ----- */
    const gPlot = add(root2, 'g');
    if (plotSp) paint(add(gPlot, 'rect', { x: r2(P.x), y: r2(P.y), width: r2(P.w), height: r2(P.h) }), fillFrom(plotSp, { none: true }), lineFrom(plotSp, null));
    const clipId = uid();
    const cp = add(defs, 'clipPath', { id: clipId });
    add(cp, 'rect', { x: r2(P.x - 0.5), y: r2(P.y - 0.5), width: r2(P.w + 1), height: r2(P.h + 1) });
    const gGrid = add(gPlot, 'g');
    const gSer = add(gPlot, 'g', { 'clip-path': `url(#${clipId})` });
    const gMk = add(gPlot, 'g');
    const gAxis = add(gPlot, 'g');
    const gLbl = add(gPlot, 'g');

    const line = (parent, x1, y1, x2, y2, l) => paint(add(parent, 'line', { x1: r2(x1), y1: r2(y1), x2: r2(x2), y2: r2(y2) }), null, l);
    const path = (parent, d, f, l) => paint(add(parent, 'path', { d }), f, l);

    // a の線が、交わる軸 other の上のどこにあるか（0〜1）
    function crossPos(a, other) {
      if (!other) return 0;
      if (other.sc) {
        const s = other.sc;
        let v = a.crosses === 'max' ? s.max : a.crosses === 'min' ? s.min : a.crossesAt != null ? a.crossesAt : clamp(s.log ? s.min : 0, s.min, s.max);
        if (a.crosses === 'autoZero' && s.log) v = s.min;
        return clamp(vT(other, v), 0, 1);
      }
      if (a.crosses === 'max') return 1;
      if (a.crossesAt != null) return clamp(other.n > 1 ? (other.between ? (a.crossesAt - 1) / other.n : (a.crossesAt - 1) / (other.n - 1)) : 0, 0, 1);
      return 0;
    }
    // 目盛り線・軸の線・目盛りの文字・軸の題名
    function drawAxis(p, a, other) {
      const vert = a.vert;
      const L = lblSide(a, other, P), posPx = L.line;
      const gl = a.grid ? lineFrom(kid(a.grid, 'spPr'), { color: tx1, w: 1 }) : null;
      const mgl = a.mgrid ? lineFrom(kid(a.mgrid, 'spPr'), { color: tx1, w: 1 }) : null;
      const al = lineFrom(a.spPr, { color: tx1, w: 1 }); // 指定がなければ黒い線（PowerPoint で確かめた）
      // 目盛りの位置（0〜1）
      let majors = [], minors = [];
      if (a.ticks) {
        majors = a.ticks.map(k => vT(a, k.v));
        if (a.mgrid || a.minorTick !== 'none') {
          const s = a.sc;
          if (s.log) { for (let v = s.min; v < s.max; v *= s.log) for (let k = 2; k < s.log; k++) minors.push(vT(a, v * k)); }
          else { const n = Math.round((s.max - s.min) / s.minor); for (let k = 1; k < n && k < 2000; k++) minors.push((k * s.minor) / (s.max - s.min)); }
        }
      } else if (a.lpos) majors = a.lpos.map(sl => (a.between ? sl / a.n : cT(a, sl)));
      else if (a.n != null) {
        const n = a.n, sk = a.markSkip || 1;
        if (a.between) for (let i = 0; i <= n; i += sk) majors.push(i / n);
        else for (let i = 0; i < n; i += sk) majors.push(cT(a, i));
      }
      const at2 = t => toPx(a, t, P);
      const span = vert ? [P.x, P.x + P.w] : [P.y, P.y + P.h];
      if (mgl && !mgl.none) for (const t of minors) { const q = at2(t); if (vert) line(gGrid, span[0], q, span[1], q, mgl); else line(gGrid, q, span[0], q, span[1], mgl); }
      if (gl && !gl.none) for (const t of majors) { const q = at2(t); if (vert) line(gGrid, span[0], q, span[1], q, gl); else line(gGrid, q, span[0], q, span[1], gl); }
      if (a.del) return;
      // 軸の線
      if (al && !al.none) { if (vert) line(gAxis, posPx, P.y, posPx, P.y + P.h, al); else line(gAxis, P.x, posPx, P.x + P.w, posPx, al); }
      // 目盛りの印
      const out = a.pos === 'l' || a.pos === 't' ? -1 : 1;
      const tick = (list, kind, len) => {
        if (kind === 'none' || !al || al.none) return;
        const a0 = kind === 'in' ? 0 : out * len, a1 = kind === 'in' || kind === 'cross' ? -out * len : 0;
        for (const t of list) { const q = at2(t); if (vert) line(gAxis, posPx + a0, q, posPx + a1, q, al); else line(gAxis, q, posPx + a0, q, posPx + a1, al); }
      };
      tick(majors, a.majorTick, pt(4));
      tick(minors, a.minorTick, pt(2));
      // 文字
      const s = a.st, sz = s.sz, edge = L.edge, dir = L.dir;
      const gap = vert ? pt(0.88 * sz) : pt(0.55 * sz);
      if (a.lblPos !== 'none') {
        if (a.ticks) {
          for (const k of a.ticks) {
            const q = at2(vT(a, k.v));
            const B = plain(k.t, s);
            if (vert) drawBlock(gLbl, B, edge + dir * (gap + B.w / 2), q);
            else drawBlock(gLbl, B, q, edge + dir * (gap + B.h / 2));
          }
        } else if (a.labels) {
          const Lc = catLayout(a, P);
          for (let i = 0; i < a.labels.length; i += Lc.skip) {
            const q = at2(cT(a, a.lpos ? a.lpos[i] : i));
            const B = block(Lc.lines[i].map(t => [{ t, s }]));
            if (vert) drawBlock(gLbl, B, edge + dir * (gap + B.w / 2), q, dir < 0 ? 'end' : 'start');
            else if (Lc.mode === 'r') {
              // 文字の終わり（上向きなら右の端）を目盛りの下に合わせる
              const r = Lc.rot * Math.PI / 180, c = Math.cos(r), sn = Math.sin(r);
              const ax = q, ay = edge + dir * gap, half = B.w / 2, hh = B.h / 2;
              let cx2, cy2;
              if (Lc.rot < 0) { cx2 = ax - c * half - sn * hh; cy2 = ay - sn * half + c * hh; }
              else { cx2 = ax + c * half - sn * hh; cy2 = ay + sn * half + c * hh; }
              drawBlock(gLbl, B, cx2, dir > 0 ? cy2 : 2 * ay - cy2, 'middle', Lc.rot);
            } else drawBlock(gLbl, B, q, edge + dir * (gap + B.h / 2));
          }
        }
      }
      // 軸の題名：四角のふちから（縦は 3.7pt、横は 2.35pt）
      const tb = axisTitle(a);
      if (tb) {
        const m = ML(tb.layout);
        let x, y;
        if (vert) { x = L.side === 'l' ? O.x + pt(3.7) + tb.thick / 2 : O.x + O.w - pt(3.7) - tb.thick / 2; y = P.y + P.h / 2; }
        else { y = L.side === 'b' ? O.y + O.h - pt(2.35) - tb.thick / 2 : O.y + pt(2.35) + tb.thick / 2; x = P.x + P.w / 2; }
        const side = Math.abs(Math.abs(tb.rot) - 90) < 1;
        if (m && m.x != null) x = m.xMode === 'edge' ? m.x * W + (side ? tb.h : tb.w) / 2 : x + m.x * W;
        if (m && m.y != null) y = m.yMode === 'edge' ? m.y * H + (side ? tb.w : tb.h) / 2 : y + m.y * H;
        drawBlock(gLbl, tb, x, y, 'middle', tb.rot);
      }
      if (a.disp && a.disp.label && a.lblPos !== 'none') {
        const B = plain(a.disp.label, txs(s, kid(a.disp.el, 'txPr')));
        if (vert) drawBlock(gLbl, B, edge + dir * (gap + B.w / 2), P.y - B.h / 2 - pt(2));
        else drawBlock(gLbl, B, P.x + P.w - B.w / 2, edge + dir * (gap + lineH(s, '0') + B.h / 2 + pt(2)));
      }
    }

    // 系列を描く
    const blanks = val(chart, 'dispBlanksAs', 'gap');
    const labels = []; // 値の文字（あとでまとめて描く）
    function dl(g, s) {
      const e = s.dLbls || g.dLbls;
      if (!e) return null;
      const one = (x, d) => ({ val: bool(x, 'showVal', d.val), cat: bool(x, 'showCatName', d.cat), ser: bool(x, 'showSerName', d.ser), pct: bool(x, 'showPercent', d.pct), bub: bool(x, 'showBubbleSize', d.bub), key: bool(x, 'showLegendKey', d.key) });
      const show = one(e, { val: false, cat: false, ser: false, pct: false, bub: false, key: false });
      const nf = kid(e, 'numFmt');
      const sep = kid(e, 'separator');
      const cfg = { show, pos: val(e, 'dLblPos', null), fmt: nf && !truthy(at(nf, 'sourceLinked')) ? at(nf, 'formatCode') : null, sep: sep ? sep.textContent : ', ', st: txs(base, kid(e, 'txPr')), spPr: kid(e, 'spPr'), del: bool(e, 'delete', false), per: new Map() };
      for (const d of kids(e, 'dLbl')) {
        const i = num(d, 'idx', -1), nf2 = kid(d, 'numFmt'), sep2 = kid(d, 'separator');
        cfg.per.set(i, { show: one(d, show), del: bool(d, 'delete', false), pos: val(d, 'dLblPos', cfg.pos), fmt: nf2 && !truthy(at(nf2, 'sourceLinked')) ? at(nf2, 'formatCode') : cfg.fmt, sep: sep2 ? sep2.textContent : cfg.sep, st: kid(d, 'txPr') ? txs(cfg.st, kid(d, 'txPr')) : cfg.st, spPr: kid(d, 'spPr') || cfg.spPr, rich: pathOf(d, 'tx', 'rich'), layout: pathOf(d, 'layout', 'manualLayout') });
      }
      return cfg;
    }
    function dlText(cfg, g, s, i, v, extra = {}) {
      if (!cfg || cfg.del) return null;
      const c = cfg.per.get(i) || cfg;
      if (c.del) return null;
      if (c.rich) return { B: richBlock(c.rich, c.st), c };
      const sh = c.show, parts = [];
      if (sh.ser) parts.push(s.name);
      if (sh.cat) parts.push(g.type === 'scatter' || g.type === 'bubble' ? fmtNum(extra.x, (s.cat && s.cat.fmt) || 'General') : catText(s, i, g.ax && g.ax[0]));
      if (sh.val && v != null) parts.push(fmtNum(v, c.fmt || (s.val && ((s.val.m.get(i) && s.val.m.get(i).fmt) || s.val.fmt)) || 'General'));
      if (sh.pct && extra.pct != null) parts.push(fmtNum(extra.pct, c.fmt && sh.val === false ? c.fmt : '0%'));
      if (sh.bub && extra.size != null) parts.push(fmtNum(extra.size, 'General'));
      if (!parts.length) return null;
      return { B: plain(parts.join(c.sep), c.st), c };
    }
    // 値の文字をためる：x, y は文字の中心
    const pushLbl = (L, x, y) => { if (L) labels.push({ ...L, x, y }); };

    function smoothD(pts) {
      if (pts.length < 3) return 'M' + pts.map(p => `${r2(p[0])},${r2(p[1])}`).join('L');
      let d = `M${r2(pts[0][0])},${r2(pts[0][1])}`;
      for (let i = 0; i < pts.length - 1; i++) {
        const p0 = pts[i - 1] || pts[i], p1 = pts[i], p2 = pts[i + 1], p3 = pts[i + 2] || p2;
        const t = 1 / 6;
        d += `C${r2(p1[0] + (p2[0] - p0[0]) * t)},${r2(p1[1] + (p2[1] - p0[1]) * t)} ${r2(p2[0] - (p3[0] - p1[0]) * t)},${r2(p2[1] - (p3[1] - p1[1]) * t)} ${r2(p2[0])},${r2(p2[1])}`;
      }
      return d;
    }
    const polyD = pts => 'M' + pts.map(p => `${r2(p[0])},${r2(p[1])}`).join('L');
    function drawMarker(mk, x, y) {
      if (!mk || x < P.x - 1 || x > P.x + P.w + 1 || y < P.y - 1 || y > P.y + P.h + 1) return;
      path(gMk, markD(mk.sym, x, y, mk.r), mk.fill, mk.line);
    }
    // 近似曲線
    function drawTrend(tr, xs, ys, toXY, lk) {
      const type = val(tr, 'trendlineType', 'linear');
      const pts = xs.map((x, i) => [x, ys[i]]).filter(p => p[0] != null && p[1] != null && isFinite(p[0]) && isFinite(p[1]));
      if (pts.length < 2) return;
      const l = lineFrom(kid(tr, 'spPr'), { color: lk.auto, w: 2.25 * PT / 1.0 / 1.5, dash: 'sysDot' });
      const x0 = Math.min(...pts.map(p => p[0])) - num(tr, 'backward', 0), x1 = Math.max(...pts.map(p => p[0])) + num(tr, 'forward', 0);
      let f = null;
      const fit = (X, Y) => {
        const n = X.length, mx = X.reduce((a, b) => a + b, 0) / n, my = Y.reduce((a, b) => a + b, 0) / n;
        let sxy = 0, sxx = 0; for (let i = 0; i < n; i++) { sxy += (X[i] - mx) * (Y[i] - my); sxx += (X[i] - mx) ** 2; }
        const b = sxx ? sxy / sxx : 0; return [my - b * mx, b];
      };
      if (type === 'linear') { const [a0, b] = fit(pts.map(p => p[0]), pts.map(p => p[1])); f = x => a0 + b * x; }
      else if (type === 'exp') { const q = pts.filter(p => p[1] > 0); if (q.length < 2) return; const [a0, b] = fit(q.map(p => p[0]), q.map(p => Math.log(p[1]))); f = x => Math.exp(a0 + b * x); }
      else if (type === 'log') { const q = pts.filter(p => p[0] > 0); if (q.length < 2) return; const [a0, b] = fit(q.map(p => Math.log(p[0])), q.map(p => p[1])); f = x => (x > 0 ? a0 + b * Math.log(x) : null); }
      else if (type === 'power') { const q = pts.filter(p => p[0] > 0 && p[1] > 0); if (q.length < 2) return; const [a0, b] = fit(q.map(p => Math.log(p[0])), q.map(p => Math.log(p[1]))); f = x => (x > 0 ? Math.exp(a0) * Math.pow(x, b) : null); }
      else if (type === 'poly') {
        const k = clamp(num(tr, 'order', 2), 2, 6), n = k + 1;
        const M = Array.from({ length: n }, () => new Array(n + 1).fill(0));
        for (const [x, y] of pts) for (let r = 0; r < n; r++) { for (let c = 0; c < n; c++) M[r][c] += Math.pow(x, r + c); M[r][n] += y * Math.pow(x, r); }
        for (let c = 0; c < n; c++) {
          let piv = c; for (let r = c + 1; r < n; r++) if (Math.abs(M[r][c]) > Math.abs(M[piv][c])) piv = r;
          [M[c], M[piv]] = [M[piv], M[c]];
          if (!M[c][c]) return;
          for (let r = 0; r < n; r++) if (r !== c) { const q = M[r][c] / M[c][c]; for (let j = c; j <= n; j++) M[r][j] -= q * M[c][j]; }
        }
        const co = M.map((r, i) => r[n] / r[i]);
        f = x => co.reduce((a, c, i) => a + c * Math.pow(x, i), 0);
      } else if (type === 'movingAvg') {
        const per = Math.max(2, num(tr, 'period', 2)), q = [];
        for (let i = per - 1; i < pts.length; i++) { let s = 0; for (let j = 0; j < per; j++) s += pts[i - j][1]; q.push(toXY(pts[i][0], s / per)); }
        if (q.length > 1) path(gSer, polyD(q), { none: true }, l);
        return;
      }
      if (!f) return;
      const out = [];
      for (let k = 0; k <= 60; k++) { const x = x0 + (x1 - x0) * k / 60, y = f(x); if (y != null && isFinite(y)) out.push(toXY(x, y)); }
      if (out.length > 1) path(gSer, polyD(out), { none: true }, l);
    }

    for (const p of pairs) {
      const c = p.c, v = p.v;
      // 横の位置・縦の位置
      const catPx = t => toPx(c, t, P), valPx = x => toPx(v, vT(v, x), P);
      const XY = (ct, x) => (c.vert ? [valPx(x), catPx(ct)] : [catPx(ct), valPx(x)]);
      const baseV = clamp(v.sc.log ? v.sc.min : 0, v.sc.min, v.sc.max);
      // 縦の軸がどこで交わるか（棒の根もと）
      const crossV = (() => { const s = v.sc; const x = c.crosses === 'max' ? s.max : c.crosses === 'min' ? s.min : c.crossesAt != null ? c.crossesAt : baseV; return clamp(x, s.min, s.max); })();
      // 点の数と、その点のカテゴリの位置（日付の軸では日付の場所）
      const NP = c.idx ? c.idx.length : c.n || 1, at = i => (c.idx ? c.idx[i] : i), ctr = i => cT(c, at(i));
      for (const g of p.groups) {
        const N = c.n || 1;
        if (g.type === 'bar') {
          const ss = g.series, k = g.stacked ? 1 : ss.length;
          const ov = g.overlap != null ? g.overlap : g.stacked ? 100 : 0;
          const slot = 1 / N;
          const bw = slot / (k - (k - 1) * ov / 100 + g.gap / 100);
          const posAcc = new Array(NP).fill(0), negAcc = new Array(NP).fill(0);
          const tot = new Array(NP).fill(0);
          if (g.pct) for (const s of ss) s.v.forEach((x, i) => { tot[i] += Math.abs(x || 0); });
          const cfgs = ss.map(s => dl(g, s));
          ss.forEach((s, j) => {
            for (let i = 0; i < NP; i++) {
              let x = s.v[i];
              if (x == null || at(i) == null) continue;
              const raw = x;
              if (g.pct) x = tot[i] ? x / tot[i] : 0;
              let b0, b1;
              if (g.stacked) { if (x >= 0) { b0 = posAcc[i]; b1 = b0 + x; posAcc[i] = b1; } else { b0 = negAcc[i]; b1 = b0 + x; negAcc[i] = b1; } }
              else { b0 = crossV; b1 = x; }
              const t0 = (at(i) + 0.5) * slot - (bw * (k - (k - 1) * ov / 100)) / 2 + (g.stacked ? 0 : j * bw * (1 - ov / 100));
              const t1 = t0 + bw;
              const lk = look(g, s, i);
              const ca = catPx(t0), cb = catPx(t1), va = valPx(b0), vb = valPx(b1);
              const xx = c.vert ? Math.min(va, vb) : Math.min(ca, cb), yy = c.vert ? Math.min(ca, cb) : Math.min(va, vb);
              const ww = c.vert ? Math.abs(vb - va) : Math.abs(cb - ca), hh = c.vert ? Math.abs(cb - ca) : Math.abs(vb - va);
              paint(add(gSer, 'rect', { x: r2(xx), y: r2(yy), width: r2(Math.max(ww, 0)), height: r2(Math.max(hh, 0)) }), lk.fill, lk.line);
              const L = dlText(cfgs[j], g, s, i, raw);
              if (L) {
                // 棒のはしから 3pt（文字の枠のすき間もふくめて）
                const pos = (L.c.pos) || (g.stacked ? 'ctr' : 'outEnd');
                const B = L.B;
                const along = (c.vert ? B.w + pt(5.9) : B.h + pt(3.2)) / 2 + pt(3);
                const mid = (catPx(t0) + catPx(t1)) / 2;
                let q;
                const e0 = valPx(b0), e1 = valPx(b1), sgn = Math.sign(e1 - e0) || (c.vert ? 1 : -1);
                if (pos === 'ctr') q = (e0 + e1) / 2;
                else if (pos === 'inEnd') q = e1 - sgn * along;
                else if (pos === 'inBase') q = e0 + sgn * along;
                else q = e1 + sgn * along;
                if (c.vert) pushLbl(L, q, mid); else pushLbl(L, mid, q);
              }
            }
          });
        } else if (g.type === 'line' || g.type === 'stock') {
          const acc = new Array(NP).fill(0), tot = new Array(NP).fill(0);
          if (g.pct) for (const s of g.series) s.v.forEach((x, i) => { tot[i] += Math.abs(x || 0); });
          const all = [];
          g.series.forEach(s => {
            const lk = look(g, s, -1), pts = [], segs = [[]];
            for (let i = 0; i < NP; i++) {
              let x = s.v[i];
              if (at(i) == null) continue;
              if (x == null) { if (blanks === 'zero') x = 0; else { if (blanks !== 'span' && segs[segs.length - 1].length) segs.push([]); continue; } }
              if (g.stacked) { acc[i] += g.pct ? (tot[i] ? x / tot[i] : 0) : x; x = acc[i]; }
              const xy = XY(ctr(i), x);
              pts.push({ i, xy, raw: s.v[i] });
              segs[segs.length - 1].push(xy);
            }
            if (lk.line && !lk.line.none) for (const sg of segs) if (sg.length > 1) path(gSer, s.smooth ? smoothD(sg) : polyD(sg), { none: true }, lk.line);
            all.push({ s, pts, lk });
            for (const tr of s.trend) drawTrend(tr, s.v.map((_, i) => (at(i) == null ? null : at(i) + 1)), s.v, (x, y) => XY(cT(c, x - 1), y), lk);
          });
          // 高低線・ドロップ線・ローソク
          if (g.hiLow) {
            const l = lineFrom(kid(g.hiLow, 'spPr'), { color: tx1, w: 1 });
            for (let i = 0; i < NP; i++) {
              const ys = g.series.map(s => s.v[i]).filter(x => x != null);
              if (ys.length < 2 || at(i) == null) continue;
              const a0 = XY(ctr(i), Math.min(...ys)), a1 = XY(ctr(i), Math.max(...ys));
              line(gSer, a0[0], a0[1], a1[0], a1[1], l);
            }
          }
          if (g.drop) {
            const l = lineFrom(kid(g.drop, 'spPr'), { color: tx1, w: 1 });
            for (const a of all) for (const q of a.pts) { const b = XY(ctr(q.i), crossV); line(gSer, q.xy[0], q.xy[1], b[0], b[1], l); }
          }
          if (g.upDown && g.series.length >= 2) {
            const gw = num(g.upDown, 'gapWidth', 150), bw = (1 / N) / (1 + gw / 100);
            const up = kid(g.upDown, 'upBars'), dn = kid(g.upDown, 'downBars');
            const s0 = g.series[0], s1 = g.series[g.series.length - 1];
            for (let i = 0; i < NP; i++) {
              const o2 = s0.v[i], cl = s1.v[i];
              if (o2 == null || cl == null || at(i) == null) continue;
              const isUp = cl >= o2, el = isUp ? up : dn;
              const f = fillFrom(kid(el, 'spPr'), { color: isUp ? { r: 255, g: 255, b: 255, a: 1 } : { r: 0, g: 0, b: 0, a: 1 } });
              const l = lineFrom(kid(el, 'spPr'), { color: tx1, w: 1 });
              const t = ctr(i), p0 = XY(t - bw / 2, o2), p1 = XY(t + bw / 2, cl);
              paint(add(gSer, 'rect', { x: r2(Math.min(p0[0], p1[0])), y: r2(Math.min(p0[1], p1[1])), width: r2(Math.abs(p1[0] - p0[0])), height: r2(Math.abs(p1[1] - p0[1])) }), f, l);
            }
          }
          for (const a of all) {
            const cfg = dl(g, a.s);
            for (const q of a.pts) {
              const lk = a.s.dPt.has(q.i) ? look(g, a.s, q.i) : a.lk;
              drawMarker(lk.mk, q.xy[0], q.xy[1]);
              const L = dlText(cfg, g, a.s, q.i, q.raw);
              if (L) placeNear(L, q.xy, L.c.pos || 'r');
            }
          }
        } else if (g.type === 'area') {
          const acc = new Array(NP).fill(0), tot = new Array(NP).fill(0);
          if (g.pct) for (const s of g.series) s.v.forEach((x, i) => { tot[i] += Math.abs(x || 0); });
          const layers = [];
          const ok = i => at(i) != null;
          g.series.forEach(s => {
            const prev = acc.slice(), top = [];
            for (let i = 0; i < NP; i++) {
              let x = s.v[i] == null ? 0 : s.v[i];
              if (g.stacked) { acc[i] += g.pct ? (tot[i] ? x / tot[i] : 0) : x; x = acc[i]; }
              top.push(x);
            }
            layers.push({ s, top, bot: g.stacked ? prev : new Array(NP).fill(crossV) });
          });
          for (const L of layers) {
            const lk = look(g, L.s, -1);
            const up = L.top.map((x, i) => ok(i) && XY(ctr(i), x)).filter(Boolean), dn = L.bot.map((x, i) => ok(i) && XY(ctr(i), x)).filter(Boolean).reverse();
            if (N === 1) { const w = 0.5 / N; up.splice(0, 1, XY(ctr(0) - w, L.top[0]), XY(ctr(0) + w, L.top[0])); dn.splice(0, 1, XY(ctr(0) + w, L.bot[0]), XY(ctr(0) - w, L.bot[0])); }
            path(gSer, polyD([...up, ...dn]) + 'Z', lk.fill, lk.line);
            const cfg = dl(g, L.s);
            L.top.forEach((x, i) => { const L2 = ok(i) && dlText(cfg, g, L.s, i, L.s.v[i]); if (L2) { const a0 = XY(ctr(i), x), a1 = XY(ctr(i), L.bot[i]); pushLbl(L2, (a0[0] + a1[0]) / 2, (a0[1] + a1[1]) / 2); } });
          }
        } else if (g.type === 'scatter' || g.type === 'bubble') {
          const xOf = (s, i) => (s.cat && s.cat.num ? numAt(s.cat, i) : i + 1);
          const XY2 = (x, y) => (c.vert ? [valPx(y), toPx(c, vT(c, x), P)] : [toPx(c, vT(c, x), P), valPx(y)]);
          if (g.type === 'bubble') {
            const sizes = g.series.flatMap(s => s.v.map((_, i) => numAt(s.size, i))).filter(x => x != null);
            const mx = Math.max(...sizes.map(Math.abs), 1e-12);
            const maxD = Math.min(P.w, P.h) * 0.25 * g.bubbleScale / 100;
            g.series.forEach(s => {
              const cfg = dl(g, s);
              s.v.forEach((y, i) => {
                const sz = numAt(s.size, i), x = xOf(s, i);
                if (y == null || x == null || sz == null || (sz < 0 && !g.showNeg)) return;
                const r = (g.sizeRep === 'w' ? Math.abs(sz) / mx : Math.sqrt(Math.abs(sz) / mx)) * maxD / 2;
                const [px, py] = XY2(x, y);
                const lk = look(g, s, i);
                path(gSer, markD('circle', px, py, r), sz < 0 ? { none: true } : lk.fill, lk.line);
                const L = dlText(cfg, g, s, i, y, { x, size: sz });
                if (L) pushLbl(L, px, py);
              });
            });
          } else {
            g.series.forEach(s => {
              const lk = look(g, s, -1), segs = [[]], pts = [];
              s.v.forEach((y, i) => {
                const x = xOf(s, i);
                if (y == null || x == null) { if (blanks !== 'span' && segs[segs.length - 1].length) segs.push([]); return; }
                const xy = XY2(x, y);
                segs[segs.length - 1].push(xy); pts.push({ i, xy, x, y });
              });
              if (lk.line && !lk.line.none) for (const sg of segs) if (sg.length > 1) path(gSer, s.smooth ? smoothD(sg) : polyD(sg), { none: true }, lk.line);
              for (const tr of s.trend) drawTrend(tr, s.v.map((_, i) => xOf(s, i)), s.v, (x, y) => XY2(x, y), lk);
              const cfg = dl(g, s);
              for (const q of pts) {
                const lk2 = s.dPt.has(q.i) ? look(g, s, q.i) : lk;
                drawMarker(lk2.mk, q.xy[0], q.xy[1]);
                const L = dlText(cfg, g, s, q.i, q.y, { x: q.x });
                if (L) placeNear(L, q.xy, L.c.pos || 'r');
              }
            });
          }
        }
      }
      drawAxis(p, c, v);
      drawAxis(p, v, c);
    }
    // 点のそばの値の文字：文字の枠（すき間 横 2.95pt・縦 1.6pt）が点から 5.5pt
    function placeNear(L, xy, pos) {
      const B = L.B, gx = pt(5.5 + 2.95) + B.w / 2, gy = pt(5.5 + 1.6) + B.h / 2;
      const [x, y] = xy;
      if (pos === 'l') pushLbl(L, x - gx, y);
      else if (pos === 't') pushLbl(L, x, y - gy);
      else if (pos === 'b') pushLbl(L, x, y + gy);
      else if (pos === 'ctr') pushLbl(L, x, y);
      else pushLbl(L, x + gx, y);
    }

    // 円・ドーナツ：使える所の短いほうの辺の円（値の文字があると 0.86 倍）
    const pies = groups.filter(g => g.type === 'pie' || g.type === 'doughnut');
    for (const g of pies) {
      const rings = g.type === 'doughnut' ? g.series : g.series.slice(0, 1);
      if (!rings.length) continue;
      const cfgs = rings.map(s => dl(g, s));
      const anyLbl = rings.some((s, ri) => { const c = cfgs[ri]; return c && !c.del && s.v.some((x, i) => dlText(c, g, s, i, x, { pct: 0 })); });
      const maxEx = Math.max(0, ...rings.flatMap(s => [s.explosion, ...[...s.dPt.values()].map(d => num(d, 'explosion', 0))]));
      let R = Math.max(4, (Math.min(P.w, P.h) * (anyLbl ? 0.86 : 1)) / 2);
      R /= 1 + maxEx / 100;
      const cxp = P.x + P.w / 2, cyp = P.y + P.h / 2;
      const hole = g.type === 'doughnut' ? clamp(g.hole, 10, 90) / 100 : 0;
      const ringW = (R * (1 - hole)) / rings.length;
      rings.forEach((s, ri) => {
        const vals = s.v.map(x => (x == null ? 0 : Math.abs(x)));
        const tot = vals.reduce((a, b) => a + b, 0);
        if (!tot) return;
        const r1 = R - ringW * (rings.length - 1 - ri), r0 = g.type === 'doughnut' ? r1 - ringW : 0;
        let ang = (g.first || 0) * Math.PI / 180;
        vals.forEach((x, i) => {
          const a0 = ang, a1 = ang + (x / tot) * Math.PI * 2;
          ang = a1;
          if (!x) return;
          const lk = look(g, s, i);
          const dp = s.dPt.get(i), ex = (dp && kid(dp, 'explosion') ? num(dp, 'explosion', 0) : s.explosion) / 100 * R;
          const mid = (a0 + a1) / 2, ox = Math.sin(mid) * ex, oy = -Math.cos(mid) * ex;
          const P2 = (rr, a) => [cxp + ox + Math.sin(a) * rr, cyp + oy - Math.cos(a) * rr];
          let d;
          if (a1 - a0 >= Math.PI * 2 - 1e-9) {
            d = markD('circle', cxp, cyp, r1);
            if (r0 > 0) d += markD('circle', cxp, cyp, r0);
          } else {
            const large = a1 - a0 > Math.PI ? 1 : 0;
            const [x0, y0] = P2(r1, a0), [x1, y1] = P2(r1, a1);
            d = `M${r2(x0)},${r2(y0)}A${r2(r1)},${r2(r1)} 0 ${large} 1 ${r2(x1)},${r2(y1)}`;
            if (r0 > 0) { const [x2, y2] = P2(r0, a1), [x3, y3] = P2(r0, a0); d += `L${r2(x2)},${r2(y2)}A${r2(r0)},${r2(r0)} 0 ${large} 0 ${r2(x3)},${r2(y3)}Z`; }
            else { const [xc, yc] = P2(0, 0); d += `L${r2(xc)},${r2(yc)}Z`; }
          }
          const el = path(gMk, d, lk.fill, lk.line);
          el.setAttribute('fill-rule', 'evenodd');
          const L = dlText(cfgs[ri], g, s, i, s.v[i], { pct: x / tot });
          if (L) {
            const pos = L.c.pos || (g.type === 'doughnut' ? 'ctr' : 'bestFit');
            const bw = L.B.w + pt(5.9), bh = L.B.h + pt(3.2); // 文字の枠
            const outR = r1 + pt(1.6) + Math.abs(Math.sin(mid)) * bw / 2 + Math.abs(Math.cos(mid)) * bh / 2;
            let rr;
            if (g.type === 'doughnut' || pos === 'ctr') rr = (r0 + r1) / 2;
            else if (pos === 'inEnd') rr = r1 - Math.max(bw, bh) / 2 - pt(2);
            else if (pos === 'outEnd') rr = outR;
            else rr = (a1 - a0) * r1 * 0.79 >= bw && bh < r1 * 0.4 ? r1 * 0.79 : outR; // ちょうどよい所：入るなら中、入らなければ外
            const [lx, ly] = P2(rr, mid);
            pushLbl(L, lx, ly);
          }
        });
      });
    }

    // レーダー：上下に「文字の行の 0.8 倍」、左右に文字の幅をあけた円
    const radars = groups.filter(g => g.type === 'radar');
    if (radars.length) {
      const g0 = radars[0], ca = axes.get(g0.axIds[0]), va = axes.get(g0.axIds[1]) || { kind: 'val', rev: false, linked: true, del: true };
      const n = Math.max(1, ...radars.flatMap(g => g.series.map(s => Math.max(s.v.length, s.cat ? s.cat.n : 0))));
      const cs = ca ? axisStyle(ca) : base, vs = axisStyle(va);
      const s0 = radars.flatMap(g => g.series).find(s => s.cat) || g0.series[0];
      const labelsR = [];
      for (let i = 0; i < n; i++) labelsR.push(s0 ? catText(s0, i, ca) : String(i + 1));
      const lh = Math.max(lineH(cs, '0'), ...labelsR.map(t => lineH(cs, t)));
      const showC = ca && !ca.del && ca.lblPos !== 'none';
      // 横（90°・270° の近く）にある文字の幅
      let side = 0;
      labelsR.forEach((t, i) => { const sn = Math.abs(Math.sin((i / n) * Math.PI * 2)); if (sn > 0.3) side = Math.max(side, tw(cs, t) * sn); });
      const R = Math.max(5, showC ? Math.min((P.h - 1.6 * lh) / 2, (P.w - 2 * (side + pt(3))) / 2) : Math.min(P.w, P.h) / 2);
      const cxp = P.x + P.w / 2, cyp = P.y + P.h / 2;
      let lo = Infinity, hi = -Infinity;
      for (const g of radars) for (const s of g.series) for (const x of s.v) if (x != null) { lo = Math.min(lo, x); hi = Math.max(hi, x); }
      va.st = vs;
      va.sc = scaleOf(va, lo, hi, R, vs.sz, { noHead: true });
      va.code = valFmt(va, { groups: radars });
      va.ticks = valTicks(va, va.sc, va.code, null);
      const rOf = x => clamp(vT(va, x), 0, 1.05) * R;
      const PT2 = (i, rr) => { const a = (i / n) * Math.PI * 2; return [cxp + Math.sin(a) * rr, cyp - Math.cos(a) * rr]; };
      const gl = va.grid ? lineFrom(kid(va.grid, 'spPr'), { color: tx1, w: 1 }) : null;
      if (gl && !gl.none) for (const k of va.ticks) { const rr = vT(va, k.v) * R; if (rr <= 0) continue; const q = []; for (let i = 0; i < n; i++) q.push(PT2(i, rr)); path(gGrid, polyD(q) + 'Z', { none: true }, gl); }
      const sl = ca ? lineFrom(ca.grid ? kid(ca.grid, 'spPr') : ca.spPr, { color: tx1, w: 1 }) : { color: tx1, w: 1 };
      if (sl && !sl.none) for (let i = 0; i < n; i++) { const q = PT2(i, R); line(gGrid, cxp, cyp, q[0], q[1], sl); }
      for (const g of radars) for (const s of g.series) {
        const lk = look(g, s, -1);
        const q = [];
        for (let i = 0; i < n; i++) { const x = s.v[i]; q.push(PT2(i, rOf(x == null ? va.sc.min : x))); }
        if (g.radarStyle === 'filled') path(gSer, polyD(q) + 'Z', lk.fill, lk.line);
        else if (lk.line && !lk.line.none) path(gSer, polyD(q) + 'Z', { none: true }, lk.line);
        const cfg = dl(g, s);
        q.forEach((xy, i) => { drawMarker(lk.mk, xy[0], xy[1]); const L = dlText(cfg, g, s, i, s.v[i]); if (L) placeNear(L, xy, L.c.pos || 't'); });
      }
      if (showC) for (let i = 0; i < n; i++) {
        const a = (i / n) * Math.PI * 2, q = PT2(i, R + pt(2));
        const B = plain(labelsR[i], cs);
        drawBlock(gLbl, B, q[0] + Math.sin(a) * B.w / 2, q[1] - Math.cos(a) * B.h / 2);
      }
      if (!va.del && va.lblPos !== 'none') {
        const al = lineFrom(va.spPr, { color: tx1, w: 1 });
        if (al && !al.none) line(gAxis, cxp, cyp, cxp, cyp - R, al);
        for (const k of va.ticks) { const rr = vT(va, k.v) * R; const B = plain(k.t, vs); drawBlock(gLbl, B, cxp - pt(0.88 * vs.sz) - B.w / 2, cyp - rr); }
      }
    }

    if (groups.some(g => g.type === 'surface')) {
      const B = plain('等高線のグラフは表示できません', { ...base, color: GRAY });
      drawBlock(gLbl, B, P.x + P.w / 2, P.y + P.h / 2);
    }

    // 値の文字（枠は、文字のまわりに 横 2.95pt・縦 1.6pt）
    for (const L of labels) {
      const B = L.B;
      const sp = L.c.spPr;
      if (sp) {
        const f = fillFrom(sp, { none: true }), l = lineFrom(sp, null);
        if ((f && !f.none) || (l && !l.none)) paint(add(gLbl, 'rect', { x: r2(L.x - B.w / 2 - pt(2.95)), y: r2(L.y - B.h / 2 - pt(1.6)), width: r2(B.w + pt(5.9)), height: r2(B.h + pt(3.2)) }), f, l);
      }
      drawBlock(gLbl, B, L.x, L.y);
    }

    // 題名
    if (TB && titleBox) {
      if (TB.spPr) paint(add(root2, 'rect', { x: r2(titleBox.x), y: r2(titleBox.y), width: r2(titleBox.w), height: r2(titleBox.h) }), fillFrom(TB.spPr, { none: true }), lineFrom(TB.spPr, null));
      drawBlock(root2, TB, titleBox.x + titleBox.w / 2, titleBox.y + titleBox.h / 2);
    }
    // 凡例：印（四角は 文字 × 0.49 の正方形、線は 文字 × 2.33 の長さ）＋ 名前
    if (legend && legBox && LG) {
      const gL = add(root2, 'g');
      if (legend.spPr) paint(add(gL, 'rect', { x: r2(legBox.x), y: r2(legBox.y), width: r2(legBox.w), height: r2(legBox.h) }), fillFrom(legend.spPr, { none: true }), lineFrom(legend.spPr, null));
      const sz = legend.s.sz;
      let y = legBox.y + (legBox.h - LG.rows.length * LG.rowH) / 2;
      for (const row of LG.rows) {
        const rw = row.reduce((a, i, k) => a + LG.ws[i] + (k ? LG.sep : 0), 0);
        let x = LG.vertical ? legBox.x + LG.padX[0] : legBox.x + (legBox.w - rw) / 2;
        for (const i of row) {
          const it = legend.items[i];
          const cy = y + LG.rowH / 2;
          const lk = it.kind === 'trend' ? { line: lineFrom(kid(it.tr, 'spPr'), { color: look(it.g, it.s, -1).auto, w: 1.5, dash: 'sysDot' }) } : look(it.g, it.s, it.i);
          if (it.kind === 'rect') {
            const kh = pt(0.49 * sz), kw = LG.mixed ? LG.keyW - pt(0.27 * sz) : kh;
            paint(add(gL, 'rect', { x: r2(x + (LG.keyW - kw) / 2), y: r2(cy - kh / 2), width: r2(kw), height: r2(kh) }), lk.fill, lk.line && !lk.line.none ? { ...lk.line, w: Math.min(lk.line.w || 1, 2) } : null);
          } else {
            const cap = lk.line && lk.line.cap === 'round' ? (lk.line.w || 1) / 2 : 0;
            if (lk.line && !lk.line.none) line(gL, x + cap, cy, x + LG.keyW - cap, cy, lk.line);
            if (lk.mk) path(gL, markD(lk.mk.sym, x + LG.keyW / 2, cy, lk.mk.r), lk.mk.fill, lk.mk.line);
          }
          const B = plain(it.text, it.st);
          drawBlock(gL, B, x + LG.keyW + LG.gapK + B.w / 2, cy);
          x += LG.ws[i] + LG.sep;
        }
        y += LG.rowH;
      }
    }

    // テスト用：場所と目盛り
    svg.oc = { P, title: titleBox, legend: legBox, axes: [...axes.values()].map(a => ({ id: a.id, kind: a.kind, pos: a.pos, min: a.sc && a.sc.min, max: a.sc && a.sc.max, major: a.sc && a.sc.major })) };
    return svg;
  }

  // グラフの XML の場所：graphicData の c:chart の r:id から
  function chartPathOf(gd, rels) {
    const c = kid(gd, 'chart');
    const id = c && DML.rAttr(c, 'id'), r = id && rels && rels.get(id);
    return r ? r.target : null;
  }

  return { render, chartPathOf };
})();
