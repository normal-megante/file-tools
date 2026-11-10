'use strict';
// =========================================================
// PowerPoint（.pptx）を見る部品（自作）。js/zip.js・js/drawingml.js を先に読む
//   const deck = await Pptx.open(zip)   … zip は Zip.open の結果
//   deck.width / deck.height            … スライドの大きさ（px、96 dpi。16:9 なら 1280×720）
//   deck.count                          … スライドの枚数
//   const s = await deck.slide(i)       … { el: スライドの要素（width×height px）, notes: ノートの文, hidden: 非表示のスライドか }
//   deck.dispose()                      … 画像の URL を片付ける
// しくみ：
//   ・中身は ZIP に入った XML。スライド（slideN.xml）→ レイアウト → マスター → テーマ と、足りない情報
//     （位置・文字の大きさ・色・箇条書きの記号）を順に引き継ぐ（プレースホルダー＝「タイトルを入力」の枠）
//   ・図形は SVG、文字は HTML で、スライドの大きさのまま描く（見るときに全体を縮める）
//   ・できないもの：EMF・WMF・TIFF の図、アニメーション、影・3D などの効果。グラフは js/chart.js
// =========================================================
const Pptx = (() => {
  const { EMU, px, PT, parseXml, kid, kids, pathOf, at, atNum, truthy, rAttr, desc, h, SVGNS, sv, uid, resolvePath,
    relsPath, PRESET, COLOR_TAGS, hex, rgb2hsl, hsl2rgb, clamp, s2lin, lin2s, tintColor, colorEl, color, mods, css, mix,
    FILL_TAGS, fillEl, fillOf, lineOf, mergeLine, xfrmOf, placeBox, adjOf, poly, ellD, rrectD, starD, preset, custom,
    geometry, presetTextRect, textRect, EMPTY_R, EMPTY_P, mergeRPr, mergePPr, levelProps, bodyPrChain, SERIF,
    fontFamily, WINGDINGS, bulletChar, roman, autoNum, drawText, spacing, runCss, drawPara, DASH, shapeSvg, blipImage,
  } = DML;

  /* ---------- スライドを描く ---------- */
  async function open(zip) {
    const pk = DML.pkg(zip), { xml, rels, relOfType, recolor, media, theme } = pk;

    const root = await rels('');
    const main = relOfType(root, 'officeDocument');
    const presPath = main ? main.target : 'ppt/presentation.xml';
    const pres = await xml(presPath);
    if (!pres) throw new Error('PowerPoint のファイルとして読めませんでした');
    const P = pres.documentElement, presRels = await rels(presPath);
    const sz = kid(P, 'sldSz');
    const W = px(atNum(sz, 'cx', 9144000)), H = px(atNum(sz, 'cy', 6858000));
    const slidePaths = kids(kid(P, 'sldIdLst'), 'sldId').map(s => presRels.get(rAttr(s, 'id'))).filter(Boolean).map(r => r.target);
    const defaultTextStyle = kid(P, 'defaultTextStyle');
    const tsRel = relOfType(presRels, 'tableStyles');
    const tableStyles = tsRel ? await xml(tsRel.target) : null;

    // マスター・レイアウト・テーマ（何枚ものスライドで同じものを使うので、1回だけ読む）
    const partCache = new Map();
    function part(path) {
      if (!partCache.has(path)) partCache.set(path, (async () => {
        const doc = await xml(path), r = await rels(path);
        const el = doc ? doc.documentElement : null;
        const phs = [];
        for (const sp of kids(pathOf(el, 'cSld', 'spTree'))) {
          const ph = phOf(sp);
          if (ph) phs.push({ ...ph, el: sp });
        }
        return { path, el, rels: r, phs };
      })());
      return partCache.get(path);
    }

    function phOf(sp) {
      const nv = kids(sp).find(c => /^nv/.test(c.localName));
      const ph = nv && pathOf(nv, 'nvPr', 'ph');
      if (!ph) return null;
      return { type: at(ph, 'type') || 'obj', idx: at(ph, 'idx'), hasIdx: at(ph, 'idx') != null };
    }
    const masterType = t => (t === 'ctrTitle' ? 'title' : /^(subTitle|obj|body|pic|tbl|chart|clipArt|dgm|media)$/.test(t) ? 'body' : t);
    function matchPh(list, ph, byIdx) {
      if (!list) return null;
      if (byIdx && ph.hasIdx) { const m = list.find(x => x.idx === ph.idx); if (m) return m; }
      return list.find(x => x.type === ph.type) || list.find(x => masterType(x.type) === masterType(ph.type)) || null;
    }

    // 図形の木（spTree）を描く。src：{ path, rels }（画像などの場所の基準）、skipPh：プレースホルダーは描かない（マスター・レイアウト）
    async function drawTree(tree, ctx, parent, skipPh) {
      for (let el of kids(tree)) {
        if (el.localName === 'AlternateContent') el = kids(kid(el, 'Fallback'))[0] || kids(kid(el, 'Choice'))[0];
        if (!el) continue;
        try {
          const n = el.localName;
          if (skipPh && (n === 'sp' || n === 'pic') && phOf(el)) continue;
          if (n === 'sp' || n === 'cxnSp') await drawSp(el, ctx, parent);
          else if (n === 'pic') await drawPic(el, ctx, parent);
          else if (n === 'grpSp') await drawGroup(el, ctx, parent, skipPh);
          else if (n === 'graphicFrame') await drawFrame(el, ctx, parent);
        } catch (e) { console.warn('図形を描けませんでした', e); }
      }
    }
    // プレースホルダーの引き継ぎ（スライド → レイアウト → マスター）
    function inherit(sp, ctx) {
      const ph = phOf(sp);
      if (!ph || !ctx.layout) return { ph: null, list: [] };
      const lay = matchPh(ctx.layout.phs, ph, true);
      const mas = matchPh(ctx.master && ctx.master.phs, lay || ph, false);
      return { ph, list: [lay, mas].filter(Boolean).map(x => x.el) };
    }
    function styleOf(sp, ctx) {
      const st = kid(sp, 'style'), out = {};
      if (!st) return out;
      const T = ctx.theme, lr = kid(st, 'lnRef'), fr = kid(st, 'fillRef'), fo = kid(st, 'fontRef');
      if (lr) { const i = atNum(lr, 'idx', 0); if (i > 0 && T.lines[i - 1]) out.line = lineOf(T.lines[i - 1], { ...ctx, ph: color(lr, ctx) }); }
      if (fr) {
        const i = atNum(fr, 'idx', 0), f = i >= 1001 ? T.bgFills[i - 1001] : T.fills[i - 1];
        if (i > 0) out.fill = f ? fillEl(f, { ...ctx, ph: color(fr, ctx) }) : { color: color(fr, ctx) };
        if (out.fill && out.fill.blip) out.fill = { color: color(fr, ctx) };
      }
      if (fo) out.font = { font: at(fo, 'idx') === 'major' ? 'major' : 'minor', color: colorEl(fo) ? fo : null };
      return out;
    }
    function textChain(sp, inh, ctx, extra) {
      const own = kid(kid(sp, 'txBody'), 'lstStyle');
      const style = extra && (extra.color || extra.font || extra.bold != null) ? extra : null;
      if (!inh.ph) return [defaultTextStyle, style, own];
      const t = inh.ph.type, ts = ctx.master && pathOf(ctx.master.el, 'txStyles');
      const base = ts && kid(ts, /^(title|ctrTitle)$/.test(t) ? 'titleStyle' : /^(body|subTitle|obj|pic|tbl|chart|clipArt|dgm|media)$/.test(t) ? 'bodyStyle' : 'otherStyle');
      const lst = inh.list.map(e => kid(kid(e, 'txBody'), 'lstStyle')).reverse();
      return [defaultTextStyle, base, ...lst, style, own];
    }
    async function drawSp(sp, ctx, parent) {
      const nv = kids(sp).find(c => /^nv/.test(c.localName));
      if (truthy(at(kid(nv, 'cNvPr'), 'hidden'))) return;
      const inh = inherit(sp, ctx);
      const spPr = kid(sp, 'spPr');
      const xf = xfrmOf(spPr) || inh.list.map(e => xfrmOf(kid(e, 'spPr'))).find(Boolean);
      if (!xf) return;
      const st = styleOf(sp, ctx);
      const box = placeBox(xf);
      const fill = fillOf(spPr, ctx) || inh.list.map(e => fillOf(kid(e, 'spPr'), ctx)).find(Boolean) || st.fill || null;
      const line = mergeLine(st.line, ...inh.list.map(e => lineOf(kid(kid(e, 'spPr'), 'ln'), ctx)).reverse(), lineOf(kid(spPr, 'ln'), ctx));
      const paths = geometry(spPr, xf.w, xf.h) || inh.list.map(e => geometry(kid(e, 'spPr'), xf.w, xf.h)).find(Boolean)
        || (sp.localName === 'cxnSp' ? preset('line', xf.w, xf.h, {}) : preset('rect', xf.w, xf.h, {}));
      if ((fill && !fill.none) || (line && !line.none && line.color)) box.appendChild(await shapeSvg(paths, fill, line, xf.w, xf.h, { h: xf.flipH, v: xf.flipV }, ctx));
      const tx = kid(sp, 'txBody');
      if (tx) {
        const extra = st.font ? { font: st.font.font, color: st.font.color } : null;
        const bodies = inh.list.map(e => kid(kid(e, 'txBody'), 'bodyPr')).reverse();
        let tbox = textRect(spPr, xf) || inh.list.map(e => textRect(kid(e, 'spPr'), xf)).find(Boolean) || { x: 0, y: 0, w: xf.w, h: xf.h };
        const txX = kid(sp, 'txXfrm');
        if (txX) {
          const o = kid(txX, 'off'), e = kid(txX, 'ext');
          if (o && e) tbox = { x: px(at(o, 'x')) - xf.x, y: px(at(o, 'y')) - xf.y, w: px(at(e, 'cx')), h: px(at(e, 'cy')) };
        }
        const t = drawText(tx, textChain(sp, inh, ctx, extra), bodies, ctx, tbox);
        if (t) box.appendChild(t);
      }
      if (box.childNodes.length) parent.appendChild(box);
    }
    async function drawPic(pic, ctx, parent, frameXf) {
      const nv = kid(pic, 'nvPicPr');
      if (truthy(at(kid(nv, 'cNvPr'), 'hidden'))) return;
      const inh = inherit(pic, ctx), spPr = kid(pic, 'spPr');
      const xf = xfrmOf(spPr) || frameXf || inh.list.map(e => xfrmOf(kid(e, 'spPr'))).find(Boolean);
      if (!xf) return;
      const box = placeBox(xf);
      box.classList.add('pp-pic');
      const bf = kid(pic, 'blipFill'), m = await blipImage(bf, ctx.src, ctx);
      const paths = geometry(spPr, xf.w, xf.h);
      const prst = at(kid(spPr, 'prstGeom'), 'prst');
      if (m && m.url) {
        const img = h('img', null);
        img.alt = ''; img.draggable = false; img.decoding = 'async';
        img.src = m.url;
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
      } else {
        box.appendChild(missing(m && m.unsupported ? `この図は表示できません（${m.unsupported} 形式）` : '図を読めませんでした'));
      }
      const line = mergeLine(styleOf(pic, ctx).line, lineOf(kid(spPr, 'ln'), ctx));
      if (line && !line.none && line.color) box.appendChild(await shapeSvg(paths || preset('rect', xf.w, xf.h, {}), { none: true }, line, xf.w, xf.h, {}, ctx));
      parent.appendChild(box);
    }
    function missing(text) {
      const d = h('div', 'pp-missing');
      d.textContent = text;
      return d;
    }
    async function drawGroup(g, ctx, parent, skipPh) {
      const gp = kid(g, 'grpSpPr'), xf = xfrmOf(gp);
      if (truthy(at(pathOf(g, 'nvGrpSpPr', 'cNvPr'), 'hidden'))) return;
      const c2 = { ...ctx, grpFill: fillOf(gp, ctx) || ctx.grpFill };
      if (!xf) { await drawTree(g, c2, parent, skipPh); return; }
      const box = placeBox(xf, 'pp-grp');
      const fl = [xf.flipH ? 'scaleX(-1)' : '', xf.flipV ? 'scaleY(-1)' : ''].join(' ').trim();
      if (fl) box.style.transform = (box.style.transform + ' ' + fl).trim();
      const cw = xf.cw || xf.w, ch = xf.ch || xf.h;
      const inner = h('div', 'pp-grp', { left: '0', top: '0', width: cw + 'px', height: ch + 'px', transformOrigin: '0 0', transform: `scale(${cw ? xf.w / cw : 1},${ch ? xf.h / ch : 1}) translate(${-(xf.cx || 0)}px,${-(xf.cy || 0)}px)` });
      box.appendChild(inner);
      await drawTree(g, c2, inner, skipPh);
      parent.appendChild(box);
    }
    async function drawFrame(gf, ctx, parent) {
      if (truthy(at(pathOf(gf, 'nvGraphicFramePr', 'cNvPr'), 'hidden'))) return;
      const inh = inherit(gf, ctx);
      const xf = xfrmOf(gf) || inh.list.map(e => xfrmOf(kid(e, 'spPr'))).find(Boolean);
      if (!xf) return;
      const gd = pathOf(gf, 'graphic', 'graphicData'), uri = at(gd, 'uri') || '';
      const tbl = kid(gd, 'tbl');
      if (tbl) return drawTable(tbl, xf, ctx, parent);
      if (/\/chart$/.test(uri)) return drawChartBox(gd, xf, ctx, parent);
      if (/\/diagram$/.test(uri)) return drawDiagram(gd, xf, ctx, parent);
      const pic = desc(gd, 'pic')[0];
      if (pic) return drawPic(pic, ctx, parent, xf);
      const box = placeBox(xf);
      box.appendChild(missing('この部分は表示できません（埋め込まれたもの）'));
      parent.appendChild(box);
    }
    async function drawChartBox(gd, xf, ctx, parent) {
      const box = placeBox(xf);
      const cp = OChart.chartPathOf(gd, ctx.src.rels);
      try {
        if (!cp) throw new Error('no chart');
        box.appendChild(await OChart.render(cp, { P: { xml }, theme: ctx.theme, map: ctx.map, w: xf.w, h: xf.h, host: 'pp' }));
      } catch (e) {
        console.warn('グラフを描けませんでした', e);
        box.appendChild(missing('グラフ（表示できませんでした）'));
      }
      parent.appendChild(box);
    }
    // SmartArt：PowerPoint が一緒に保存している「描いたもの」（drawingN.xml）を図形として描く
    async function drawDiagram(gd, xf, ctx, parent) {
      const ids = kid(gd, 'relIds'), dm = rAttr(ids, 'dm'), dataRel = dm && ctx.src.rels.get(dm);
      const data = dataRel && await xml(dataRel.target);
      let drawingPath = null;
      if (data) {
        const ext = desc(data.documentElement, 'dataModelExt')[0], rid = at(ext, 'relId');
        const r = rid && ctx.src.rels.get(rid);
        if (r) drawingPath = r.target;
      }
      const dr = drawingPath && await xml(drawingPath);
      const box = placeBox(xf, 'pp-grp');
      if (dr) {
        const r = await rels(drawingPath);
        await drawTree(pathOf(dr.documentElement, 'spTree'), { ...ctx, src: { path: drawingPath, rels: r }, layout: null }, box, false);
      } else {
        const texts = data ? [...desc(data.documentElement, 't')].map(t => t.textContent).filter(Boolean) : [];
        box.appendChild(missing('SmartArt（この画面では表示できません）' + (texts.length ? '：' + texts.join(' / ') : '')));
      }
      parent.appendChild(box);
    }
    // 表
    function tableStyle(id, ctx) {
      const el = tableStyles && kids(tableStyles.documentElement, 'tblStyle').find(s => at(s, 'styleId') === id);
      const parts = {};
      if (el) {
        for (const p of kids(el)) {
          const o = {};
          const tx = kid(p, 'tcTxStyle');
          if (tx) { if (at(tx, 'b')) o.bold = at(tx, 'b') === 'on'; const c = color(tx, ctx) || (kid(tx, 'fontRef') && color(kid(tx, 'fontRef'), ctx)); if (c) o.color = c; }
          const tc = kid(p, 'tcStyle');
          if (tc) {
            const f = kid(tc, 'fill'), fr = kid(tc, 'fillRef');
            if (f) o.fill = fillOf(f, ctx); else if (fr) o.fill = { color: color(fr, ctx) };
            const bd = kid(tc, 'tcBdr');
            if (bd) { o.bd = {}; for (const s of kids(bd)) { const ln = kid(s, 'ln'), lr = kid(s, 'lnRef'); o.bd[s.localName] = ln ? lineOf(ln, ctx) : lr ? { color: color(lr, ctx), w: 1 } : null; } }
          }
          parts[p.localName] = o;
        }
        return parts;
      }
      // ファイルの中に定義がないとき：PowerPoint の標準（中間スタイル 2 - アクセント 1）に近いもの
      const acc = hex(ctx.theme.colors.accent1 || '4472C4'), white = { r: 255, g: 255, b: 255, a: 1 }, wl = { color: white, w: 1.33 };
      const tint = v => tintColor(acc, v);
      const all = { left: wl, right: wl, top: wl, bottom: wl, insideH: wl, insideV: wl };
      return {
        wholeTbl: { fill: { color: tint(0.2) }, bd: all }, band1H: { fill: { color: tint(0.4) } }, band1V: { fill: { color: tint(0.4) } },
        firstRow: { fill: { color: acc }, bold: true, color: white, bd: { bottom: { color: white, w: 4 } } }, lastRow: { fill: { color: acc }, bold: true, color: white, bd: { top: { color: white, w: 4 } } },
        firstCol: { fill: { color: acc }, bold: true, color: white }, lastCol: { fill: { color: acc }, bold: true, color: white },
      };
    }
    async function drawTable(tbl, xf, ctx, parent) {
      const box = placeBox(xf);
      const tp = kid(tbl, 'tblPr'), flag = k => truthy(at(tp, k));
      const sid = (kid(tp, 'tableStyleId') || {}).textContent;
      const style = sid || tableStyles ? tableStyle((sid || '').trim(), ctx) : {};
      const cols = kids(kid(tbl, 'tblGrid'), 'gridCol').map(c => px(at(c, 'w')));
      const rows = kids(tbl, 'tr'), nR = rows.length, nC = cols.length;
      const t = h('table', 'pp-tbl', { width: cols.reduce((a, b) => a + b, 0) + 'px' });
      const cg = h('colgroup');
      for (const w of cols) cg.appendChild(h('col', null, { width: w + 'px' }));
      t.appendChild(cg);
      const first = flag('firstRow'), last = flag('lastRow');
      for (let r = 0; r < nR; r++) {
        const tr = h('tr', null, { height: px(at(rows[r], 'h')) + 'px' });
        let c = 0;
        for (const tc of kids(rows[r], 'tc')) {
          const ci = c;
          c += Math.max(1, atNum(tc, 'gridSpan', 1));
          if (truthy(at(tc, 'hMerge')) || truthy(at(tc, 'vMerge'))) continue;
          // 表のスタイルの重なり（弱い順）
          const parts = ['wholeTbl'];
          const br = r - (first ? 1 : 0);
          if (flag('bandCol')) parts.push(ci % 2 ? 'band2V' : 'band1V');
          if (flag('bandRow') && br >= 0 && !(last && r === nR - 1)) parts.push(br % 2 ? 'band2H' : 'band1H');
          if (flag('lastCol') && ci === nC - 1) parts.push('lastCol');
          if (flag('firstCol') && ci === 0) parts.push('firstCol');
          if (last && r === nR - 1) parts.push('lastRow');
          if (first && r === 0) parts.push('firstRow');
          let fill = null, bold = null, tcolor = null;
          const bd = { left: null, right: null, top: null, bottom: null };
          const span = Math.max(1, atNum(tc, 'gridSpan', 1)), rspan = Math.max(1, atNum(tc, 'rowSpan', 1));
          for (const pn of parts) {
            const p = style[pn];
            if (!p) continue;
            if (p.fill) fill = p.fill;
            if (p.bold != null) bold = p.bold;
            if (p.color) tcolor = p.color;
            if (p.bd) {
              const b = p.bd;
              if (ci === 0 ? b.left : b.insideV) bd.left = ci === 0 ? b.left : b.insideV;
              if (ci + span >= nC ? b.right : b.insideV) bd.right = ci + span >= nC ? b.right : b.insideV;
              if (r === 0 ? b.top : b.insideH) bd.top = r === 0 ? b.top : b.insideH;
              if (r + rspan >= nR ? b.bottom : b.insideH) bd.bottom = r + rspan >= nR ? b.bottom : b.insideH;
              if (pn !== 'wholeTbl') { if (b.top) bd.top = b.top; if (b.bottom) bd.bottom = b.bottom; if (b.left) bd.left = b.left; if (b.right) bd.right = b.right; }
            }
          }
          const tcPr = kid(tc, 'tcPr');
          const own = fillOf(tcPr, ctx);
          if (own) fill = own;
          for (const [k, n] of [['left', 'lnL'], ['right', 'lnR'], ['top', 'lnT'], ['bottom', 'lnB']]) { const l = lineOf(kid(tcPr, n), ctx); if (l) bd[k] = { ...(bd[k] || {}), ...l }; }
          const td = h('td');
          if (span > 1) td.colSpan = span;
          if (rspan > 1) td.rowSpan = rspan;
          if (fill && fill.color) td.style.background = css(fill.color);
          else if (fill && fill.stops) td.style.background = css(fill.color);
          for (const k in bd) { const l = bd[k]; td.style['border' + k[0].toUpperCase() + k.slice(1)] = l && !l.none && l.color ? `${Math.max(0.75, l.w || 1)}px solid ${css(l.color)}` : '0'; }
          const ins = (k, d) => px(at(tcPr, k) != null ? at(tcPr, k) : d);
          td.style.padding = `${ins('marT', 45720)}px ${ins('marR', 91440)}px ${ins('marB', 45720)}px ${ins('marL', 91440)}px`;
          const an = at(tcPr, 'anchor');
          td.style.verticalAlign = an === 'ctr' ? 'middle' : an === 'b' ? 'bottom' : 'top';
          const tx = kid(tc, 'txBody');
          if (tx) {
            const chain = [defaultTextStyle, tcolor ? { rgb: tcolor } : null, bold != null ? { bold } : null, kid(tx, 'lstStyle')];
            const counters = [];
            for (const p of kids(tx, 'p')) {
              const pPr = kid(p, 'pPr'), lvl = clamp(atNum(pPr, 'lvl', 0), 0, 8);
              td.appendChild(drawPara(p, mergePPr(levelProps(chain, lvl), pPr), ctx, { fontScale: 1, lnRed: 0, counters, lvl }));
            }
          }
          tr.appendChild(td);
        }
        t.appendChild(tr);
      }
      box.appendChild(t);
      parent.appendChild(box);
    }

    function bgOf(el, ctx) {
      const bg = pathOf(el, 'cSld', 'bg');
      if (!bg) return null;
      const pr = kid(bg, 'bgPr');
      if (pr) return fillOf(pr, ctx);
      const ref = kid(bg, 'bgRef');
      if (ref) {
        const i = atNum(ref, 'idx', 0), c = color(ref, ctx), f = i >= 1001 ? ctx.theme.bgFills[i - 1001] : ctx.theme.fills[i - 1];
        // テーマの背景の画像は、テーマのファイルからの場所で探す
        const r = f ? fillEl(f, { ...ctx, ph: c, src: f.localName === 'blipFill' ? ctx.theme.src : ctx.src }) : null;
        return r || { color: c };
      }
      return null;
    }
    async function paintBg(slideEl, fill, ctx) {
      if (!fill || fill.none) return;
      if (fill.color && !fill.stops) slideEl.style.background = css(fill.color);
      else if (fill.stops) {
        const list = fill.stops.map(s => `${css(s.c)} ${s.pos * 100}%`).join(',');
        slideEl.style.background = fill.radial ? `radial-gradient(ellipse at center,${list})` : `linear-gradient(${fill.ang + 90}deg,${list})`;
      } else if (fill.blip) {
        const m = await blipImage(fill.blip, fill.src || ctx.src, ctx, fill.ph);
        if (m && m.url) {
          slideEl.style.backgroundImage = `url("${m.url}")`;
          const tile = kid(fill.blip, 'tile');
          slideEl.style.backgroundSize = tile ? 'auto' : '100% 100%';
          slideEl.style.backgroundRepeat = tile ? 'repeat' : 'no-repeat';
        }
      }
    }
    const mapOf = (masterEl, ...ovr) => {
      const m = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
      const cm = kid(masterEl, 'clrMap');
      if (cm) for (const x of cm.attributes) m[x.name] = x.value;
      for (const o of ovr) { const e = pathOf(o, 'clrMapOvr', 'overrideClrMapping'); if (e) for (const x of e.attributes) m[x.name] = x.value; }
      return m;
    };

    async function slide(i) {
      const path = slidePaths[i];
      const s = await part(path);
      const el = h('div', 'pp-slide', { width: W + 'px', height: H + 'px' });
      if (!s.el) { el.appendChild(missing('このスライドを読めませんでした')); return { el, notes: '', hidden: false }; }
      const layRel = relOfType(s.rels, 'slideLayout'), layout = layRel ? await part(layRel.target) : null;
      const masRel = layout && relOfType(layout.rels, 'slideMaster'), master = masRel ? await part(masRel.target) : null;
      const thRel = master && relOfType(master.rels, 'theme'), th = await theme(thRel ? thRel.target : '');
      const map = mapOf(master && master.el, layout && layout.el, s.el);
      const base = { theme: th, map, layout, master, slideNo: i + 1, media, recolor };
      const ctxOf = p => ({ ...base, src: { path: p.path, rels: p.rels } });
      // 背景（スライド → レイアウト → マスターの順に、最初に見つかったもの）
      const bg = bgOf(s.el, ctxOf(s)) || (layout && bgOf(layout.el, ctxOf(layout))) || (master && bgOf(master.el, ctxOf(master)));
      el.style.background = '#fff';
      await paintBg(el, bg, ctxOf(bg && bg.src && layout && bg.src.path === layout.path ? layout : s));
      // マスター・レイアウトの飾り（プレースホルダー以外）→ スライドの図形
      const showMaster = at(s.el, 'showMasterSp') !== '0' && at(s.el, 'showMasterSp') !== 'false';
      if (showMaster && master && layout && at(layout.el, 'showMasterSp') !== '0') await drawTree(pathOf(master.el, 'cSld', 'spTree'), { ...ctxOf(master), layout: null }, el, true);
      if (showMaster && layout) await drawTree(pathOf(layout.el, 'cSld', 'spTree'), { ...ctxOf(layout), layout: null }, el, true);
      await drawTree(pathOf(s.el, 'cSld', 'spTree'), ctxOf(s), el, false);
      // ノート
      let notes = '';
      const nRel = relOfType(s.rels, 'notesSlide');
      if (nRel) {
        const nd = await xml(nRel.target);
        const body = nd && kids(pathOf(nd.documentElement, 'cSld', 'spTree'), 'sp').find(sp => { const ph = phOf(sp); return ph && ph.type === 'body'; });
        if (body) notes = kids(kid(body, 'txBody'), 'p').map(p => [...desc(p, 't')].map(t => t.textContent).join('')).join('\n').trim();
      }
      return { el, notes, hidden: at(s.el, 'show') === '0' };
    }

    return {
      width: W, height: H, count: slidePaths.length,
      async slide(i) {
        try { return await slide(i); } catch (e) {
          console.error(e);
          const el = h('div', 'pp-slide', { width: W + 'px', height: H + 'px', background: '#fff' });
          el.appendChild(missing('このスライドを描けませんでした'));
          return { el, notes: '', hidden: false, error: e };
        }
      },
      dispose() { pk.dispose(); },
    };
  }

  // 見た目（スライドの中だけで使う）
  if (typeof document !== 'undefined' && !document.getElementById('pptx-css')) {
    const st = document.createElement('style');
    st.id = 'pptx-css';
    st.textContent = `
.pp-slide{position:relative;overflow:hidden;background:#fff;color:#000;transform-origin:0 0;font-family:sans-serif;line-height:1.2;text-align:left;-webkit-text-size-adjust:none;text-size-adjust:none}
.pp-slide *{box-sizing:border-box}
.pp-tbl{border-collapse:collapse;table-layout:fixed;color:#000}
.pp-tbl td{overflow:hidden;white-space:pre-wrap;overflow-wrap:break-word;line-height:1.2}`;
    document.head.appendChild(st);
  }

  return { open };
})();
