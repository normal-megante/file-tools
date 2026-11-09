'use strict';
// =========================================================
// Word（.docx）を見る部品（自作）。js/zip.js・js/drawingml.js を先に読む
//   const doc = await Docx.open(zip)          … zip は Zip.open の結果
//   doc.info                                  … { pageW, pageH（px・1ページ目）, sections, paper }
//   const pages = await doc.paginate(opts)    … 紙と同じページに分ける → [{ el, w, h }]（el は w×h px）
//   const el = await doc.flow(width, opts)    … スマホ向け：画面の幅で折り返す（ページに分けない）
//   Docx.layoutTabs(root)                     … flow の要素を画面に入れたあとで、タブの幅を決める
//   doc.dispose()                             … 画像の URL を片付ける
//   opts.alive()：false を返したら途中でやめる。opts.progress(text)：進み具合
// しくみ：
//   ・document.xml の段落・表を、styles.xml（スタイル）・numbering.xml（箇条書きの番号）の書式を重ねて HTML にする
//   ・行の高さは Word と同じ計算（書体ごとの「1行」の高さ・行グリッド）。書体が違っても行数が合いやすいように
//   ・ページ分け：見えない所に並べて高さをはかり、行の切れ目で切って紙に分ける（切った段落は上下をかくして2回置く）
//   ・できないもの：段組み（1段で表示）、縦書き（横書きで表示）、グラフ（枠だけ）、EMF・WMF の図、変更履歴の色分け、コメント
// =========================================================
const Docx = (() => {
  const { px, PT, kid, kids, pathOf, truthy, rAttr, h, color, css, fillOf, fillEl, lineOf, mergeLine,
    geometry, textRect, shapeSvg, blipImage, SERIF, WINGDINGS } = DML;
  // Word の属性は「w:val」のように接頭辞つき（図の部分は接頭辞なし）。どちらでも名前だけで読む
  const WNS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main', WNS2 = 'http://purl.oclc.org/ooxml/wordprocessingml/main';
  const at = (el, n) => { if (!el) return null; let v = el.getAttributeNS(WNS, n); if (v == null) v = el.getAttribute(n); if (v == null) v = el.getAttributeNS(WNS2, n); return v; };
  const atNum = (el, n, d) => { const v = at(el, n); return v == null || v === '' || isNaN(+v) ? d : +v; };
  const TW = 15;                     // 1px が何 twip か（1440 twip = 1 インチ = 96px）
  const tw = v => (+v || 0) / TW;    // twip → px
  const hp = v => (+v || 0) / 2;     // 半ポイント → pt
  const onOff = el => { if (!el) return null; const v = at(el, 'val'); return v == null || truthy(v) || v === 'On'; };
  const W_ = el => (el ? at(el, 'val') : null);
  // 和文と欧文の間の 1/4 字のすき間（Word の autoSpaceDE・autoSpaceDN。この PC の Word ではかった）。
  // E＝かな・漢字・半角カナ、L＝文字（ラテン・ギリシャ・キリル…。全角の英字は入らない）、D＝半角数字。記号・ー・々・全角の英数字の横はあけない
  const AS_E = /[ぁ-ゖゝゞァ-ヺヽヾ㐀-䶿一-鿿豈-﫿ｦ-ｯｱ-ﾝ\u{20000}-\u{3134f}]/u;
  const AS_WIDE = /[ᄀ-ᇿ⺀-㏿㐀-鿿가-힯豈-﫿︰-﹏＀-￯\u{20000}-\u{3ffff}]/u;
  const asKind = c => !c ? '' : AS_E.test(c) ? 'E' : c >= '0' && c <= '9' ? 'D' : /\p{L}/u.test(c) && !AS_WIDE.test(c) ? 'L' : '';
  function asGap(a, b, P) {
    const x = asKind(a), y = asKind(b);
    if (!x || !y || x === y || (x !== 'E' && y !== 'E')) return false;
    return (x === 'D' || y === 'D') ? P.autoSpaceN !== false : P.autoSpace !== false;
  }
  // 画面に並べたときの文字の幅（px、1文字あたり）。canvas の measureText は ＭＳ 明朝などで画面と違う幅を返すので、本物の要素ではかる
  const widthCache = new Map();
  function domWidth(family, px, ch, weight) {
    const k = family + '|' + px + '|' + ch + '|' + (weight || '');
    if (!widthCache.has(k)) {
      let w = ch === ' ' ? px / 4 : px;
      if (document.body) {
        const s = document.createElement('span');
        s.style.cssText = `position:absolute;left:-9999px;top:0;visibility:hidden;white-space:pre;font-size:${px}px`;
        s.style.fontFamily = family || 'serif';
        if (weight) s.style.fontWeight = weight;
        s.textContent = ch.repeat(20);
        document.body.appendChild(s);
        w = s.getBoundingClientRect().width / 20;
        s.remove();
      }
      widthCache.set(k, w);
    }
    return widthCache.get(k);
  }
  // その書体の半角スペースの幅
  const spaceWidth = (family, px, weight) => domWidth(family, px, ' ', weight);
  // Windows の Chrome は ＭＳ 明朝・ＭＳ ゴシックを、小さい字（10.5pt＝14px など）で少し広く並べる（14px の字が 14.67px。
  // 小さい字のための「ヒント」の幅を使うため。CSS では切りかえられない）。Word は書体のままの幅なので、広がった分だけ字間を詰める。
  // 等幅の書体なので、全角の字は 1 字分、半角の字は半分の比で広がる。ほかの端末（書体がない・広がらない）では何もしない
  const MS_MONO = /^(ＭＳ|MS)\s*(明朝|Mincho|ゴシック|Gothic)$/i;
  function hintRatio(family, px) {
    if (px >= 24) return 1;
    const r = domWidth(`"${family}"`, px, '国') / px;
    return r > 0.95 && r < 1.15 && Math.abs(r - 1) > 1e-3 ? r : 1;
  }
  const isEA = ch => { const u = ch.codePointAt(0); return u >= 0x2100 || u === 0x203b || u === 0x2025; };   // 和文の書体で出る字
  const isHalf = ch => { const u = ch.codePointAt(0); return u < 0x1100 || (u >= 0xff61 && u <= 0xffdc) || (u >= 0xffe8 && u <= 0xffee); };
  function hintCls(hf, ch) {
    if (isEA(ch) ? !hf.eaMS : !hf.asMS) return 'n';
    return isHalf(ch) ? 'h' : 'f';
  }
  const CJK = /[\u3000-\u30ff\u3400-\u9fff\uf900-\ufaff\uff00-\uffef\u{20000}-\u{2ffff}]/u;
  const LATIN = /[\u0021-\u024f\u0370-\u03ff\u0400-\u04ff]/;

  /* ---------- 書体ごとの「1行」の高さ（Word の行間「1行」÷ 文字の大きさ。この PC の Word ではかった値） ---------- */
  const LINE_FACTOR = [
    [/^(ＭＳ|MS)\s*(Ｐ|P)?\s*(明朝|ゴシック|Mincho|Gothic)|^MS UI Gothic|^HG|^BIZ\s*UD|UD\s*デジタル|^UD Digi|教科書|^AR/i, 1.3],
    [/Yu Gothic UI/i, 1.729], [/^(游|Yu (Mincho|Gothic))/i, 1.673], [/^(メイリオ|Meiryo)$/i, 1.95], [/Meiryo UI/i, 1.651],
    [/Noto (Sans|Serif) (CJK|JP)|Source Han|源ノ|^Zen /i, 1.882], [/^Century Gothic/i, 1.226], [/^Century/i, 1.202],
    [/^Times New Roman/i, 1.15], [/^Arial Black/i, 1.41], [/^Arial/i, 1.15], [/^Calibri/i, 1.221], [/^Cambria/i, 1.172],
    [/^Segoe UI/i, 1.33], [/^Tahoma/i, 1.207], [/^Verdana/i, 1.215], [/^Georgia/i, 1.136], [/^Courier New/i, 1.133],
    [/^Consolas/i, 1.171], [/^Garamond/i, 1.125], [/^Book Antiqua/i, 1.206], [/^Aptos/i, 1.2], [/^Symbol|^Wingdings/i, 1.2],
  ];
  const factorOf = name => { for (const [re, f] of LINE_FACTOR) if (re.test(name || '')) return f; return null; };
  // 文字列 text を書体 fonts で書いたときの「1行」の高さ（pt）。日本語を含めば和文の書体、英数字を含めば欧文の書体で、大きい方
  function singlePt(size, f, text) {
    const cj = text == null || CJK.test(text), la = text == null || LATIN.test(text) || !cj;
    let k = 0;
    if (cj) k = Math.max(k, 1.3, factorOf(f.ea) || 1.3);
    if (la) k = Math.max(k, factorOf(f.ascii) || (/[^\x00-\x7f]/.test(f.ascii || '') ? 1.3 : 1.2));
    return size * k;
  }

  /* ---------- 段落・文字の書式（XML を、あとから来たものが勝つように重ねる） ---------- */
  const P0 = { jc: 'left', indL: 0, indR: 0, indF: 0, indLC: null, indRC: null, indFC: null, before: 0, after: 0, beforeL: null, afterL: null,
    line: 240, rule: 'auto', widow: true, snap: true, tabs: [], keepNext: false, keepLines: false, pbb: false, contextual: false,
    numId: null, ilvl: null, bdr: null, shd: null, beforeAuto: false, afterAuto: false, kinsoku: true, hang: true, autoSpace: true, autoSpaceN: true, textDir: null };
  const R0 = { b: false, i: false, u: null, uColor: null, strike: false, dstrike: false, color: null, sz: 20, f: { ascii: '', ea: '', asciiT: '', eaT: '' },
    hl: null, shd: null, va: null, caps: false, smallCaps: false, vanish: false, em: null, spc: 0, pos: 0, bdr: null, fit: 0, outline: false, shadow: false };
  function border(el) {
    if (!el) return null;
    const v = at(el, 'val') || 'single';
    if (v === 'nil' || v === 'none') return { none: true };
    const c = at(el, 'color');
    return { val: v, w: Math.max(0.5, atNum(el, 'sz', 4) / 8 * PT), color: c && c !== 'auto' ? '#' + c : '#000', space: atNum(el, 'space', 0) * PT };
  }
  const shdColor = el => { if (!el) return null; const f = at(el, 'fill'); if (f && f !== 'auto') return '#' + f; const v = at(el, 'val'); if (v === 'solid') { const c = at(el, 'color'); return c && c !== 'auto' ? '#' + c : '#000'; } return f === 'auto' ? 'none' : null; };
  function mergeP(acc, el) {
    if (!el) return acc;
    const o = { ...acc };
    for (const c of kids(el)) {
      const n = c.localName;
      switch (n) {
        case 'jc': o.jc = W_(c); break;
        case 'ind': {
          const a = k => at(c, k);
          if (a('left') != null || a('start') != null) { o.indL = tw(a('left') ?? a('start')); o.indLC = null; }
          if (a('leftChars') != null || a('startChars') != null) o.indLC = +(a('leftChars') ?? a('startChars')) || null;
          if (a('right') != null || a('end') != null) { o.indR = tw(a('right') ?? a('end')); o.indRC = null; }
          if (a('rightChars') != null || a('endChars') != null) o.indRC = +(a('rightChars') ?? a('endChars')) || null;
          if (a('hanging') != null) { o.indF = -tw(a('hanging')); o.indFC = null; } else if (a('firstLine') != null) { o.indF = tw(a('firstLine')); o.indFC = null; }
          if (a('hangingChars') != null) o.indFC = -(+a('hangingChars')) || null; else if (a('firstLineChars') != null) o.indFC = +a('firstLineChars') || null;
          break;
        }
        case 'spacing': {
          const a = k => at(c, k);
          if (a('before') != null) { o.before = +a('before'); o.beforeL = null; }
          if (a('beforeLines') != null) o.beforeL = +a('beforeLines');
          if (a('after') != null) { o.after = +a('after'); o.afterL = null; }
          if (a('afterLines') != null) o.afterL = +a('afterLines');
          if (a('line') != null) o.line = +a('line');
          if (a('lineRule') != null) o.rule = a('lineRule'); else if (a('line') != null) o.rule = 'auto';
          if (a('beforeAutospacing') != null) o.beforeAuto = truthy(a('beforeAutospacing'));
          if (a('afterAutospacing') != null) o.afterAuto = truthy(a('afterAutospacing'));
          break;
        }
        case 'widowControl': o.widow = onOff(c); break;
        case 'snapToGrid': o.snap = onOff(c); break;
        case 'keepNext': o.keepNext = onOff(c); break;
        case 'keepLines': o.keepLines = onOff(c); break;
        case 'pageBreakBefore': o.pbb = onOff(c); break;
        case 'contextualSpacing': o.contextual = onOff(c); break;
        case 'kinsoku': o.kinsoku = onOff(c); break;
        case 'overflowPunct': o.hang = onOff(c); break;
        case 'autoSpaceDE': o.autoSpace = onOff(c); break;
        case 'autoSpaceDN': o.autoSpaceN = onOff(c); break;
        case 'textDirection': o.textDir = W_(c); break;
        case 'numPr': {
          const id = kid(c, 'numId'), lv = kid(c, 'ilvl');
          if (id) o.numId = W_(id);
          if (lv) o.ilvl = +W_(lv) || 0;
          break;
        }
        case 'tabs': {
          let t = o.tabs.slice();
          for (const s of kids(c, 'tab')) {
            const pos = atNum(s, 'pos', 0), v = at(s, 'val');
            t = t.filter(x => Math.abs(x.pos - pos) > 1);
            if (v !== 'clear') t.push({ pos, val: v === 'start' ? 'left' : v === 'end' ? 'right' : v, leader: at(s, 'leader') || 'none' });
          }
          o.tabs = t.sort((a, b) => a.pos - b.pos);
          break;
        }
        case 'pBdr': {
          const b = { ...(o.bdr || {}) };
          for (const e of kids(c)) b[e.localName === 'start' ? 'left' : e.localName === 'end' ? 'right' : e.localName] = border(e);
          o.bdr = b;
          break;
        }
        case 'shd': o.shd = shdColor(c); break;
      }
    }
    return o;
  }
  function mergeR(acc, el, theme) {
    if (!el) return acc;
    const o = { ...acc };
    for (const c of kids(el)) {
      const n = c.localName;
      switch (n) {
        case 'b': o.b = onOff(c); break;
        case 'i': o.i = onOff(c); break;
        case 'u': { const v = W_(c) || 'single'; o.u = v === 'none' ? null : v; const uc = at(c, 'color'); o.uColor = uc && uc !== 'auto' ? '#' + uc : null; break; }
        case 'strike': o.strike = onOff(c); break;
        case 'dstrike': o.dstrike = onOff(c); break;
        case 'color': {
          const v = W_(c), tc = at(c, 'themeColor');
          if (v && v !== 'auto' && /^[0-9a-f]{6}$/i.test(v)) o.color = '#' + v;
          else if (tc && theme) o.color = themeColor(theme, tc, at(c, 'themeTint'), at(c, 'themeShade'));
          else o.color = null;
          break;
        }
        case 'sz': o.sz = atNum(c, 'val', o.sz); break;
        case 'rFonts': {
          const f = { ...o.f };
          const a = k => at(c, k);
          if (a('ascii') != null) { f.ascii = a('ascii'); f.asciiT = ''; } else if (a('hAnsi') != null && !f.ascii) f.ascii = a('hAnsi');
          if (a('eastAsia') != null) { f.ea = a('eastAsia'); f.eaT = ''; }
          if (a('asciiTheme') != null) f.asciiT = a('asciiTheme');
          if (a('eastAsiaTheme') != null) f.eaT = a('eastAsiaTheme');
          o.f = f;
          break;
        }
        case 'highlight': o.hl = W_(c) === 'none' ? null : W_(c); break;
        case 'shd': o.shd = shdColor(c); break;
        case 'vertAlign': o.va = W_(c) === 'baseline' ? null : W_(c); break;
        case 'caps': o.caps = onOff(c); break;
        case 'smallCaps': o.smallCaps = onOff(c); break;
        case 'vanish': o.vanish = onOff(c); break;
        case 'specVanish': break;
        case 'em': o.em = W_(c) === 'none' ? null : W_(c); break;
        case 'spacing': o.spc = atNum(c, 'val', 0); break;
        case 'kern': o.kern = atNum(c, 'val', 0); break;
        case 'position': o.pos = atNum(c, 'val', 0); break;
        case 'bdr': o.bdr = border(c); break;
        case 'fitText': o.fit = atNum(c, 'val', 0); break;
        case 'outline': o.outline = onOff(c); break;
        case 'shadow': o.shadow = onOff(c); break;
      }
    }
    return o;
  }
  // テーマの色（Word の themeTint・themeShade は明るさ（HSL）で調整）
  function themeColor(theme, name, tint, shade) {
    const map = { text1: 'dk1', background1: 'lt1', text2: 'dk2', background2: 'lt2', dark1: 'dk1', light1: 'lt1', dark2: 'dk2', light2: 'lt2', hyperlink: 'hlink', followedHyperlink: 'folHlink' };
    const v = theme.colors[map[name] || name];
    if (!v) return null;
    let c = { r: parseInt(v.slice(0, 2), 16), g: parseInt(v.slice(2, 4), 16), b: parseInt(v.slice(4, 6), 16) };
    if (tint || shade) {
      const [hh, s, l] = DML.rgb2hsl(c);
      const t = tint ? parseInt(tint, 16) / 255 : null, sh = shade ? parseInt(shade, 16) / 255 : null;
      const l2 = t != null ? l * t + (1 - t) : l * sh;
      c = DML.hsl2rgb(hh, s, Math.min(1, Math.max(0, l2)));
    }
    return css({ ...c, a: 1 });
  }
  // 書体の名前（テーマの書体「+本文」なども、ふつうの名前にする）
  function faces(R, theme) {
    const tf = (t, ea) => {
      if (!t || !theme) return '';
      const g = /^major/.test(t) ? theme.major : theme.minor;
      return ea || /EastAsia/.test(t) ? (g.ea || '') : g.latin || '';
    };
    const ascii = (R.f.asciiT && tf(R.f.asciiT, false)) || R.f.ascii || '';
    const ea = (R.f.eaT && tf(R.f.eaT, true)) || R.f.ea || '';
    return { ascii, ea };
  }
  const ALIAS = [[/^(ＭＳ|MS)\s*(Ｐ|P)?\s*(明朝|Mincho)/i, ['MS Mincho', 'Hiragino Mincho ProN']], [/^(ＭＳ|MS)\s*(Ｐ|P)?\s*(ゴシック|Gothic)/i, ['MS Gothic', 'Hiragino Sans']],
    [/^(游明朝|Yu Mincho)/i, ['Yu Mincho', 'YuMincho', 'Hiragino Mincho ProN']], [/^(游ゴシック|Yu Gothic)/i, ['Yu Gothic', 'YuGothic', 'Hiragino Sans']], [/^(メイリオ|Meiryo)$/i, ['Meiryo']]];
  const GOTHIC = /ゴシック|gothic|sans|meiryo|メイリオ|ﾎﾟｯﾌﾟ|ポップ|丸|arial|calibri|segoe|verdana|tahoma|helvetica|aptos/i;
  // Word で太字の本体がなく、字の幅を変えずに太くする書体（和文と欧文の両方がこれのときだけ、まねをする）
  const SYNTH_BOLD = /^(游明朝|Yu Mincho|Century|MS UI Gothic)$|^(ＭＳ|MS)\s*(Ｐ|P)?\s*(明朝|Mincho|ゴシック|Gothic)$|^HG/i;
  function synthBold(fc) {
    const ea = (fc.ea || '').trim(), as = (fc.ascii || '').trim();
    return SYNTH_BOLD.test(ea || as) && (!as || SYNTH_BOLD.test(as)) && (!ea || SYNTH_BOLD.test(ea));
  }
  function fontCss(fc) {
    const list = [];
    for (const n of [fc.ascii, fc.ea]) {
      if (!n || list.includes(n)) continue;
      list.push(n);
      for (const [re, al] of ALIAS) if (re.test(n)) for (const a of al) if (!list.includes(a)) list.push(a);
    }
    const key = fc.ea || fc.ascii || '';
    const serif = SERIF.test(key) || (/明朝|mincho|教科書|楷書|行書/i.test(key)) || (!GOTHIC.test(key) && /century|times|cambria|georgia/i.test(fc.ascii || ''));
    return list.map(n => `"${n.replace(/["\\]/g, '')}"`).concat(serif ? 'serif' : 'sans-serif').join(',');
  }

  /* ---------- 番号（箇条書き・ページ番号） ---------- */
  const AIUEO = 'アイウエオカキクケコサシスセソタチツテトナニヌネノハヒフヘホマミムメモヤユヨラリルレロワヲン';
  const AIUEO_H = 'ｱｲｳｴｵｶｷｸｹｺｻｼｽｾｿﾀﾁﾂﾃﾄﾅﾆﾇﾈﾉﾊﾋﾌﾍﾎﾏﾐﾑﾒﾓﾔﾕﾖﾗﾘﾙﾚﾛﾜｦﾝ';
  const IROHA = 'イロハニホヘトチリヌルヲワカヨタレソツネナラムウヰノオクヤマケフコエテアサキユメミシヱヒモセス';
  const KANJI = '〇一二三四五六七八九';
  function kanjiNum(n) {
    if (n <= 0) return KANJI[0];
    let s = '';
    const units = [[10000, '万'], [1000, '千'], [100, '百'], [10, '十']];
    for (const [u, c] of units) { const k = Math.floor(n / u); if (k) { s += (k > 1 || u === 10000 ? kanjiNum(k) : '') + c; n %= u; } }
    return s + (n ? KANJI[n] : '');
  }
  const roman = n => DML.roman(n);
  const alpha = n => { let k = n, t = ''; while (k > 0) { k--; t = String.fromCharCode(65 + (k % 26)) + t; k = Math.floor(k / 26); } return t; };
  const full = s => s.replace(/[0-9]/g, d => String.fromCharCode(0xff10 + +d));
  function fmtNum(n, f) {
    switch (f) {
      case 'none': return '';
      case 'bullet': return '';
      case 'decimalZero': return n < 10 ? '0' + n : String(n);
      case 'upperRoman': return roman(n);
      case 'lowerRoman': return roman(n).toLowerCase();
      case 'upperLetter': return alpha(n);
      case 'lowerLetter': return alpha(n).toLowerCase();
      case 'ordinal': return n + (n % 100 >= 11 && n % 100 <= 13 ? 'th' : ['th', 'st', 'nd', 'rd'][n % 10] || 'th');
      case 'decimalEnclosedCircle': case 'decimalEnclosedCircleChinese': return n >= 1 && n <= 20 ? String.fromCodePoint(0x2460 + n - 1) : n <= 35 ? String.fromCodePoint(0x3251 + n - 21) : n <= 50 ? String.fromCodePoint(0x32b1 + n - 36) : `(${n})`;
      case 'decimalEnclosedParen': return n >= 1 && n <= 20 ? String.fromCodePoint(0x2474 + n - 1) : `(${n})`;
      case 'decimalEnclosedFullstop': return n >= 1 && n <= 20 ? String.fromCodePoint(0x2488 + n - 1) : n + '.';
      case 'decimalFullWidth': case 'decimalFullWidth2': return full(String(n));
      case 'aiueo': return AIUEO_H[(n - 1) % AIUEO_H.length];
      case 'aiueoFullWidth': return AIUEO[(n - 1) % AIUEO.length];
      case 'iroha': case 'irohaFullWidth': return IROHA[(n - 1) % IROHA.length];
      case 'japaneseCounting': case 'chineseCounting': case 'chineseCountingThousand': case 'taiwaneseCounting': case 'taiwaneseCountingThousand': return kanjiNum(n);
      case 'ideographDigital': case 'japaneseDigitalTenThousand': return String(n).replace(/\d/g, d => KANJI[+d]);
      case 'ideographTraditional': return '甲乙丙丁戊己庚辛壬癸'[(n - 1) % 10];
      case 'ideographZodiac': return '子丑寅卯辰巳午未申酉戌亥'[(n - 1) % 12];
      case 'japaneseLegal': return String(n).replace(/\d/g, d => '〇壱弐参四伍六七八九'[+d]);
      case 'chicago': return ['*', '†', '‡', '§'][(n - 1) % 4].repeat(Math.floor((n - 1) / 4) + 1);
      default: return String(n);
    }
  }
  // Symbol・Wingdings の字（F0xx）を、ふつうの文字に
  const SYMBOL = { a: 'α', b: 'β', g: 'γ', d: 'δ', e: 'ε', z: 'ζ', h: 'η', q: 'θ', i: 'ι', k: 'κ', l: 'λ', m: 'μ', n: 'ν', x: 'ξ', p: 'π', r: 'ρ', s: 'σ', t: 'τ', u: 'υ', f: 'φ', c: 'χ', y: 'ψ', w: 'ω',
    D: 'Δ', G: 'Γ', Q: 'Θ', L: 'Λ', P: 'Π', S: 'Σ', F: 'Φ', Y: 'Ψ', W: 'Ω', '\xae': '→', '\xac': '←', '\xad': '↑', '\xaf': '↓', '\xab': '↔', '\xb4': '×', '\xb8': '÷', '\xa3': '≤', '\xb3': '≥', '\xb9': '≠',
    '\xbb': '≈', '\xb1': '±', '\xd6': '√', '\xa5': '∞', '\xb6': '∂', '\xe5': '∑', '\xf2': '∫', '\xb7': '•', '\xd7': '·', '\xb0': '°', '\xde': '⇒', '\xdb': '⇔', '\xa7': '♣', '\xa8': '♦', '\xa9': '♥', '\xaa': '♠', '\xc6': '∅', '\xce': '∈', '\xc7': '∩', '\xc8': '∪', '\xcc': '⊂', '\xd9': '∧', '\xda': '∨', '\x22': '∀', '\x24': '∃', '\x27': '∋' };
  function symChar(font, code) {
    if (code >= 0xf000) code -= 0xf000;
    const ch = String.fromCharCode(code);
    if (/symbol/i.test(font || '')) return SYMBOL[ch] || ch;
    if (/wingdings|webdings/i.test(font || '')) return WINGDINGS[ch] || ({ '\xa8': '□', '\xfe': '☑', '\xfd': '☒', '\x6f': '□', '\x9f': '•', '\xa1': '○', '\xa4': '◉', '\x4a': '☺', '\x4c': '☹', '\x28': '☎', '\x2a': '✉', '\xe0': '→', '\xdf': '←', '\xf0': '⇨' })[ch] || '•';
    return ch;
  }

  /* ---------- タブ：並べたあとで、次のタブ位置までの幅を決める（左・中央・右揃え・リーダー） ---------- */
  // 段落ごとに、タブの幅 → 行末の詰め（変わったらタブをもう一度）
  function layoutTabs(root) {
    for (const p of root.querySelectorAll('.dx-p')) {
      if (p._tabs) tabsOf(p);
      if (p._fit && fitPara(p) && p._tabs) tabsOf(p);
      if (p._noEnd && unTrimEnd(p) && p._tabs) tabsOf(p);
    }
  }
  function tabsOf(p) {
    {
      const T = p._tabs;
      const list = [...p.querySelectorAll('.dx-tab')].filter(t => t.closest('.dx-p') === p);
      if (!list.length) return;
      for (const t of list) t.style.width = '0px';
      const pr = p.getBoundingClientRect(), k = pr.width / (p.offsetWidth || 1) || 1;
      for (let i = 0; i < list.length; i++) {
        const t = list[i], r = t.getBoundingClientRect(), x = (r.left - pr.left) / k + (T.o || 0);
        let stop = null;
        const pt = t.dataset.ptab;
        if (pt) { const pw = p.offsetWidth + (T.o || 0); stop = { x: pt === 'center' ? pw / 2 : pt === 'right' ? pw : x, val: pt }; }
        else for (const s of T.stops) if (s.x > x + 0.5 && s.val !== 'bar') { stop = s; break; }
        if (!pt && T.hang != null && T.hang > x + 0.5 && (!stop || T.hang < stop.x)) stop = { x: T.hang, val: 'left' };
        if (!stop) { const d = T.def > 4 ? T.def : 56; stop = { x: (Math.floor((x + 0.5) / d) + 1) * d, val: 'left' }; }
        let w = stop.x - x;
        if (stop.val === 'right' || stop.val === 'center' || stop.val === 'decimal' || stop.val === 'end') {
          const seg = segWidth(t, list[i + 1] || null, p, r) / k;
          w = stop.val === 'center' ? w - seg / 2 : w - seg;
        }
        t.style.width = Math.max(0, w) + 'px';
        if (stop.leader && stop.leader !== 'none') t.dataset.ld = stop.leader; else delete t.dataset.ld;
      }
    }
  }
  // タブのうしろ（次のタブか段落の終わりまで）の、同じ行にある文字の幅
  function segWidth(t, next, p, tr) {
    const rg = document.createRange();
    rg.setStartAfter(t);
    if (next) rg.setEndBefore(next); else rg.setEnd(p, p.childNodes.length);
    const y = tr.top;
    let l = Infinity, r = -Infinity;
    for (const q of rg.getClientRects()) if (q.width && q.top < y + 0.5 && q.bottom > y - 0.5) { l = Math.min(l, q.left); r = Math.max(r, q.right); }
    return r > l ? r - l : 0;
  }

  /* ---------- 行末の詰め（Word の「句読点のぶら下げ」と「句読点の圧縮」） ----------
     両端揃え（均等割り付け）の段落だけ。Word は、行に入りきらない文字があっても
       ・行の中の 、。「」（）など を最大半分まで詰めて、もう1文字入れる（characterSpacingControl が compressPunctuation のとき）
       ・行末の 、。」） などは半分の幅にする（同じく compressPunctuation のとき）
       ・それでも入らなければ、行末の 、。」） などを右の余白にはみ出させる（ぶら下げ。overflowPunct、初めからオン）
     左揃え・中央・右揃えでは、どれもせず次の行へ送る（Word で確かめた）。
     ブラウザはどれもしないので、並べたあとで行の切れ目を調べ、Word なら前の行に入る文字を引き上げる */
  // 行頭禁則：Word の「標準」（小さい かな と ー は行頭に来てよい）と「上級」（strictFirstAndLastChars）
  const NO_START = /[!%),.:;?\]}¢°’”‰′″℃、。々〉》」』】〕゛゜ゝゞ・ヽヾ！％），．：；？］｝｡｣､･ﾞﾟ￠]/;
  const NO_START_STRICT = /[!%),.:;?\]}¢°’”‰′″℃、。々〉》」』】〕゛゜ゝゞ・ヽヾーァィゥェォッャュョヮヵヶぁぃぅぇぉっゃゅょゎゕゖㇰ-ㇿ！％），．：；？］｝｡｣､･ｧ-ｯｰﾞﾟ￠]/;
  const NO_END = /[$(\[{£¥‘“〈《「『【〔＄（［｛｢￡￥]/;
  const CLOSE = /[、。，．」』）］】〕〉》｝]/;
  const WORD = /[A-Za-z0-9À-ɏ'’]/;   // 欧文の単語の字（この間では行を分けない）   // 行末で半分にしたり、ぶら下げたりできる記号
  const SQUEEZE = /[、。，．「」『』（）［］【】〔〕〈〉《》｛｝]/;
  function fitPara(p) {
    const F = p._fit, ps = p.style, ta = ps.textAlign, tal = ps.textAlignLast;
    ps.textAlign = 'left'; ps.textAlignLast = 'auto';   // 字間が広がっていない、もとの幅ではかる
    const rg = document.createRange();
    const rect = c => { rg.setStart(c.n, c.i); rg.setEnd(c.n, c.i + c.len); return rg.getBoundingClientRect(); };
    const collect = () => {
      const out = [];
      let forced = false;
      const w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT | NodeFilter.SHOW_ELEMENT);
      for (let n = w.nextNode(); n; n = w.nextNode()) {
        if (n.nodeType === 1) {
          if (n.classList.contains('dx-jbr')) forced = true;
          else if (n.matches('.dx-p,.dx-inl,.dx-tbx,.dx-fit,rt,.dx-float,.dx-abs')) { forced = true; n = skipKids(w, n); }
          continue;
        }
        let i = 0;
        for (const ch of n.data) {
          if (ch === '\n' || ch === '​') forced = true;
          else { out.push({ n, i, ch, len: ch.length, forced }); forced = false; }
          i += ch.length;
        }
      }
      return out;
    };
    let chars = collect(), changed = false;
    const done = new Set();   // 詰めた行（1行に1回だけ）
    const pr = p.getBoundingClientRect(), k = pr.width / (p.offsetWidth || 1) || 1, cs = getComputedStyle(p);
    const right = pr.right - (parseFloat(cs.paddingRight) + parseFloat(cs.borderRightWidth)) * k;
    let prev = null, lineStart = 0;
    for (let b = 0; b < chars.length; b++) {
      const r = rect(chars[b]);
      if (!r.width && !r.height) continue;
      const brk = prev && (r.top + r.bottom) / 2 > prev.r.bottom;
      if (!brk) { prev = { i: b, r }; continue; }
      const e = prev.i, s = b;
      prev = { i: b, r };
      const ls = lineStart; lineStart = s;
      if (chars[s].forced || done.has(ls)) continue;
      // 引き上げる文字のかたまり（次の文字が行頭禁則なら一緒に。欧文の単語は途中で切れないので丸ごと）
      let u = s;
      const ns = F.strict ? NO_START_STRICT : NO_START;
      while (u + 1 < chars.length && u - s < 20 && !chars[u + 1].forced && (ns.test(chars[u + 1].ch) || (WORD.test(chars[u].ch) && WORD.test(chars[u + 1].ch)))) u++;
      if (NO_END.test(chars[u].ch)) continue;
      const kin = u > s && ns.test(chars[u].ch);   // 行頭禁則の字があるので、まとめて送られたかたまり
      if (!CJK.test(chars[s].ch) && !kin) continue;   // 欧文の単語だけなら引き上げない
      const adv = [];
      let sum = 0;
      for (let j = s; j <= u; j++) { const q = rect(chars[j]); adv.push(q.width); sum += q.width; }
      const slack = right - rect(chars[e]).right;
      const need = sum - slack + 0.3 * k;
      if (need <= 0) continue;   // 幅は足りている（ほかの理由で切れた）
      const lw = adv[u - s], close = CLOSE.test(chars[u].ch);
      // ふつうの字1つは、詰める量が「送ったときに行に残るすき間」より少ないときだけ（Word で確かめた。半分以上はみ出す字は送る）
      if (!kin && !close && sum - slack > slack + 0.05 * lw) continue;
      // 詰められる記号（行の頭から、引き上げる文字の最後の1つ前まで）。それぞれ半分まで
      const cand = [];
      let room = 0;
      if (F.comp) for (let j = ls; j < u; j++) if (SQUEEZE.test(chars[j].ch)) { const w = j >= s ? adv[j - s] : rect(chars[j]).width; cand.push({ j, w }); room += w / 2; }
      // Word の順：行末の閉じる記号を半分に → ほかの記号を詰める → その両方 → 行末の記号をぶら下げる（はみ出させる）
      let cut = 0, hang = false;
      if (close && F.comp && need <= lw / 2) cut = lw / 2;
      else if (F.comp && need <= room) cut = 0;
      else if (close && F.comp && need - lw / 2 <= room) cut = lw / 2;
      else if (close && F.hang && need - lw <= (F.comp ? room : 0)) { cut = lw; hang = true; }
      else continue;
      const rest = Math.max(0, need - cut), f = rest > 0 ? rest / room : 0;
      // 後ろから包む（前の文字の場所がずれないように）
      const ops = rest > 0 ? cand.map(x => ({ j: x.j, cls: 'dx-sq', d: x.w / 2 * f / k })) : [];
      if (cut) ops.push({ j: u, cls: hang ? 'dx-hang' : 'dx-sq', d: cut / k });
      ops.sort((a, b2) => b2.j - a.j);
      for (const o of ops) {
        const sp = wrapChar(chars[o.j], o.cls);
        // 始めのかっこは左側を、ほかは右側を詰める（ぶら下げは幅 0 に。右の margin は行を分ける計算に入らないので字間で）
        if (/[「『（［【〔〈《｛]/.test(chars[o.j].ch)) sp.style.marginLeft = -o.d + 'px';
        else sp.style.letterSpacing = (parseFloat(getComputedStyle(sp).letterSpacing) || 0) - o.d + 'px';
      }
      changed = true; done.add(ls);
      chars = collect();
      prev = null; lineStart = ls; b = ls - 1;   // この行からはかり直す
    }
    ps.textAlign = ta; ps.textAlignLast = tal;
    return changed;
  }
  // 両端揃えでない段落：ブラウザは、行に入りきらない行末の 」。 などを半分の幅にして入れる（text-spacing-trim: normal）が、
  // Word はそうせず次の行へ送る。半分になった行末の記号を、詰めない span（dx-nt）で包む（並んだ記号の詰めは Word もする）
  const PUNCT = /[、。，．「」『』（）［］【】〔〕〈〉《》｛｝・：；！？]/;
  function unTrimEnd(p) {
    if (!CLOSE.test(p.textContent)) return false;
    const rg = document.createRange(), k = p.getBoundingClientRect().width / (p.offsetWidth || 1) || 1;
    const collect = () => {
      const out = [], w = document.createTreeWalker(p, NodeFilter.SHOW_TEXT);
      for (let n; (n = w.nextNode());) if (n.parentElement.closest('.dx-p') === p) for (let i = 0; i < n.data.length; i++) out.push({ n, i, ch: n.data[i], len: 1 });
      return out;
    };
    let chars = collect(), changed = false;
    for (let j = 0; j < chars.length; j++) {
      const c = chars[j], nx = chars[j + 1];
      if (!CLOSE.test(c.ch) || (nx && PUNCT.test(nx.ch)) || c.n.parentElement.classList.contains('dx-nt')) continue;
      rg.setStart(c.n, c.i); rg.setEnd(c.n, c.i + 1);
      const q = rg.getBoundingClientRect();
      if (!q.width || q.width > parseFloat(getComputedStyle(c.n.parentElement).fontSize) * k * 0.75) continue;
      wrapChar(c, 'dx-nt');
      changed = true; chars = collect();   // 文字の並びは同じなので j はそのまま使える
    }
    return changed;
  }
  // TreeWalker で n の中を飛ばす（次は n のうしろから）
  function skipKids(w, n) { let x = n; while (x.lastChild) x = x.lastChild; w.currentNode = x; return x; }
  function wrapChar(c, cls) {
    let n = c.n;
    if (c.i > 0) n = n.splitText(c.i);
    if (n.data.length > c.len) n.splitText(c.len);
    const s = document.createElement('span');
    s.className = cls;
    n.before(s); s.appendChild(n);
    return s;
  }

  /* ---------- ページを切ってよい所（行の切れ目・表の行の切れ目）をはかる ---------- */
  // 行の上下（el の上からの px）。絵（dx-inl）も行の一部。浮いている図（dx-abs・dx-float）は数えない
  function lineBoxes(el, top) {
    const rects = [];
    const walk = n => {
      for (let c = n.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 3) {
          if (!c.data) continue;
          const rg = document.createRange();
          rg.selectNodeContents(c);
          for (const q of rg.getClientRects()) if (q.height > 0) rects.push(q);
        } else if (c.nodeType === 1) {
          const cl = c.classList;
          if (cl.contains('dx-abs') || cl.contains('dx-float') || cl.contains('dx-tab')) continue;
          if (cl.contains('dx-inl') || c.localName === 'math') { const q = c.getBoundingClientRect(); if (q.height > 0) rects.push(q); continue; }
          walk(c);
        }
      }
    };
    walk(el);
    rects.sort((a, b) => a.top - b.top);
    const lines = [];
    for (const q of rects) {
      const L = lines[lines.length - 1];
      if (L && q.top < L.b - Math.min(q.height, L.b - L.t) * 0.5) { L.t = Math.min(L.t, q.top); L.b = Math.max(L.b, q.bottom); }
      else lines.push({ t: q.top, b: q.bottom });
    }
    return lines.map(L => ({ t: L.t - top, b: L.b - top }));
  }
  // 切ってよい所 [{ y, n: その上の行数, of: 全部の行数 }]（el の上からの px、小さい順）
  function cutsOf(el) {
    const top = el.getBoundingClientRect().top;
    if (el.classList.contains('dx-p')) {
      const L = lineBoxes(el, top), out = [];
      for (let i = 0; i + 1 < L.length; i++) out.push({ y: (L[i].b + L[i + 1].t) / 2, n: i + 1, of: L.length });
      return out;
    }
    const tb = el.localName === 'table' ? el : el.querySelector(':scope > table');
    if (!tb) return [];
    const out = [], rows = [...tb.rows];
    rows.forEach((tr, i) => {
      const r = tr.getBoundingClientRect();
      if (!tr.classList.contains('dx-cs')) for (const y of rowInnerCuts(tr, r)) out.push({ y: y - top, inRow: i });
      if (i + 1 < rows.length) out.push({ y: r.bottom - top, row: i + 1 });
    });
    return out.sort((a, b) => a.y - b.y);
  }
  function contentCuts(box) {
    const ys = [];
    for (const b of box.children) {
      if (b.classList.contains('dx-abs')) continue;
      const r = b.getBoundingClientRect();
      if (b !== box.firstElementChild) ys.push(r.top);
      for (const c of cutsOf(b)) ys.push(r.top + c.y);
    }
    return ys;
  }
  // 表の1行の途中で切るとき：どのセルでも行の切れ目になっている所だけ
  function rowInnerCuts(tr, rr) {
    const cells = [...tr.cells].map(td => {
      const inner = td.querySelector(':scope > .dx-cell') || td, r = inner.getBoundingClientRect();
      return { top: r.top, bottom: r.bottom, ys: contentCuts(inner) };
    });
    const cand = new Set();
    for (const c of cells) for (const y of c.ys) if (y > rr.top + 2 && y < rr.bottom - 2) cand.add(y);
    return [...cand].filter(y => cells.every(c => y <= c.top + 1 || y >= c.bottom - 1 || c.ys.some(v => Math.abs(v - y) < 2))).sort((a, b) => a - b);
  }

  const SUPPORTED = new Set(['wps', 'wpg', 'w14', 'w15', 'w16', 'w16se', 'w16cid', 'a14', 'wp14', 'm']);
  function choose(ac) {
    for (const c of kids(ac, 'Choice')) if ((at(c, 'Requires') || '').split(/\s+/).filter(Boolean).every(p => SUPPORTED.has(p))) return c;
    return kid(ac, 'Fallback');
  }
  // 段落・表の並び（入れものの sdt・customXml・挿入の印などはほどく。削除の印の中は出さない）
  function blockKids(el, out = []) {
    for (const c of kids(el)) {
      const n = c.localName;
      if (n === 'p' || n === 'tbl' || n === 'sectPr') out.push(c);
      else if (n === 'sdt') blockKids(kid(c, 'sdtContent'), out);
      else if (n === 'customXml' || n === 'ins' || n === 'moveTo' || n === 'smartTag') blockKids(c, out);
      else if (n === 'AlternateContent') { const ch = choose(c); if (ch) blockKids(ch, out); }
      else if (n === 'altChunk') out.push(c);
    }
    return out;
  }

  /* ---------- 文書を開く ---------- */
  async function open(zip) {
    const pk = DML.pkg(zip), { xml, rels, relOfType, media, recolor } = pk;
    const rootRels = await rels('');
    const main = relOfType(rootRels, 'officeDocument');
    const docPath = main ? main.target : 'word/document.xml';
    const docXml = await xml(docPath);
    const body = docXml && pathOf(docXml.documentElement, 'body');
    if (!body) throw new Error('Word のファイルとして読めませんでした');
    const docRels = await rels(docPath);
    const partX = async type => { const r = relOfType(docRels, type); return r ? { path: r.target, x: await xml(r.target) } : null; };
    const [stylesP, numP, setP, fnP, enP] = await Promise.all(['styles', 'numbering', 'settings', 'footnotes', 'endnotes'].map(partX));
    const thRel = relOfType(docRels, 'theme');
    const theme = thRel ? await pk.theme(thRel.target) : { colors: {}, major: { latin: '', ea: '' }, minor: { latin: '', ea: '' }, fills: [], lines: [], bgFills: [] };
    const cmap = { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' };
    const dctx = src => ({ theme, map: cmap, media, recolor, src });

    // 設定
    const S = setP && setP.x ? setP.x.documentElement : null;
    const defTab = tw(atNum(kid(S, 'defaultTabStop'), 'val', 720));
    const evenOdd = onOff(kid(S, 'evenAndOddHeaders'));
    let compat = 12;
    let otf = false;   // OpenType の機能を使う（Word 2013 以降の文書）。このときだけ、文字の組み合わせで間を詰める（カーニング）
    for (const c of kids(pathOf(S, 'compat'), 'compatSetting')) {
      if (at(c, 'name') === 'compatibilityMode') compat = +at(c, 'val') || 12;
      if (at(c, 'name') === 'enableOpenTypeFeatures') otf = at(c, 'val') === '1' || at(c, 'val') === 'true';
    }
    // 両端そろえの段落で Shift+Enter の前の行も広げる（Word の標準。互換オプションで止められる）
    const expandBr = !onOff(kid(pathOf(S, 'compat'), 'doNotExpandShiftReturn'));
    // 句読点を詰めて1文字多く入れるか（日本語の Word の標準は compressPunctuation）
    const squeeze = /^compressPunctuation/.test(W_(kid(S, 'characterSpacingControl')) || '');
    // 禁則の強さ（上級なら小さい かな・ー も行頭に来ない）と、和文のとなりの半角スペースを半角（0.5 字）にするか
    const strictKinsoku = !!onOff(kid(S, 'strictFirstAndLastChars'));
    const balance = !!onOff(kid(pathOf(S, 'compat'), 'balanceSingleByteDoubleByteWidth'));
    const bgEl = kid(docXml.documentElement, 'background');
    const pageBg = bgEl && onOff(kid(S, 'displayBackgroundShape')) && at(bgEl, 'color') && at(bgEl, 'color') !== 'auto' ? '#' + at(bgEl, 'color') : null;

    // スタイル
    const ST = stylesP && stylesP.x ? stylesP.x.documentElement : null;
    const styles = new Map();
    let defP = null, defT = null;
    for (const s of kids(ST, 'style')) {
      const id = at(s, 'styleId'), type = at(s, 'type');
      const o = { id, type, based: W_(kid(s, 'basedOn')), pPr: kid(s, 'pPr'), rPr: kid(s, 'rPr'), tblPr: kid(s, 'tblPr'), trPr: kid(s, 'trPr'), tcPr: kid(s, 'tcPr'), cond: {} };
      for (const c of kids(s, 'tblStylePr')) o.cond[at(c, 'type')] = { pPr: kid(c, 'pPr'), rPr: kid(c, 'rPr'), tblPr: kid(c, 'tblPr'), trPr: kid(c, 'trPr'), tcPr: kid(c, 'tcPr') };
      styles.set(id, o);
      if (truthy(at(s, 'default'))) { if (type === 'paragraph') defP = id; else if (type === 'table') defT = id; }
    }
    const chainCache = new Map();
    function chain(id) {
      if (!id || !styles.has(id)) return [];
      if (!chainCache.has(id)) {
        const out = [], seen = new Set();
        for (let cur = id; cur && styles.has(cur) && !seen.has(cur); cur = styles.get(cur).based) { seen.add(cur); out.unshift(styles.get(cur)); }
        chainCache.set(id, out);
      }
      return chainCache.get(id);
    }
    const dd = pathOf(ST, 'docDefaults');
    const PD = mergeP(P0, pathOf(dd, 'pPrDefault', 'pPr')), RD = mergeR(R0, pathOf(dd, 'rPrDefault', 'rPr'), theme);

    // 番号
    const NR = numP && numP.x ? numP.x.documentElement : null;
    const lvlOf = l => ({
      start: kid(l, 'start') ? atNum(kid(l, 'start'), 'val', 1) : 1, fmt: W_(kid(l, 'numFmt')) || 'decimal', text: kid(l, 'lvlText') ? (at(kid(l, 'lvlText'), 'val') || '') : '',
      jc: W_(kid(l, 'lvlJc')) || 'left', pPr: kid(l, 'pPr'), rPr: kid(l, 'rPr'), suff: W_(kid(l, 'suff')) || 'tab', lgl: !!onOff(kid(l, 'isLgl')),
      restart: kid(l, 'lvlRestart') ? atNum(kid(l, 'lvlRestart'), 'val', -1) : null,
    });
    const abstract = new Map(), nums = new Map();
    for (const a of kids(NR, 'abstractNum')) {
      const lv = [];
      for (const l of kids(a, 'lvl')) lv[atNum(l, 'ilvl', 0)] = lvlOf(l);
      abstract.set(at(a, 'abstractNumId'), { lv, link: W_(kid(a, 'numStyleLink')) });
    }
    for (const n of kids(NR, 'num')) {
      const ov = {};
      for (const o of kids(n, 'lvlOverride')) ov[atNum(o, 'ilvl', 0)] = { start: kid(o, 'startOverride') ? atNum(kid(o, 'startOverride'), 'val', 1) : null, lvl: kid(o, 'lvl') ? lvlOf(kid(o, 'lvl')) : null };
      nums.set(at(n, 'numId'), { abs: W_(kid(n, 'abstractNumId')), ov, hasOv: Object.keys(ov).length > 0 });
    }
    function levelDef(numId, ilvl) {
      const n = nums.get(numId);
      if (!n) return null;
      let a = abstract.get(n.abs);
      if (a && a.link && !a.lv.length) {
        const st = styles.get(a.link), np = st && pathOf(st.pPr, 'numPr'), n2 = np && nums.get(W_(kid(np, 'numId')));
        if (n2 && abstract.get(n2.abs)) a = abstract.get(n2.abs);
      }
      if (!a) return null;
      const o = n.ov[ilvl], L = (o && o.lvl) || a.lv[ilvl];
      return L ? { L, n, start: o && o.start != null ? o.start : L.start } : null;
    }
    const counters = new Map();
    function nextNumber(numId, ilvl) {
      const d = levelDef(numId, ilvl);
      if (!d) return null;
      const key = d.n.hasOv ? 'n' + numId : 'a' + d.n.abs;
      let c = counters.get(key);
      if (!c) counters.set(key, c = []);
      c[ilvl] = c[ilvl] == null ? d.start : c[ilvl] + 1;
      for (let k = ilvl + 1; k < 10; k++) { const dk = levelDef(numId, k); if (!(dk && dk.L.restart === 0)) c[k] = null; }
      let text;
      if (d.L.fmt === 'bullet') text = d.L.text;
      else text = d.L.text.replace(/%(\d)/g, (_, m) => {
        const k = +m - 1, dk = levelDef(numId, k);
        return fmtNum(c[k] != null ? c[k] : dk ? dk.start : 1, d.L.lgl ? 'decimal' : dk ? dk.L.fmt : 'decimal');
      });
      return { text, L: d.L };
    }

    // 書式を決める
    function paraProps(pPr, tctx) {
      let P = PD, R = RD;
      if (tctx) { for (const e of tctx.pEls) P = mergeP(P, e); for (const e of tctx.rEls) R = mergeR(R, e, theme); }
      const sid = W_(kid(pPr, 'pStyle')) || defP;
      for (const s of chain(sid)) { P = mergeP(P, s.pPr); R = mergeR(R, s.rPr, theme); }
      const dn = pathOf(pPr, 'numPr');
      let numId = P.numId, ilvl = P.ilvl;
      if (dn) { if (kid(dn, 'numId')) numId = W_(kid(dn, 'numId')); if (kid(dn, 'ilvl')) ilvl = +W_(kid(dn, 'ilvl')) || 0; }
      let num = null;
      if (numId && numId !== '0') {
        const d = levelDef(numId, ilvl || 0);
        if (d) { P = mergeP(P, d.L.pPr); num = { numId, ilvl: ilvl || 0 }; }
      }
      const numL = P.indL;
      P = mergeP(P, pPr);
      // 番号つきの段落に「字」単位の左インデント（leftChars）を直接つけると、Word は番号の位置からの字数として並べる
      // （単一レベルの番号で「インデントを増やす」とこうなる。Word で確かめた）
      const di = kid(pPr, 'ind');
      if (num && di && +(at(di, 'leftChars') ?? at(di, 'startChars') ?? 0)) P = { ...P, indLCbase: numL };
      return { P, R, sid, num, markR: mergeR(R, kid(pPr, 'rPr'), theme) };
    }
    function runProps(rPr, R) {
      const rs = W_(kid(rPr, 'rStyle'));
      if (rs) for (const s of chain(rs)) R = mergeR(R, s.rPr, theme);
      return mergeR(R, rPr, theme);
    }

    // 区切り（セクション）：用紙・余白・行グリッド・ヘッダー
    function sectOf(sp) {
      const sz = kid(sp, 'pgSz'), mg = kid(sp, 'pgMar'), grid = kid(sp, 'docGrid'), pn = kid(sp, 'pgNumType');
      const m = (k, d) => tw(atNum(mg, k, d));
      const s = {
        W: tw(atNum(sz, 'w', 11906)), H: tw(atNum(sz, 'h', 16838)), top: Math.abs(m('top', 1440)), bottom: Math.abs(m('bottom', 1440)),
        left: m('left', 1440) + m('gutter', 0), right: m('right', 1440), header: m('header', 708), footer: m('footer', 708),
        type: W_(kid(sp, 'type')) || 'nextPage', titlePg: !!onOff(kid(sp, 'titlePg')), cols: atNum(kid(sp, 'cols'), 'num', 1),
        pnStart: pn && at(pn, 'start') != null ? +at(pn, 'start') : null, pnFmt: (pn && at(pn, 'fmt')) || 'decimal',
        vert: /^tbRl|^tbLrV/.test(W_(kid(sp, 'textDirection')) || ''), hdr: {}, ftr: {}, bdr: null,
      };
      const gt = at(grid, 'type');
      s.grid = gt && gt !== 'default' ? { pitch: atNum(grid, 'linePitch', 360) / 20, chars: gt === 'linesAndChars' || gt === 'snapToChars', charSpace: atNum(grid, 'charSpace', 0) } : null;
      for (const r of kids(sp, 'headerReference')) s.hdr[at(r, 'type') || 'default'] = rAttr(r, 'id');
      for (const r of kids(sp, 'footerReference')) s.ftr[at(r, 'type') || 'default'] = rAttr(r, 'id');
      const pb = kid(sp, 'pgBorders');
      if (pb) { s.bdr = { from: at(pb, 'offsetFrom') || 'text' }; for (const e of kids(pb)) s.bdr[e.localName] = border(e); }
      s.cw = Math.max(48, s.W - s.left - s.right);
      return s;
    }
    const secs = [];
    {
      let cur = [];
      for (const el of blockKids(body)) {
        if (el.localName === 'sectPr') continue;
        cur.push(el);
        const sp = el.localName === 'p' && pathOf(el, 'pPr', 'sectPr');
        if (sp) { secs.push({ els: cur, s: sectOf(sp) }); cur = []; }
      }
      secs.push({ els: cur, s: sectOf(kid(body, 'sectPr')) });
      // ヘッダー・フッターの指定がない区切りは、前の区切りのものを使う
      for (let i = 1; i < secs.length; i++) for (const k of ['hdr', 'ftr']) for (const t of ['default', 'first', 'even']) if (!secs[i].s[k][t] && secs[i - 1].s[k][t]) secs[i].s[k][t] = secs[i - 1].s[k][t];
    }

    /* ---------- 段落を描く ---------- */
    // X：{ src: { path, rels }, s: 区切り, mode: 'paper'|'flow', hk: 横の縮み, tctx: 表の書式, fld: フィールドの重なり, hdr: ヘッダーか }
    const HL = { yellow: '#ff0', green: '#0f0', cyan: '#0ff', magenta: '#f0f', blue: '#00f', red: '#f00', darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000', darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000', white: '#fff' };
    const notes = { fn: new Map(), en: new Map(), fnN: 0, enN: 0 };
    function lineH(single, P, s) {
      const g = s && s.grid && P.snap !== false ? s.grid : null;
      const mult = P.rule === 'auto' ? (P.line || 240) / 240 : 1;
      let pt;
      if (P.rule === 'exact') pt = Math.abs(P.line) / 20;
      else if (g) { const n = Math.max(1, Math.ceil(single / g.pitch - 1e-3)); pt = P.rule === 'atLeast' ? Math.max(P.line / 20, g.pitch * n) : g.pitch * Math.max(mult, n); }
      else pt = P.rule === 'atLeast' ? Math.max(P.line / 20, single) : single * mult;
      return pt * PT;
    }
    function runStyle(R, X, P, st, text) {
      const fc = faces(R, theme), size = hp(R.sz);
      st.fontFamily = fontCss(fc);
      // 上付き・下付きは、ずらしても行を高くしない（Word と同じ。vertical-align だと行が高くなるので、位置だけずらす）
      let shift = 0;
      if (R.va === 'superscript' || R.va === 'subscript') { st.fontSize = size * 0.65 * PT + 'px'; shift = R.va === 'superscript' ? -(size * PT / 3 + 1) : size * PT / 5 + 1; }
      else st.fontSize = size * PT + 'px';
      // 太字の本体がない書体（游明朝・ＭＳ 明朝など）は、Word は字の幅を変えずに太く見せる。
      // ブラウザは別の太さの書体（游明朝 Demibold）を使ったり幅を広げたりするので、ふつうの太さに線を足して太く見せる
      const fake = R.b && !R.outline && synthBold(fc);
      st.fontWeight = R.b && !fake ? 'bold' : 'normal';
      if (fake) st.webkitTextStroke = +(parseFloat(st.fontSize) * 0.035).toFixed(2) + 'px currentColor';
      st.fontStyle = R.i ? 'italic' : 'normal';
      const deco = [];
      if (R.u) deco.push('underline');
      if (R.strike || R.dstrike) deco.push('line-through');
      if (deco.length) {
        st.textDecorationLine = deco.join(' ');
        const u = R.u || '';
        st.textDecorationStyle = /double/i.test(u) || R.dstrike ? 'double' : /dot/i.test(u) ? 'dotted' : /dash/i.test(u) ? 'dashed' : /wav/i.test(u) ? 'wavy' : 'solid';
        if (/thick|heavy/i.test(u)) st.textDecorationThickness = '2px';
        if (R.uColor && R.u) st.textDecorationColor = R.uColor;
      }
      if (R.color) st.color = R.color;
      if (R.hl && HL[R.hl]) st.backgroundColor = HL[R.hl]; else if (R.shd && R.shd !== 'none') st.backgroundColor = R.shd;
      if (R.caps) st.textTransform = 'uppercase'; else if (R.smallCaps) st.fontVariant = 'small-caps';
      // カーニング（w:kern：この大きさ（半ポイント）以上の字で、書体の組み合わせの詰めを使う。游明朝の「す。」など）
      if (otf && R.kern && size * 2 >= R.kern) st.fontKerning = 'normal';
      let ls = R.spc / 20 * PT;
      if (X.s && X.s.grid && X.s.grid.chars && P.snap !== false) ls += X.s.grid.charSpace / 4096 * PT;
      if (ls) { st.letterSpacing = ls + 'px'; if (st.setProperty) st.setProperty('--ls', ls + 'px'); }
      if (R.pos || shift) { st.position = 'relative'; st.top = shift - hp(R.pos || 0) * PT + 'px'; }
      if (R.em) { st.textEmphasis = (R.em === 'comma' ? 'sesame' : R.em === 'circle' ? 'open circle' : 'filled dot'); if (R.em === 'underDot') st.textEmphasisPosition = 'under right'; }
      if (R.bdr && !R.bdr.none) st.border = `${R.bdr.w}px solid ${R.bdr.color}`;
      if (R.shadow) st.textShadow = '1px 1px 1px rgba(0,0,0,.35)';
      if (R.outline) { st.webkitTextStroke = '0.6px ' + (R.color || '#000'); st.color = '#fff'; }
      st.lineHeight = lineH(singlePt(size, fc, text), P, X.s) + 'px';
      return size;
    }
    const isSym = R => /symbol|wingdings|webdings/i.test(faces(R, theme).ascii);

    function renderPara(p, X) {
      const pPr = kid(p, 'pPr'), pp = paraProps(pPr, X.tctx), P = pp.P, hk = X.hk || 1;
      const fs = hp(pp.markR.sz) * PT;
      const indL = (P.indLC != null ? (P.indLCbase || 0) + P.indLC / 100 * fs : P.indL) * hk, indR = (P.indRC != null ? P.indRC / 100 * fs : P.indR) * hk;
      const indF = (P.indFC != null ? P.indFC / 100 * fs : P.indF) * hk;
      const unit = X.s && X.s.grid ? X.s.grid.pitch * PT : 12 * PT;
      const el = h('div', 'dx-p');
      const st = el.style;
      if (indL >= 0) st.paddingLeft = indL + 'px'; else st.marginLeft = indL + 'px';
      if (indR >= 0) st.paddingRight = indR + 'px'; else st.marginRight = indR + 'px';
      if (indF) st.textIndent = indF + 'px';
      st.textAlign = { both: 'justify', distribute: 'justify', center: 'center', right: 'right', end: 'right', thaiDistribute: 'justify' }[P.jc] || 'left';
      if (P.jc === 'distribute') st.textAlignLast = 'justify';
      const mst = {};
      runStyle(pp.markR, X, P, mst, null);
      // 行の高さは文字（span）ごとに持たせ、段落そのものは 0（段落の書体の「支柱」で行が 1px ほど高くなるのを防ぐ。
      // Word も段落記号の書体では行を高くしない）。数式だけは段落の行の高さ（--lh）
      st.fontSize = mst.fontSize; st.fontFamily = mst.fontFamily; st.lineHeight = '0'; st.setProperty('--lh', mst.lineHeight);
      if (P.shd && P.shd !== 'none') st.backgroundColor = P.shd;
      if (P.bdr) for (const k of ['top', 'bottom', 'left', 'right']) { const b = P.bdr[k]; if (b && !b.none) { st['border' + k[0].toUpperCase() + k.slice(1)] = `${b.w}px ${/double/.test(b.val) ? 'double' : /dot/.test(b.val) ? 'dotted' : /dash/.test(b.val) ? 'dashed' : 'solid'} ${b.color}`; if (k === 'top' || k === 'bottom') st['padding' + (k === 'top' ? 'Top' : 'Bottom')] = b.space + 'px'; } }
      if (!P.kinsoku) st.lineBreak = 'loose'; else if (strictKinsoku) st.lineBreak = 'strict';
      const before = P.beforeAuto ? 14 * PT : P.beforeL != null ? P.beforeL / 100 * unit : P.before / 20 * PT;
      const after = P.afterAuto ? 14 * PT : P.afterL != null ? P.afterL / 100 * unit : P.after / 20 * PT;
      const tabs = { stops: P.tabs.map(t => ({ x: tw(t.pos) * hk, val: t.val, leader: t.leader })), def: defTab * hk, hang: indF < 0 ? Math.max(0, indL) : null, o: Math.min(0, indL) };
      el._tabs = tabs;
      if (st.textAlign !== 'justify') el._noEnd = true;
      else if (P.hang !== false || squeeze) el._fit = { hang: P.hang !== false, comp: squeeze, strict: strictKinsoku };
      const meta = { el, kind: 'p', before, after, sid: pp.sid, contextual: P.contextual, keepNext: P.keepNext, keepLines: P.keepLines, widow: P.widow, pbb: P.pbb, fns: [], bdr: P.bdr ? JSON.stringify(P.bdr) : null, between: P.bdr && P.bdr.between };
      const ctx = { X, pp, P, root: el, out: el, last: null, meta, content: false };
      // 番号・記号
      if (pp.num) {
        const nn = nextNumber(pp.num.numId, pp.num.ilvl);
        if (nn && (nn.text || nn.L.fmt !== 'none')) {
          const R = mergeR(pp.markR, nn.L.rPr, theme);
          let t = nn.text;
          if (nn.L.fmt === 'bullet' && isSym(R)) t = [...t].map(ch => symChar(faces(R, theme).ascii, ch.charCodeAt(0))).join('');
          else if (nn.L.fmt === 'bullet' && /[-]/.test(t)) t = [...t].map(ch => symChar('Symbol', ch.charCodeAt(0))).join('');
          const sp = h('span', 'dx-num');
          runStyle({ ...R, u: null, hl: null }, X, P, sp.style, t);
          sp.textContent = t;
          el.appendChild(sp);
          if (nn.L.suff === 'tab') el.appendChild(h('span', 'dx-tab'));
          else if (nn.L.suff === 'space') sp.textContent += ' ';
          ctx.content = true;
        }
      }
      inline(p, ctx);
      // 改ページで分かれた段落
      const marks = [...el.children].filter(c => c.classList.contains('dx-pbk'));
      const parts = [el];
      for (const m of marks) {
        const cur = parts[parts.length - 1], nx = cur.cloneNode(false);
        nx._tabs = tabs; nx._fit = el._fit; nx._noEnd = el._noEnd;
        nx.style.textIndent = '0px';
        while (m.nextSibling) nx.appendChild(m.nextSibling);
        m.remove();
        parts.push(nx);
      }
      const metas = parts.map((e, i) => ({ ...meta, el: e, before: i ? 0 : before, after: i === parts.length - 1 ? after : 0, brk: i < parts.length - 1 ? 'page' : null, fns: i ? [] : meta.fns, cont: i > 0 }));
      if (parts.length > 1 && compat >= 15 && !hasContent(parts[parts.length - 1])) { metas.pop(); metas[metas.length - 1].after = after; metas[metas.length - 1].brk = 'page'; }
      for (const m of metas) { m.el.style.paddingBottom = m.after + (parseFloat(m.el.style.paddingBottom) || 0) + 'px'; if (!hasContent(m.el)) addEmpty(m.el, pp.markR, X, P); }
      return metas;
    }
    const hasContent = e => !!(e.textContent.replace(/[​\s]/g, '') || e.querySelector('.dx-inl,.dx-tab,math,img,.dx-float,.dx-fld') || /\S|\n/.test(e.textContent));
    function addEmpty(e, R, X, P) { const s = h('span'); runStyle(R, X, P, s.style, null); s.textContent = '​'; e.appendChild(s); }

    // 段落の中身（文字の並び）
    function inline(parent, ctx) {
      for (const c of kids(parent)) {
        switch (c.localName) {
          case 'r': doRun(c, ctx); break;
          case 'hyperlink': {
            const id = rAttr(c, 'id'), rel = id && ctx.X.src.rels.get(id);
            const url = rel && rel.external && /^(https?:|mailto:)/i.test(rel.target) ? rel.target : null;
            if (url && ctx.X.mode && !ctx.X.hdr) {
              const a = h('a', 'dx-a');
              a.href = url; a.target = '_blank'; a.rel = 'noopener noreferrer';
              ctx.out.appendChild(a);
              const o = ctx.out; ctx.out = a; ctx.last = null;
              inline(c, ctx);
              ctx.out = o; ctx.last = null;
            } else inline(c, ctx);
            break;
          }
          case 'fldSimple': {
            const kind = fieldKind(at(c, 'instr'));
            if (kind && ctx.X.hdr) { placeholder(kind, runProps(pathOf(c, 'r', 'rPr'), ctx.pp.R), ctx); break; }
            inline(c, ctx);
            break;
          }
          case 'smartTag': case 'customXml': case 'ins': case 'moveTo': case 'dir': case 'bdo': inline(c, ctx); break;
          case 'sdt': inline(kid(c, 'sdtContent'), ctx); break;
          case 'AlternateContent': { const ch = choose(c); if (ch) inline(ch, ctx); break; }
          case 'oMath': case 'oMathPara': if (showing(ctx.X)) { ctx.out.appendChild(mathOf(c)); ctx.last = null; ctx.content = true; } break;
        }
      }
    }
    const fieldKind = instr => { const m = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(instr || ''); return m ? m[1].toUpperCase() : null; };
    const showing = X => X.fld.every(f => f.res && !f.hide);
    function placeholder(kind, R, ctx) {
      const s = h('span', 'dx-fld');
      s.dataset.f = kind;
      runStyle(R, ctx.X, ctx.P, s.style, '0');
      s.textContent = '1';
      ctx.out.appendChild(s); ctx.last = null; ctx.content = true;
    }
    function doRun(r, ctx) {
      const R = runProps(kid(r, 'rPr'), ctx.pp.R), X = ctx.X;
      const each = list => {
        for (const c of list) {
          const n = c.localName;
          if (n === 'fldChar') {
            const t = at(c, 'fldCharType');
            if (t === 'begin') X.fld.push({ instr: '', res: false });
            else if (t === 'separate') { const f = X.fld[X.fld.length - 1]; if (f) { f.res = true; f.kind = fieldKind(f.instr); if (f.kind && X.hdr && X.fld.length === 1) { f.hide = true; placeholder(f.kind, R, ctx); } } }
            else if (t === 'end') { const f = X.fld.pop(); if (f && !f.res && X.hdr && fieldKind(f.instr) && showing(X)) placeholder(fieldKind(f.instr), R, ctx); }
            continue;
          }
          if (n === 'instrText') { const f = X.fld[X.fld.length - 1]; if (f && !f.res) f.instr += c.textContent; continue; }
          if (!showing(X) || (R.vanish && n !== 'drawing')) continue;
          switch (n) {
            case 't': text(c.textContent, R, ctx); break;
            case 'tab': ctx.out.appendChild(h('span', 'dx-tab')); ctx.last = null; ctx.content = true; ctx.lineTab = true; break;
            case 'ptab': { const s = h('span', 'dx-tab'); s.dataset.ptab = at(c, 'alignment') || 'left'; ctx.out.appendChild(s); ctx.last = null; break; }
            case 'br': {
              const bt = at(c, 'type');
              if (bt === 'page' && ctx.out === ctx.root && X.mode) { ctx.out.appendChild(h('span', 'dx-pbk')); ctx.last = null; }
              else if (bt !== 'page' && bt !== 'column' && expandBr && ctx.P.jc === 'both' && ctx.lineText && !ctx.lineTab && !R.fit) {
                // 行の幅いっぱいの見えない箱で次の行へ送る → その前の行は「最後の行」ではないので両端そろえで広がる（タブのある行は広げない）
                ctx.out.appendChild(h('span', 'dx-jbr')); ctx.last = null; ctx.lineText = false;
              } else text('\n', R, ctx);
              break;
            }
            case 'cr': text('\n', R, ctx); break;
            case 'noBreakHyphen': text('‑', R, ctx); break;
            case 'softHyphen': text('­', R, ctx); break;
            case 'sym': text(symChar(at(c, 'font'), parseInt(at(c, 'char') || '20', 16)), R, ctx, true); break;
            case 'drawing': drawing(c, R, ctx); break;
            case 'pict': case 'object': vml(c, R, ctx); break;
            case 'footnoteReference': case 'endnoteReference': noteRef(c, R, ctx, n === 'footnoteReference' ? 'fn' : 'en'); break;
            case 'footnoteRef': case 'endnoteRef': text(String(X.noteNo || ''), { ...R, va: 'superscript' }, ctx); break;
            case 'pgNum': if (X.hdr) placeholder('PAGE', R, ctx); break;
            case 'ruby': ruby(c, R, ctx); break;
            case 'AlternateContent': { const ch = choose(c); if (ch) each(kids(ch)); break; }
          }
        }
      };
      each(kids(r));
    }
    function text(str, R, ctx, mapped) {
      if (!str) return;
      if (!mapped && isSym(R)) str = [...str].map(ch => symChar(faces(R, theme).ascii, ch.charCodeAt(0))).join('');
      const L = ctx.last;
      if (L && L.R === R && L.span.parentNode === ctx.out && !R.fit) {
        L.text += str; fillText(L.span, L.text, ctx.P, L.before);
        L.span.style.lineHeight = lineH(singlePt(hp(R.sz), faces(R, theme), L.text), ctx.P, ctx.X.s) + 'px';
      } else {
        const s = h('span');
        runStyle(R, ctx.X, ctx.P, s.style, str);
        if (!R.fit) { const hf = hintFix(R, parseFloat(s.style.letterSpacing) || 0); if (hf) { s._hf = hf; s.style.letterSpacing = hf.ls + hf.c[hf.base] + 'px'; } }
        if (R.fit) { s.className = 'dx-fit'; s.style.width = R.fit / 20 * PT * (ctx.X.hk || 1) + 'px'; s.textContent = str; }
        let before = null;
        if (!R.fit) {
          // 前の文字の並び（すぐ前にある別の書式の文字）とのすき間
          const pv = ctx.prevText;
          if (pv && !pv.R.fit && endsWith(ctx.root, pv.span)) {
            const a = [...pv.text].pop(), b = [...str][0];
            before = a;
            if (asGap(a, b, ctx.P)) wrapLast(pv.span);
            else if (balance && a === ' ' && CJK.test(b)) balanceLast(pv.span);
          }
          fillText(s, str, ctx.P, before);
        }
        ctx.out.appendChild(s);
        ctx.last = { R, span: s, text: str, before };
      }
      ctx.prevText = ctx.last;
      ctx.content = true;
      // いまの行（最後の改行のあと）に文字があるか・タブがあるか（Shift+Enter の行を広げるかどうか）
      const nl = str.lastIndexOf('\n');
      if (nl >= 0) { ctx.lineTab = false; ctx.lineText = /\S/.test(str.slice(nl + 1)); }
      else if (/\S/.test(str)) ctx.lineText = true;
    }
    // 文字を入れる。和文と欧文・数字の間は 1/4 字あける（Word の autoSpaceDE / DN。前の文字の字間を広げて作る）
    // balance（balanceSingleByteDoubleByteWidth）のときは、和文のとなりの半角スペースを 0.5 字にする（語間で足す）
    // ＭＳ 明朝・ＭＳ ゴシックの字の広がりを詰める設定（hintRatio）。いらなければ null
    function hintFix(R, ls) {
      const fc = faces(R, theme), ea = (fc.ea || '').trim(), as = (fc.ascii || '').trim();
      const eaMS = MS_MONO.test(ea), asMS = MS_MONO.test(as) || (!as && eaMS);
      if (!eaMS && !asMS) return null;
      const px = hp(R.sz) * PT * (R.va === 'superscript' || R.va === 'subscript' ? 0.65 : 1);
      const r = hintRatio(eaMS ? ea : as, px);
      if (r === 1) return null;
      const d = (1 - r) * px;
      return { ls, eaMS, asMS, c: { n: 0, f: d, h: d / 2 }, base: eaMS ? 'f' : 'h' };
    }
    function fillText(span, str, P, before) {
      span.textContent = '';
      const hf = span._hf, cs = [...str];
      let buf = '', bc = hf ? hf.base : null;
      // 字間の詰めが span と違う字（和文の書体の中の半角、欧文の書体の字）は、別の span に入れる
      const flush = () => {
        if (!buf) return;
        if (hf && bc !== hf.base) { const w = h('span', 'dx-hw'); w.style.letterSpacing = hf.ls + hf.c[bc] + 'px'; w.textContent = buf; span.appendChild(w); }
        else span.append(buf);
        buf = '';
      };
      for (let i = 0; i < cs.length; i++) {
        const c = hf ? hintCls(hf, cs[i]) : null;
        if (i + 1 < cs.length && asGap(cs[i], cs[i + 1], P)) { flush(); span.appendChild(asWrap(cs[i], hf ? hf.ls + hf.c[c] : null)); }
        else if (balance && cs[i] === ' ' && (CJK.test((i ? cs[i - 1] : before) || '') || CJK.test(cs[i + 1] || ''))) { flush(); span.appendChild(halfSpace(span)); }
        else { if (hf && c !== bc) { flush(); bc = c; } buf += cs[i]; }
      }
      flush();
    }
    // 1/4 字あける字（ls：詰めを足した字間。null なら CSS のまま）
    function asWrap(ch, ls) {
      const w = h('span', 'dx-as');
      w.textContent = ch;
      if (ls != null) w.style.letterSpacing = `calc(${ls}px + .25em)`;
      return w;
    }
    function halfSpace(span) {
      const w = h('span', 'dx-bsp'), px = parseFloat(span.style.fontSize) || 14, hf = span._hf;
      const corr = hf ? hf.c[hintCls(hf, ' ')] : 0;
      w.textContent = ' ';
      if (hf) w.style.letterSpacing = hf.ls + corr + 'px';
      w.style.wordSpacing = px / 2 - spaceWidth(span.style.fontFamily, px, span.style.fontWeight) - corr + 'px';
      return w;
    }
    // span の最後の文字が入っている所（字間の詰めで分けた span の中のこともある）
    const lastHost = span => { const t = span.lastChild; return t && t.nodeType === 1 && t.classList.contains('dx-hw') ? t : span; };
    function balanceLast(span) {
      const host = lastHost(span), t = host.lastChild;
      if (!t || t.nodeType !== 3 || !t.data.endsWith(' ')) return;
      t.data = t.data.slice(0, -1);
      span.appendChild(halfSpace(span));
      if (!t.data) t.remove();
      if (host !== span && !host.firstChild) host.remove();
    }
    function wrapLast(span) {
      const host = lastHost(span), t = host.lastChild;
      if (!t || t.nodeType !== 3) return;
      const cs = [...t.data];
      const w = asWrap(cs.pop(), span._hf ? parseFloat(host.style.letterSpacing) || 0 : null);
      t.data = cs.join('');
      host.appendChild(w);
      if (!t.data) t.remove();
    }
    // span が root のいちばん最後の要素か（あいだに絵・タブなどが入っていないか）
    const endsWith = (root, span) => { let e = root; while (e && e !== span) e = e.lastElementChild; return e === span; };
    function noteRef(c, R, ctx, kind) {
      if (truthy(at(c, 'customMarkFollows'))) return;
      const id = at(c, 'id'), map = notes[kind];
      if (!map.has(id)) map.set(id, ++notes[kind + 'N']);
      const n = map.get(id);
      const s = h('span', 'dx-fnref');
      runStyle({ ...R, va: 'superscript' }, ctx.X, ctx.P, s.style, '1');
      s.textContent = kind === 'en' ? fmtNum(n, 'lowerRoman') : String(n);
      s.dataset.fn = kind + ':' + id;
      ctx.out.appendChild(s); ctx.last = null; ctx.content = true;
      ctx.meta.fns.push({ kind, id, n });
    }
    function ruby(c, R, ctx) {
      const rb = h('ruby', 'dx-ruby'), rt = h('rt');
      const o = ctx.out;
      ctx.out = rb; ctx.last = null;
      inline(kid(c, 'rubyBase'), ctx);
      ctx.out = rt; ctx.last = null;
      inline(kid(c, 'rt'), ctx);
      const hps = atNum(pathOf(c, 'rubyPr', 'hps'), 'val', 0);
      if (hps) for (const s of rt.querySelectorAll('span')) s.style.fontSize = hp(hps) * PT + 'px';
      rb.appendChild(rt);
      ctx.out = o; ctx.last = null;
      o.appendChild(rb);
    }

    /* ---------- 図（写真・図形・テキストボックス・グループ） ---------- */
    // 中身は外の箱（W×H）に対する % で置く（スマホ向け表示で縮めても崩れないように）
    const missing = t => { const e = h('span', 'dx-miss'); e.textContent = t; return e; };
    function picture(pic, W, H, X) {
      const el = h('span', 'dx-pic'), bf = kid(pic, 'blipFill');
      const sr = kid(bf, 'srcRect'), q = k => atNum(sr, k, 0) / 100000;
      const cw = Math.max(0.01, 1 - q('l') - q('r')), chh = Math.max(0.01, 1 - q('t') - q('b'));
      const img = h('img', null, { width: 100 / cw + '%', height: 100 / chh + '%', left: -q('l') / cw * 100 + '%', top: -q('t') / chh * 100 + '%' });
      img.alt = ''; img.draggable = false;
      el.appendChild(img);
      X.wait.push(blipImage(bf, X.src, dctx(X.src), null).then(m => {
        if (m && m.url) { img.src = m.url; if (m.opacity != null) img.style.opacity = m.opacity; }
        else el.replaceChildren(missing(m && m.unsupported ? `${m.unsupported} の図（表示できません）` : '図（表示できません）'));
      }).catch(() => {}));
      transformOf(pathOf(pic, 'spPr', 'xfrm'), el);
      return el;
    }
    function transformOf(xf, el) {
      if (!xf) return;
      const t = [], rot = atNum(xf, 'rot', 0) / 60000;
      if (rot) t.push(`rotate(${rot}deg)`);
      if (truthy(at(xf, 'flipH'))) t.push('scaleX(-1)');
      if (truthy(at(xf, 'flipV'))) t.push('scaleY(-1)');
      if (t.length) el.style.transform = t.join(' ');
    }
    function shape(sp, W, H, X) {
      const spPr = kid(sp, 'spPr'), stl = kid(sp, 'style'), ctx = dctx(X.src), el = h('span', 'dx-shp');
      const xf = kid(spPr, 'xfrm');
      let fill = fillOf(spPr, ctx), line = lineOf(kid(spPr, 'ln'), ctx);
      if (stl) {
        const fr = kid(stl, 'fillRef'), lr = kid(stl, 'lnRef');
        const fi = atNum(fr, 'idx', 0), li = atNum(lr, 'idx', 0);
        if (!fill && fi > 0) { const f = fi >= 1001 ? theme.bgFills[fi - 1001] : theme.fills[fi - 1]; fill = f ? fillEl(f, { ...ctx, ph: color(fr, ctx) }) : null; }
        if (li > 0 && theme.lines[li - 1]) line = mergeLine(lineOf(theme.lines[li - 1], { ...ctx, ph: color(lr, ctx) }), line);
      }
      const paths = geometry(spPr, W, H);
      if (paths && ((fill && !fill.none) || (line && !line.none && line.color))) {
        X.wait.push(shapeSvg(paths, fill, line, W, H, { h: truthy(at(xf, 'flipH')), v: truthy(at(xf, 'flipV')) }, ctx).then(svg => {
          svg.setAttribute('width', '100%'); svg.setAttribute('height', '100%'); svg.setAttribute('preserveAspectRatio', 'none');
          el.prepend(svg);
        }).catch(() => {}));
      }
      const tb = pathOf(sp, 'txbx', 'txbxContent');
      if (tb) {
        const bp = kid(sp, 'bodyPr'), ins = (k, d) => px(at(bp, k) != null ? at(bp, k) : d);
        const r = textRect(spPr, { w: W, h: H, flipH: false, flipV: false }) || { x: 0, y: 0, w: W, h: H };
        const tx = h('div', 'dx-tbx', { left: r.x / W * 100 + '%', top: r.y / H * 100 + '%', width: r.w / W * 100 + '%', height: r.h / H * 100 + '%', padding: `${ins('tIns', 45720)}px ${ins('rIns', 91440)}px ${ins('bIns', 45720)}px ${ins('lIns', 91440)}px` });
        const an = at(bp, 'anchor');
        // 文字が箱に入りきらないときは、Word は上から並べる（safe）
        tx.style.justifyContent = an === 'ctr' ? 'safe center' : an === 'b' ? 'safe flex-end' : 'flex-start';
        if (/vert|eaVert/.test(at(bp, 'vert') || '') && at(bp, 'vert') !== 'horz') tx.style.writingMode = 'vertical-rl';
        const fc = stl && color(kid(stl, 'fontRef'), ctx);
        if (fc) tx.style.color = css(fc);
        // 箱の中の文字も、本文と同じ行グリッドに合わせる（Word で測った）
        renderBlocks(blockKids(tb), { ...X, tctx: null, fld: [], inBox: true }, tx);
        el.appendChild(tx);
      }
      transformOf(xf && atNum(xf, 'rot', 0) ? xf : null, el);
      return el;
    }
    function group(g, W, H, X) {
      const el = h('span', 'dx-grp'), gx = pathOf(g, 'grpSpPr', 'xfrm');
      const co = kid(gx, 'chOff'), ce = kid(gx, 'chExt');
      const cx = px(at(co, 'x')), cy = px(at(co, 'y')), cw = px(at(ce, 'cx')) || W, chh = px(at(ce, 'cy')) || H;
      for (const c of kids(g)) {
        const n = c.localName;
        if (!/^(wsp|pic|grpSp|graphicFrame)$/.test(n)) continue;
        const xe = n === 'grpSp' ? pathOf(c, 'grpSpPr', 'xfrm') : n === 'graphicFrame' ? kid(c, 'xfrm') : pathOf(c, 'spPr', 'xfrm');
        if (!xe) continue;
        const off = kid(xe, 'off'), ext = kid(xe, 'ext');
        const x = px(at(off, 'x')), y = px(at(off, 'y')), w = px(at(ext, 'cx')), hh = px(at(ext, 'cy'));
        const rw = w * W / cw, rh = hh * H / chh;
        const b = h('span', 'dx-gi', { left: (x - cx) / cw * 100 + '%', top: (y - cy) / chh * 100 + '%', width: w / cw * 100 + '%', height: hh / chh * 100 + '%' });
        b.appendChild(n === 'wsp' ? shape(c, rw, rh, X) : n === 'pic' ? picture(c, rw, rh, X) : n === 'grpSp' ? group(c, rw, rh, X) : missing('図'));
        if (n === 'grpSp') transformOf(xe, b);
        el.appendChild(b);
      }
      return el;
    }
    function graphic(d, W, H, X) {
      const g = pathOf(d, 'graphic', 'graphicData'), c = g && g.firstElementChild, n = c ? c.localName : '';
      if (n === 'pic') return picture(c, W, H, X);
      if (n === 'wsp') return shape(c, W, H, X);
      if (n === 'wgp') return group(c, W, H, X);
      if (n === 'chart') return missing('グラフ（表示できません）');
      if (n === 'relIds') return missing('SmartArt（表示できません）');
      return missing('図（表示できません）');
    }
    function drawing(c, R, ctx) {
      const d = kid(c, 'inline') || kid(c, 'anchor');
      if (!d) return;
      const ext = kid(d, 'extent'), W = Math.max(1, px(at(ext, 'cx'))), H = Math.max(1, px(at(ext, 'cy')));
      const box = h('span', 'dx-inl');
      box.appendChild(graphic(d, W, H, ctx.X));
      sizeBox(box, W, H, ctx.X);
      if (d.localName === 'inline') {
        // 行グリッドのある文書：行より高い絵の行は、グリッドの行の倍数の高さにして、絵を上下まん中に置く（Word で確かめた）
        const g = ctx.X.s && ctx.X.s.grid, P = ctx.P;
        if (g && ctx.X.mode !== 'flow' && P && P.snap !== false && P.rule !== 'exact') {
          const pitch = g.pitch * PT;
          if (H > pitch) { const ex = Math.ceil(H / pitch - 1e-3) * pitch - H; box.style.verticalAlign = 'bottom'; box.style.marginTop = box.style.marginBottom = ex / 2 + 'px'; }
        }
        ctx.out.appendChild(box); ctx.last = null; ctx.content = true; return;
      }
      const wrapEl = kids(d).find(k => /^wrap/.test(k.localName)), ph = kid(d, 'positionH'), pv = kid(d, 'positionV');
      const tx = (e, n) => { const k = kid(e, n); return k ? k.textContent.trim() : null; };
      place(box, {
        W, H, wrap: wrapEl ? wrapEl.localName.replace(/^wrap/, '').toLowerCase() : 'none', behind: truthy(at(d, 'behindDoc')), z: atNum(d, 'relativeHeight', 0),
        hRel: at(ph, 'relativeFrom') || 'column', hOff: px(tx(ph, 'posOffset')), hAl: tx(ph, 'align'),
        vRel: at(pv, 'relativeFrom') || 'paragraph', vOff: px(tx(pv, 'posOffset')), vAl: tx(pv, 'align'),
        dl: px(at(d, 'distL')), dr: px(at(d, 'distR')), dt: px(at(d, 'distT')), db: px(at(d, 'distB')),
      }, ctx);
    }
    function sizeBox(box, W, H, X) {
      box.style.width = W + 'px';
      if (X.mode === 'flow') { box.style.maxWidth = '100%'; box.style.aspectRatio = `${W} / ${H}`; }
      else box.style.height = H + 'px';
    }
    // 文章の外に浮いている図：ページの上の位置はページ分けのあとで決める（dx-abs）。文字を回り込ませるものは float でまねる
    function place(box, pos, ctx) {
      const X = ctx.X, s = X.s || { cw: 600, left: 0, right: 0, W: 600 };
      if (X.mode === 'flow' || X.inBox) {
        if (X.mode === 'flow' && !box.querySelector('img,.dx-tbx,.dx-miss')) return; // 飾りの図形は出さない
        if (X.mode === 'flow') {
          box.classList.add('dx-blk');
          if (box.querySelector('.dx-tbx')) { box.classList.add('dx-tbxb'); box.style.aspectRatio = 'auto'; box.style.width = 'auto'; }
          ctx.root.insertBefore(box, ctx.root.firstChild);
        }
        else { box.classList.add('dx-abs'); box.style.left = pos.hOff + 'px'; box.style.top = pos.vOff + 'px'; ctx.root.appendChild(box); }
        return;
      }
      const cw = s.cw, al = (a, area, w) => (a === 'center' ? (area - w) / 2 : a === 'right' || a === 'outside' ? area - w : 0);
      let x;
      switch (pos.hRel) {
        case 'page': x = pos.hAl ? al(pos.hAl, s.W, pos.W) - s.left : pos.hOff - s.left; break;
        case 'leftMargin': x = pos.hAl ? al(pos.hAl, s.left, pos.W) - s.left : pos.hOff - s.left; break;
        case 'rightMargin': x = cw + (pos.hAl ? al(pos.hAl, s.right, pos.W) : pos.hOff); break;
        default: x = pos.hAl ? al(pos.hAl, cw, pos.W) : pos.hOff;
      }
      if (pos.wrap === 'none' || (pos.wrap !== 'topandbottom' && pos.behind)) {
        box.classList.add('dx-abs');
        box._pos = pos;
        box.style.left = x + 'px';
        box.style.top = (/paragraph|line/.test(pos.vRel) ? pos.vOff : 0) + 'px';
        box.style.zIndex = pos.behind ? 0 : 3;
        ctx.root.appendChild(box);
        return;
      }
      box.classList.add('dx-float');
      const top = /paragraph|line/.test(pos.vRel) ? Math.max(0, pos.vOff) : 0;
      if (pos.wrap === 'topandbottom') Object.assign(box.style, { display: 'block', marginLeft: Math.max(0, x) + 'px', marginTop: top + pos.dt + 'px', marginBottom: pos.db + 'px' });
      // 余白にはみ出す図も、Word と同じ位置に置く（margin をマイナスにして、文字の回り込みは図の内側の端に合わせる）
      else if (x + pos.W / 2 < cw / 2) Object.assign(box.style, { float: 'left', marginLeft: x + 'px', marginRight: pos.dr + 'px', marginTop: top + 'px', marginBottom: pos.db + 'px' });
      else Object.assign(box.style, { float: 'right', marginRight: cw - x - pos.W + 'px', marginLeft: pos.dl + 'px', marginTop: top + 'px', marginBottom: pos.db + 'px' });
      ctx.root.insertBefore(box, ctx.root.firstChild);
    }
    // 古い形の図（VML：w:pict・w:object）。写真とテキストボックスだけ
    function vlen(v) {
      const m = /^(-?[\d.]+)\s*(pt|in|cm|mm|px|emu)?$/.exec((v || '').trim());
      if (!m) return 0;
      const n = +m[1];
      return { pt: n * PT, in: n * 96, cm: n * 96 / 2.54, mm: n * 96 / 25.4, emu: n / 9525 }[m[2]] ?? n;
    }
    function vml(c, R, ctx) {
      const X = ctx.X;
      for (const s of kids(c)) {
        const n = s.localName;
        if (!/^(shape|rect|roundrect|oval|image)$/.test(n)) { if (n === 'group') { const st = styleMap(at(s, 'style')); const b = h('span', 'dx-inl'); b.appendChild(missing('図（表示できません）')); sizeBox(b, vlen(st.width) || 100, vlen(st.height) || 40, X); ctx.out.appendChild(b); } continue; }
        const st = styleMap(at(s, 'style')), W = Math.max(1, vlen(st.width)), H = Math.max(1, vlen(st.height));
        const box = h('span', 'dx-inl'), inner = h('span', 'dx-shp');
        const im = kid(s, 'imagedata'), tb = pathOf(s, 'textbox', 'txbxContent');
        if (im) {
          const id = rAttr(im, 'id') || rAttr(im, 'pict'), rel = id && X.src.rels.get(id);
          const img = h('img', null, { width: '100%', height: '100%', left: 0, top: 0 });
          img.alt = '';
          inner.className = 'dx-pic'; inner.appendChild(img);
          if (rel && !rel.external) X.wait.push(media(rel.target).then(m => { if (m && m.url) img.src = m.url; else inner.replaceChildren(missing(m && m.unsupported ? `${m.unsupported} の図（表示できません）` : '図（表示できません）')); }).catch(() => {}));
        } else {
          const fc = at(s, 'fillcolor'), sc = at(s, 'strokecolor');
          if (at(s, 'filled') !== 'f' && at(s, 'filled') !== 'false') inner.style.background = vcolor(fc) || '#fff';
          if (at(s, 'stroked') !== 'f' && at(s, 'stroked') !== 'false') inner.style.border = `${Math.max(0.75, vlen(at(s, 'strokeweight') || '0.75pt'))}px solid ${vcolor(sc) || '#000'}`;
          if (n === 'oval') inner.style.borderRadius = '50%'; else if (n === 'roundrect') inner.style.borderRadius = '12%';
        }
        if (tb) {
          const tx = h('div', 'dx-tbx', { left: 0, top: 0, width: '100%', height: '100%', padding: '4.8px 9.6px' });
          renderBlocks(blockKids(tb), { ...X, tctx: null, fld: [], inBox: true }, tx);
          inner.appendChild(tx);
        }
        if (!im && !tb && X.mode === 'flow') continue;
        box.appendChild(inner);
        sizeBox(box, W, H, X);
        if (st.position !== 'absolute') { ctx.out.appendChild(box); ctx.last = null; ctx.content = true; continue; }
        const rel = v => ({ page: 'page', margin: 'margin', text: null, char: 'character', line: 'line', 'left-margin-area': 'leftMargin', 'right-margin-area': 'rightMargin', 'top-margin-area': 'topMargin', 'bottom-margin-area': 'bottomMargin' })[v];
        const wr = kid(s, 'wrap'), wt = at(wr, 'type');
        place(box, {
          W, H, wrap: wt === 'square' || wt === 'tight' || wt === 'through' ? 'square' : wt === 'topAndBottom' ? 'topandbottom' : 'none', behind: +(st['z-index'] || 0) < 0, z: +(st['z-index'] || 0),
          hRel: rel(st['mso-position-horizontal-relative']) || 'column', hOff: vlen(st['margin-left']) + vlen(st.left), hAl: /^(left|center|right)$/.test(st['mso-position-horizontal'] || '') ? st['mso-position-horizontal'] : null,
          vRel: rel(st['mso-position-vertical-relative']) || 'paragraph', vOff: vlen(st['margin-top']) + vlen(st.top), vAl: /^(top|center|bottom)$/.test(st['mso-position-vertical'] || '') ? st['mso-position-vertical'] : null,
          dl: 9.6, dr: 9.6, dt: 0, db: 0,
        }, ctx);
      }
    }
    const styleMap = s => { const o = {}; for (const kv of (s || '').split(';')) { const i = kv.indexOf(':'); if (i > 0) o[kv.slice(0, i).trim().toLowerCase()] = kv.slice(i + 1).trim(); } return o; };
    const vcolor = v => { if (!v) return null; const m = /^#?([0-9a-f]{6}|[0-9a-f]{3})\b/i.exec(v.trim()); return m ? '#' + m[1] : /^[a-z]+$/i.test(v.trim()) ? v.trim() : null; };

    /* ---------- 数式（OMML → MathML） ---------- */
    const MNS = 'http://www.w3.org/1998/Math/MathML';
    const me = (tag, ...ch) => { const e = document.createElementNS(MNS, tag); for (const c of ch) if (c) e.appendChild(typeof c === 'string' ? document.createTextNode(c) : c); return e; };
    const mo = t => me('mo', t);
    function mathOf(el) {
      const m = me('math');
      if (el.localName === 'oMathPara') { m.setAttribute('display', 'block'); for (const o of kids(el, 'oMath')) m.appendChild(mrow(o)); }
      else m.appendChild(mrow(el));
      return m;
    }
    function mrow(el) { const r = me('mrow'); if (el) for (const c of kids(el)) { const x = mnode(c); if (x) r.appendChild(x); } return r; }
    function mtext(t) {
      const r = me('mrow');
      for (const tok of t.match(/\d+(?:\.\d+)?|[A-Za-zͰ-Ͽ]|\s+|./gu) || []) {
        if (/^\s+$/.test(tok)) continue;
        r.appendChild(me(/^\d/.test(tok) ? 'mn' : /^[A-Za-zͰ-Ͽ]$/.test(tok) ? 'mi' : /[぀-鿿]/.test(tok) ? 'mtext' : 'mo', tok));
      }
      return r;
    }
    function mnode(c) {
      const k = n => kid(c, n), pv = (pr, n) => at(pathOf(c, pr, n), 'val');
      const hasKids = e => e && kids(e).length;
      switch (c.localName) {
        case 'r': return mtext(kids(c, 't').map(t => t.textContent).join(''));
        case 'f': return me('mfrac', mrow(k('num')), mrow(k('den')));
        case 'sSup': return me('msup', mrow(k('e')), mrow(k('sup')));
        case 'sSub': return me('msub', mrow(k('e')), mrow(k('sub')));
        case 'sSubSup': return me('msubsup', mrow(k('e')), mrow(k('sub')), mrow(k('sup')));
        case 'sPre': return me('mmultiscripts', mrow(k('e')), me('mprescripts'), mrow(k('sub')), mrow(k('sup')));
        case 'rad': return hasKids(k('deg')) && !truthy(pv('radPr', 'degHide') || '') ? me('mroot', mrow(k('e')), mrow(k('deg'))) : me('msqrt', mrow(k('e')));
        case 'd': {
          const ch = n => { const e = pathOf(c, 'dPr', n); return e ? (at(e, 'val') ?? '') : null; };
          const beg = ch('begChr') ?? '(', end = ch('endChr') ?? ')', sep = ch('sepChr') ?? '|';
          const r = me('mrow');
          if (beg) r.appendChild(mo(beg));
          kids(c, 'e').forEach((e, i) => { if (i) r.appendChild(mo(sep)); r.appendChild(mrow(e)); });
          if (end) r.appendChild(mo(end));
          return r;
        }
        case 'nary': {
          const chr = pv('naryPr', 'chr') || '∫', op = mo(chr), sub = k('sub'), sup = k('sup');
          op.setAttribute('largeop', 'true');
          const hs = hasKids(sub), hu = hasKids(sup), und = (pv('naryPr', 'limLoc') || (chr === '∫' ? 'subSup' : 'undOvr')) === 'undOvr';
          const base = hs || hu ? me(und ? (hs && hu ? 'munderover' : hs ? 'munder' : 'mover') : (hs && hu ? 'msubsup' : hs ? 'msub' : 'msup'), op, hs ? mrow(sub) : null, hu ? mrow(sup) : null) : op;
          return me('mrow', base, mrow(k('e')));
        }
        case 'func': return me('mrow', mrow(k('fName')), mo('⁡'), mrow(k('e')));
        case 'acc': { const o = mo(pv('accPr', 'chr') || '̂'); const m = me('mover', mrow(k('e')), o); m.setAttribute('accent', 'true'); return m; }
        case 'bar': { const top = pv('barPr', 'pos') === 'top'; return me(top ? 'mover' : 'munder', mrow(k('e')), mo(top ? '‾' : '_')); }
        case 'groupChr': { const top = pv('groupChrPr', 'pos') === 'top'; return me(top ? 'mover' : 'munder', mrow(k('e')), mo(pv('groupChrPr', 'chr') || '⏟')); }
        case 'limLow': return me('munder', mrow(k('e')), mrow(k('lim')));
        case 'limUpp': return me('mover', mrow(k('e')), mrow(k('lim')));
        case 'm': { const t = me('mtable'); for (const r of kids(c, 'mr')) { const tr = me('mtr'); for (const e of kids(r, 'e')) tr.appendChild(me('mtd', mrow(e))); t.appendChild(tr); } return t; }
        case 'eqArr': { const t = me('mtable'); for (const e of kids(c, 'e')) t.appendChild(me('mtr', me('mtd', mrow(e)))); return t; }
        case 'box': case 'borderBox': case 'phant': case 'e': return mrow(c.localName === 'e' ? c : k('e'));
        default: return null;
      }
    }

    /* ---------- 表 ---------- */
    const bcss = b => {
      if (!b || b.none) return 'none';
      const v = b.val || 'single';
      const dbl = /double|triple|thinThick|thickThin/.test(v);
      return `${dbl ? Math.max(3, b.w * 3) : b.w}px ${dbl ? 'double' : /^dot/.test(v) && !/dash/i.test(v) ? 'dotted' : /dash/i.test(v) ? 'dashed' : 'solid'} ${b.color}`;
    };
    function bordersOf(e, into) {
      if (e) for (const c of kids(e)) into[c.localName === 'start' ? 'left' : c.localName === 'end' ? 'right' : c.localName] = border(c);
      return into;
    }
    function renderTable(tbl, X) {
      const hk = X.hk || 1, tblPr = kid(tbl, 'tblPr');
      const sch = chain(W_(kid(tblPr, 'tblStyle')) || defT);
      const tp = { bdr: {}, mar: { top: 0, bottom: 0, left: 108, right: 108 }, jc: null, ind: 0, rowBand: 1, colBand: 1, shd: null };
      const applyTP = e => {
        if (!e) return;
        bordersOf(kid(e, 'tblBorders'), tp.bdr);
        for (const m of kids(kid(e, 'tblCellMar'))) tp.mar[m.localName === 'start' ? 'left' : m.localName === 'end' ? 'right' : m.localName] = atNum(m, 'w', 0);
        if (kid(e, 'jc')) tp.jc = W_(kid(e, 'jc'));
        if (kid(e, 'tblInd')) tp.ind = atNum(kid(e, 'tblInd'), 'w', 0);
        if (kid(e, 'tblStyleRowBandSize')) tp.rowBand = atNum(kid(e, 'tblStyleRowBandSize'), 'val', 1) || 1;
        if (kid(e, 'tblStyleColBandSize')) tp.colBand = atNum(kid(e, 'tblStyleColBandSize'), 'val', 1) || 1;
        if (kid(e, 'shd')) tp.shd = shdColor(kid(e, 'shd'));
      };
      for (const s of sch) applyTP(s.tblPr);
      applyTP(tblPr);
      const lk = kid(tblPr, 'tblLook'), lv = parseInt(at(lk, 'val') || '04A0', 16);
      const flag = (a, bit) => (at(lk, a) != null ? truthy(at(lk, a)) : !!(lv & bit));
      const look = { firstRow: flag('firstRow', 0x20), lastRow: flag('lastRow', 0x40), firstCol: flag('firstColumn', 0x80), lastCol: flag('lastColumn', 0x100), noH: flag('noHBand', 0x200), noV: flag('noVBand', 0x400) };
      const cond = type => sch.map(s => s.cond[type]).filter(Boolean);

      // 行とセル（横の結合 gridSpan・縦の結合 vMerge）
      const trs = [];
      const collect = (e, name, out) => { for (const c of kids(e)) { if (c.localName === name) out.push(c); else if (c.localName === 'sdt') collect(kid(c, 'sdtContent'), name, out); else if (/^(customXml|ins|moveTo)$/.test(c.localName)) collect(c, name, out); } return out; };
      collect(tbl, 'tr', trs);
      const rows = [];
      let ncol = 0;
      trs.forEach((tr, r) => {
        const trPr = kid(tr, 'trPr'), row = { r, cells: [], byCol: [], trPr, before: atNum(kid(trPr, 'gridBefore'), 'val', 0), after: atNum(kid(trPr, 'gridAfter'), 'val', 0) };
        const th = kid(trPr, 'trHeight');
        row.h = th ? tw(at(th, 'val')) : 0; row.hRule = at(th, 'hRule') || 'atLeast';
        row.header = !!onOff(kid(trPr, 'tblHeader')); row.cantSplit = !!onOff(kid(trPr, 'cantSplit'));
        let col = row.before;
        for (const tc of collect(tr, 'tc', [])) {
          const tcPr = kid(tc, 'tcPr'), span = Math.max(1, atNum(kid(tcPr, 'gridSpan'), 'val', 1)), vm = kid(tcPr, 'vMerge');
          const prev = rows[r - 1] && rows[r - 1].byCol[col];
          if (vm && (at(vm, 'val') || 'continue') === 'continue' && prev) { prev.rowSpan++; for (let k = 0; k < span; k++) row.byCol[col + k] = prev; }
          else { const cell = { tc, tcPr, col, span, rowSpan: 1, r }; row.cells.push(cell); for (let k = 0; k < span; k++) row.byCol[col + k] = cell; }
          col += span;
        }
        row.end = col;
        ncol = Math.max(ncol, col + row.after);
        rows.push(row);
      });
      let grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map(g => tw(at(g, 'w')));
      const avail = (X.s ? X.s.cw : 600);
      if (grid.length < ncol || grid.every(w => !w)) { const w = avail / Math.max(1, ncol); grid = Array.from({ length: Math.max(ncol, grid.length) }, (_, i) => grid[i] || w); }
      const cols = grid.map(w => w * hk), nRows = rows.length, nCols = grid.length;

      const t = h('table', 'dx-tbl'), cg = h('colgroup');
      for (const w of cols) cg.appendChild(h('col', null, { width: w + 'px' }));
      t.appendChild(cg);
      t.style.width = cols.reduce((a, b) => a + b, 0) + 'px';
      const tb = h('tbody');
      t.appendChild(tb);
      const mpx = v => tw(v) * Math.min(1, hk * 1.5);
      let headerRows = 0, inHeader = true;
      for (const row of rows) {
        const tr = h('tr');
        if (row.header && inHeader) { tr.className = 'dx-th'; headerRows++; } else inHeader = false;
        if (row.cantSplit) tr.classList.add('dx-cs');
        if (row.h) tr.style.height = row.h + 'px';
        if (row.before) { const td = h('td', 'dx-gb'); td.colSpan = row.before; tr.appendChild(td); }
        for (const cell of row.cells) {
          const r0 = row.r, rEnd = r0 + cell.rowSpan;
          const fr = look.firstRow && r0 === 0, lr = look.lastRow && rEnd === nRows, fc = look.firstCol && cell.col === 0, lc = look.lastCol && cell.col + cell.span >= nCols;
          const types = ['wholeTable'];
          if (!look.noV && !fc && !lc) { const ci = cell.col - (look.firstCol ? 1 : 0); types.push(Math.floor(ci / tp.colBand) % 2 ? 'band2Vert' : 'band1Vert'); }
          if (!look.noH && !fr && !lr) { const ri = r0 - (look.firstRow ? 1 : 0); types.push(Math.floor(ri / tp.rowBand) % 2 ? 'band2Horz' : 'band1Horz'); }
          if (fc) types.push('firstCol'); if (lc) types.push('lastCol'); if (fr) types.push('firstRow'); if (lr) types.push('lastRow');
          if (fr && fc) types.push('nwCell'); if (fr && lc) types.push('neCell'); if (lr && fc) types.push('swCell'); if (lr && lc) types.push('seCell');
          const parts = types.flatMap(cond);
          const tctx = { pEls: [...sch.map(s => s.pPr), ...parts.map(p => p.pPr)].filter(Boolean), rEls: [...sch.map(s => s.rPr), ...parts.map(p => p.rPr)].filter(Boolean) };
          const cb = {}, tcs = [...sch.map(s => s.tcPr), ...parts.map(p => p.tcPr)].filter(Boolean);
          let shd = tp.shd, va = null, mar = { ...tp.mar }, dir = null, noWrap = false;
          const applyTc = e => {
            if (!e) return;
            if (kid(e, 'shd')) shd = shdColor(kid(e, 'shd'));
            if (kid(e, 'vAlign')) va = W_(kid(e, 'vAlign'));
            for (const m of kids(kid(e, 'tcMar'))) mar[m.localName === 'start' ? 'left' : m.localName === 'end' ? 'right' : m.localName] = atNum(m, 'w', 0);
            if (kid(e, 'textDirection')) dir = W_(kid(e, 'textDirection'));
            if (kid(e, 'noWrap')) noWrap = !!onOff(kid(e, 'noWrap'));
          };
          for (const e of tcs) { applyTc(e); bordersOf(kid(e, 'tcBorders'), cb); }
          for (const p of parts) bordersOf(kid(p.tblPr, 'tblBorders'), cb);
          const own = bordersOf(kid(cell.tcPr, 'tcBorders'), {});
          applyTc(cell.tcPr);
          const edge = (name, outer) => {
            if (own[name]) return own[name];
            if (cb[name]) return cb[name];
            const inner = name === 'top' || name === 'bottom' ? 'insideH' : 'insideV';
            if (!outer && cb[inner]) return cb[inner];
            return outer ? tp.bdr[name] : tp.bdr[inner];
          };
          const td = h('td');
          if (cell.span > 1) td.colSpan = cell.span;
          if (cell.rowSpan > 1) td.rowSpan = cell.rowSpan;
          const s = td.style;
          s.borderTop = bcss(edge('top', r0 === 0)); s.borderBottom = bcss(edge('bottom', rEnd === nRows));
          s.borderLeft = bcss(edge('left', cell.col === row.before)); s.borderRight = bcss(edge('right', cell.col + cell.span >= row.end));
          s.padding = `${mpx(mar.top)}px ${mpx(mar.right)}px ${mpx(mar.bottom)}px ${mpx(mar.left)}px`;
          if (shd && shd !== 'none') s.backgroundColor = shd;
          s.verticalAlign = va === 'center' ? 'middle' : va === 'bottom' ? 'bottom' : 'top';
          const box = h('div', 'dx-cell');
          if (dir === 'tbRl' || dir === 'tbRlV') box.style.writingMode = 'vertical-rl';
          else if (dir === 'btLr') box.style.writingMode = 'sideways-lr';
          if (noWrap) box.style.whiteSpace = 'nowrap';
          if (row.hRule === 'exact' && row.h) { box.style.maxHeight = row.h + 'px'; box.style.overflow = 'hidden'; }
          const cw = cols.slice(cell.col, cell.col + cell.span).reduce((a, b) => a + b, 0) - mpx(mar.left) - mpx(mar.right);
          renderBlocks(blockKids(cell.tc), { ...X, tctx, s: X.s && { ...X.s, cw: Math.max(10, cw / hk) } }, box);
          td.appendChild(box);
          tr.appendChild(td);
        }
        if (row.after) { const td = h('td', 'dx-gb'); td.colSpan = row.after; tr.appendChild(td); }
        tb.appendChild(tr);
      }
      const wrap = h('div', 'dx-tw');
      wrap.appendChild(t);
      const left = (tp.ind - (compat >= 15 ? 0 : tp.mar.left)) * hk / TW;
      if (tp.jc === 'center') { t.style.marginLeft = 'auto'; t.style.marginRight = 'auto'; }
      else if (tp.jc === 'right' || tp.jc === 'end') t.style.marginLeft = 'auto';
      else if (left) t.style.marginLeft = left + 'px';
      return { el: wrap, kind: 'tbl', before: 0, after: 0, headerRows, fns: [...wrap.querySelectorAll('.dx-fnref')].map(e => { const [kind, id] = e.dataset.fn.split(':'); return { kind, id }; }) };
    }

    /* ---------- 段落と表を順に並べる ---------- */
    // prevAfter：前の区切り（セクション）の最後の段落の「後の間隔」（区切りをまたいでも、Word は前後の間隔の大きい方だけあける）
    function renderBlocks(els, X, box, prevAfter) {
      const metas = [];
      let prev = prevAfter ? { kind: 'p', after: prevAfter, sid: null } : null;
      for (const el of els) {
        let ms;
        if (el.localName === 'p') ms = renderPara(el, X);
        else if (el.localName === 'tbl') ms = [renderTable(el, X)];
        else if (el.localName === 'altChunk') { const d = h('div', 'dx-p dx-note'); d.textContent = '（ここに埋め込まれた文書は表示できません）'; ms = [{ el: d, kind: 'p', before: 0, after: 0, fns: [] }]; }
        else continue;
        for (const m of ms) {
          if (m.kind === 'p' && !m.cont) {
            let bef = m.before;
            if (prev && prev.kind === 'p') {
              if (prev.sid === m.sid && m.sid != null) {
                if (prev.contextual && prev.after) { prev.el.style.paddingBottom = Math.max(0, parseFloat(prev.el.style.paddingBottom) - prev.after) + 'px'; prev.after = 0; }
                if (m.contextual) bef = 0;
              }
              bef = Math.max(0, bef - prev.after); // Word は前の段落の「後」と次の段落の「前」の大きい方をあける
              if (m.bdr && prev.bdr === m.bdr) { prev.el.style.borderBottom = 'none'; m.el.style.borderTop = m.between ? bcss(m.between) : 'none'; }
            }
            m.el.style.paddingTop = bef + (parseFloat(m.el.style.paddingTop) || 0) + 'px';
            m.before = bef;
          }
          box.appendChild(m.el);
          metas.push(m);
          prev = m;
        }
      }
      return metas;
    }

    /* ---------- 脚注・文末脚注 ---------- */
    const noteRoots = { fn: fnP && fnP.x ? fnP.x.documentElement : null, en: enP && enP.x ? enP.x.documentElement : null };
    const noteSrc = {};
    for (const k of ['fn', 'en']) { const p = k === 'fn' ? fnP : enP; noteSrc[k] = p ? { path: p.path, rels: await rels(p.path) } : null; }
    function renderNote(kind, id, n, s, mode, hk, waits) {
      const box = h('div', 'dx-fn');
      const root = noteRoots[kind], el = root && kids(root).find(e => at(e, 'id') === id);
      if (el) { const X = { src: noteSrc[kind], s, mode, hk, tctx: null, fld: [], wait: waits, noteNo: kind === 'en' ? fmtNum(n, 'lowerRoman') : n }; renderBlocks(blockKids(el), X, box); }
      return box;
    }
    const resetState = () => { counters.clear(); notes.fn.clear(); notes.en.clear(); notes.fnN = notes.enN = 0; };

    // ヘッダー・フッターの部品（先に読んでおく）
    const hfCache = new Map();
    async function loadHF() {
      for (const sec of secs) for (const k of ['hdr', 'ftr']) for (const t in sec.s[k]) {
        const id = sec.s[k][t];
        if (hfCache.has(id)) continue;
        const rel = docRels.get(id), x = rel && await xml(rel.target);
        hfCache.set(id, x ? { x, src: { path: rel.target, rels: await rels(rel.target) } } : null);
      }
    }
    const docSrc = { path: docPath, rels: docRels };
    const paperName = s => {
      const mm = v => v * 25.4 / 96, a = Math.min(mm(s.W), mm(s.H)), b = Math.max(mm(s.W), mm(s.H));
      const P = [['A3', 297, 420], ['A4', 210, 297], ['A5', 148, 210], ['B4', 257, 364], ['B5', 182, 257], ['レター', 216, 279], ['リーガル', 216, 356], ['はがき', 100, 148]];
      const f = P.find(([, x, y]) => Math.abs(a - x) < 3 && Math.abs(b - y) < 3);
      return (f ? f[0] : `${Math.round(mm(s.W))}×${Math.round(mm(s.H))}mm`) + (s.W > s.H + 1 ? ' 横' : '');
    };

    /* ---------- 紙と同じページに分ける ---------- */
    async function paginate(opts = {}) {
      const alive = opts.alive || (() => true), progress = opts.progress || (() => {});
      resetState();
      await loadHF();
      const host = h('div', 'dx-doc dx-host');
      document.body.appendChild(host);
      const waits = [], pages = [];
      let page = null;
      try {
        const sepH = 12;
        const noteCache = new Map();
        const noteEl = (f, s) => {
          const key = f.kind + ':' + f.id;
          if (!noteCache.has(key)) { const e = renderNote(f.kind, f.id, f.n, s, 'paper', 1, waits); e.style.width = s.cw + 'px'; host.appendChild(e); layoutTabs(e); noteCache.set(key, { el: e, h: e.offsetHeight }); }
          return noteCache.get(key);
        };
        const pageNotes = fns => fns.filter(f => f.kind === 'fn').map(f => ({ ...f, n: f.n || notes.fn.get(f.id) }));
        const notesH = (fns, pg) => { const l = pageNotes(fns); if (!l.length) return 0; return (pg.fnList.length ? 0 : sepH) + l.reduce((a, f) => a + noteEl(f, pg.s).h, 0); };
        const renderHF = (id, s) => {
          const box = h('div', 'dx-hf');
          box.style.width = s.cw + 'px';
          const part = id && hfCache.get(id);
          if (part) { const X = { src: part.src, s, mode: 'paper', hk: 1, tctx: null, fld: [], wait: waits, hdr: true }; renderBlocks(blockKids(part.x.documentElement), X, box); }
          host.appendChild(box);
          return box;
        };
        function newPage(s, si, natural, secStart) {
          let num = page ? page.num + 1 : 1;
          if (secStart && s.pnStart != null) num = s.pnStart;
          if (secStart && page && ((s.type === 'evenPage' && num % 2) || (s.type === 'oddPage' && !(num % 2)))) { // 空白のページを入れる
            pages.push(page = { s, si, num, frags: [], used: 0, fnList: [], fnH: 0, blank: true, top: s.top, bodyH: s.H - s.top - s.bottom, hdrEl: renderHF(null, s), ftrEl: renderHF(null, s) });
            num++;
          }
          const type = secStart && s.titlePg ? 'first' : evenOdd && num % 2 === 0 ? 'even' : 'default';
          const pick = k => (type === 'default' ? s[k].default : s[k][type]) || null;
          const hdrEl = renderHF(pick('hdr'), s), ftrEl = renderHF(pick('ftr'), s);
          for (const e of [hdrEl, ftrEl]) for (const f of e.querySelectorAll('[data-f="PAGE"]')) f.textContent = fmtNum(num, s.pnFmt) || String(num);
          layoutTabs(hdrEl); layoutTabs(ftrEl);
          const top = Math.max(s.top, hdrEl.firstChild ? s.header + hdrEl.offsetHeight : 0), bottom = Math.max(s.bottom, ftrEl.firstChild ? s.footer + ftrEl.offsetHeight : 0);
          page = { s, si, num, frags: [], used: 0, fnList: [], fnH: 0, natural, secStart: !!secStart, top, bodyH: Math.max(40, s.H - top - bottom), hdrEl, ftrEl };
          pages.push(page);
        }
        function addFrag(m, off, end, fns, hdr) {
          page.frags.push({ m, off, end, fns, hdr });
          page.used += end - off + (hdr ? m.hdrH : 0);
          for (const f of pageNotes(fns)) { const ne = noteEl(f, page.s); if (!page.fnList.length) page.fnH += sepH; page.fnList.push({ ...f, el: ne.el }); page.fnH += ne.h; }
        }
        function popChain() {
          const out = [];
          while (page.frags.length > 1) {
            const f = page.frags[page.frags.length - 1];
            if (!(f.m.keepNext && f.off < 0.5 && f.end >= f.m.h - 0.5)) break;
            page.frags.pop(); page.used -= f.end - f.off;
            for (const n of pageNotes(f.fns)) { const i = page.fnList.findIndex(x => x.id === n.id); if (i >= 0) { page.fnH -= noteEl(n, page.s).h; page.fnList.splice(i, 1); } }
            if (!page.fnList.length) page.fnH = 0;
            out.unshift(f);
          }
          return out;
        }
        function placeBlock(m) {
          const end = m.h, fitEnd = end - (m.after || 0);
          let off = 0, fns = m.fns || [], cuts = null;
          const firstOff = () => (m.kind === 'p' && !m.cont ? Math.min(m.before || 0, fitEnd) : 0);
          if (!page.frags.length && page.natural) off = firstOff();
          for (let guard = 0; guard < 2000; guard++) {
            const hdr = m.kind === 'tbl' && off > 0.5 && m.hdrH > 0 && off > m.hdrEnd;
            const room = page.bodyH - page.used - page.fnH - notesH(fns, page) - (hdr ? m.hdrH : 0);
            if (fitEnd - off <= room + 0.5) { addFrag(m, off, end, fns, hdr); return; }
            if (!cuts) cuts = cutsOf(m.el).filter(c => c.y > 0.5 && c.y < fitEnd - 0.5);
            const startN = (cuts.find(c => Math.abs(c.y - off) < 0.6) || { n: 0 }).n;
            const fresh = off <= firstOff() + 0.5;
            let best = null;
            if (!(m.keepLines && fresh && page.frags.length)) {
              for (const c of cuts) {
                if (c.y <= off + 0.5) continue;
                if (c.y - off > room + 0.5) break;
                if (m.kind === 'p' && m.widow && c.of) { if (fresh && c.n - startN < 2) continue; if (c.of - c.n < 2) continue; }
                if (m.kind === 'tbl' && c.row != null && c.row <= m.headerRows) continue;
                best = c;
              }
            }
            if (best) { addFrag(m, off, best.y, fns, hdr); fns = []; off = best.y; newPage(page.s, page.si, true); continue; }
            if (page.frags.length) {
              const moved = fresh ? popChain() : [];
              newPage(page.s, page.si, true);
              for (const f of moved) { const ro = f.m.kind === 'p' && !f.m.cont ? Math.min(f.m.before || 0, f.m.h) : 0; addFrag(f.m, page.frags.length ? 0 : ro, f.m.h, f.fns); }
              if (fresh && !page.frags.length) off = firstOff();
              continue;
            }
            const c = cuts.find(x => x.y > off + 0.5);
            if (!c) { addFrag(m, off, end, fns, hdr); return; }
            addFrag(m, off, c.y, fns, hdr); fns = []; off = c.y; newPage(page.s, page.si, true);
          }
        }

        let lastAfter = 0;
        for (let si = 0; si < secs.length; si++) {
          if (!alive()) return null;
          progress(`ページに分けています… ${si + 1} / ${secs.length}`);
          const sec = secs[si], s = sec.s;
          const col = h('div', 'dx-col');
          col.style.width = s.cw + 'px';
          host.appendChild(col);
          const X = { src: docSrc, s, mode: 'paper', hk: 1, tctx: null, fld: [], wait: waits };
          const metas = renderBlocks(sec.els, X, col, lastAfter);
          { const lm = metas[metas.length - 1]; lastAfter = lm && lm.kind === 'p' ? lm.after || 0 : 0; }
          if (si === secs.length - 1 && notes.en.size) {
            const sep = h('div', 'dx-ensep');
            col.appendChild(sep);
            metas.push({ el: sep, kind: 'x', fns: [] });
            for (const [id, n] of notes.en) { const e = renderNote('en', id, n, s, 'paper', 1, waits); col.appendChild(e); metas.push({ el: e, kind: 'x', fns: [] }); }
          }
          layoutTabs(col);
          await new Promise(r => setTimeout(r));
          const ct = col.getBoundingClientRect().top;
          for (const m of metas) {
            const r = m.el.getBoundingClientRect();
            m.h = r.height; m.top = r.top - ct;
            if (m.kind === 'tbl' && m.headerRows) { const rows = m.el.querySelector('table').rows, hr = rows[m.headerRows - 1].getBoundingClientRect(); m.hdrEnd = hr.bottom - r.top; m.hdrH = hr.bottom - rows[0].getBoundingClientRect().top; }
          }
          const same = page && Math.abs(page.s.W - s.W) < 1 && Math.abs(page.s.H - s.H) < 1 && Math.abs(page.s.cw - s.cw) < 1;
          if (!page || s.type !== 'continuous' || !same) newPage(s, si, false, true);
          else page.s = s;
          // ページの頭の段落の「前の間隔」は、自然な改ページ・改ページ・「段落前で改ページ」では消し、区切り（次のページから）では残す（Word で確かめた）
          for (const m of metas) {
            if (m.pbb && page.frags.length) newPage(s, si, true);
            placeBlock(m);
            if (m.brk === 'page') newPage(s, si, true);
          }
        }
        if (!alive()) return null;
        progress('仕上げています…');
        await Promise.all(waits);
        // ページを組み立てる
        const total = pages.length;
        const secCount = {};
        for (const p of pages) secCount[p.si] = (secCount[p.si] || 0) + 1;
        const out = [];
        for (const pg of pages) {
          const s = pg.s, el = h('div', 'dx-doc dx-page', { width: s.W + 'px', height: s.H + 'px' });
          if (pageBg) el.style.background = pageBg;
          for (const e of [pg.hdrEl, pg.ftrEl]) {
            for (const f of e.querySelectorAll('[data-f="NUMPAGES"]')) f.textContent = String(total);
            for (const f of e.querySelectorAll('[data-f="SECTIONPAGES"]')) f.textContent = String(secCount[pg.si]);
          }
          Object.assign(pg.hdrEl.style, { left: s.left + 'px', top: s.header + 'px' });
          Object.assign(pg.ftrEl.style, { left: s.left + 'px', bottom: s.footer + 'px' });
          const bodyEl = h('div', 'dx-body', { left: s.left + 'px', top: pg.top + 'px', width: s.cw + 'px', height: pg.bodyH + 'px' });
          for (const f of pg.frags) {
            const m = f.m, whole = f.off < 0.5 && f.end >= m.h - 0.5;
            if (whole && !m.taken) { m.taken = true; bodyEl.appendChild(m.el); continue; }
            if (f.hdr) {
              const hc = m.el.cloneNode(true), rows = [...hc.querySelector('table').rows];
              rows.slice(m.headerRows).forEach(r => r.remove());
              hc.querySelectorAll('.dx-abs').forEach(e => e.remove());
              bodyEl.appendChild(hc);
            }
            const wrap = h('div', 'dx-frag', { height: Math.max(0, f.end - f.off) + 'px' });
            let inner;
            if (!m.taken) { inner = m.el; m.taken = true; } else { inner = m.el.cloneNode(true); inner.querySelectorAll('.dx-abs').forEach(e => e.remove()); }
            inner.style.position = 'relative'; inner.style.top = -f.off + 'px';
            wrap.appendChild(inner);
            bodyEl.appendChild(wrap);
          }
          el.append(pg.hdrEl, bodyEl, pg.ftrEl);
          if (pg.fnList.length) {
            const fe = h('div', 'dx-fns', { left: s.left + 'px', width: s.cw + 'px', top: pg.top + pg.bodyH - pg.fnH + 'px' });
            fe.appendChild(h('div', 'dx-fnsep'));
            for (const f of pg.fnList) fe.appendChild(f.el.parentNode === host ? f.el : f.el.cloneNode(true));
            el.appendChild(fe);
          }
          if (s.bdr) {
            const b = s.bdr, byText = b.from !== 'page', g = k => (b[k] && !b[k].none ? b[k].space : 0);
            const box = h('div', 'dx-pbdr', byText
              ? { left: s.left - g('left') + 'px', top: pg.top - g('top') + 'px', width: s.cw + g('left') + g('right') + 'px', height: pg.bodyH + g('top') + g('bottom') + 'px' }
              : { left: g('left') + 'px', top: g('top') + 'px', right: g('right') + 'px', bottom: g('bottom') + 'px' });
            for (const k of ['top', 'bottom', 'left', 'right']) box.style['border' + k[0].toUpperCase() + k.slice(1)] = bcss(b[k]);
            el.appendChild(box);
          }
          host.appendChild(el);
          layoutTabs(pg.hdrEl); layoutTabs(pg.ftrEl);
          positionAnchors(el, pg);
          el.remove();
          out.push({ el, w: s.W, h: s.H });
        }
        return out;
      } finally { host.remove(); }
    }
    // 浮いている図を、ページの上の位置に置く
    function positionAnchors(el, pg) {
      const s = pg.s, pr = el.getBoundingClientRect();
      const al = (a, area, w) => (a === 'center' ? (area - w) / 2 : a === 'right' || a === 'bottom' || a === 'outside' ? area - w : 0);
      for (const a of [...el.querySelectorAll('.dx-abs')]) {
        const p = a._pos;
        if (!p || a.closest('.dx-tbx')) continue;
        const para = a.parentElement, r = para.getBoundingClientRect();
        let x, y;
        switch (p.hRel) {
          case 'page': x = p.hAl ? al(p.hAl, s.W, p.W) : p.hOff; break;
          case 'leftMargin': x = p.hAl ? al(p.hAl, s.left, p.W) : p.hOff; break;
          case 'rightMargin': x = s.W - s.right + (p.hAl ? al(p.hAl, s.right, p.W) : p.hOff); break;
          default: x = s.left + (p.hAl ? al(p.hAl, s.cw, p.W) : p.hOff);
        }
        switch (p.vRel) {
          case 'page': y = p.vAl ? al(p.vAl, s.H, p.H) : p.vOff; break;
          case 'margin': y = s.top + (p.vAl ? al(p.vAl, s.H - s.top - s.bottom, p.H) : p.vOff); break;
          case 'topMargin': y = p.vAl ? al(p.vAl, s.top, p.H) : p.vOff; break;
          case 'bottomMargin': y = s.H - s.bottom + (p.vAl ? al(p.vAl, s.bottom, p.H) : p.vOff); break;
          default: y = r.top - pr.top + (parseFloat(para.style.paddingTop) || 0) + p.vOff;
        }
        a.style.left = x + 'px'; a.style.top = y + 'px';
        el.appendChild(a);
      }
    }

    /* ---------- スマホ向け：画面の幅で折り返す ---------- */
    async function flow(width, opts = {}) {
      resetState();
      const root = h('div', 'dx-doc dx-flow'), waits = [];
      secs.forEach((sec, si) => {
        const hk = Math.min(1, Math.max(0.3, width / sec.s.cw));
        const X = { src: docSrc, s: sec.s, mode: 'flow', hk, tctx: null, fld: [], wait: waits };
        const metas = renderBlocks(sec.els, X, root);
        for (const m of metas) if (m.brk === 'page' || m.pbb) { const hr = h('div', 'dx-brk'); if (m.pbb) root.insertBefore(hr, m.el); else m.el.after(hr); }
        if (si < secs.length - 1) root.appendChild(h('div', 'dx-brk'));
      });
      const s0 = secs[0].s, hk0 = Math.min(1, width / s0.cw);
      for (const [kind, map] of [['fn', notes.fn], ['en', notes.en]]) {
        if (!map.size) continue;
        const box = h('div', 'dx-notes');
        box.appendChild(h('div', 'dx-notes-h')).textContent = kind === 'fn' ? '脚注' : '文末脚注';
        for (const [id, n] of map) box.appendChild(renderNote(kind, id, n, s0, 'flow', hk0, waits));
        root.appendChild(box);
      }
      await Promise.all(waits);
      return root;
    }

    const s0 = secs[0].s;
    return {
      info: { pageW: s0.W, pageH: s0.H, paper: paperName(s0), sections: secs.length, vert: secs.some(x => x.s.vert), cols: secs.some(x => x.s.cols > 1), comments: !!relOfType(docRels, 'comments') },
      paginate, flow,
      dispose() { pk.dispose(); },
    };
  }

  // 見た目（文書の中だけで使う）
  if (typeof document !== 'undefined' && !document.getElementById('docx-css')) {
    const st = document.createElement('style');
    st.id = 'docx-css';
    st.textContent = `
.dx-doc{color:#000;background:#fff;text-align:left;-webkit-text-size-adjust:none;text-size-adjust:none;font-kerning:none;font-size:14px;line-height:normal}
.dx-doc *{box-sizing:border-box}
.dx-host{position:fixed;left:-40000px;top:0;visibility:hidden;pointer-events:none}
.dx-host>*{position:absolute;left:0;top:0}
.dx-page{position:relative;overflow:hidden;transform-origin:0 0}
.dx-body,.dx-hf,.dx-fns,.dx-pbdr{position:absolute}
.dx-body,.dx-hf{z-index:1}
.dx-pbdr{pointer-events:none;z-index:2}
.dx-p{white-space:pre-wrap;overflow-wrap:break-word;word-break:normal;line-break:normal;margin:0;position:relative;text-autospace:no-autospace;text-spacing-trim:normal}
.dx-as{letter-spacing:calc(var(--ls, 0px) + .25em)}
.dx-nt{text-spacing-trim:space-all}
.dx-p math{line-height:var(--lh)}
.dx-frag{overflow-x:visible;overflow-y:clip;position:relative}
.dx-tab{display:inline-block;height:0;white-space:pre;text-indent:0}
.dx-jbr{display:inline-block;width:100%;height:0;vertical-align:top}
.dx-tab[data-ld=dot],.dx-tab[data-ld=middleDot]{border-bottom:2px dotted currentColor}
.dx-tab[data-ld=hyphen]{border-bottom:1px dashed currentColor}
.dx-tab[data-ld=underscore]{border-bottom:1px solid currentColor}
.dx-tab[data-ld=heavy]{border-bottom:2px solid currentColor}
.dx-inl{display:inline-block;position:relative;vertical-align:baseline;line-height:normal;text-indent:0;white-space:normal}
.dx-inl.dx-abs{position:absolute}
.dx-pic{position:absolute;inset:0;overflow:hidden;display:block}
.dx-pic img{position:absolute;max-width:none;display:block}
.dx-shp,.dx-grp{position:absolute;inset:0;display:block}
.dx-gi{position:absolute;display:block}
.dx-shp>svg{position:absolute;left:0;top:0;overflow:visible}
.dx-tbx{position:absolute;display:flex;flex-direction:column;overflow:hidden;text-indent:0;text-align:left}
.dx-miss{position:absolute;inset:0;display:flex;align-items:center;justify-content:center;text-align:center;border:1px dashed #b0b6c0;background:rgba(240,242,245,.85);color:#5b6474;font:12px/1.3 sans-serif;white-space:normal;text-indent:0;padding:2px;overflow:hidden}
.dx-tw{position:relative}
.dx-tbl{border-collapse:collapse;table-layout:fixed;text-indent:0}
.dx-tbl td{vertical-align:top}
.dx-gb{border:none!important;padding:0!important}
.dx-fit{display:inline-block;text-align:justify;text-align-last:justify;white-space:nowrap;overflow:hidden;vertical-align:bottom;text-indent:0}
.dx-ruby rt{font-size:50%;line-height:1}
.dx-a{color:inherit;text-decoration:inherit}
.dx-fnsep,.dx-ensep{width:33%;height:0;border-top:1px solid #000;margin:5px 0 6px}
.dx-ensep{margin-top:14px}
.dx-flow{background:#fff}
.dx-flow .dx-tw{overflow-x:auto;max-width:100%}
.dx-blk{display:block;margin:6px auto;position:relative}
.dx-tbxb{border:1px solid #c8ccd4;border-radius:6px;padding:4px 8px}
.dx-tbxb .dx-shp{position:relative;inset:auto}
.dx-tbxb .dx-shp>svg{display:none}
.dx-tbxb .dx-tbx{position:static;width:auto!important;height:auto!important;padding:0!important;overflow:visible;color:inherit!important}
.dx-brk{border-top:2px dashed #c8ccd4;margin:16px 0}
.dx-notes{border-top:1px solid #c8ccd4;margin-top:16px;padding-top:8px}
.dx-notes-h{font:600 12px sans-serif;color:#5b6474;margin-bottom:4px}
.dx-note{color:#5b6474;font:13px sans-serif}`;
    document.head.appendChild(st);
  }

  return { open, layoutTabs };
})();
