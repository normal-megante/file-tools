'use strict';
// =========================================================
// Excel の数値の表示形式（自作）。画面に依存しないので Node でもテストできる
//   NumFmt.format(value, code, opts) → { text, parts, color, kind }
//     value：数（日付も数）・文字・true / false
//     code ：表示形式（"#,##0.00" など）。番号だけのものは NumFmt.builtin(id) で文字にする
//     opts ：{ date1904, fit(text) }。fit は「その文字がセルの幅に入るか」。
//            General（G/標準）は入るまで小数を減らし、ほかの数は入らなければ ### にする
//   parts：[{ s: 字 } | { pad: 'x' }（x の幅の空白） | { fill: 'x' }（x をくり返して幅いっぱいに）]
//   text ：parts をつないだもの（pad は空白、fill は空）。Excel の Range.Text と比べるのに使う
//   kind ：'number' | 'date' | 'text' | 'general'（右に寄せるかなどに使う）
// =========================================================
const NumFmt = (() => {
  // 番号だけで書かれる表示形式（日本語の Excel のもの。この PC の Excel で確かめた）
  const BUILTIN = {
    0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00',
    5: '"¥"#,##0;"¥"\\-#,##0', 6: '"¥"#,##0;[Red]"¥"\\-#,##0', 7: '"¥"#,##0.00;"¥"\\-#,##0.00', 8: '"¥"#,##0.00;[Red]"¥"\\-#,##0.00',
    9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
    14: 'yyyy/m/d', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'h:mm', 21: 'h:mm:ss', 22: 'yyyy/m/d h:mm',
    27: '[$-411]ge.m.d', 28: '[$-411]ggge"年"m"月"d"日"', 29: '[$-411]ggge"年"m"月"d"日"', 30: 'm/d/yy', 31: 'yyyy"年"m"月"d"日"',
    32: 'h"時"mm"分"', 33: 'h"時"mm"分"ss"秒"', 34: 'yyyy"年"m"月"', 35: 'm"月"d"日"', 36: '[$-411]ge.m.d',
    37: '#,##0;-#,##0', 38: '#,##0;[Red]-#,##0', 39: '#,##0.00;-#,##0.00', 40: '#,##0.00;[Red]-#,##0.00',
    41: '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"_ ;_ @_ ', 42: '_ "¥"* #,##0_ ;_ "¥"* \\-#,##0_ ;_ "¥"* "-"_ ;_ @_ ',
    43: '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ ', 44: '_ "¥"* #,##0.00_ ;_ "¥"* \\-#,##0.00_ ;_ "¥"* "-"??_ ;_ @_ ',
    45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
    50: '[$-411]ge.m.d', 51: '[$-411]ggge"年"m"月"d"日"', 52: 'yyyy"年"m"月"', 53: 'm"月"d"日"', 54: '[$-411]ggge"年"m"月"d"日"',
    55: 'yyyy"年"m"月"', 56: 'm"月"d"日"', 57: '[$-411]ge.m.d', 58: '[$-411]ggge"年"m"月"d"日"',
  };
  const builtin = id => BUILTIN[id] != null ? BUILTIN[id] : 'General';

  const COLORS = { black: '#000000', blue: '#0000FF', cyan: '#00FFFF', green: '#00FF00', magenta: '#FF00FF', red: '#FF0000', white: '#FFFFFF', yellow: '#FFFF00',
    黒: '#000000', 青: '#0000FF', 水: '#00FFFF', 緑: '#00FF00', 紫: '#FF00FF', 赤: '#FF0000', 白: '#FFFFFF', 黄: '#FFFF00' };
  // [ColorN] の色（Excel の標準の 56 色）
  const PALETTE = ['000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080',
    '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF', '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF',
    '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99', '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696',
    '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333'];

  /* ---------- 表示形式を部品に分ける ---------- */
  const cache = new Map();
  function compile(code) {
    let c = cache.get(code);
    if (!c) {
      c = parse(code);
      if (cache.size > 2000) cache.clear();
      cache.set(code, c);
    }
    return c;
  }
  function parse(code) {
    const secs = [];
    let cur = newSec(), i = 0;
    const s = String(code == null ? 'General' : code);
    const push = tk => cur.tokens.push(tk);
    while (i < s.length) {
      const ch = s[i];
      if (ch === '"') { const j = s.indexOf('"', i + 1); const e = j < 0 ? s.length : j; push({ t: 'lit', s: s.slice(i + 1, e) }); i = e + 1; continue; }
      if (ch === '\\' || ch === '!') { if (i + 1 < s.length) push({ t: 'lit', s: s[i + 1] }); i += 2; continue; }
      if (ch === '_') { if (i + 1 < s.length) push({ t: 'pad', ch: s[i + 1] }); i += 2; continue; }
      if (ch === '*') { if (i + 1 < s.length) push({ t: 'fill', ch: s[i + 1] }); i += 2; continue; }
      if (ch === ';') { secs.push(cur); cur = newSec(); i++; continue; }
      if (ch === '[') {
        const j = s.indexOf(']', i);
        if (j < 0) { push({ t: 'lit', s: s.slice(i) }); break; }
        bracket(cur, s.slice(i + 1, j));
        i = j + 1;
        continue;
      }
      const rest = s.slice(i);
      let m;
      if ((m = /^(General|G\/標準)/i.exec(rest))) { push({ t: 'gen' }); i += m[0].length; continue; }
      if ((m = /^(AM\/PM|am\/pm|A\/P|a\/p|午前\/午後)/.exec(rest))) { push({ t: 'ampm', s: m[0] }); i += m[0].length; continue; }
      if (ch === '0' || ch === '#' || ch === '?') { push({ t: 'd', ch }); i++; continue; }
      if (ch === '.') { push({ t: '.' }); i++; continue; }
      if (ch === ',') { push({ t: ',' }); i++; continue; }
      if (ch === '%') { push({ t: '%' }); i++; continue; }
      if (ch === '@') { push({ t: '@' }); i++; continue; }
      if ((ch === 'E' || ch === 'e') && (s[i + 1] === '+' || s[i + 1] === '-')) { push({ t: 'E', sign: s[i + 1] }); i += 2; continue; }
      if (ch === '/') {
        // 分数の分母が決まった数（?/8・0 ??/100）
        const dm = /^\/([1-9]\d*)/.exec(rest);
        if (dm) { push({ t: '/' }); push({ t: 'den', n: +dm[1] }); i += dm[0].length; continue; }
        push({ t: '/' }); i++; continue;
      }
      const run = re => { let j = i; while (j < s.length && re.test(s[j])) j++; return j - i; };
      if (/[yY]/.test(ch)) { const n = run(/[yY]/); push({ t: 'y', n }); i += n; continue; }
      if (/[mM]/.test(ch)) { const n = run(/[mM]/); push({ t: 'm', n }); i += n; continue; }
      if (/[dD]/.test(ch)) { const n = run(/[dD]/); push({ t: 'dd', n }); i += n; continue; }
      if (/[hH]/.test(ch)) { const n = run(/[hH]/); push({ t: 'h', n }); i += n; continue; }
      if (/[sS]/.test(ch)) { const n = run(/[sS]/); push({ t: 's', n }); i += n; continue; }
      if (ch === 'g' || ch === 'G') { const n = run(/[gG]/); push({ t: 'g', n }); i += n; continue; }
      if (ch === 'e') { const n = run(/e/); push({ t: 'e', n }); i += n; continue; }
      if (ch === 'r' || ch === 'R') { const n = run(/[rR]/); push({ t: 'r', n }); i += n; continue; }
      if (ch === 'a' || ch === 'A') { const n = run(/[aA]/); if (n >= 3) { push({ t: 'a', n }); i += n; continue; } }
      if (ch === 'b' || ch === 'B') { const n = run(/[bB]/); push({ t: 'b', n }); i += n; continue; }
      push({ t: 'lit', s: ch });
      i++;
    }
    secs.push(cur);
    for (const sec of secs) finish(sec);
    return secs;
  }
  function newSec() { return { tokens: [], cond: null, color: null, dbnum: 0, locale: null, elapsed: false }; }
  function bracket(sec, b) {
    let m;
    if ((m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+(?:E[+-]?\d+)?)$/i.exec(b))) { sec.cond = { op: m[1], v: +m[2] }; return; }
    const low = b.toLowerCase();
    if (COLORS[low]) { sec.color = COLORS[low]; return; }
    if (COLORS[b]) { sec.color = COLORS[b]; return; }
    if ((m = /^(?:Color|色)\s*(\d+)$/i.exec(b))) { const k = +m[1]; sec.color = '#' + (PALETTE[k - 1] || '000000'); return; }
    if ((m = /^DBNum(\d)$/i.exec(b))) { sec.dbnum = +m[1]; return; }
    if (b[0] === '$') {
      const k = b.indexOf('-'), sym = k < 0 ? b.slice(1) : b.slice(1, k);
      if (sym) sec.tokens.push({ t: 'lit', s: sym });
      sec.locale = k < 0 ? null : b.slice(k + 1).toLowerCase();
      return;
    }
    if ((m = /^(h+|m+|s+)$/i.exec(b))) { sec.tokens.push({ t: 'el', u: m[1][0].toLowerCase(), n: m[1].length }); sec.elapsed = true; }
  }
  // 部品を見て、日付か・数か・文字かを決める。m（月・分）や秒の小数もここで決める
  function finish(sec) {
    const tk = sec.tokens;
    const isDate = tk.some(x => /^(y|m|dd|h|s|ampm|el|g|e|r|a|b)$/.test(x.t));
    sec.hasAt = tk.some(x => x.t === '@');
    sec.gen = tk.some(x => x.t === 'gen');
    sec.kind = sec.hasAt ? 'text' : sec.gen ? 'general' : isDate ? 'date' : tk.some(x => x.t === 'd') ? 'number' : 'literal';
    if (sec.kind === 'date') {
      // 分の m：時（h）のすぐあとか、秒（s）のすぐ前
      const dt = tk.map((x, k) => ({ x, k })).filter(o => /^(y|m|dd|h|s|el|ampm|g|e|r|a|b)$/.test(o.x.t));
      dt.forEach((o, j) => {
        if (o.x.t !== 'm' || o.x.n > 2) return;
        const p = dt[j - 1], n = dt[j + 1];
        if ((p && (p.x.t === 'h' || (p.x.t === 'el' && p.x.u === 'h'))) || (n && (n.x.t === 's' || (n.x.t === 'el' && n.x.u === 's')))) o.x.t = 'min';
      });
      // 秒のあとの .0 / .00 / .000 は秒の小数
      for (let k = 0; k < tk.length; k++) {
        if ((tk[k].t === 's' || (tk[k].t === 'el' && tk[k].u === 's')) && tk[k + 1] && tk[k + 1].t === '.') {
          let n = 0;
          while (tk[k + 2 + n] && tk[k + 2 + n].t === 'd' && tk[k + 2 + n].ch === '0') n++;
          if (n) { tk.splice(k + 1, n + 1, { t: 'sf', n }); sec.frac = Math.max(sec.frac || 0, n); }
        }
      }
      sec.ampm = tk.some(x => x.t === 'ampm');
      // 日付の中の数字の 0 などは、ただの字
      for (const x of tk) if (x.t === 'd' || x.t === '.' || x.t === ',' || x.t === '%' || x.t === '/') { x.s = x.t === 'd' ? x.ch : x.t; x.t = 'lit'; }
      for (const x of tk) if (x.t === 'den') { x.s = String(x.n); x.t = 'lit'; }
      for (const x of tk) if (x.t === 'E') { x.s = 'E' + x.sign; x.t = 'lit'; }
    }
    if (sec.kind === 'number') analyzeNumber(sec);
    if (sec.kind === 'text' || sec.kind === 'literal' || sec.kind === 'general') {
      for (const x of tk) if (x.t === 'd' || x.t === '.' || x.t === ',' || x.t === '%' || x.t === '/') { x.s = x.t === 'd' ? x.ch : x.t; x.t = 'lit'; }
    }
  }
  function analyzeNumber(sec) {
    const tk = sec.tokens;
    const e = tk.findIndex(x => x.t === 'E');
    const slash = tk.findIndex(x => x.t === '/');
    sec.pct = tk.filter(x => x.t === '%').length;
    // 最後の数字の場所より後ろのカンマは「1000 で割る」
    let lastD = -1;
    tk.forEach((x, k) => { if (x.t === 'd' && (e < 0 || k < e)) lastD = k; });
    let scale = 0;
    for (let k = lastD + 1; k < tk.length && tk[k].t === ','; k++) { scale++; tk[k].t = 'nil'; }
    // 小数点の直前のカンマも 1000 で割る（#,##0,.0 のような形）
    const dot = tk.findIndex(x => x.t === '.');
    if (dot > 0) for (let k = dot - 1; k >= 0 && tk[k].t === ','; k--) { scale++; tk[k].t = 'nil'; }
    sec.scale = scale;
    sec.group = tk.some((x, k) => x.t === ',' && (dot < 0 || k < dot));
    for (const x of tk) if (x.t === ',') x.t = 'nil';
    if (e >= 0) { sec.sci = true; sec.e = e; }
    else if (slash >= 0) {
      // 分数：分子の数字の並び・分母
      sec.frac = true;
      let k = slash - 1;
      while (k >= 0 && tk[k].t !== 'd') k--;
      let ns = k;
      while (ns - 1 >= 0 && tk[ns - 1].t === 'd') ns--;
      sec.num = { from: ns, to: k };
      const intD = tk.slice(0, ns).map((x, j) => ({ x, j })).filter(o => o.x.t === 'd');
      sec.whole = intD.length ? { from: intD[0].j, to: intD[intD.length - 1].j } : null;
      const den = tk[slash + 1];
      if (den && den.t === 'den') sec.den = { fixed: den.n, at: slash + 1 };
      else { let a = slash + 1; while (tk[a] && tk[a].t !== 'd') a++; let b = a; while (tk[b + 1] && tk[b + 1].t === 'd') b++; sec.den = tk[a] ? { from: a, to: b } : { fixed: 1, at: slash }; }
    }
  }

  /* ---------- 数の字（15 けたで丸める。Excel と同じく 0.5 は切り上げ） ---------- */
  // x（0 以上）→ { d: 数字の並び, e: 整数部のけた数 }（x = 0.d × 10^e）
  function digits(x) {
    if (!x) return { d: '', e: 0 };
    const s = x.toPrecision(15);
    let [m, ex] = s.split('e');
    const [a, b = ''] = m.split('.');
    let d = a + b, e = a.length + (ex ? +ex : 0);
    while (d.length > 1 && d[0] === '0') { d = d.slice(1); e--; }
    d = d.replace(/0+$/, '');
    if (!d) return { d: '', e: 0 };
    return { d, e };
  }
  // { d, e } を小数 n けたで丸める → { i: 整数部の字（0 なら ''）, f: 小数部の字（n けた） }
  function roundAt(D, n) {
    const keep = D.e + n;
    if (keep < 0 || !D.d) return { i: '', f: '0'.repeat(n) };
    let ds = D.d.padEnd(keep + 1, '0');
    let head = ds.slice(0, keep), e = D.e;
    if (ds[keep] >= '5') {
      const arr = head.split('');
      let k = arr.length - 1;
      while (k >= 0) { if (arr[k] === '9') { arr[k] = '0'; k--; } else { arr[k] = String(+arr[k] + 1); break; } }
      head = arr.join('');
      if (k < 0) { head = '1' + head; e++; }
    }
    if (e < 0) { head = '0'.repeat(-e) + head; e = 0; }
    let i = head.slice(0, e), f = head.slice(e);
    i = i.replace(/^0+/, '');
    f = f.padEnd(n, '0');
    return { i, f };
  }
  const group3 = s => s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

  /* ---------- 選ぶ（正・負・ゼロ・文字の区切り） ---------- */
  function pick(secs, v) {
    if (typeof v === 'string') {
      if (secs.length >= 4) return { sec: secs[3] };
      const t = secs.find(x => x.hasAt);
      return t ? { sec: t } : null;
    }
    const nums = secs.length >= 4 ? secs.slice(0, 3) : secs.filter((x, k) => !(k > 0 && x.hasAt));
    if (nums.some(x => x.cond)) {
      const ok = c => { switch (c.op) { case '<': return v < c.v; case '>': return v > c.v; case '=': return v === c.v; case '<=': return v <= c.v; case '>=': return v >= c.v; case '<>': return v !== c.v; } return false; };
      if (nums[0].cond && ok(nums[0].cond)) return { sec: nums[0], minus: v < 0 && nums[0].cond.v >= 0 && !(nums[0].cond.op === '<' || nums[0].cond.op === '<=') };
      if (nums[1]) {
        if (nums[1].cond) {
          if (ok(nums[1].cond)) return { sec: nums[1], minus: v < 0 && nums[1].cond.v >= 0 && !(nums[1].cond.op === '<' || nums[1].cond.op === '<=') };
          return nums[2] ? { sec: nums[2], minus: v < 0 } : { sec: null };
        }
        return { sec: nums[1], minus: v < 0 && (nums.length >= 3 || !nums[0].cond || nums[0].cond.v < 0 || (nums[0].cond.op !== '<' && nums[0].cond.op !== '<=')) };
      }
      return { sec: null };
    }
    if (nums.length === 1 || v > 0 || (v === 0 && nums.length === 2)) return { sec: nums[0], minus: v < 0 };
    if (v < 0) return { sec: nums[1], minus: false };
    return { sec: nums[2] || nums[0], minus: false };
  }

  /* ---------- 形にする ---------- */
  function format(value, code, opts = {}) {
    if (value == null || value === '') return out([], null, 'text');
    if (typeof value === 'boolean') return out([{ s: value ? 'TRUE' : 'FALSE' }], null, 'bool');
    const secs = compile(code);
    if (typeof value === 'string') {
      const p = pick(secs, value);
      if (!p || !p.sec) return out([{ s: value }], null, 'text');
      return out(textSec(p.sec, value), p.sec.color, 'text');
    }
    if (!isFinite(value)) return out([{ s: '#NUM!' }], null, 'error');
    const p = pick(secs, value);
    if (!p.sec) return hashes(opts);
    // 文字用の形式（@）の数は General で出す
    const sec = p.sec.hasAt ? compile('General')[0] : p.sec;
    let parts;
    if (sec.kind === 'date') {
      parts = dateSec(sec, value, opts);
      if (!parts) return hashes(opts);
    } else if (sec.kind === 'general' && sec.dbnum) {
      const s = dbGeneral(p.minus ? value : Math.abs(value), sec.dbnum);
      parts = sec.tokens.map(x => (x.t === 'gen' ? { s } : x.t === 'lit' ? { s: x.s } : x.t === 'pad' ? { pad: x.ch } : x.t === 'fill' ? { fill: x.ch } : { s: '' }));
      const r = out(parts, sec.color, 'general');
      if (opts.fit && !opts.fit(r.text)) return hashes(opts);
      return r;
    } else if (sec.kind === 'general') {
      parts = generalSec(sec, value, p.minus, opts);
      if (!parts) return hashes(opts);
      return out(parts, sec.color, 'general');
    } else if (sec.kind === 'number') {
      parts = numberSec(sec, value, p.minus);
    } else {
      parts = litSec(sec, p.minus);
    }
    let r = out(parts, sec.color, sec.kind === 'date' ? 'date' : 'number');
    if (sec.dbnum) r = dbnum(r, sec.dbnum);
    if (opts.fit && !opts.fit(r.text)) return hashes(opts);
    return r;
  }
  function out(parts, color, kind) {
    const merged = [];
    for (const p of parts) {
      if (p.s != null && merged.length && merged[merged.length - 1].s != null) merged[merged.length - 1] = { s: merged[merged.length - 1].s + p.s };
      else if (p.s !== '') merged.push(p);
    }
    return { parts: merged, text: merged.map(p => (p.s != null ? p.s : p.pad != null ? ' ' : '')).join(''), color: color || null, kind };
  }
  // 入りきらないときの ####
  function hashes(opts) {
    let n = 1;
    if (opts.fit) { while (n < 255 && opts.fit('#'.repeat(n + 1))) n++; }
    return { parts: [{ s: '#'.repeat(n) }], text: '#'.repeat(n), color: null, kind: 'number', hash: true };
  }
  function textSec(sec, v) {
    const parts = [];
    for (const x of sec.tokens) {
      if (x.t === '@') parts.push({ s: v });
      else if (x.t === 'lit') parts.push({ s: x.s });
      else if (x.t === 'pad') parts.push({ pad: x.ch });
      else if (x.t === 'fill') parts.push({ fill: x.ch });
      else if (x.s != null) parts.push({ s: x.s });
    }
    return parts;
  }
  function litSec(sec, minus) {
    const parts = minus ? [{ s: '-' }] : [];
    for (const x of sec.tokens) {
      if (x.t === 'lit') parts.push({ s: x.s });
      else if (x.t === 'pad') parts.push({ pad: x.ch });
      else if (x.t === 'fill') parts.push({ fill: x.ch });
    }
    return parts;
  }

  // General（G/標準）：符号のほかは 11 字まで。ok(s) → その字（符号なし）が使えるか（字数・セルの幅）
  //   入らなければ小数を減らす。1 より小さい数は、指数の形のほうが大事な数字を多く出せるならそちら（Excel の細い列で確かめた）
  function generalPick(v, ok) {
    const sign = v < 0 ? '-' : '', a = Math.abs(v);
    if (!a) return ok('0') ? '0' : null;
    const D = digits(a);
    const plain = dec => { const r = roundAt(D, dec), f = r.f.replace(/0+$/, ''); return (r.i || '0') + (f ? '.' + f : ''); };
    const sci = dec => {
      let r = roundAt({ d: D.d, e: 1 }, dec), ex = D.e - 1;
      if (r.i.length > 1) { r = { i: '1', f: '0'.repeat(dec) }; ex++; }
      const f = r.f.replace(/0+$/, '');
      return (r.i || '0') + (f ? '.' + f : '') + 'E' + (ex < 0 ? '-' : '+') + String(Math.abs(ex)).padStart(2, '0');
    };
    const okS = s => s.length <= 11 && ok(s, sign);
    let sciBest = null;
    for (let dec = 5; dec >= 0; dec--) { const s = sci(dec); if (okS(s)) { sciBest = s; break; } }
    if (D.e > 11) return sciBest && sign + sciBest;
    if (D.e >= 1) {
      for (let dec = Math.max(0, 10 - D.e); dec >= 0; dec--) { const s = plain(dec); if (okS(s)) return sign + s; }
      return sciBest && sign + sciBest;
    }
    let plainBest = null;
    for (let dec = 10; dec >= 0; dec--) { const s = plain(dec); if (okS(s)) { plainBest = s; break; } }
    const sig = s => s.split('E')[0].replace(/\./g, '').replace(/^0+/, '').length;
    const best = sciBest && (plainBest == null || sig(sciBest) > sig(plainBest)) ? sciBest : plainBest;
    return best && sign + best;
  }
  function generalStr(v, maxLen = 11) {
    return generalPick(v, s => s.length <= maxLen) || generalPick(v, () => true);
  }
  function generalSec(sec, v, minus, opts) {
    const a = minus ? v : Math.abs(v);
    const wrap = s => {
      const parts = [];
      for (const x of sec.tokens) {
        if (x.t === 'gen') parts.push({ s });
        else if (x.t === 'lit') parts.push({ s: x.s });
        else if (x.t === 'pad') parts.push({ pad: x.ch });
        else if (x.t === 'fill') parts.push({ fill: x.ch });
      }
      return parts;
    };
    if (!opts.fit) return wrap(generalStr(a));
    const s = generalPick(a, (t, sign) => opts.fit(out(wrap(sign + t), null, 'general').text));
    return s == null ? null : wrap(s);
  }

  function numberSec(sec, v, minus) {
    const tk = sec.tokens;
    let a = Math.abs(v);
    for (let k = 0; k < sec.pct; k++) a *= 100;
    for (let k = 0; k < sec.scale; k++) a /= 1000;
    if (sec.sci) return sciSec(sec, a, minus);
    if (sec.frac) return fracSec(sec, a, minus);
    const dot = tk.findIndex(x => x.t === '.');
    const fracPh = dot < 0 ? [] : tk.slice(dot + 1).filter(x => x.t === 'd');
    const r = roundAt(digits(a), fracPh.length);
    const parts = [];
    // 丸めて 0 になったときは - をつけない（Excel で確かめた）
    if (minus && (r.i || /[1-9]/.test(r.f))) parts.push({ s: '-' });
    // 整数部
    const intEnd = dot < 0 ? tk.length : dot;
    const intPh = [];
    for (let k = 0; k < intEnd; k++) if (tk[k].t === 'd') intPh.push(k);
    const fill = {};
    if (intPh.length) {
      let ds = r.i;
      const zeros = intPh.filter(k => tk[k].ch !== '#').length;
      const first0 = intPh.findIndex(k => tk[k].ch === '0');
      const minLen = first0 < 0 ? 0 : intPh.length - first0;
      if (sec.group && intPh.every((k, j) => j === 0 || intPh[j - 1] === k - 1 || tk.slice(intPh[j - 1] + 1, k).every(x => x.t === 'nil'))) {
        // ふつうの #,##0：けたをそろえてからカンマ
        const qs = intPh.map(k => tk[k].ch);
        let body = ds.padStart(minLen, '0');
        // ? の所は空白
        const padQ = Math.max(0, qs.length - body.length);
        const qpad = qs.slice(0, padQ).filter(c => c === '?').length;
        body = group3(body);
        fill[intPh[intPh.length - 1]] = [...Array(qpad).fill({ pad: '0' }), { s: body }];
        for (const k of intPh.slice(0, -1)) fill[k] = [];
      } else {
        // 000-0000 のように間に字があるとき：右から 1 字ずつ入れる
        let rest = ds;
        for (let j = intPh.length - 1; j >= 0; j--) {
          const k = intPh[j], ph = tk[k].ch;
          if (j === 0) {
            const take = rest;
            rest = '';
            if (take) fill[k] = [{ s: take }];
            else fill[k] = ph === '0' ? [{ s: '0' }] : ph === '?' ? [{ pad: '0' }] : [];
          } else if (rest) { fill[k] = [{ s: rest.slice(-1) }]; rest = rest.slice(0, -1); }
          else fill[k] = ph === '0' ? [{ s: '0' }] : ph === '?' ? [{ pad: '0' }] : [];
        }
        void zeros;
      }
    }
    // 小数部（# は末尾の 0 を消す、? は空白に）
    if (dot >= 0) {
      let f = r.f.split('');
      for (let j = fracPh.length - 1; j >= 0; j--) {
        if (f[j] !== '0') break;
        if (fracPh[j].ch === '#') f[j] = '';
        else if (fracPh[j].ch === '?') f[j] = { pad: '0' };
        else break;
      }
      let j = 0;
      for (let k = dot + 1; k < tk.length; k++) if (tk[k].t === 'd') { const c = f[j++]; fill[k] = c === '' ? [] : typeof c === 'object' ? [c] : [{ s: c }]; }
    }
    for (let k = 0; k < tk.length; k++) {
      const x = tk[k];
      if (fill[k]) parts.push(...fill[k]);
      else if (x.t === '.') parts.push({ s: '.' });
      else if (x.t === 'lit') parts.push({ s: x.s });
      else if (x.t === 'pad') parts.push({ pad: x.ch });
      else if (x.t === 'fill') parts.push({ fill: x.ch });
      else if (x.t === '%') parts.push({ s: '%' });
    }
    return parts;
  }
  function sciSec(sec, a, minus) {
    const tk = sec.tokens, e = sec.e;
    const dot = tk.findIndex((x, k) => x.t === '.' && k < e);
    const intPh = [], fracPh = [], expPh = [];
    tk.forEach((x, k) => { if (x.t !== 'd') return; if (k > e) expPh.push(k); else if (dot >= 0 && k > dot) fracPh.push(k); else intPh.push(k); });
    const nInt = Math.max(1, intPh.length), eng = intPh.some(k => tk[k].ch === '#') && intPh.length > 1;
    const D = digits(a);
    let ex = 0, r;
    for (let tries = 0; tries < 3; tries++) {
      if (D.d) {
        const top = D.e - 1 + (tries ? 1 : 0);
        ex = eng ? Math.floor(top / nInt) * nInt : top - (nInt - 1);
      }
      r = roundAt({ d: D.d, e: D.e - ex }, fracPh.length);
      const lim = eng ? nInt : nInt;
      if (r.i.length <= lim) break;
    }
    const parts = minus ? [{ s: '-' }] : [];
    const fill = {};
    // 仮数の整数部（0 のときは # も 0 になる：##0.0E+0 → 000.0E+0）
    let rest = r.i;
    for (let j = intPh.length - 1; j >= 0; j--) {
      const k = intPh[j], ph = !D.d && tk[k].ch === '#' ? '0' : tk[k].ch;
      if (j === 0) { fill[k] = rest ? [{ s: rest }] : ph === '0' ? [{ s: '0' }] : ph === '?' ? [{ pad: '0' }] : []; rest = ''; }
      else if (rest) { fill[k] = [{ s: rest.slice(-1) }]; rest = rest.slice(0, -1); }
      else fill[k] = ph === '0' ? [{ s: '0' }] : ph === '?' ? [{ pad: '0' }] : [];
    }
    const f = r.f.split('');
    for (let j = fracPh.length - 1; j >= 0; j--) { if (f[j] !== '0') break; if (tk[fracPh[j]].ch === '#') f[j] = ''; else if (tk[fracPh[j]].ch === '?') f[j] = { pad: '0' }; else break; }
    fracPh.forEach((k, j) => { const c = f[j]; fill[k] = c === '' ? [] : typeof c === 'object' ? [c] : [{ s: c }]; });
    // 指数
    const es = String(Math.abs(ex));
    const minE = expPh.filter(k => tk[k].ch === '0').length;
    const ebody = es.padStart(minE, '0');
    expPh.forEach((k, j) => { fill[k] = j === expPh.length - 1 ? [{ s: ebody }] : []; });
    for (let k = 0; k < tk.length; k++) {
      const x = tk[k];
      if (fill[k]) parts.push(...fill[k]);
      else if (x.t === 'E') parts.push({ s: 'E' + (ex < 0 ? '-' : x.sign === '+' ? '+' : '') });
      else if (x.t === '.') parts.push({ s: '.' });
      else if (x.t === 'lit') parts.push({ s: x.s });
      else if (x.t === 'pad') parts.push({ pad: x.ch });
      else if (x.t === 'fill') parts.push({ fill: x.ch });
      else if (x.t === '%') parts.push({ s: '%' });
    }
    return parts;
  }
  // 分数（# ?/? など）
  function fracSec(sec, a, minus) {
    const tk = sec.tokens;
    const numPh = [];
    for (let k = sec.num.from; k <= sec.num.to && k >= 0; k++) numPh.push(k);
    const denPh = [];
    if (sec.den.fixed == null) for (let k = sec.den.from; k <= sec.den.to; k++) denPh.push(k);
    let whole = 0, fr = a;
    if (sec.whole) { whole = Math.floor(a); fr = a - whole; }
    let n, d;
    if (sec.den.fixed != null) { d = sec.den.fixed; n = Math.round(fr * d); }
    else {
      const maxD = Math.pow(10, denPh.length) - 1;
      [n, d] = approx(fr, maxD);
    }
    if (sec.whole && n === d && d) { whole++; n = 0; }
    const parts = minus ? [{ s: '-' }] : [];
    const fill = {};
    const zeroFrac = n === 0 && sec.whole;
    if (sec.whole) {
      const wph = [];
      for (let k = sec.whole.from; k <= sec.whole.to; k++) if (tk[k].t === 'd') wph.push(k);
      const ws = whole ? String(whole) : '';
      wph.forEach((k, j) => {
        if (j < wph.length - 1) { fill[k] = []; return; }
        const has0 = wph.some(q => tk[q].ch === '0');
        fill[k] = ws ? [{ s: ws }] : (has0 || zeroFrac) ? [{ s: '0' }] : [];
      });
      if (!whole && !wph.some(q => tk[q].ch === '0') && !zeroFrac) {
        // 整数部がない（0 で # のとき）は、間の空白も詰めない（Excel と同じく空白のまま）
      }
    }
    const pad = (s, k, left) => {
      const len = k.length;
      const extra = Math.max(0, len - s.length);
      const pads = k.slice(0, extra).map(q => (tk[q].ch === '0' ? { s: '0' } : tk[q].ch === '?' ? { pad: '0' } : null)).filter(Boolean);
      return left ? [...pads, { s }] : [{ s }, ...pads];
    };
    if (zeroFrac) {
      // 分数の所は空白
      for (const k of numPh) fill[k] = [{ pad: '0' }];
      const sl = tk.findIndex(x => x.t === '/');
      fill[sl] = [{ pad: '/' }];
      if (sec.den.fixed != null) fill[sec.den.at] = String(sec.den.fixed).split('').map(() => ({ pad: '0' }));
      else for (const k of denPh) fill[k] = [{ pad: '0' }];
    } else {
      const ns = String(n), ds = String(d);
      numPh.forEach((k, j) => { fill[k] = j === numPh.length - 1 ? pad(ns, numPh, true) : []; });
      if (sec.den.fixed != null) fill[sec.den.at] = [{ s: String(sec.den.fixed) }];
      else {
        // 分母：左に寄せ、残りは右に空白
        const extra = Math.max(0, denPh.length - ds.length);
        denPh.forEach((k, j) => { fill[k] = j === 0 ? [{ s: ds }, ...denPh.slice(denPh.length - extra).map(q => (tk[q].ch === '0' ? { s: '0' } : tk[q].ch === '?' ? { pad: '0' } : null)).filter(Boolean)] : []; });
      }
    }
    for (let k = 0; k < tk.length; k++) {
      const x = tk[k];
      if (fill[k]) parts.push(...fill[k]);
      else if (x.t === '/') parts.push({ s: '/' });
      else if (x.t === 'lit') parts.push({ s: x.s });
      else if (x.t === 'pad') parts.push({ pad: x.ch });
      else if (x.t === 'fill') parts.push({ fill: x.ch });
      else if (x.t === '.') parts.push({ s: '.' });
      else if (x.t === '%') parts.push({ s: '%' });
    }
    return parts;
  }
  // 分母が maxD 以下でいちばん近い分数
  function approx(x, maxD) {
    if (x <= 0) return [0, 1];
    let best = [Math.round(x), 1], err = Math.abs(x - best[0]);
    // 連分数で近いものを探し、途中の分母も試す
    let h0 = 1, h1 = 0, k0 = 0, k1 = 1, b = x;
    for (let it = 0; it < 64; it++) {
      const ai = Math.floor(b);
      const h2 = ai * h0 + h1, k2 = ai * k0 + k1;
      if (k2 > maxD) {
        const m = Math.floor((maxD - k1) / k0);
        if (m > 0) { const hh = m * h0 + h1, kk = m * k0 + k1; const e2 = Math.abs(x - hh / kk); if (e2 < err - 1e-12) { best = [hh, kk]; err = e2; } }
        break;
      }
      const e2 = Math.abs(x - h2 / k2);
      if (e2 < err - 1e-12) { best = [h2, k2]; err = e2; }
      h1 = h0; h0 = h2; k1 = k0; k0 = k2;
      if (b - ai < 1e-12) break;
      b = 1 / (b - ai);
    }
    return best;
  }

  /* ---------- 日付・時刻 ---------- */
  const MON = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
  const DAY = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  const JDAY = ['日', '月', '火', '水', '木', '金', '土'];
  const ERAS = [[2019, 5, 1, 'R', '令', '令和'], [1989, 1, 8, 'H', '平', '平成'], [1926, 12, 25, 'S', '昭', '昭和'], [1912, 7, 30, 'T', '大', '大正'], [1868, 1, 1, 'M', '明', '明治']];
  // シリアル値 → 年月日・曜日（1900 年をうるう年とする Excel のくせも同じに）
  function ymd(days, date1904) {
    if (date1904) days += 1462;
    if (days === 0) return { y: 1900, m: 1, d: 0, w: 6 };
    if (days === 60) return { y: 1900, m: 2, d: 29, w: 3 };
    const base = days < 60 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
    const dt = new Date(base + days * 86400000);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), w: (days + 6) % 7 };
  }
  function eraOf(y, m, d) {
    const key = y * 10000 + m * 100 + d;
    for (const e of ERAS) if (key >= e[0] * 10000 + e[1] * 100 + e[2]) return { name: e, year: y - e[0] + 1 };
    return { name: null, year: y };
  }
  function dateSec(sec, v, opts) {
    // 経過時間だけの形式（[h]:mm など）は、負の数も「-」をつけて出す
    if (v < 0 && sec.elapsed && !sec.tokens.some(x => /^(y|m|dd|a|g|e|r|b)$/.test(x.t))) { const p = dateSec(sec, -v, opts); return p && [{ s: '-' }, ...p]; }
    if (v < 0 || v >= 2958466) return null;
    // 日本の書式（[$-411]）では、曜日・月の名前も日本語
    const ja = /^(411|ja)/i.test(sec.locale || '') || /gannen/.test(sec.locale || '');
    const fd = sec.frac || 0;
    const unit = Math.pow(10, fd);
    // 表示するけたで丸める（秒の小数がなければ秒で）
    let total = Math.round(v * 86400 * unit);
    let days = Math.floor(total / (86400 * unit));
    let rem = total - days * 86400 * unit;
    const D = ymd(days, opts.date1904);
    const H = Math.floor(rem / (3600 * unit)), Mi = Math.floor(rem / (60 * unit)) % 60, S = Math.floor(rem / unit) % 60, F = rem % unit;
    const pad2 = (x, n) => String(x).padStart(n, '0');
    const er = eraOf(D.y, D.m, D.d);
    const gannen = /gannen/.test(sec.locale || '');
    const parts = [];
    for (const x of sec.tokens) {
      switch (x.t) {
        case 'y': parts.push({ s: x.n <= 2 ? pad2(D.y % 100, 2) : String(D.y) }); break;
        case 'm':
          if (x.n === 1) parts.push({ s: String(D.m) });
          else if (x.n === 2) parts.push({ s: pad2(D.m, 2) });
          else if (ja) parts.push({ s: x.n === 5 ? String(D.m) : D.m + '月' });
          else if (x.n === 3) parts.push({ s: MON[D.m - 1].slice(0, 3) });
          else if (x.n === 5) parts.push({ s: MON[D.m - 1][0] });
          else parts.push({ s: MON[D.m - 1] });
          break;
        case 'min': parts.push({ s: x.n >= 2 ? pad2(Mi, 2) : String(Mi) }); break;
        case 'dd':
          if (x.n === 1) parts.push({ s: String(D.d) });
          else if (x.n === 2) parts.push({ s: pad2(D.d, 2) });
          else if (x.n === 3) parts.push({ s: ja ? JDAY[D.w] : DAY[D.w].slice(0, 3) });
          else parts.push({ s: ja ? JDAY[D.w] + '曜日' : DAY[D.w] });
          break;
        case 'a': parts.push({ s: x.n === 3 ? JDAY[D.w] : JDAY[D.w] + '曜日' }); break;
        case 'h': {
          let hh = H;
          if (sec.ampm) { hh = H % 12; if (hh === 0) hh = 12; }
          parts.push({ s: x.n >= 2 ? pad2(hh, 2) : String(hh) });
          break;
        }
        case 's': parts.push({ s: x.n >= 2 ? pad2(S, 2) : String(S) }); break;
        case 'sf': parts.push({ s: '.' + pad2(Math.floor(F / Math.pow(10, fd - x.n)), x.n) }); break;
        case 'el': {
          const tot = { h: Math.floor(total / (3600 * unit)), m: Math.floor(total / (60 * unit)), s: Math.floor(total / unit) }[x.u];
          parts.push({ s: pad2(tot, x.n) });
          break;
        }
        case 'ampm': {
          const pm = H >= 12;
          const [am, p] = ja && x.s.length > 3 ? ['午前', '午後'] : x.s.split('/');
          parts.push({ s: pm ? p : am });
          break;
        }
        case 'g': parts.push({ s: er.name ? er.name[Math.min(x.n, 3) + 2] : '' }); break;
        case 'e':
          if (gannen && er.year === 1 && er.name) parts.push({ s: '元' });
          else parts.push({ s: er.name ? (x.n >= 2 ? pad2(er.year, 2) : String(er.year)) : String(D.y) });
          break;
        case 'r':
          if (x.n >= 2) parts.push({ s: (er.name ? er.name[5] : '') + (er.name ? pad2(er.year, 2) : String(D.y)) });
          else parts.push({ s: er.name ? pad2(er.year, 2) : String(D.y) });
          break;
        case 'b': parts.push({ s: String(D.y + 543) }); break;
        case 'lit': parts.push({ s: x.s }); break;
        case 'pad': parts.push({ pad: x.ch }); break;
        case 'fill': parts.push({ fill: x.ch }); break;
        case 'gen': parts.push({ s: generalStr(v, 11) }); break;
      }
    }
    return parts;
  }

  /* ---------- 漢数字（[DBNum1]〜[DBNum3]） ----------
     General のときは 千百十・万億兆 をつける（兆より上が 5 けた以上なら 1 字ずつ）。ほかの形式は 1 字ずつ置きかえる（Excel で確かめた） */
  const KAN = { 1: '〇一二三四五六七八九', 2: '〇壱弐参四伍六七八九', 3: '０１２３４５６７８９' };
  function kanjiInt(s, k) {
    const dg = KAN[k];
    const units = k === 2 ? ['', '拾', '百', '阡'] : ['', '十', '百', '千'];
    const big = k === 2 ? ['', '萬', '億', '兆'] : ['', '万', '億', '兆'];
    s = s.replace(/^0+(?=\d)/, '');
    if (/^0*$/.test(s)) return dg[0];
    const groups = [];
    for (let e = s.length; e > 0; e -= 4) groups.unshift(s.slice(Math.max(0, e - 4), e));
    // 兆より上
    if (groups.length > 4) groups.splice(0, groups.length - 3, groups.slice(0, groups.length - 3).join(''));
    let res = '';
    groups.forEach((g, gi) => {
      let part = '';
      if (g.length > 4) part = [...g].map(c => dg[+c]).join('');
      else for (let j = 0; j < g.length; j++) {
        const d = +g[j], u = g.length - 1 - j;
        if (!d) continue;
        part += (d === 1 && u > 0 && k !== 2 ? '' : dg[d]) + units[u];
      }
      if (part) res += part + big[groups.length - 1 - gi];
    });
    return res;
  }
  // General を漢数字に（指数の形にはしない）
  function dbGeneral(v, k) {
    const a = Math.abs(v);
    let s = generalStr(a, 11);
    if (/E/.test(s)) s = plainStr(a);
    const [i, f] = s.split('.');
    return (v < 0 ? '-' : '') + kanjiInt(i, k) + (f ? '.' + [...f].map(c => KAN[k][+c]).join('') : '');
  }
  function plainStr(a) {
    const D = digits(a);
    if (!D.d) return '0';
    if (D.e >= D.d.length) return D.d + '0'.repeat(D.e - D.d.length);
    if (D.e <= 0) return '0.' + '0'.repeat(-D.e) + D.d;
    return D.d.slice(0, D.e) + '.' + D.d.slice(D.e);
  }
  function dbnum(r, k) {
    const conv = s => s.replace(/[0-9]/g, c => KAN[k][+c]);
    const parts = r.parts.map(p => (p.s != null ? { s: conv(p.s) } : p));
    return { ...r, parts, text: parts.map(p => (p.s != null ? p.s : p.pad != null ? ' ' : '')).join('') };
  }

  // 日付の形式か（セルの値を日付として見せるか）
  function isDate(code) { const secs = compile(code); return secs[0].kind === 'date'; }

  return { format, builtin, isDate, compile, generalStr, digits, roundAt, PALETTE };
})();
if (typeof module !== 'undefined') module.exports = NumFmt;
