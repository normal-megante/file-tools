'use strict';
// =========================================================
// QRコードを作るしくみ（規格 JIS X 0510 / ISO 18004 に沿って自前で計算）
//
//   const qr = QR.encode('こんにちは', 'M');
//   qr.size        … 1辺のマスの数（21〜177）
//   qr.get(x, y)   … そのマスが黒なら true
//   qr.version     … 大きさの段階（1〜40）
//
// 文字は UTF-8 のバイト列として入れます（日本語もOK）。
// 誤り訂正（汚れ・欠けへの強さ）は L(7%) / M(15%) / Q(25%) / H(30%)。
// 同じ大きさのまま強くできるときは、自動で一段強くします。
// =========================================================
const QR = (() => {
  const LEVELS = ['L', 'M', 'Q', 'H'];
  const FORMAT_BITS = { L: 1, M: 0, Q: 3, H: 2 };
  // 規格の表：[誤り訂正][バージョン] → 1ブロックあたりの誤り訂正コード数 / ブロック数
  const ECC_PER_BLOCK = {
    L: [-1, 7, 10, 15, 20, 26, 18, 20, 24, 30, 18, 20, 24, 26, 30, 22, 24, 28, 30, 28, 28, 28, 28, 30, 30, 26, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
    M: [-1, 10, 16, 26, 18, 24, 16, 18, 22, 22, 26, 30, 22, 22, 24, 24, 28, 28, 26, 26, 26, 26, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28, 28],
    Q: [-1, 13, 22, 18, 26, 18, 24, 18, 22, 20, 24, 28, 26, 24, 20, 30, 24, 28, 28, 26, 30, 28, 30, 30, 30, 30, 28, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
    H: [-1, 17, 28, 22, 16, 22, 28, 26, 26, 24, 28, 24, 28, 22, 24, 24, 30, 28, 28, 26, 28, 30, 24, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30, 30],
  };
  const NUM_BLOCKS = {
    L: [-1, 1, 1, 1, 1, 1, 2, 2, 2, 2, 4, 4, 4, 4, 4, 6, 6, 6, 6, 7, 8, 8, 9, 9, 10, 12, 12, 12, 13, 14, 15, 16, 17, 18, 19, 19, 20, 21, 22, 24, 25],
    M: [-1, 1, 1, 1, 2, 2, 4, 4, 4, 5, 5, 5, 8, 9, 9, 10, 10, 11, 13, 14, 16, 17, 17, 18, 20, 21, 23, 25, 26, 28, 29, 31, 33, 35, 37, 38, 40, 43, 45, 47, 49],
    Q: [-1, 1, 1, 2, 2, 4, 4, 6, 6, 8, 8, 8, 10, 12, 16, 12, 17, 16, 18, 21, 20, 23, 23, 25, 27, 29, 34, 34, 35, 38, 40, 43, 45, 48, 51, 53, 56, 59, 62, 65, 68],
    H: [-1, 1, 1, 2, 4, 4, 4, 5, 6, 8, 8, 11, 11, 16, 16, 18, 16, 19, 21, 25, 25, 25, 34, 30, 32, 35, 37, 40, 42, 45, 48, 51, 54, 57, 60, 63, 66, 70, 74, 77, 81],
  };

  // データを置けるマスの数（位置合わせの模様などを除いた数）
  function rawDataModules(ver) {
    let n = (16 * ver + 128) * ver + 64;
    if (ver >= 2) {
      const align = Math.floor(ver / 7) + 2;
      n -= (25 * align - 10) * align - 55;
      if (ver >= 7) n -= 36;
    }
    return n;
  }
  const dataCodewords = (ver, ecl) => Math.floor(rawDataModules(ver) / 8) - ECC_PER_BLOCK[ecl][ver] * NUM_BLOCKS[ecl][ver];
  const countBits = ver => ver <= 9 ? 8 : 16; // バイトモードの文字数欄の長さ
  const usedBits = (ver, len) => 4 + countBits(ver) + len * 8;

  /* ---- リード・ソロモン符号（誤り訂正） ---- */
  function gfMul(x, y) {
    let z = 0;
    for (let i = 7; i >= 0; i--) {
      z = (z << 1) ^ ((z >>> 7) * 0x11D);
      z ^= ((y >>> i) & 1) * x;
    }
    return z & 0xFF;
  }
  function rsDivisor(degree) {
    const r = new Array(degree).fill(0);
    r[degree - 1] = 1;
    let root = 1;
    for (let i = 0; i < degree; i++) {
      for (let j = 0; j < degree; j++) {
        r[j] = gfMul(r[j], root);
        if (j + 1 < degree) r[j] ^= r[j + 1];
      }
      root = gfMul(root, 0x02);
    }
    return r;
  }
  function rsRemainder(data, divisor) {
    const r = divisor.map(() => 0);
    for (const b of data) {
      const factor = b ^ r.shift();
      r.push(0);
      for (let i = 0; i < divisor.length; i++) r[i] ^= gfMul(divisor[i], factor);
    }
    return r;
  }

  function encode(text, ecl = 'M', { boost = true } = {}) {
    const bytes = new TextEncoder().encode(text);
    // 入る大きさ（バージョン）のうち、いちばん小さいものを選ぶ
    let ver = 1;
    for (; ver <= 40; ver++) {
      if (bytes.length < 2 ** countBits(ver) && usedBits(ver, bytes.length) <= dataCodewords(ver, ecl) * 8) break;
    }
    if (ver > 40) return null; // 長すぎる
    if (boost) for (const e of LEVELS.slice(LEVELS.indexOf(ecl) + 1)) {
      if (usedBits(ver, bytes.length) <= dataCodewords(ver, e) * 8) ecl = e;
    }

    // ---- ビット列を作る ----
    const bits = [];
    const put = (val, len) => { for (let i = len - 1; i >= 0; i--) bits.push((val >>> i) & 1); };
    put(0b0100, 4);                          // バイトモード
    put(bytes.length, countBits(ver));        // 文字（バイト）数
    for (const b of bytes) put(b, 8);
    const capacity = dataCodewords(ver, ecl) * 8;
    put(0, Math.min(4, capacity - bits.length)); // 終わりの印
    put(0, (8 - bits.length % 8) % 8);
    for (let pad = 0xEC; bits.length < capacity; pad ^= 0xEC ^ 0x11) put(pad, 8); // 埋め草
    const data = [];
    for (let i = 0; i < bits.length; i += 8) data.push(parseInt(bits.slice(i, i + 8).join(''), 2));

    // ---- ブロックに分けて誤り訂正コードを付け、交互に並べる ----
    const numBlocks = NUM_BLOCKS[ecl][ver], eccLen = ECC_PER_BLOCK[ecl][ver];
    const raw = Math.floor(rawDataModules(ver) / 8);
    const numShort = numBlocks - raw % numBlocks, shortLen = Math.floor(raw / numBlocks);
    const divisor = rsDivisor(eccLen), blocks = [];
    for (let i = 0, k = 0; i < numBlocks; i++) {
      const dat = data.slice(k, k + shortLen - eccLen + (i < numShort ? 0 : 1));
      k += dat.length;
      const ecc = rsRemainder(dat, divisor);
      if (i < numShort) dat.push(0);
      blocks.push(dat.concat(ecc));
    }
    const codewords = [];
    for (let i = 0; i < blocks[0].length; i++) {
      blocks.forEach((blk, j) => { if (i !== shortLen - eccLen || j >= numShort) codewords.push(blk[i]); });
    }

    // ---- マス目に描く ----
    const size = ver * 4 + 17;
    const mod = Array.from({ length: size }, () => new Array(size).fill(false));
    const fixed = Array.from({ length: size }, () => new Array(size).fill(false));
    const setF = (x, y, dark) => { mod[y][x] = dark; fixed[y][x] = true; };

    for (let i = 0; i < size; i++) { setF(6, i, i % 2 === 0); setF(i, 6, i % 2 === 0); } // タイミングパターン
    const finder = (cx, cy) => {
      for (let dy = -4; dy <= 4; dy++) for (let dx = -4; dx <= 4; dx++) {
        const x = cx + dx, y = cy + dy, d = Math.max(Math.abs(dx), Math.abs(dy));
        if (x >= 0 && x < size && y >= 0 && y < size) setF(x, y, d !== 2 && d !== 4);
      }
    };
    finder(3, 3); finder(size - 4, 3); finder(3, size - 4);
    const pos = alignPositions(ver, size), n = pos.length;
    for (let i = 0; i < n; i++) for (let j = 0; j < n; j++) {
      if ((i === 0 && j === 0) || (i === 0 && j === n - 1) || (i === n - 1 && j === 0)) continue;
      for (let dy = -2; dy <= 2; dy++) for (let dx = -2; dx <= 2; dx++) setF(pos[i] + dx, pos[j] + dy, Math.max(Math.abs(dx), Math.abs(dy)) !== 1);
    }
    const drawFormat = mask => {
      const d = FORMAT_BITS[ecl] << 3 | mask;
      let rem = d;
      for (let i = 0; i < 10; i++) rem = (rem << 1) ^ ((rem >>> 9) * 0x537);
      const b = (d << 10 | rem) ^ 0x5412, bit = i => ((b >>> i) & 1) === 1;
      for (let i = 0; i <= 5; i++) setF(8, i, bit(i));
      setF(8, 7, bit(6)); setF(8, 8, bit(7)); setF(7, 8, bit(8));
      for (let i = 9; i < 15; i++) setF(14 - i, 8, bit(i));
      for (let i = 0; i < 8; i++) setF(size - 1 - i, 8, bit(i));
      for (let i = 8; i < 15; i++) setF(8, size - 15 + i, bit(i));
      setF(8, size - 8, true);
    };
    drawFormat(0); // 場所の予約（あとで正しい値で描き直す）
    if (ver >= 7) {
      let rem = ver;
      for (let i = 0; i < 12; i++) rem = (rem << 1) ^ ((rem >>> 11) * 0x1F25);
      const b = ver << 12 | rem;
      for (let i = 0; i < 18; i++) {
        const dark = ((b >>> i) & 1) === 1, a = size - 11 + i % 3, c = Math.floor(i / 3);
        setF(a, c, dark); setF(c, a, dark);
      }
    }
    // データを右下から2列ずつジグザグに置く
    let k = 0;
    for (let right = size - 1; right >= 1; right -= 2) {
      if (right === 6) right = 5;
      for (let v = 0; v < size; v++) for (let j = 0; j < 2; j++) {
        const x = right - j, up = ((right + 1) & 2) === 0, y = up ? size - 1 - v : v;
        if (!fixed[y][x] && k < codewords.length * 8) { mod[y][x] = ((codewords[k >>> 3] >>> (7 - (k & 7))) & 1) === 1; k++; }
      }
    }

    // ---- 8種類の模様（マスク）から、読み取りやすいものを選ぶ ----
    const applyMask = m => {
      for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) {
        if (fixed[y][x]) continue;
        let inv;
        switch (m) {
          case 0: inv = (x + y) % 2 === 0; break;
          case 1: inv = y % 2 === 0; break;
          case 2: inv = x % 3 === 0; break;
          case 3: inv = (x + y) % 3 === 0; break;
          case 4: inv = (Math.floor(x / 3) + Math.floor(y / 2)) % 2 === 0; break;
          case 5: inv = x * y % 2 + x * y % 3 === 0; break;
          case 6: inv = (x * y % 2 + x * y % 3) % 2 === 0; break;
          default: inv = ((x + y) % 2 + x * y % 3) % 2 === 0;
        }
        if (inv) mod[y][x] = !mod[y][x];
      }
    };
    let best = 0, bestScore = Infinity;
    for (let m = 0; m < 8; m++) {
      applyMask(m); drawFormat(m);
      const s = penalty(mod, size);
      if (s < bestScore) { bestScore = s; best = m; }
      applyMask(m); // 元に戻す
    }
    applyMask(best); drawFormat(best);

    return { version: ver, ecl, size, mask: best, get: (x, y) => x >= 0 && y >= 0 && x < size && y < size && mod[y][x] };
  }

  function alignPositions(ver, size) {
    if (ver === 1) return [];
    const n = Math.floor(ver / 7) + 2;
    const step = Math.floor((ver * 8 + n * 3 + 5) / (n * 4 - 4)) * 2;
    const r = [6];
    for (let p = size - 7; r.length < n; p -= step) r.splice(1, 0, p);
    return r;
  }

  // 読み取りにくい模様ほど点数が高い（規格の採点方法）
  function penalty(mod, size) {
    let score = 0;
    const lines = (get) => {
      for (let a = 0; a < size; a++) {
        let color = false, run = 0;
        const hist = [0, 0, 0, 0, 0, 0, 0];
        const add = len => { if (hist[0] === 0) len += size; hist.pop(); hist.unshift(len); };
        const count = () => {
          const n = hist[1], core = n > 0 && hist[2] === n && hist[3] === n * 3 && hist[4] === n && hist[5] === n;
          return (core && hist[0] >= n * 4 && hist[6] >= n ? 1 : 0) + (core && hist[6] >= n * 4 && hist[0] >= n ? 1 : 0);
        };
        for (let b = 0; b < size; b++) {
          if (get(a, b) === color) {
            run++;
            if (run === 5) score += 3; else if (run > 5) score++;
          } else {
            add(run);
            if (!color) score += count() * 40;
            color = get(a, b); run = 1;
          }
        }
        if (color) { add(run); run = 0; }
        run += size; add(run);
        score += count() * 40;
      }
    };
    lines((a, b) => mod[a][b]); // 横
    lines((a, b) => mod[b][a]); // 縦
    let dark = 0;
    for (let y = 0; y < size; y++) for (let x = 0; x < size; x++) {
      if (mod[y][x]) dark++;
      if (x < size - 1 && y < size - 1) {
        const c = mod[y][x];
        if (c === mod[y][x + 1] && c === mod[y + 1][x] && c === mod[y + 1][x + 1]) score += 3;
      }
    }
    const total = size * size;
    score += (Math.ceil(Math.abs(dark * 20 - total * 10) / total) - 1) * 10;
    return score;
  }

  // そのバージョン・誤り訂正で入る最大のバイト数
  const capacity = (ver, ecl) => Math.floor((dataCodewords(ver, ecl) * 8 - 4 - countBits(ver)) / 8);
  const maxBytes = ecl => capacity(40, ecl);

  // 読み取り（js/qrread.js）でも同じ規格の表を使う
  const spec = { ECC_PER_BLOCK, NUM_BLOCKS, rawDataModules, alignPositions };

  return { encode, capacity, maxBytes, spec };
})();

if (typeof module === 'object' && module.exports) module.exports = QR; // テスト用
