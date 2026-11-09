'use strict';
// =========================================================
// ZIP を読む部品（自作。画面に依存しないので Node でもテストできる）
//   const zip = await Zip.open(blob)   … 中身の一覧（entries）だけを先に読む（大きな ZIP でも速い）
//   zip.get('ppt/slides/slide1.xml')   … 名前で探す（大文字・小文字は区別しない）
//   await zip.bytes(entry) / zip.text(entry) / zip.blob(entry, type)  … 1つ取り出す
// Word・Excel・PowerPoint のファイルも、中身は ZIP（XML と画像が入っている）。
// 圧縮は「そのまま」と「deflate」だけ（ほとんどの ZIP はこれ）。伸ばすのはブラウザの DecompressionStream。
// ファイル名は、UTF-8 の印があれば UTF-8、なければ UTF-8 として読めるか試して、だめならシフト JIS（日本の Windows で作った ZIP）。
// =========================================================
const Zip = (() => {
  const u16 = (d, o) => d.getUint16(o, true), u32 = (d, o) => d.getUint32(o, true);
  const u64 = (d, o) => d.getUint32(o, true) + d.getUint32(o + 4, true) * 4294967296;
  const view = async blob => new DataView(await blob.arrayBuffer());

  function decodeName(bytes, utf8) {
    if (utf8) return new TextDecoder('utf-8').decode(bytes);
    try { return new TextDecoder('utf-8', { fatal: true }).decode(bytes); } catch (e) { /* UTF-8 ではない */ }
    try { return new TextDecoder('shift_jis').decode(bytes); } catch (e) { return new TextDecoder('latin1').decode(bytes); }
  }

  class ZipError extends Error { constructor(msg, code) { super(msg); this.code = code; } }

  async function open(blob) {
    const size = blob.size;
    if (size < 22) throw new ZipError('ZIP ではありません', 'notzip');
    // 終わりの印（End of central directory）を、うしろから探す（コメントがあると少し前にある）
    const tailLen = Math.min(size, 65557), tail = await view(blob.slice(size - tailLen));
    let eocd = -1;
    for (let i = tailLen - 22; i >= 0; i--) if (u32(tail, i) === 0x06054b50) { eocd = i; break; }
    if (eocd < 0) throw new ZipError('ZIP ではないか、こわれています', 'notzip');
    let count = u16(tail, eocd + 10), cdSize = u32(tail, eocd + 12), cdOff = u32(tail, eocd + 16);
    // 大きな ZIP（ZIP64）
    if ((cdOff === 0xFFFFFFFF || count === 0xFFFF || cdSize === 0xFFFFFFFF) && eocd >= 20 && u32(tail, eocd - 20) === 0x07064b50) {
      const recOff = u64(tail, eocd - 12), rec = await view(blob.slice(recOff, recOff + 56));
      if (u32(rec, 0) === 0x06064b50) { count = u64(rec, 32); cdSize = u64(rec, 40); cdOff = u64(rec, 48); }
    }
    if (cdOff + cdSize > size) throw new ZipError('ZIP がこわれています（途中で切れているかもしれません）', 'broken');
    const cd = await view(blob.slice(cdOff, cdOff + cdSize)), entries = [];
    for (let p = 0, k = 0; k < count && p + 46 <= cd.byteLength; k++) {
      if (u32(cd, p) !== 0x02014b50) break;
      const flags = u16(cd, p + 8), method = u16(cd, p + 10);
      let csize = u32(cd, p + 20), usize = u32(cd, p + 24), off = u32(cd, p + 42);
      const nlen = u16(cd, p + 28), xlen = u16(cd, p + 30), clen = u16(cd, p + 32);
      const nameBytes = new Uint8Array(cd.buffer, cd.byteOffset + p + 46, nlen);
      let name = decodeName(nameBytes, flags & 0x800);
      // おまけの情報：ZIP64 の大きさ・場所、Info-ZIP の UTF-8 の名前
      for (let x = p + 46 + nlen, xe = x + xlen; x + 4 <= xe;) {
        const id = u16(cd, x), len = u16(cd, x + 2), d = x + 4;
        if (id === 0x0001) {
          let q = d;
          if (usize === 0xFFFFFFFF) { usize = u64(cd, q); q += 8; }
          if (csize === 0xFFFFFFFF) { csize = u64(cd, q); q += 8; }
          if (off === 0xFFFFFFFF) { off = u64(cd, q); q += 8; }
        } else if (id === 0x7075 && len > 5) {
          name = new TextDecoder('utf-8').decode(new Uint8Array(cd.buffer, cd.byteOffset + d + 5, len - 5));
        }
        x = d + len;
      }
      entries.push({ name, size: usize, csize, method, offset: off, encrypted: !!(flags & 1), dir: name.endsWith('/') });
      p += 46 + nlen + xlen + clen;
    }
    const byName = new Map(entries.map(e => [e.name.toLowerCase(), e]));

    // 1つ取り出す
    async function bytes(e) {
      if (typeof e === 'string') e = byName.get(e.toLowerCase());
      if (!e) throw new ZipError('中にそのファイルがありません', 'missing');
      if (e.encrypted) throw new ZipError('パスワードつきの ZIP は開けません', 'encrypted');
      const lh = await view(blob.slice(e.offset, e.offset + 30));
      if (u32(lh, 0) !== 0x04034b50) throw new ZipError('ZIP がこわれています', 'broken');
      const start = e.offset + 30 + u16(lh, 26) + u16(lh, 28), raw = blob.slice(start, start + e.csize);
      if (e.method === 0) return new Uint8Array(await raw.arrayBuffer());
      if (e.method !== 8) throw new ZipError('この圧縮のしかたには対応していません', 'method');
      const out = await new Response(raw.stream().pipeThrough(new DecompressionStream('deflate-raw'))).arrayBuffer();
      return new Uint8Array(out);
    }
    const text = async e => new TextDecoder('utf-8').decode(await bytes(e));
    const blobOf = async (e, type = '') => new Blob([await bytes(e)], { type });
    return { entries, get: name => byName.get(String(name).toLowerCase()) || null, bytes, text, blob: blobOf };
  }

  // ZIP の頭（PK\3\4）か、空の ZIP（PK\5\6）か
  async function isZip(blob) {
    const h = new Uint8Array(await blob.slice(0, 4).arrayBuffer());
    return h[0] === 0x50 && h[1] === 0x4B && ((h[2] === 3 && h[3] === 4) || (h[2] === 5 && h[3] === 6));
  }

  return { open, isZip, ZipError };
})();
if (typeof module === 'object' && module.exports) module.exports = Zip; // テスト用
