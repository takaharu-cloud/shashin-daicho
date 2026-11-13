// 写真のファイルの先頭から、種類・撮影日時（DateTimeOriginal）・向きを読む。
// DOM に触らない。入力は途中で切れていてもよい。例外は外へ投げない。

const HEIC_BRANDS = ['heic', 'heix', 'hevc', 'hevx', 'heim', 'heis', 'mif1', 'msf1'];
const EMPTY = Object.freeze({ kind: 'other', dateTimeOriginal: null, orientation: null });

const ascii = (b, start, len) => {
  let s = '';
  for (let i = start; i < start + len && i < b.length; i++) s += String.fromCharCode(b[i]);
  return s;
};

function detectKind(b) {
  if (b.length >= 3 && b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return 'jpeg';
  if (b.length >= 8 && [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a].every((v, i) => b[i] === v)) return 'png';
  if (b.length >= 12 && ascii(b, 4, 4) === 'ftyp' && HEIC_BRANDS.includes(ascii(b, 8, 4))) return 'heic';
  return 'other';
}

/** 'YYYY:MM:DD HH:MM:SS' の形と範囲を確かめ、正しければその文字列、だめなら null。 */
function validateDateTime(s) {
  const m = /^(\d{4}):(\d{2}):(\d{2}) (\d{2}):(\d{2}):(\d{2})$/.exec(s);
  if (!m) return null;
  const [y, mo, d, h, mi, se] = m.slice(1).map(Number);
  if (y < 1900 || y > 2100 || mo < 1 || mo > 12) return null;
  const leap = (y % 4 === 0 && y % 100 !== 0) || y % 400 === 0;
  const days = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31][mo - 1];
  if (d < 1 || d > days || h > 23 || mi > 59 || se > 59) return null;
  return s;
}

/** JPEG の区切りをたどり、Exif の TIFF 部分（切れていれば有る所まで）を返す。無ければ null。 */
function findExifTiff(b) {
  let pos = 2;
  while (pos + 4 <= b.length) {
    if (b[pos] !== 0xff) return null;
    let marker = b[pos + 1];
    if (marker === 0xff) { pos++; continue; } // 詰め物
    if (marker === 0xd9 || marker === 0xda) return null; // EOI・SOS
    if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd8)) { pos += 2; continue; } // 長さの無い区切り
    const len = (b[pos + 2] << 8) | b[pos + 3];
    if (len < 2) return null;
    const start = pos + 4;
    const end = Math.min(pos + 2 + len, b.length);
    if (marker === 0xe1 && end - start >= 6 && ascii(b, start, 6) === 'Exif\0\0') {
      return b.subarray(start + 6, end);
    }
    pos += 2 + len;
  }
  return null;
}

/** TIFF の中から Orientation と DateTimeOriginal を読む。範囲の外は読まない。 */
function readTiff(t) {
  const out = { dateTimeOriginal: null, orientation: null };
  if (t.length < 8) return out;
  const head = ascii(t, 0, 2);
  const le = head === 'II';
  if (!le && head !== 'MM') return out;
  const dv = new DataView(t.buffer, t.byteOffset, t.byteLength);
  const u16 = (o) => (o >= 0 && o + 2 <= t.length ? dv.getUint16(o, le) : null);
  const u32 = (o) => (o >= 0 && o + 4 <= t.length ? dv.getUint32(o, le) : null);
  if (u16(2) !== 42) return out;

  /** IFD の項目を (tag, type, count, 項目の位置) で順に渡す。個数は配列に収まる分だけ。 */
  const eachEntry = (ifdOff, fn) => {
    const n = u16(ifdOff);
    if (n === null) return;
    const room = Math.floor((t.length - ifdOff - 2) / 12);
    for (let i = 0; i < Math.min(n, room); i++) {
      const p = ifdOff + 2 + i * 12;
      if (fn(u16(p), u16(p + 2), u32(p + 4), p + 8) === true) return;
    }
  };

  let exifOff = null;
  eachEntry(u32(4), (tag, type, count, valPos) => {
    if (tag === 0x0112 && type === 3 && count === 1) {
      const v = u16(valPos);
      if (v !== null && v >= 1 && v <= 8) out.orientation = v;
    } else if (tag === 0x8769 && type === 4) {
      exifOff = u32(valPos);
    }
  });
  if (exifOff === null) return out;

  eachEntry(exifOff, (tag, type, count, valPos) => {
    if (tag !== 0x9003 || type !== 2 || count === null || count < 19) return false;
    const at = count <= 4 ? valPos : u32(valPos);
    if (at === null || at + count > t.length) return true; // 文字データが切れている
    let s = ascii(t, at, count);
    const nul = s.indexOf('\0');
    if (nul >= 0) s = s.slice(0, nul);
    out.dateTimeOriginal = validateDateTime(s);
    return true;
  });
  return out;
}

/**
 * @param {Uint8Array} bytes ファイルの先頭（全部でなくてよい）
 * @returns {{kind:'jpeg'|'png'|'heic'|'other', dateTimeOriginal:string|null, orientation:number|null}}
 */
export function readPhotoInfo(bytes) {
  try {
    if (!(bytes instanceof Uint8Array)) return { ...EMPTY };
    const kind = detectKind(bytes);
    const info = { kind, dateTimeOriginal: null, orientation: null };
    if (kind !== 'jpeg') return info;
    const tiff = findExifTiff(bytes);
    if (tiff) Object.assign(info, readTiff(tiff));
    return info;
  } catch {
    return { ...EMPTY };
  }
}

function parts(dt) {
  if (typeof dt !== 'string' || validateDateTime(dt) === null) return null;
  return dt.split(/[: ]/).map(Number);
}

/** 'YYYY:MM:DD HH:MM:SS' → '2026年4月1日'（0 を詰めない）。null・不正は ''。 */
export function formatDateJa(dateTimeOriginal) {
  const p = parts(dateTimeOriginal);
  return p ? `${p[0]}年${p[1]}月${p[2]}日` : '';
}

/** 'YYYY:MM:DD HH:MM:SS' → '2026/04/01 09:30'。null・不正は ''。 */
export function formatDateTimeShort(dateTimeOriginal) {
  const p = parts(dateTimeOriginal);
  if (!p) return '';
  const z = (n) => String(n).padStart(2, '0');
  return `${p[0]}/${z(p[1])}/${z(p[2])} ${z(p[3])}:${z(p[4])}`;
}
