import { test } from 'node:test';
import assert from 'node:assert/strict';
import { readPhotoInfo, formatDateJa, formatDateTimeShort } from '../src/exif.js';

// ---- テスト用の JPEG を自分で組む小さな関数 ----
const enc = new TextEncoder();

function u16(v, le) { return le ? [v & 255, (v >> 8) & 255] : [(v >> 8) & 255, v & 255]; }
function u32(v, le) {
  const b = [v & 255, (v >> 8) & 255, (v >> 16) & 255, (v >>> 24) & 255];
  return le ? b : b.reverse();
}
function ascii20(s) {
  const a = [...enc.encode(s)];
  while (a.length < 19) a.push(0x20);
  a.length = 19;
  a.push(0);
  return a;
}

/** TIFF 部分を組む。entries＝{tag,type,count,value(数) | data(バイト列＝後ろに置く)} */
function buildTiff(le, ifd0, exifIfd, opts = {}) {
  const head = [...enc.encode(le ? 'II' : 'MM'), ...u16(42, le), ...u32(8, le)];
  const ifd0Size = 2 + 12 * ifd0.length + 4;
  const exifOff = opts.exifOffset ?? (8 + ifd0Size);
  const exifSize = exifIfd ? 2 + 12 * exifIfd.length + 4 : 0;
  let dataPos = 8 + ifd0Size + exifSize;
  const dataArea = [];
  const entryBytes = (e) => {
    let valueBytes;
    if (e.data) {
      valueBytes = u32(dataPos, le);
      dataPos += e.data.length;
      dataArea.push(...e.data);
    } else if (e.type === 3) {
      valueBytes = [...u16(e.value, le), 0, 0];
    } else {
      valueBytes = u32(e.value, le);
    }
    return [...u16(e.tag, le), ...u16(e.type, le), ...u32(e.count ?? 1, le), ...valueBytes];
  };
  const ifdBytes = (list, count) => {
    const out = [...u16(count ?? list.length, le)];
    for (const e of list) out.push(...entryBytes(e));
    out.push(...u32(0, le));
    return out;
  };
  const list0 = ifd0.map((e) => (e.tag === 0x8769 ? { ...e, value: exifOff } : e));
  const b0 = ifdBytes(list0, opts.ifd0Count);
  const b1 = exifIfd ? ifdBytes(exifIfd, opts.exifCount) : [];
  return [...head, ...b0, ...b1, ...dataArea];
}

function segment(marker, payload) {
  return [0xff, marker, ...u16(payload.length + 2, false), ...payload];
}

/**
 * @param {{le?:boolean, orientation?:number, dateTime?:string, dateTimeOriginal?:string,
 *   dateTimeDigitized?:string, before?:string[], noExif?:boolean, tiffOpts?:object}} o
 */
function buildJpeg(o = {}) {
  const le = o.le ?? true;
  const out = [0xff, 0xd8];
  for (const b of o.before ?? []) {
    if (b === 'jfif') out.push(...segment(0xe0, [...enc.encode('JFIF\0'), 1, 1, 0, 0, 1, 0, 1, 0, 0]));
    if (b === 'xmp') out.push(...segment(0xe1, [...enc.encode('http://ns.adobe.com/xap/1.0/\0'), ...enc.encode('<x:xmpmeta></x:xmpmeta>')]));
    if (b === 'icc') out.push(...segment(0xe2, [...enc.encode('ICC_PROFILE\0'), 1, 1, 0, 0, 0, 0]));
    if (b === 'comment') out.push(...segment(0xfe, [...enc.encode('hello')]));
  }
  if (!o.noExif) {
    const ifd0 = [];
    if (o.orientation !== undefined) ifd0.push({ tag: 0x0112, type: 3, value: o.orientation });
    if (o.dateTime) ifd0.push({ tag: 0x0132, type: 2, count: 20, data: ascii20(o.dateTime) });
    const exif = [];
    if (o.dateTimeOriginal !== undefined) exif.push({ tag: 0x9003, type: 2, count: 20, data: ascii20(o.dateTimeOriginal) });
    if (o.dateTimeDigitized) exif.push({ tag: 0x9004, type: 2, count: 20, data: ascii20(o.dateTimeDigitized) });
    if (exif.length) ifd0.push({ tag: 0x8769, type: 4, value: 0 });
    ifd0.sort((a, b) => a.tag - b.tag);
    const tiff = buildTiff(le, ifd0, exif.length ? exif : null, o.tiffOpts);
    out.push(...segment(0xe1, [...enc.encode('Exif\0\0'), ...tiff]));
  }
  out.push(...segment(0xdb, new Array(65).fill(1)));
  out.push(0xff, 0xda, 0x00, 0x02, 0x11, 0x22, 0x33, 0xff, 0xd9);
  return Uint8Array.from(out);
}

// ---- 読み取り ----
for (const le of [true, false]) {
  const name = le ? 'little-endian' : 'big-endian';
  for (const ori of [1, 3, 6, 8]) {
    test(`${name}：撮影日時と向き ${ori} が取れる`, () => {
      const r = readPhotoInfo(buildJpeg({ le, orientation: ori, dateTimeOriginal: '2026:04:01 09:30:00' }));
      assert.deepEqual(r, { kind: 'jpeg', dateTimeOriginal: '2026:04:01 09:30:00', orientation: ori });
    });
  }
  test(`${name}：向きだけ有って日時が無い`, () => {
    const r = readPhotoInfo(buildJpeg({ le, orientation: 6 }));
    assert.equal(r.orientation, 6);
    assert.equal(r.dateTimeOriginal, null);
  });
}

test('EXIF の無い JPEG は両方 null で kind は jpeg', () => {
  assert.deepEqual(readPhotoInfo(buildJpeg({ noExif: true })), { kind: 'jpeg', dateTimeOriginal: null, orientation: null });
});

test('DateTime と DateTimeDigitized だけでは日時を返さない', () => {
  const r = readPhotoInfo(buildJpeg({ dateTime: '2026:04:01 09:30:00', dateTimeDigitized: '2026:04:02 10:00:00', orientation: 1 }));
  assert.equal(r.dateTimeOriginal, null);
  assert.equal(r.orientation, 1);
});

test('DateTimeDigitized が有っても DateTimeOriginal の方を返す', () => {
  const r = readPhotoInfo(buildJpeg({ dateTime: '2026:05:05 05:05:05', dateTimeOriginal: '2026:04:01 09:30:00', dateTimeDigitized: '2026:04:02 10:00:00' }));
  assert.equal(r.dateTimeOriginal, '2026:04:01 09:30:00');
});

test('Exif の前に JFIF と XMP の APP1 と ICC とコメントが有っても取れる', () => {
  for (const le of [true, false]) {
    const r = readPhotoInfo(buildJpeg({ le, before: ['jfif', 'xmp', 'icc', 'comment'], orientation: 8, dateTimeOriginal: '2025:12:31 23:59:59' }));
    assert.equal(r.dateTimeOriginal, '2025:12:31 23:59:59');
    assert.equal(r.orientation, 8);
  }
});

test('XMP の APP1 しか無ければ null', () => {
  const r = readPhotoInfo(buildJpeg({ before: ['jfif', 'xmp'], noExif: true }));
  assert.equal(r.dateTimeOriginal, null);
});

test('うるう日は通り、うるう年でない 2 月 29 日は null', () => {
  assert.equal(readPhotoInfo(buildJpeg({ dateTimeOriginal: '2024:02:29 10:00:00' })).dateTimeOriginal, '2024:02:29 10:00:00');
  assert.equal(readPhotoInfo(buildJpeg({ dateTimeOriginal: '2026:02:29 10:00:00' })).dateTimeOriginal, null);
});

test('不正な日時は null', () => {
  for (const bad of ['0000:00:00 00:00:00', '2026:02:30 10:00:00', '2026:13:01 00:00:00', ' '.repeat(19),
    '2026:04:01 24:00:00', '2026:04:01 09:60:00', '2026:04:01 09:30:60', '1899:01:01 00:00:00', '2101:01:01 00:00:00', '2026:04:01 09:3']) {
    assert.equal(readPhotoInfo(buildJpeg({ dateTimeOriginal: bad })).dateTimeOriginal, null, bad);
  }
});

test('向きが 0 や 9 は null', () => {
  for (const o of [0, 9, 65535]) {
    assert.equal(readPhotoInfo(buildJpeg({ orientation: o })).orientation, null, String(o));
  }
});

// ---- 壊れた入力 ----
test('正しい JPEG を 1 バイトずつ切ったすべての長さで例外を投げない', () => {
  for (const le of [true, false]) {
    const full = buildJpeg({ le, before: ['jfif', 'xmp'], orientation: 6, dateTime: '2026:01:01 00:00:00', dateTimeOriginal: '2026:04:01 09:30:00', dateTimeDigitized: '2026:04:01 09:30:01' });
    for (let n = 0; n <= full.length; n++) {
      const r = readPhotoInfo(full.subarray(0, n));
      assert.ok(['jpeg', 'other'].includes(r.kind));
      if (r.dateTimeOriginal !== null) assert.equal(r.dateTimeOriginal, '2026:04:01 09:30:00');
    }
  }
});

test('でたらめなバイト列で例外を投げない', () => {
  assert.deepEqual(readPhotoInfo(new Uint8Array(0)), { kind: 'other', dateTimeOriginal: null, orientation: null });
  assert.equal(readPhotoInfo(new Uint8Array(1)).kind, 'other');
  assert.equal(readPhotoInfo(new Uint8Array(2)).kind, 'other');
  let seed = 12345;
  const rnd = () => ((seed = (seed * 1103515245 + 12345) & 0x7fffffff) >> 8) & 255;
  for (let t = 0; t < 50; t++) {
    const b = Uint8Array.from({ length: 1000 }, rnd);
    readPhotoInfo(b);
    b[0] = 0xff; b[1] = 0xd8; b[2] = 0xff;
    const r = readPhotoInfo(b);
    assert.equal(r.kind, 'jpeg');
  }
  assert.equal(readPhotoInfo(null).kind, 'other');
});

test('Exif IFD の指し先が配列の外なら null', () => {
  const r = readPhotoInfo(buildJpeg({ orientation: 6, dateTimeOriginal: '2026:04:01 09:30:00', tiffOpts: { exifOffset: 100000 } }));
  assert.equal(r.dateTimeOriginal, null);
  assert.equal(r.orientation, 6);
});

test('IFD の個数が巨大でも止まって例外を投げない', () => {
  for (const le of [true, false]) {
    const r = readPhotoInfo(buildJpeg({ le, orientation: 6, dateTimeOriginal: '2026:04:01 09:30:00', tiffOpts: { ifd0Count: 65535, exifCount: 65535 } }));
    assert.equal(r.kind, 'jpeg');
    assert.equal(r.orientation, 6);
  }
});

test('Exif IFD の指し先が自分（IFD0）を指しても止まる', () => {
  const r = readPhotoInfo(buildJpeg({ orientation: 6, dateTimeOriginal: '2026:04:01 09:30:00', tiffOpts: { exifOffset: 8 } }));
  assert.equal(r.dateTimeOriginal, null);
  assert.equal(r.orientation, 6);
});

test('日時の文字列の指し先が配列の外なら null', () => {
  const j = buildJpeg({ dateTimeOriginal: '2026:04:01 09:30:00' });
  // 末尾の文字データ（20 バイト）を含む所まで切る
  const cut = j.subarray(0, j.length - 100);
  assert.equal(readPhotoInfo(cut).dateTimeOriginal, null);
});

// ---- 種類の見分け ----
test('kind：PNG・HEIC・HEIF・ほかの物', () => {
  const png = Uint8Array.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0]);
  assert.deepEqual(readPhotoInfo(png), { kind: 'png', dateTimeOriginal: null, orientation: null });
  const ftyp = (brand) => Uint8Array.from([0, 0, 0, 0x18, ...enc.encode('ftyp'), ...enc.encode(brand), 0, 0, 0, 0]);
  for (const b of ['heic', 'heix', 'hevc', 'hevx', 'heim', 'heis', 'mif1', 'msf1']) {
    assert.equal(readPhotoInfo(ftyp(b)).kind, 'heic', b);
  }
  assert.equal(readPhotoInfo(ftyp('mp42')).kind, 'other');
  assert.equal(readPhotoInfo(enc.encode('GIF89a......')).kind, 'other');
  assert.equal(readPhotoInfo(Uint8Array.from([0xff, 0xd8, 0xff, 0xe0])).kind, 'jpeg');
  assert.equal(readPhotoInfo(Uint8Array.from([0xff, 0xd8, 0x00])).kind, 'other');
});

// ---- 書式 ----
test('formatDateJa と formatDateTimeShort', () => {
  assert.equal(formatDateJa('2026:04:01 09:30:00'), '2026年4月1日');
  assert.equal(formatDateJa('2026:12:31 00:00:00'), '2026年12月31日');
  assert.equal(formatDateTimeShort('2026:04:01 09:30:00'), '2026/04/01 09:30');
  for (const bad of [null, undefined, '', 'abc', '2026:13:01 00:00:00', '0000:00:00 00:00:00']) {
    assert.equal(formatDateJa(bad), '');
    assert.equal(formatDateTimeShort(bad), '');
  }
});
