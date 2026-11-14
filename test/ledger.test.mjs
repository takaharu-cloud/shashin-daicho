import { test } from 'node:test';
import assert from 'node:assert/strict';
import { inflateRawSync } from 'node:zlib';
import { loadExcelJS } from '../tools/load_exceljs.mjs';
import { buildLedger } from '../src/ledger.js';
import { SHEET, fitImage, imageAnchor, blockRows, hasPageBreakAfter, fitFileName, blockLayout } from '../src/layout.js';

const ExcelJS = await loadExcelJS('bare');

// 架空の小さな絵（Pillow で作った 8×6 の単色 JPEG と 4×3 の単色 PNG）
const JPEG = Uint8Array.from(Buffer.from('/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwgIyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCj/wAARCAAGAAgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDAooor70+JP//Z', 'base64'));
// 段と絵の対応を見るため、色の違う 8×6 の JPEG をあと 2 つ（赤・緑）
const JPEG_R = Uint8Array.from(Buffer.from('/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwgIyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCj/wAARCAAGAAgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDnKKKK8Q/UD//Z', 'base64'));
const JPEG_G = Uint8Array.from(Buffer.from('/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwgIyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCj/wAARCAAGAAgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwCnRRRXzJ8ef//Z', 'base64'));
const JPEGS = [JPEG, JPEG_R, JPEG_G];
const asBuf = (b) => (b instanceof ArrayBuffer ? Buffer.from(new Uint8Array(b)) : Buffer.from(b.buffer, b.byteOffset, b.byteLength));
const PNG = Uint8Array.from(Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAQAAAADCAIAAAA7ljmRAAAAE0lEQVR4nGM8EaDBAANMcBY6BwAvfAFGB1gaKwAAAABJRU5ErkJggg==', 'base64'));

const NOW = new Date('2026-04-01T00:00:00Z');
const SIZES = [[1200, 900], [900, 1200], [1200, 675], [1200, 1200], [1200, 400], [8, 6], [320, 240]];
const DATES = ['2026:04:01 09:30:00', '2026:12:31 23:59:59', null];

function makeEntries(n) {
  return Array.from({ length: n }, (_, i) => {
    const [width, height] = SIZES[i % SIZES.length];
    return {
      name: `IMG_${String(i + 1).padStart(4, '0')}.jpg`,
      dateTimeOriginal: DATES[i % DATES.length],
      image: { bytes: JPEGS[i % JPEGS.length], width, height, type: 'jpeg' },
    };
  });
}

async function reload(bytes) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  return wb;
}

// xlsx（zip）の中の 1 つのファイルを文字列で読む。ExcelJS が読み直さない物（改ページ・テーマ）を見るため。
function zipText(bytes, name) {
  const b = Buffer.from(bytes);
  let eocd = b.length - 22;
  while (eocd >= 0 && b.readUInt32LE(eocd) !== 0x06054b50) eocd--;
  const count = b.readUInt16LE(eocd + 10);
  let p = b.readUInt32LE(eocd + 16);
  for (let k = 0; k < count; k++) {
    const method = b.readUInt16LE(p + 10);
    const csize = b.readUInt32LE(p + 20);
    const nlen = b.readUInt16LE(p + 28);
    const xlen = b.readUInt16LE(p + 30);
    const clen = b.readUInt16LE(p + 32);
    const local = b.readUInt32LE(p + 42);
    const fname = b.toString('utf8', p + 46, p + 46 + nlen);
    if (fname === name) {
      const start = local + 30 + b.readUInt16LE(local + 26) + b.readUInt16LE(local + 28);
      const data = b.subarray(start, start + csize);
      return (method === 8 ? inflateRawSync(data) : data).toString('utf8');
    }
    p += 46 + nlen + xlen + clen;
  }
  throw new Error(`zip に ${name} が無い`);
}

const imagesByRow = (ws) => ws.getImages().slice().sort((a, b) => a.range.tl.nativeRow - b.range.tl.nativeRow);

/** 台帳の形を全部確かめる */
async function checkLedger(res, entries, { startNumber = 1, title = '工事写真台帳', header } = {}) {
  const n = entries.length;
  assert.equal(res.ok, true, JSON.stringify(res));
  assert.ok(res.bytes instanceof Uint8Array);
  assert.equal(res.pages, Math.ceil(n / 3));
  const wb = await reload(res.bytes);
  assert.equal(wb.worksheets.length, 1);
  const ws = wb.worksheets[0];
  assert.equal(ws.name, '写真台帳');

  // 列の幅＝px÷8
  assert.deepEqual([1, 2, 3, 4].map((c) => ws.getColumn(c).width), [52.5, 1, 8, 20.75]);

  // 結合
  const merges = [...ws.model.merges].sort();
  const wantMerges = entries.map((_, i) => `A${blockRows(i).first}:A${blockRows(i).last}`).sort();
  assert.deepEqual(merges, wantMerges);

  // 絵
  const imgs = imagesByRow(ws);
  assert.equal(imgs.length, n);

  for (let i = 0; i < n; i++) {
    const { first, gap } = blockRows(i);
    // 行の高さ
    SHEET.rowPt.forEach((pt, k) => assert.equal(ws.getRow(first + k).height, pt, `行 ${first + k}`));
    assert.equal(ws.getRow(gap).height, 7.5);
    // 見出しと値
    SHEET.labels.forEach((label, k) => {
      const c = ws.getCell(first + k, 3);
      assert.equal(c.value, label);
      assert.equal(c.alignment.horizontal, 'distributed');
      assert.equal(c.font.name, '游ゴシック');
      assert.equal(c.font.size, 9);
      assert.equal(c.border.top.style, 'thin');
      assert.equal(c.border.left.style, 'thin');
      const d = ws.getCell(first + k, 4);
      assert.equal(d.font.name, '游ゴシック');
      assert.equal(d.font.size, 9);
      assert.equal(d.border.right.style, 'thin');
      assert.equal(d.border.bottom.style, 'thin');
      // B 列には罫線を引かない
      assert.ok(!ws.getCell(first + k, 2).border || !ws.getCell(first + k, 2).border.top);
    });
    assert.equal(ws.getCell(first, 4).value, startNumber + i);
    const e = entries[i];
    const dateText = ws.getCell(first + 1, 4).value;
    if (e.dateTimeOriginal === null) assert.equal(dateText ?? '', '');
    assert.equal(ws.getCell(first + 2, 4).alignment.wrapText, true);
    assert.equal(ws.getCell(first + 2, 4).alignment.vertical, 'middle');
    assert.equal(ws.getCell(first + 7, 4).alignment.wrapText, true);
    assert.equal(ws.getCell(first + 7, 4).alignment.vertical, 'top');
    for (let k = 3; k < 8; k++) assert.equal(ws.getCell(first + k, 4).value ?? '', '');
    // 枠の罫線（結合した A の外周）
    const a = ws.getCell(first, 1);
    for (const side of ['top', 'left', 'bottom', 'right']) assert.equal(a.border[side].style, 'thin', side);

    // 絵の置き方＝layout の計算と同じ
    const fit = fitImage(e.image.width, e.image.height);
    const want = imageAnchor(i, fit);
    const got = imgs[i].range;
    for (const k of ['nativeCol', 'nativeColOff', 'nativeRow', 'nativeRowOff']) assert.equal(got.tl[k], want[k], `${i} ${k}`);
    assert.equal(got.br, undefined);
    assert.ok(Math.abs(got.ext.width - fit.w) < 1 / 9525 + 1e-9, `ext.width ${got.ext.width} ${fit.w}`);
    assert.ok(Math.abs(got.ext.height - fit.h) < 1 / 9525 + 1e-9, `ext.height ${got.ext.height} ${fit.h}`);
    // 段 i の絵のバイト列が entries[i] の物と同じ（段の順に対応している）
    assert.ok(Buffer.from(wb.getImage(imgs[i].imageId).buffer).equals(asBuf(e.image.bytes)), `段 ${i} の絵`);
  }

  // 改ページ（ExcelJS は読み直さないので XML を見る）
  const sheetXml = zipText(res.bytes, 'xl/worksheets/sheet1.xml');
  const brks = [...sheetXml.matchAll(/<brk id="(\d+)"/g)].map((m) => Number(m[1]));
  const wantBrks = entries.map((_, i) => i).filter((i) => hasPageBreakAfter(i, n)).map((i) => blockRows(i).gap);
  assert.deepEqual(brks, wantBrks);

  const ps = ws.pageSetup;
  assert.equal(ps.printArea, `A1:D${n * 9}`);
  assert.equal(ps.paperSize, 9);
  assert.equal(ps.orientation, 'portrait');
  assert.equal(ps.horizontalCentered, true);
  assert.equal(ps.fitToPage, true);
  assert.equal(ps.fitToWidth, 1);
  assert.equal(ps.fitToHeight, 0);
  assert.deepEqual({ ...ps.margins }, { left: 0.71, right: 0.39, top: 0.47, bottom: 0.51, header: 0.2, footer: 0.2 });
  assert.equal(ws.views[0].showGridLines, false);

  assert.equal(ws.headerFooter.oddHeader, header ?? '&C' + title.replace(/&/g, '&&'));
  assert.ok(ws.headerFooter.oddHeader.length <= 255, 'ヘッダーは 255 字以下');
  assert.equal(ws.headerFooter.oddFooter, '&C&P / &N');

  assert.equal(wb.creator, 'shashin-daicho');
  assert.equal(wb.lastModifiedBy, 'shashin-daicho');
  return { wb, ws, sheetXml };
}

for (const n of [1, 3, 4, 7]) {
  test(`${n} 件：台帳の形が layout の計算どおり`, async () => {
    const entries = makeEntries(n);
    const res = await buildLedger(ExcelJS, entries, { now: NOW });
    const { wb, ws } = await checkLedger(res, entries);
    assert.equal(wb.created.getTime(), NOW.getTime());
    assert.equal(wb.modified.getTime(), NOW.getTime());
    // 撮影年月日
    assert.equal(ws.getCell(2, 4).value, '2026年4月1日');
    if (n > 1) assert.equal(ws.getCell(11, 4).value, '2026年12月31日');
    assert.equal(ws.getCell(3, 4).value, 'IMG_0001.jpg');
  });
}

test('テーマの標準の字体が 游ゴシック（minor・major の latin と Jpan）', async () => {
  const res = await buildLedger(ExcelJS, makeEntries(2), { now: NOW });
  assert.equal(res.ok, true);
  const theme = zipText(res.bytes, 'xl/theme/theme1.xml');
  for (const kind of ['minorFont', 'majorFont']) {
    const block = theme.match(new RegExp(`<a:${kind}>([\\s\\S]*?)</a:${kind}>`))[1];
    assert.match(block, /<a:latin typeface="游ゴシック"/, kind);
    assert.match(block, /<a:font script="Jpan" typeface="游ゴシック"/, kind);
  }
});

test('startNumber: 151 なら番号が 151 から', async () => {
  const entries = makeEntries(4);
  const res = await buildLedger(ExcelJS, entries, { startNumber: 151, now: NOW });
  const { ws } = await checkLedger(res, entries, { startNumber: 151 });
  assert.equal(ws.getCell(1, 4).value, 151);
  assert.equal(ws.getCell(28, 4).value, 154);
});

test('名前：式に見える名前は文字列・制御文字は取って読み直せる・長い名前は … 入り', async () => {
  const long = '基礎配筋検査の状況写真_' + 'あいうえおかきくけこ'.repeat(6) + '.jpg';
  const entries = makeEntries(3);
  entries[0].name = '=1+1.jpg';
  entries[1].name = '改行\n入り\u0000の\u001F名前\uD800￾￿.jpg';
  entries[2].name = long;
  const res = await buildLedger(ExcelJS, entries, { now: NOW });
  const { ws } = await checkLedger(res, entries);
  const c = ws.getCell(3, 4);
  assert.equal(c.type, ExcelJS.ValueType.String);
  assert.equal(c.value, '=1+1.jpg');
  assert.equal(ws.getCell(12, 4).value, '改行入りの名前.jpg');
  const shown = ws.getCell(21, 4).value;
  assert.equal(shown, fitFileName(long));
  assert.ok(shown.includes('…') && shown.endsWith('.jpg'));
});

test('題：& は && にして頭に入る', async () => {
  const entries = makeEntries(1);
  const res = await buildLedger(ExcelJS, entries, { title: 'A&B 工事', now: NOW });
  await checkLedger(res, entries, { title: 'A&B 工事' });
});

test('撮影日時が null の欄は空', async () => {
  const entries = makeEntries(1);
  entries[0].dateTimeOriginal = null;
  const res = await buildLedger(ExcelJS, entries, { now: NOW });
  const { ws } = await checkLedger(res, entries);
  assert.equal(ws.getCell(2, 4).value ?? '', '');
});

test('PNG の絵も入る', async () => {
  const entries = makeEntries(2);
  entries[1].image = { bytes: PNG, width: 4, height: 3, type: 'png' };
  const res = await buildLedger(ExcelJS, entries, { now: NOW });
  const { wb, ws } = await checkLedger(res, entries);
  const exts = imagesByRow(ws).map((im) => wb.getImage(im.imageId).extension);
  assert.deepEqual(exts, ['jpeg', 'png']);
});

test('bytes が ArrayBuffer でも Uint8Array（ずらした view）でも同じ結果', async () => {
  const a = makeEntries(2);
  const big = new Uint8Array(JPEGS[1].length + 10);
  big.set(JPEGS[1], 5);
  const b = makeEntries(2).map((e, i) => ({
    ...e, image: { ...e.image, bytes: i === 0 ? JPEGS[0].slice().buffer : big.subarray(5, 5 + JPEGS[1].length) },
  }));
  const ra = await buildLedger(ExcelJS, a, { now: NOW });
  const rb = await buildLedger(ExcelJS, b, { now: NOW });
  await checkLedger(ra, a);
  await checkLedger(rb, b);
  const wa = await reload(ra.bytes);
  const wbk = await reload(rb.bytes);
  const media = (wb) => imagesByRow(wb.worksheets[0]).map((im) => Buffer.from(wb.getImage(im.imageId).buffer).toString('base64'));
  assert.deepEqual(media(wbk), media(wa));
  assert.equal(media(wa)[0], Buffer.from(JPEG).toString('base64'));
  assert.equal(zipText(rb.bytes, 'xl/worksheets/sheet1.xml'), zipText(ra.bytes, 'xl/worksheets/sheet1.xml'));
  assert.equal(zipText(rb.bytes, 'xl/drawings/drawing1.xml'), zipText(ra.bytes, 'xl/drawings/drawing1.xml'));
});

test('だめな入力：empty・too-many・bad-entry（index つき）', async () => {
  assert.deepEqual(await buildLedger(ExcelJS, [], { now: NOW }), { ok: false, reason: 'empty' });
  assert.deepEqual(await buildLedger(ExcelJS, makeEntries(151), { now: NOW }), { ok: false, reason: 'too-many' });

  const bads = [
    (e) => { e.name = 123; },
    (e) => { delete e.image; },
    (e) => { e.image = null; },
    (e) => { e.image.bytes = new Uint8Array(0); },
    (e) => { e.image.bytes = new ArrayBuffer(0); },
    (e) => { e.image.bytes = 'abc'; },
    (e) => { e.image.width = 0; },
    (e) => { e.image.height = NaN; },
    (e) => { e.image.width = Infinity; },
    (e) => { e.image.height = -3; },
    (e) => { e.image.type = 'gif'; },
    (e) => { e.image.type = undefined; },
  ];
  for (const [k, spoil] of bads.entries()) {
    const entries = makeEntries(3);
    entries[2] = { ...entries[2], image: { ...entries[2].image } };
    spoil(entries[2]);
    const r = await buildLedger(ExcelJS, entries, { now: NOW });
    assert.equal(r.ok, false, `case ${k}`);
    assert.equal(r.reason, 'bad-entry', `case ${k}`);
    assert.equal(r.index, 2, `case ${k}`);
  }
  const r = await buildLedger(ExcelJS, [null], { now: NOW });
  assert.deepEqual([r.ok, r.reason, r.index], [false, 'bad-entry', 0]);
  const r2 = await buildLedger(ExcelJS, 'nope', { now: NOW });
  assert.deepEqual([r2.ok, r2.reason], [false, 'bad-entry']);
});

test('write-failed：ExcelJS が投げても外へ出さず、中身を message に入れる', async () => {
  class Broken { constructor() { throw new Error('こわれた束'); } }
  const r = await buildLedger({ Workbook: Broken }, makeEntries(1), { now: NOW });
  assert.equal(r.ok, false);
  assert.equal(r.reason, 'write-failed');
  assert.match(r.message, /こわれた束/);

  // 書き出しの所で投げる束
  const LateBroken = {
    Workbook: class extends ExcelJS.Workbook {
      constructor() {
        super();
        this.xlsx.writeBuffer = async () => { throw new Error('書けない'); };
      }
    },
  };
  const r2 = await buildLedger(LateBroken, makeEntries(1), { now: NOW });
  assert.equal(r2.reason, 'write-failed');
  assert.match(r2.message, /書けない/);

  const r3 = await buildLedger(null, makeEntries(1), { now: NOW });
  assert.equal(r3.reason, 'write-failed');
});

test('題：頭に入る幅（印刷の幅 516pt）を超える題は … を含めて切る（サロゲートの対を割らない・ヘッダーは 255 字以下）', async () => {
  // 段 3b：本物の Excel は頭を折り返さず、印刷の幅（約 516pt）を超えた分を切る（日本語 46 字・& は 65 字ほど）
  assert.equal(SHEET.titleMaxWidth, 540);
  assert.equal(SHEET.titleMaxChars, 100);
  const entries = makeEntries(1);
  const cases = [
    ['工'.repeat(300), '&C' + '工'.repeat(44) + '…'],
    ['😀'.repeat(150), '&C' + '😀'.repeat(44) + '…'],
    ['&'.repeat(300), '&C' + '&&'.repeat(58) + '…'],
    ['工'.repeat(45), '&C' + '工'.repeat(45)], // 12×45＝540 → そのまま
    ['工'.repeat(46), '&C' + '工'.repeat(44) + '…'],
    ['&'.repeat(60), '&C' + '&&'.repeat(60)], // 9×60＝540 → そのまま
    ['&'.repeat(61), '&C' + '&&'.repeat(58) + '…'],
    ['x'.repeat(200), '&C' + 'x'.repeat(75) + '…'], // 7×75＋12＝537（76 字目を足すと 544）
    ['W'.repeat(100), '&C' + 'W'.repeat(48) + '…'], // 11×48＋12＝540
  ];
  for (const [title, header] of cases) {
    const res = await buildLedger(ExcelJS, entries, { title, now: NOW });
    await checkLedger(res, entries, { header });
  }
});

test('題：cleanText の後に空になったら既定の題', async () => {
  const entries = makeEntries(1);
  const res = await buildLedger(ExcelJS, entries, { title: '\u0000\u0007￾\uD800', now: NOW });
  await checkLedger(res, entries, { header: '&C工事写真台帳' });
});

test('startNumber：1 以上の整数の時だけ使い、ほかは 1', async () => {
  const entries = makeEntries(1);
  for (const bad of [1.5, -3, 0, NaN, Infinity, '5', null]) {
    const res = await buildLedger(ExcelJS, entries, { startNumber: bad, now: NOW });
    const { ws } = await checkLedger(res, entries);
    assert.equal(ws.getCell(1, 4).value, 1, String(bad));
  }
  const res = await buildLedger(ExcelJS, entries, { startNumber: 7, now: NOW });
  await checkLedger(res, entries, { startNumber: 7 });
});

test('options を省いても動く（番号は 1 から・題は既定）', async () => {
  const entries = makeEntries(2);
  const res = await buildLedger(ExcelJS, entries);
  await checkLedger(res, entries);
});

// ---- 選んだ欄（fields） ----
const heights = (ws, first = 1) => Array.from({ length: 8 }, (_, k) => ws.getRow(first + k).height);
const mergesOf = (ws) => [...ws.model.merges].sort();
const want = (...a) => a.sort();
const labelsOf = (ws, first = 1) => Array.from({ length: 8 }, (_, k) => ws.getCell(first + k, 3).value ?? '');
async function buildWith(entries, fields) {
  const res = await buildLedger(ExcelJS, entries, { now: NOW, fields });
  assert.equal(res.ok, true, JSON.stringify(res));
  const wb = await reload(res.bytes);
  return { res, wb, ws: wb.worksheets[0] };
}
const LABELS8 = ['写真番号', '撮影年月日', 'ファイル名', '工種', '種別', '細別', '撮影箇所', '備考'];

test('fields なしは v0.1 と同じ', async () => {
  const { ws } = await buildWith(makeEntries(2), undefined);
  assert.deepEqual(heights(ws), [18, 18, 45, 18, 18, 18, 18, 83.25]);
  assert.deepEqual(heights(ws, 10), [18, 18, 45, 18, 18, 18, 18, 83.25]);
  assert.deepEqual(labelsOf(ws), LABELS8);
  assert.deepEqual(mergesOf(ws), want('A1:A8', 'A10:A17'));
  assert.equal(ws.getCell(1, 4).value, 1);
  assert.equal(ws.getCell(3, 4).value, 'IMG_0001.jpg');
});

test('種別を外す', async () => {
  const fields = SHEET.fields.map((f) => f.key).filter((k) => k !== 'shubetsu');
  const { ws } = await buildWith(makeEntries(1), fields);
  assert.equal(ws.getCell(5, 3).value, '細別');
  assert.equal(ws.getCell(6, 3).value, '撮影箇所');
  assert.equal(ws.getCell(7, 3).value, '備考');
  assert.deepEqual(mergesOf(ws), want('A1:A8', 'C7:C8', 'D7:D8'));
  assert.equal(ws.getRow(7).height, 18);
  assert.equal(ws.getRow(8).height, 83.25);
  assert.equal(ws.getCell(7, 3).alignment.vertical, 'top');
  assert.equal(ws.getCell(7, 4).alignment.vertical, 'top');
});

test('ファイル名を外す', async () => {
  const fields = SHEET.fields.map((f) => f.key).filter((k) => k !== 'file');
  const { ws } = await buildWith(makeEntries(1), fields);
  assert.deepEqual(heights(ws), [18, 18, 18, 18, 18, 18, 18, 110.25]);
  for (let r = 1; r <= 8; r++) assert.ok(!String(ws.getCell(r, 4).value ?? '').includes('.jpg'), `D${r}`);
  assert.deepEqual(labelsOf(ws).slice(0, 7), ['写真番号', '撮影年月日', '工種', '種別', '細別', '撮影箇所', '備考']);
  assert.deepEqual(mergesOf(ws), want('A1:A8', 'C7:C8', 'D7:D8')); // 備考が 7・8 行目を受け持つ
});

test('自動の 3 つだけ', async () => {
  const { ws } = await buildWith(makeEntries(1), ['number', 'date', 'file']);
  assert.deepEqual(labelsOf(ws).slice(0, 3), ['写真番号', '撮影年月日', 'ファイル名']);
  assert.deepEqual(mergesOf(ws), want('A1:A8', 'C4:D8'));
  assert.equal(ws.getCell(4, 3).value ?? '', '');
  assert.equal(ws.getCell(4, 4).value ?? '', '');
  assert.equal(ws.getCell(4, 3).alignment.vertical, 'top');
  assert.deepEqual(heights(ws), [18, 18, 45, 18, 18, 18, 18, 83.25]);
});

test('0 個', async () => {
  const entries = makeEntries(1);
  const { ws } = await buildWith(entries, []);
  assert.deepEqual(mergesOf(ws), want('A1:A8', 'C1:D8'));
  for (let r = 1; r <= 8; r++) {
    assert.equal(ws.getCell(r, 3).value ?? '', '', `C${r}`);
    assert.equal(ws.getCell(r, 4).value ?? '', '', `D${r}`);
  }
  assert.equal(ws.getCell(1, 1).border.top.style, 'thin');
});

test('0 個・4 枚：印刷範囲と改ページは fields なしの 4 枚と同じ', async () => {
  const entries = makeEntries(4);
  const { ws, res } = await buildWith(entries, []);
  assert.equal(ws.pageSetup.printArea, 'A1:D36');
  const brks = [...zipText(res.bytes, 'xl/worksheets/sheet1.xml').matchAll(/<brk id="(\d+)"/g)].map((m) => Number(m[1]));
  assert.deepEqual(brks, [27]);
  const ref = await buildWith(entries, undefined);
  const refBrks = [...zipText(ref.res.bytes, 'xl/worksheets/sheet1.xml').matchAll(/<brk id="(\d+)"/g)].map((m) => Number(m[1]));
  assert.deepEqual(brks, refBrks);
  assert.deepEqual(mergesOf(ws).filter((m) => m.startsWith('C1:')), ['C1:D8']);
  assert.ok(mergesOf(ws).includes('C28:D35'));
});

test('知らないキーは捨てる・配列でない物は全部', async () => {
  const entries = makeEntries(1);
  const a = await buildWith(entries, ['number', 'x', 7]);
  const b = await buildWith(entries, ['number']);
  assert.deepEqual(mergesOf(a.ws), mergesOf(b.ws));
  assert.deepEqual(labelsOf(a.ws), labelsOf(b.ws));
  assert.deepEqual(mergesOf(a.ws), want('A1:A8', 'C2:D8'));
  const c = await buildWith(entries, 'number');
  assert.deepEqual(mergesOf(c.ws), want('A1:A8'));
  assert.deepEqual(labelsOf(c.ws), LABELS8);
});

test('ファイル名を外した 16:9：絵のアンカーは、その台帳の行の高さで決まる', async () => {
  const fields = SHEET.fields.map((f) => f.key).filter((k) => k !== 'file');
  const entries = makeEntries(1);
  entries[0].image = { ...entries[0].image, width: 1600, height: 900 };
  const { ws } = await buildWith(entries, fields);
  const fit = fitImage(1600, 900);
  const want = imageAnchor(0, fit, blockLayout(fields).rowPt);
  const got = imagesByRow(ws)[0].range.tl;
  for (const k of ['nativeCol', 'nativeColOff', 'nativeRow', 'nativeRowOff']) assert.equal(got[k], want[k], k);
  assert.equal(imagesByRow(ws)[0].range.br, undefined);
});
