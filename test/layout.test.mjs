import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  SHEET, fitImage, blockRows, imageAnchor, hasPageBreakAfter, pageCount, splitBooks, fitFileName, cleanText, nameLines,
} from '../src/layout.js';

// fitFileName と同じ見積り（U+0080 未満と半角カナ＝7、ほか＝12）
const est = (s) => Array.from(s).reduce((a, ch) => {
  const c = ch.codePointAt(0);
  return a + (c < 0x80 || (c >= 0xff61 && c <= 0xff9f) ? 7 : 12);
}, 0);
// 本物の Excel の PDF で測った 9pt の字の幅（px）。太い ASCII（W・M・大文字・記号）は 7 より広い（段 3b）。
const WIDER = { 11: 'MW', 10: '%@m', 9: '&+<=>DGHNOQUw~', 8: 'ABCKPRSTVXY' };
const estReal = (s) => Array.from(s).reduce((a, ch) => {
  const hit = Object.entries(WIDER).find(([, set]) => set.includes(ch));
  return a + (hit ? Number(hit[0]) : est(ch));
}, 0);
const hasLoneSurrogate = (s) => Array.from(s).some((ch) => {
  const c = ch.codePointAt(0);
  return c >= 0xd800 && c <= 0xdfff;
});

test('SHEET：試作 1a の数字と、凍結されている事', () => {
  assert.deepEqual([...SHEET.colPx], [420, 8, 64, 166]);
  assert.deepEqual([...SHEET.rowPt], [18, 18, 45, 18, 18, 18, 18, 83.25]);
  assert.equal(SHEET.gapPt, 7.5);
  assert.equal(SHEET.rowsPerBlock, 9);
  assert.equal(SHEET.blocksPerPage, 3);
  assert.equal(SHEET.maxPerBook, 150);
  assert.equal(SHEET.frame.wPx, 420);
  assert.equal(SHEET.frame.hPx, 315);
  assert.equal(SHEET.imageBox.wPx, 396);
  assert.equal(SHEET.imageBox.hPx, 291);
  assert.equal(SHEET.fileNameMaxWidth, 440);
  // 枠の高さ（pt の合計）＝ 315px
  assert.equal(SHEET.rowPt.reduce((a, b) => a + b, 0) * 96 / 72, 315);
  assert.ok(Object.isFrozen(SHEET));
  assert.ok(Object.isFrozen(SHEET.colPx));
  assert.ok(Object.isFrozen(SHEET.rowPt));
  assert.ok(Object.isFrozen(SHEET.pageSetup.margins));
});

test('fitImage：いろいろな縦横比で箱からはみ出さず・比がずれず・真ん中', () => {
  const cases = [[4000, 3000], [3000, 4000], [1920, 1080], [1000, 1000], [3000, 1000], [8, 6], [1, 3], [1200, 400]];
  for (const [w0, h0] of cases) {
    const f = fitImage(w0, h0);
    assert.ok(f, `${w0}x${h0}`);
    assert.ok(f.w <= 396 + 1e-9 && f.h <= 291 + 1e-9, `箱に収まる ${w0}x${h0}`);
    // どちらかの辺が箱にぴったり（小さい絵も大きくする）
    assert.ok(Math.abs(f.w - 396) < 1e-9 || Math.abs(f.h - 291) < 1e-9, `箱に合わせる ${w0}x${h0}`);
    const drift = Math.abs((f.w / f.h) / (w0 / h0) - 1);
    assert.ok(drift <= 0.0001, `縦横比のずれ ${drift}`);
    assert.ok(Math.abs(f.x + f.w / 2 - 210) < 1e-9, '横の真ん中');
    assert.ok(Math.abs(f.y + f.h / 2 - 157.5) < 1e-9, '縦の真ん中');
    assert.ok(f.x >= 12 - 1e-9 && f.y >= 12 - 1e-9, '余白 12px より内側');
  }
});

test('fitImage：とても小さい絵は箱に合わせて大きくなる', () => {
  const f = fitImage(8, 6);
  assert.ok(f.w > 300);
  assert.equal(f.h, 291);
});

test('fitImage：0・負・NaN・Infinity・数でない物は null', () => {
  for (const [w, h] of [[0, 10], [10, 0], [-1, 10], [10, -5], [NaN, 10], [10, NaN], [Infinity, 10], [10, Infinity], ['10', 10], [undefined, 10], [null, 10]]) {
    assert.equal(fitImage(w, h), null, `${w},${h}`);
  }
});

test('blockRows：9 行ごと', () => {
  assert.deepEqual(blockRows(0), { first: 1, last: 8, gap: 9 });
  assert.deepEqual(blockRows(1), { first: 10, last: 17, gap: 18 });
  assert.deepEqual(blockRows(149), { first: 1342, last: 1349, gap: 1350 });
});

test('hasPageBreakAfter：n＝1〜10 の全部で、3 段ごと・最後の段には付けない', () => {
  for (let n = 1; n <= 10; n++) {
    const got = [];
    for (let i = 0; i < n; i++) if (hasPageBreakAfter(i, n)) got.push(i);
    const want = [];
    for (let i = 2; i < n - 1; i += 3) want.push(i);
    assert.deepEqual(got, want, `n=${n}`);
  }
});

test('pageCount：ceil(n/3)', () => {
  assert.deepEqual([0, 1, 2, 3, 4, 6, 7, 150].map(pageCount), [0, 1, 1, 1, 2, 2, 3, 50]);
});

test('splitBooks：0・1・150・151・300・301', () => {
  assert.deepEqual(splitBooks(0), []);
  assert.deepEqual(splitBooks(1), [{ start: 0, end: 1 }]);
  assert.deepEqual(splitBooks(150), [{ start: 0, end: 150 }]);
  assert.deepEqual(splitBooks(151), [{ start: 0, end: 150 }, { start: 150, end: 151 }]);
  assert.deepEqual(splitBooks(300), [{ start: 0, end: 150 }, { start: 150, end: 300 }]);
  assert.deepEqual(splitBooks(301), [{ start: 0, end: 150 }, { start: 150, end: 300 }, { start: 300, end: 301 }]);
  assert.deepEqual(splitBooks(7, 3), [{ start: 0, end: 3 }, { start: 3, end: 6 }, { start: 6, end: 7 }]);
});

test('imageAnchor：前の行の高さの合計＋nativeRowOff が y と合い、行の中に収まる', () => {
  const sizes = [[4000, 3000], [3000, 4000], [1920, 1080], [1000, 1000], [3000, 1000], [1200, 400], [1, 3]];
  for (const i of [0, 1, 2, 3, 149]) {
    for (const [w0, h0] of sizes) {
      const f = fitImage(w0, h0);
      const a = imageAnchor(i, f);
      assert.equal(a.nativeCol, 0);
      for (const k of ['nativeColOff', 'nativeRow', 'nativeRowOff']) assert.ok(Number.isInteger(a[k]), k);
      assert.ok(Math.abs(a.nativeColOff - f.x * 9525) <= 0.5);
      const first0 = blockRows(i).first - 1; // 0 始まり
      const rr = a.nativeRow - first0;
      assert.ok(rr >= 0 && rr < 8, `段の中の行 ${rr}`);
      const before = SHEET.rowPt.slice(0, rr).reduce((s, pt) => s + pt * 12700, 0);
      assert.ok(Math.abs(before + a.nativeRowOff - f.y * 9525) <= 1, `y が合う i=${i} ${w0}x${h0}`);
      assert.ok(a.nativeRowOff >= 0 && a.nativeRowOff < SHEET.rowPt[rr] * 12700, '行の高さより小さい');
    }
  }
});

test('imageAnchor：2 段目以降も段の中の位置は 1 段目と同じ（行が 9 ずつずれるだけ）', () => {
  const f = fitImage(3000, 4000);
  const a0 = imageAnchor(0, f);
  for (const i of [1, 5, 100]) {
    const a = imageAnchor(i, f);
    assert.equal(a.nativeRow, a0.nativeRow + 9 * i);
    assert.equal(a.nativeRowOff, a0.nativeRowOff);
    assert.equal(a.nativeColOff, a0.nativeColOff);
  }
});

test('fitFileName：短い名前はそのまま', () => {
  for (const s of ['IMG_0001.jpg', '配管の写真.jpg', '', 'a'.repeat(62)]) assert.equal(fitFileName(s), s);
  const just = '工'.repeat(33) + 'abcdef'; // 12×33＋7×6＝438 → そのまま
  assert.equal(fitFileName(just), just);
  const over = '工'.repeat(36) + 'ab'; // 12×36＋7×2＝446 → 詰める
  assert.notEqual(fitFileName(over), over);
});

test('fitFileName：長い名前は真ん中を … で詰め、440 以下・拡張子が残る・… は 1 つ', () => {
  const cases = [
    ['IMG_' + '0123456789'.repeat(12) + '.jpg', '.jpg'],
    ['基礎配筋検査の状況写真_北側_通り芯_' + 'あいうえおかきくけこ'.repeat(6) + '.JPG', '.JPG'],
    ['写真😀🏗️' + '🧱'.repeat(50) + '終わり.jpeg', '.jpeg'],
    ['ｶﾀｶﾅ'.repeat(40) + '.png', '.png'],
    ['拡張子の無い名前'.repeat(10), ''],
    ['名前'.repeat(30) + '.とても長い拡張子のような物です', ''],
  ];
  for (const [s, ext] of cases) {
    const r = fitFileName(s);
    assert.ok(est(s) > 440, '元は長い');
    assert.ok(est(r) <= 440, `幅 ${est(r)}：${r}`);
    assert.equal(r.split('…').length - 1, 1, '… は 1 つ');
    if (ext) assert.ok(r.endsWith(ext), `拡張子が残る：${r}`);
    assert.ok(!hasLoneSurrogate(r), 'サロゲートの対を割らない');
    const [head, tail] = r.split('…');
    assert.ok(s.startsWith(head) && s.endsWith(tail), '頭と尾は元の名前から');
    assert.ok(head.length > 0 && tail.length > 0);
    // 頭と尾（拡張子を除く）は、ほぼ半分ずつ
    const hw = est(head), tw = est(tail) - est(ext);
    assert.ok(Math.abs(hw - tw) <= 24, `頭 ${hw}・尾 ${tw}`);
  }
});

test('fitFileName：太い字（W・M・大文字）の名前も、実測の幅で 440 以下に詰める（段 3b・欄の 3 行に収める）', () => {
  assert.equal(SHEET.fileNameMaxWidth, 440);
  const cases = [
    'W'.repeat(60) + '.JPG',
    'M'.repeat(50) + '.jpg',
    'A'.repeat(70) + '.JPG',
    'm'.repeat(70) + '.jpg',
    'WMWM'.repeat(12) + '.jpg',
    'ABCDEFGHIJKLMNOPQRSTUVWXYZ'.repeat(2) + '.JPG',
  ];
  for (const s of cases) {
    const r = fitFileName(s);
    assert.ok(estReal(s) > 440, '元は長い');
    assert.ok(estReal(r) <= 440, `幅 ${estReal(r)}：${r}`);
    assert.ok(r.endsWith(s.slice(-4)), '拡張子が残る');
  }
  assert.equal(fitFileName('W'.repeat(30) + '.JPG'), 'W'.repeat(30) + '.JPG'); // 11×30＋... は収まる
});

test('nameLines：欄の折り返しの見積り（本物の Excel の PDF で見た事）', () => {
  // 1 行は W が 14 字・日本語が 13 字まで（W 15 字・日本語 14 字は 2 行目へ）
  assert.equal(nameLines('W'.repeat(14)), 1);
  assert.equal(nameLines('W'.repeat(15)), 2);
  assert.equal(nameLines('工'.repeat(13)), 1);
  assert.equal(nameLines('工'.repeat(14)), 2);
  assert.equal(nameLines('工'.repeat(39)), 3);
  assert.equal(nameLines('工'.repeat(40)), 4);
  // 「…」の後ろは折り返せる＝後ろの長い続きは新しい行から始まる（W 18 字＋… ＋ W 18 字＋.JPG は 4 行で、末尾が切れた）
  assert.equal(nameLines('W'.repeat(18) + '…' + 'W'.repeat(18) + '.JPG'), 4);
  // 空白で区切った単語は、収まらなければ次の行へ（1 行に 11 字の単語が 1 つしか入らない）
  assert.equal(nameLines(('ABCDEFGHIJK ').repeat(4).trim()), 4);
});

test('fitFileName：詰めた後が必ず 3 行以内（… の後ろの折り返しと、空白で区切った名前を含む）', () => {
  const names = [
    'W'.repeat(60) + '.JPG', 'M'.repeat(50) + '.jpg', 'A'.repeat(70) + '.JPG', 'm'.repeat(70) + '.jpg',
    'ABCDEFGHIJ KLMNOPQRST UVWXYZABCD EFGHIJKLMN OPQRSTUVWX YZ.JPG',
    ('ABCDEFGHIJK '.repeat(8)).trim() + '.jpg',
    'abcdefghij0123456789'.repeat(3) + 'abcdef.jpg',
    'ＡＢＣ'.repeat(12) + '.jpg',
  ];
  for (const n of names) {
    const r = fitFileName(n);
    assert.ok(nameLines(r) <= 3, `${nameLines(r)} 行：${r}`);
    assert.ok(r.endsWith(n.slice(-4)), '拡張子が残る');
  }
  const ok = 'ABCDEFGHIJK ABCDEFGHIJK ABCDEFGHIJK.jpg'; // 3 行に収まる＝そのまま
  assert.equal(fitFileName(ok), ok);
});

test('cleanText：制御文字と対の無いサロゲートを取り除く', () => {
  assert.equal(cleanText('a\u0000b\u0009c\u000Ad\u001Fe\u007Ff'), 'abcdef');
  assert.equal(cleanText('写真\uD800.jpg'), '写真.jpg');
  assert.equal(cleanText('写真\uDC00x'), '写真x');
  assert.equal(cleanText('😀ok'), '😀ok');
  assert.equal(cleanText('\u0080 あ'), '\u0080 あ');
  assert.equal(cleanText(''), '');
  assert.equal(cleanText('a￾b￿c�d'), 'abc�d'); // XML に書けない U+FFFE・U+FFFF も除く
  assert.equal(cleanText(null), '');
  assert.equal(cleanText(123), '');
});
