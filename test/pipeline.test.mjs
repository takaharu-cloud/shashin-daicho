import { test } from 'node:test';
import assert from 'node:assert/strict';
import { classify, targetSize, processAllWith } from '../src/pipeline.js';

// readPhotoInfo の返り値の形（中身の種類だけ見る）
const info = (kind) => ({ kind, dateTimeOriginal: null, orientation: null });
const JPEG = info('jpeg');
const PNG = info('png');
const HEIC = info('heic');
const OTHER = info('other');

// ---- classify（manifest に有る種類を全部）----
test('classify：ふつうの JPEG と PNG は photo', () => {
  assert.equal(classify('IMG_0001.jpg', 419681, JPEG), 'photo');
  assert.equal(classify('IMG_0005.jpeg', 417595, JPEG), 'photo');
  assert.equal(classify('IMG_0006.png', 20998, PNG), 'photo');
});

test('classify：大文字の拡張子でも同じ', () => {
  assert.equal(classify('IMG_0034.JPG', 2568511, JPEG), 'photo');
  assert.equal(classify('IMG_0007.HEIC', 3012, HEIC), 'heic');
  assert.equal(classify('A.GIF', 100, OTHER), 'unsupported');
  assert.equal(classify('A.JPE', 0, OTHER), 'broken');
});

test('classify：にせの HEIC（頭が heic）は heic', () => {
  assert.equal(classify('IMG_9001.HEIC', 3012, HEIC), 'heic');
  assert.equal(classify('IMG_9001.heif', 3012, HEIC), 'heic');
});

test('classify：0 バイトで名前が写真の拡張子なら broken', () => {
  for (const n of ['a.jpg', 'a.jpeg', 'a.jpe', 'a.png', 'a.heic', 'a.heif']) {
    assert.equal(classify(n, 0, OTHER), 'broken', n);
  }
});

test('classify：中身が文字の .jpg は broken', () => {
  assert.equal(classify('IMG_9004.jpg', 76, OTHER), 'broken');
  assert.equal(classify('x.png', 10, OTHER), 'broken');
});

test('classify：memo.txt と Thumbs.db は not-image', () => {
  assert.equal(classify('memo.txt', 50, OTHER), 'not-image');
  assert.equal(classify('Thumbs.db', 2504, OTHER), 'not-image');
});

test('classify：ほかの絵の拡張子は unsupported', () => {
  for (const n of ['a.gif', 'a.bmp', 'a.webp', 'a.tif', 'a.tiff', 'a.avif']) {
    assert.equal(classify(n, 1000, OTHER), 'unsupported', n);
  }
  assert.equal(classify('a.gif', 0, OTHER), 'unsupported'); // 0 バイトでも、写真の拡張子でなければ broken にしない
});

test('classify：拡張子なしは not-image（中身が写真なら photo）', () => {
  assert.equal(classify('README', 10, OTHER), 'not-image');
  assert.equal(classify('noext', 0, OTHER), 'not-image');
  assert.equal(classify('noext', 1000, JPEG), 'photo');
});

test('classify：点で終わる名前・点で始まる名前', () => {
  assert.equal(classify('a.', 10, OTHER), 'not-image');
  assert.equal(classify('.jpg', 0, OTHER), 'broken');
});

test('classify：info が無い・名前が文字列でなくても投げない', () => {
  assert.equal(classify('a.jpg', 10, null), 'broken');
  assert.equal(classify(undefined, 10, undefined), 'not-image');
});

// ---- targetSize ----
test('targetSize：横長・縦長・細長を長辺 1200 に縮める', () => {
  assert.deepEqual(targetSize(4000, 3000), { w: 1200, h: 900 });
  assert.deepEqual(targetSize(3000, 4000), { w: 900, h: 1200 });
  assert.deepEqual(targetSize(6000, 2000), { w: 1200, h: 400 });
});

test('targetSize：長辺が 1200 以下は拡大しない', () => {
  assert.deepEqual(targetSize(800, 600), { w: 800, h: 600 });
  assert.deepEqual(targetSize(1200, 1200), { w: 1200, h: 1200 });
});

test('targetSize：細すぎる辺は 1 にする', () => {
  assert.deepEqual(targetSize(1, 5000), { w: 1, h: 1200 });
  assert.deepEqual(targetSize(5000, 1), { w: 1200, h: 1 });
});

test('targetSize：maxLongSide を変えられる', () => {
  assert.deepEqual(targetSize(4000, 3000, 400), { w: 400, h: 300 });
});

test('targetSize：不正は null', () => {
  for (const [w, h, m] of [[0, 10], [10, 0], [-1, 10], [NaN, 10], [10, Infinity], ['4000', 3000], [null, 1], [4000, 3000, 0], [4000, 3000, NaN]]) {
    assert.equal(targetSize(w, h, m), null, `${w}×${h} max=${m}`);
  }
});

// ---- processAll（processPhoto を差し替えて Node で試す）----
const fakeFile = (name, extra = {}) => ({ name, size: 10, ...extra });
const fakeProcess = (log = []) => async (file, opt) => {
  log.push({ name: file.name, opt });
  await new Promise((r) => setTimeout(r, 1));
  if (file.name.endsWith('.txt')) {
    return { ok: false, name: file.name, relativePath: opt.relativePath ?? '', reason: 'not-image', detail: 'x' };
  }
  return { ok: true, name: file.name, relativePath: opt.relativePath ?? '' };
};

test('processAll：入力の順に、ok と ok でない物に分けて返す', async () => {
  const files = [fakeFile('b.jpg'), fakeFile('a.txt'), fakeFile('c.jpg'), fakeFile('d.txt')];
  const r = await processAllWith(fakeProcess(), files);
  assert.equal(r.aborted, false);
  assert.deepEqual(r.photos.map((p) => p.name), ['b.jpg', 'c.jpg']);
  assert.deepEqual(r.failed.map((p) => p.name), ['a.txt', 'd.txt']);
});

test('processAll：1 枚ずつ順に処理する（並列にしない）', async () => {
  let running = 0;
  let most = 0;
  const one = async (file) => {
    running++;
    most = Math.max(most, running);
    await new Promise((r) => setTimeout(r, 2));
    running--;
    return { ok: true, name: file.name, relativePath: '' };
  };
  await processAllWith(one, [fakeFile('1.jpg'), fakeFile('2.jpg'), fakeFile('3.jpg')]);
  assert.equal(most, 1);
});

test('processAll：relativePath は渡された物、無ければ webkitRelativePath、それも無ければ空', async () => {
  const log = [];
  await processAllWith(fakeProcess(log), [
    { file: fakeFile('a.jpg', { webkitRelativePath: 'x/a.jpg' }), relativePath: '上/a.jpg' },
    fakeFile('b.jpg', { webkitRelativePath: 'x/b.jpg' }),
    fakeFile('c.jpg'),
    fakeFile('d.jpg', { webkitRelativePath: '' }),
  ]);
  assert.deepEqual(log.map((l) => l.opt.relativePath), ['上/a.jpg', 'x/b.jpg', '', '']);
});

test('processAll：maxLongSide と quality を 1 枚ごとに渡す', async () => {
  const log = [];
  await processAllWith(fakeProcess(log), [fakeFile('a.jpg')], { maxLongSide: 800, quality: 0.7 });
  assert.equal(log[0].opt.maxLongSide, 800);
  assert.equal(log[0].opt.quality, 0.7);
});

test('processAll：onProgress は 1 枚ごとに (済んだ数, 全部, その結果)', async () => {
  const calls = [];
  await processAllWith(fakeProcess(), [fakeFile('a.jpg'), fakeFile('b.txt')], {
    onProgress: (done, total, last) => calls.push([done, total, last.name, last.ok]),
  });
  assert.deepEqual(calls, [[1, 2, 'a.jpg', true], [2, 2, 'b.txt', false]]);
});

test('processAll：onProgress が投げても止まらない', async () => {
  const r = await processAllWith(fakeProcess(), [fakeFile('a.jpg'), fakeFile('b.jpg')], {
    onProgress: () => { throw new Error('画面の側の失敗'); },
  });
  assert.equal(r.photos.length, 2);
});

test('processAll：途中で止めると、いまの 1 枚を終えてから止まり aborted', async () => {
  const ac = new AbortController();
  const log = [];
  const one = async (file, opt) => {
    log.push(file.name);
    if (file.name === 'b.jpg') ac.abort(); // b の処理の途中で止める
    await new Promise((r) => setTimeout(r, 2));
    return { ok: true, name: file.name, relativePath: opt.relativePath };
  };
  const r = await processAllWith(one, [fakeFile('a.jpg'), fakeFile('b.jpg'), fakeFile('c.jpg')], { signal: ac.signal });
  assert.equal(r.aborted, true);
  assert.deepEqual(log, ['a.jpg', 'b.jpg']);
  assert.deepEqual(r.photos.map((p) => p.name), ['a.jpg', 'b.jpg']);
});

test('processAll：始める前に止まっていれば 1 枚も処理しない', async () => {
  const ac = new AbortController();
  ac.abort();
  const log = [];
  const r = await processAllWith(fakeProcess(log), [fakeFile('a.jpg')], { signal: ac.signal });
  assert.equal(r.aborted, true);
  assert.equal(log.length, 0);
  assert.deepEqual(r.photos, []);
});

test('processAll：1 枚の処理が投げても、broken にして続ける', async () => {
  const one = async (file) => {
    if (file.name === 'a.jpg') throw new Error('思わぬ失敗');
    return { ok: true, name: file.name, relativePath: '' };
  };
  const r = await processAllWith(one, [fakeFile('a.jpg'), fakeFile('b.jpg')]);
  assert.deepEqual(r.photos.map((p) => p.name), ['b.jpg']);
  assert.equal(r.failed[0].name, 'a.jpg');
  assert.equal(r.failed[0].reason, 'broken');
});

test('processAll：配列でない・中身が File でない物は投げずに扱う', async () => {
  assert.deepEqual(await processAllWith(fakeProcess(), null), { photos: [], failed: [], aborted: false });
  const r = await processAllWith(fakeProcess(), [null, { file: null }]);
  assert.equal(r.photos.length, 0);
  assert.equal(r.failed.length, 2);
  assert.equal(r.failed[0].reason, 'not-image');
});
