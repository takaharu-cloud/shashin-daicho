import { test } from 'node:test';
import assert from 'node:assert/strict';
import { sortPhotos } from '../src/sort.js';

const names = (a) => a.map((x) => x.name);

test('名前順：IMG_2 は IMG_10 より先', () => {
  const r = sortPhotos([{ name: 'IMG_10.jpg', dateTimeOriginal: null }, { name: 'IMG_2.jpg', dateTimeOriginal: null }, { name: 'IMG_1.jpg', dateTimeOriginal: null }], 'name');
  assert.deepEqual(names(r), ['IMG_1.jpg', 'IMG_2.jpg', 'IMG_10.jpg']);
});

test('名前順：日本語の名前でも数字は数の大きさで並ぶ', () => {
  const r = sortPhotos([{ name: '配管10.jpg' }, { name: '配管2.jpg' }, { name: '基礎1.jpg' }], 'name');
  assert.deepEqual(names(r), ['基礎1.jpg', '配管2.jpg', '配管10.jpg']);
});

test('名前順：relativePath が有ればそちらで並べる', () => {
  const r = sortPhotos([
    { name: 'a.jpg', relativePath: 'B/a.jpg', dateTimeOriginal: null },
    { name: 'z.jpg', relativePath: 'A/z.jpg', dateTimeOriginal: null },
    { name: 'c.jpg', dateTimeOriginal: null },
  ], 'name');
  assert.deepEqual(r.map((x) => x.name), ['z.jpg', 'a.jpg', 'c.jpg']);
});

test('日付順：昇順で、同じ日時は名前の自然な順', () => {
  const r = sortPhotos([
    { name: 'IMG_10.jpg', dateTimeOriginal: '2026:04:01 09:30:00' },
    { name: 'IMG_2.jpg', dateTimeOriginal: '2026:04:01 09:30:00' },
    { name: 'IMG_5.jpg', dateTimeOriginal: '2026:03:31 18:00:00' },
    { name: 'IMG_1.jpg', dateTimeOriginal: '2026:04:02 08:00:00' },
  ], 'date');
  assert.deepEqual(names(r), ['IMG_5.jpg', 'IMG_2.jpg', 'IMG_10.jpg', 'IMG_1.jpg']);
});

test('日付順：日時が null の物は後ろにまとまり、その中は名前順', () => {
  const r = sortPhotos([
    { name: 'IMG_9.jpg', dateTimeOriginal: null },
    { name: 'IMG_3.jpg', dateTimeOriginal: '2026:04:01 09:30:00' },
    { name: 'IMG_10.jpg', dateTimeOriginal: null },
    { name: 'IMG_8.jpg', dateTimeOriginal: null },
  ], 'date');
  assert.deepEqual(names(r), ['IMG_3.jpg', 'IMG_8.jpg', 'IMG_9.jpg', 'IMG_10.jpg']);
});

test('元の配列を変えず、新しい配列を返す', () => {
  const src = [{ name: 'b', dateTimeOriginal: null }, { name: 'a', dateTimeOriginal: null }];
  const copy = src.map((x) => ({ ...x }));
  const r = sortPhotos(src, 'name');
  assert.notEqual(r, src);
  assert.deepEqual(src, copy);
  assert.deepEqual(names(r), ['a', 'b']);
});

test('知らない mode は name と同じ', () => {
  const src = [{ name: 'IMG_10.jpg', dateTimeOriginal: '2026:01:01 00:00:00' }, { name: 'IMG_2.jpg', dateTimeOriginal: '2027:01:01 00:00:00' }];
  assert.deepEqual(sortPhotos(src, 'whatever'), sortPhotos(src, 'name'));
  assert.deepEqual(names(sortPhotos(src, undefined)), ['IMG_2.jpg', 'IMG_10.jpg']);
});

test('空の配列', () => {
  assert.deepEqual(sortPhotos([], 'date'), []);
});

test('relativePath が空文字のときは name で並べる', () => {
  const r = sortPhotos([{ name: 'IMG_10.jpg', relativePath: '' }, { name: 'IMG_2.jpg', relativePath: '' }], 'name');
  assert.deepEqual(names(r), ['IMG_2.jpg', 'IMG_10.jpg']);
});
