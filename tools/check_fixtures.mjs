// manifest と src/exif.js の読み取りを突き合わせる。
// 使い方: node tools/check_fixtures.mjs smoke|mix60|bulk300|big|all
// ファイルの先頭 128KB だけを読み、kind・orientation・dateTimeOriginal が manifest と同じかを比べる。
import { openSync, readSync, closeSync, readFileSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { readPhotoInfo } from '../src/exif.js';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'fixtures');
const SETS = ['smoke', 'mix60', 'bulk300', 'big'];
const HEAD = 128 * 1024;

function readHead(file) {
  const fd = openSync(file, 'r');
  try {
    const buf = Buffer.alloc(HEAD);
    const n = readSync(fd, buf, 0, HEAD, 0);
    return new Uint8Array(buf.buffer, buf.byteOffset, n);
  } finally {
    closeSync(fd);
  }
}

function checkSet(name) {
  const mp = path.join(ROOT, `${name}.manifest.json`);
  if (!existsSync(mp)) {
    console.log(`${name}: manifest が無い`);
    return { n: 0, bad: 1 };
  }
  const m = JSON.parse(readFileSync(mp, 'utf8'));
  let bad = 0;
  for (const e of m.files) {
    const file = path.join(ROOT, name, ...e.path.split('/'));
    let info;
    try {
      info = readPhotoInfo(readHead(file));
    } catch (err) {
      console.log(`${name}/${e.path}: 読めない（${err.message}）`);
      bad++;
      continue;
    }
    for (const k of ['kind', 'orientation', 'dateTimeOriginal']) {
      if (info[k] !== e[k]) {
        console.log(`${name}/${e.path}: ${k} が ${JSON.stringify(info[k])}（manifest は ${JSON.stringify(e[k])}）`);
        bad++;
      }
    }
  }
  console.log(`${name}: ${m.files.length} 件・ちがい ${bad}`);
  return { n: m.files.length, bad };
}

const arg = process.argv[2];
if (arg !== 'all' && !SETS.includes(arg)) {
  console.error('使い方: node tools/check_fixtures.mjs smoke|mix60|bulk300|big|all');
  process.exit(2);
}
let n = 0;
let bad = 0;
for (const s of arg === 'all' ? SETS : [arg]) {
  const r = checkSet(s);
  n += r.n;
  bad += r.bad;
}
if (arg === 'all') console.log(`合計: ${n} 件・ちがい ${bad}`);
process.exit(bad ? 1 : 0);
