// 架空の写真の manifest から、Node だけで台帳（xlsx）と、次の段の検査が読む「正解」を出す。
// 使い方: node tools/build_ledger.mjs <組> [--sort date|name] [--fields <キーをカンマで。0 個は none>] [--tag <名前>]
//   --sort は既定が date。--fields が無ければ全部の欄。--tag は出力の名前の <組>_<sort> の後ろに _<名前> を付ける。
// 出力: out/<組>_<sort>[_<名前>]_<k>.xlsx（k は 1 から）と out/<組>_<sort>[_<名前>].expect.json
//   expect.json の fields は、--fields が有る時だけ書く（正規化した後の配列）。
import { readFileSync, writeFileSync, mkdirSync, existsSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { loadExcelJS } from './load_exceljs.mjs';
import { buildLedger } from '../src/ledger.js';
import { splitBooks, fitFileName, cleanText, normalizeFields, FIELD_KEYS } from '../src/layout.js';
import { sortPhotos } from '../src/sort.js';
import { formatDateJa } from '../src/exif.js';

const ROOT = path.join(path.dirname(fileURLToPath(import.meta.url)), '..');
const FIX = path.join(ROOT, 'fixtures');
const OUT = path.join(ROOT, 'out');

function usage(msg) {
  if (msg) console.error(msg);
  console.error('使い方: node tools/build_ledger.mjs <組> [--sort date|name] [--fields <キーをカンマで。0 個は none>] [--tag <名前>]');
  process.exit(2);
}

const args = process.argv.slice(2);
let set = null;
let sort = 'date';
let fieldsArg = null;
let tag = null;
for (let k = 0; k < args.length; k++) {
  if (args[k] === '--sort') sort = args[++k];
  else if (args[k] === '--fields') fieldsArg = args[++k];
  else if (args[k] === '--tag') tag = args[++k];
  else if (!set) set = args[k];
  else usage(`知らない引数：${args[k]}`);
}
if (!set || !/^[A-Za-z0-9_-]+$/.test(set)) usage('組の名前が無いか、使えない字が有る');
if (sort !== 'date' && sort !== 'name') usage(`--sort は date か name：${sort}`);
if (tag !== null && !/^[A-Za-z0-9_-]+$/.test(tag)) usage('--tag の名前が無いか、使えない字が有る');
let fields; // undefined＝全部（buildLedger に渡さず、expect.json にも書かない）
if (fieldsArg !== null) {
  if (typeof fieldsArg !== 'string' || fieldsArg === '') usage('--fields にキーが無い（0 個は none）');
  const keys = fieldsArg === 'none' ? [] : fieldsArg.split(',');
  const bad = keys.filter((k) => !FIELD_KEYS.includes(k));
  if (bad.length) usage(`知らない欄のキー：${bad.join(',')}（使えるのは ${FIELD_KEYS.join(',')}）`);
  fields = normalizeFields(keys);
}
const stemName = `${set}_${sort}${tag ? `_${tag}` : ''}`;

const manifestPath = path.join(FIX, `${set}.manifest.json`);
if (!existsSync(manifestPath)) usage(`manifest が無い：fixtures/${set}.manifest.json（python -X utf8 tools/make_fixtures.py --set ${set}）`);
const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));

const items = manifest.files
  .filter((f) => f.expect === 'ok' && f.copy)
  .map((f) => ({
    name: f.path.split('/').pop(),
    relativePath: f.path,
    dateTimeOriginal: f.dateTimeOriginal ?? null,
    label: f.label,
    copy: f.copy,
  }));
const sorted = sortPhotos(items, sort);

const ExcelJS = await loadExcelJS('bare');
mkdirSync(OUT, { recursive: true });
const books = [];
let failed = false;
for (const [k, { start, end }] of splitBooks(sorted.length).entries()) {
  const part = sorted.slice(start, end);
  // 写しの中身はブックごとに読む（全部を一度に持たない）
  const entries = part.map((it) => {
    const buf = readFileSync(path.join(FIX, it.copy.path));
    return {
      name: it.name,
      relativePath: it.relativePath,
      dateTimeOriginal: it.dateTimeOriginal,
      image: { bytes: new Uint8Array(buf.buffer, buf.byteOffset, buf.byteLength), width: it.copy.w, height: it.copy.h, type: 'jpeg' },
    };
  });
  const startNumber = start + 1;
  const res = await buildLedger(ExcelJS, entries, fields === undefined ? { startNumber } : { startNumber, fields });
  const file = `${stemName}_${k + 1}.xlsx`;
  if (!res.ok) {
    console.error(`${file}: だめ（${res.reason}${res.index !== undefined ? ` index=${res.index}` : ''}${res.message ? ` ${res.message}` : ''}）`);
    failed = true;
    continue;
  }
  writeFileSync(path.join(OUT, file), res.bytes);
  console.log(`out/${file}  ${part.length} 枚  ${res.pages} ページ  ${res.bytes.length} B`);
  books.push({
    file,
    startNumber,
    pages: res.pages,
    photos: part.map((it, j) => ({
      number: startNumber + j,
      name: it.name,
      shownName: fitFileName(cleanText(it.name)),
      dateText: formatDateJa(it.dateTimeOriginal),
      label: it.label,
      imageW: it.copy.w,
      imageH: it.copy.h,
    })),
  });
}
const expectFile = `${stemName}.expect.json`;
const expect = fields === undefined ? { set, sort, books } : { set, sort, fields, books };
writeFileSync(path.join(OUT, expectFile), JSON.stringify(expect, null, 2) + '\n');
console.log(`out/${expectFile}  ${books.length} 冊  ${sorted.length} 枚`);
if (failed) process.exit(1);
