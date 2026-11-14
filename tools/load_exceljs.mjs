// vendor/ の ExcelJS の束（ブラウザ用 UMD）を Node で読む口。テストと道具が使う。
// package.json が "type": "module" なので require() では空のモジュールになる。
// 試作 1a と同じく、中身を CommonJS の包み（module・exports を渡す関数）で直接評価する。
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import vm from 'node:vm';

const VENDOR = path.join(path.dirname(fileURLToPath(import.meta.url)), '..', 'vendor');
const FILES = { bare: 'exceljs.bare.min.js', min: 'exceljs.min.js' };
const cache = new Map();

/** @param {'bare'|'min'} which 既定はページと同じ bare 版 */
export async function loadExcelJS(which = 'bare') {
  if (!(which in FILES)) throw new Error(`束の名前は bare か min：${which}`);
  if (!cache.has(which)) {
    const file = path.join(VENDOR, FILES[which]);
    const mod = { exports: {} };
    vm.runInThisContext('(function (module, exports) {' + readFileSync(file, 'utf8') + '\n})', { filename: file })(mod, mod.exports);
    if (typeof mod.exports.Workbook !== 'function') throw new Error(`${FILES[which]} から Workbook が取れない`);
    cache.set(which, mod.exports);
  }
  return cache.get(which);
}
