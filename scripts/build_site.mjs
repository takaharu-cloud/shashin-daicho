// 置き先に置く形（dist/shashin-daicho/）を組み直す。
// 使い方: node scripts/build_site.mjs
//   web/ の index.html・app.js・app.css（まだ無ければ飛ばす）／src/*.js → lib/／vendor/ の ExcelJS（bare）と LICENSE を写す。
//   vendor の写しは 1 バイトも変えない（写した後に SHA-256 を元と比べる）。
import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const DIST = path.join(ROOT, 'dist');
export const SITE = path.join(DIST, 'shashin-daicho');
const WEB_FILES = ['index.html', 'app.js', 'app.css'];
const VENDOR_FILES = ['exceljs.bare.min.js', 'LICENSE.exceljs.txt'];

const sha256 = (file) => createHash('sha256').update(readFileSync(file)).digest('hex').toUpperCase();

/** @returns {{ files: string[], vendorOk: boolean }} */
export function buildSite({ quiet = false } = {}) {
  const log = quiet ? () => {} : (s) => console.log(s);
  // 消す前に、その道がリポの dist/ の下である事を確かめる
  const rel = path.relative(DIST, SITE);
  if (!rel || rel.startsWith('..') || path.isAbsolute(rel)) throw new Error(`消す道が dist/ の下でない：${SITE}`);
  rmSync(SITE, { recursive: true, force: true });
  mkdirSync(path.join(SITE, 'lib'), { recursive: true });
  mkdirSync(path.join(SITE, 'vendor'), { recursive: true });

  const files = [];
  const copy = (from, to) => {
    copyFileSync(from, to);
    files.push(path.relative(SITE, to).split(path.sep).join('/'));
  };

  for (const f of WEB_FILES) {
    const from = path.join(ROOT, 'web', f);
    if (existsSync(from)) copy(from, path.join(SITE, f));
    else log(`web/${f} が無い（飛ばした）`);
  }
  for (const f of readdirSync(path.join(ROOT, 'src')).filter((n) => n.endsWith('.js')).sort()) {
    copy(path.join(ROOT, 'src', f), path.join(SITE, 'lib', f));
  }
  let vendorOk = true;
  for (const f of VENDOR_FILES) {
    const from = path.join(ROOT, 'vendor', f);
    const to = path.join(SITE, 'vendor', f);
    copy(from, to);
    const a = sha256(from);
    const b = sha256(to);
    if (a !== b) vendorOk = false;
    log(`vendor/${f}  SHA-256 ${b}  ${a === b ? '元と同じ' : `元と違う（元 ${a}）`}`);
  }
  log(`dist/shashin-daicho/ に ${files.length} ファイル`);
  return { files, vendorOk };
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const { vendorOk } = buildSite();
  if (!vendorOk) process.exit(1);
}
