// 置き先に置く形（dist/shashin-daicho/）を組み直す。
// 使い方: node scripts/build_site.mjs
//   web/ の中のファイル全部（まだ無ければ飛ばす）／src/*.js → lib/／vendor/ の ExcelJS（bare）と LICENSE を写す。
//   vendor の写しは 1 バイトも変えない（写した後に SHA-256 を元と比べる）。
import { createHash } from 'node:crypto';
import { copyFileSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const DIST = path.join(ROOT, 'dist');
export const SITE = path.join(DIST, 'shashin-daicho');
const VENDOR_FILES = ['exceljs.bare.min.js', 'LICENSE.exceljs.txt'];

const sha256 = (file) => createHash('sha256').update(readFileSync(file)).digest('hex').toUpperCase();

/** @returns {{ files: string[], vendorOk: boolean }} */
export function buildSite({ quiet = false } = {}) {
  const log = quiet ? () => {} : (s) => console.log(s);
  // 消す前に、その道がリポの dist/ の下である事を確かめる
  const rel = path.relative(DIST, SITE);
  if (!rel || rel.startsWith('..') || path.isAbsolute(rel)) throw new Error(`消す道が dist/ の下でない：${SITE}`);
  // 写す前に、web/ と src/ のファイルに CR（0x0D）が混ざっていないか見る。有れば dist を作らずに止まる
  const crFiles = [];
  for (const dir of ['web', 'src']) {
    const d = path.join(ROOT, dir);
    if (!existsSync(d)) continue;
    for (const ent of readdirSync(d, { withFileTypes: true })) {
      if (ent.isFile() && readFileSync(path.join(d, ent.name)).includes(0x0d)) crFiles.push(`${dir}/${ent.name}`);
    }
  }
  if (crFiles.length) throw new Error(`改行に CR（CRLF）が混ざっているので組まずに止めた：${crFiles.join('、')}（LF にそろえてから）`);
  rmSync(SITE, { recursive: true, force: true });
  mkdirSync(path.join(SITE, 'lib'), { recursive: true });
  mkdirSync(path.join(SITE, 'vendor'), { recursive: true });

  const files = [];
  const copy = (from, to) => {
    copyFileSync(from, to);
    files.push(path.relative(SITE, to).split(path.sep).join('/'));
  };

  const web = path.join(ROOT, 'web');
  if (existsSync(web)) {
    for (const f of readdirSync(web, { withFileTypes: true }).filter((d) => d.isFile()).map((d) => d.name).sort()) {
      copy(path.join(web, f), path.join(SITE, f));
    }
  } else {
    log('web/ が無い（飛ばした）');
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
