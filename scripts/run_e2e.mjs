// 通しの試験：置き先と同じ CSP の下で、ブラウザ（headless）に架空の写真を通して台帳を作らせる。
// 使い方: node scripts/run_e2e.mjs <edge|chrome> <組> [--sort date|name]
//   build_site → サーバー（--e2e・127.0.0.1・空いている口）→ ブラウザを CDP で起こす → /__e2e/harness.html を開く →
//   報告を待つ → ブラウザとサーバーを止める → out/e2e/<ブラウザ>_<組>_<sort>.expect.json を書く。
//   次の検査：python -X utf8 tools/inspect_xlsx.py out/e2e/<ブラウザ>_<組>_<sort>.expect.json
// 前後で python -X utf8 tools/make_fixtures.py --verify <組> を回し、原本が書き換わっていない事を確かめる。
// 外れが有れば 1 行ずつ出して、終了コード 1。
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { buildSite } from './build_site.mjs';
import { createServer, OUT_E2E } from './serve_csp.mjs';
import { launch, openPage, stop } from './lib/cdp.mjs';
import { fitFileName, cleanText, pageCount } from '../src/layout.js';
import { formatDateJa } from '../src/exif.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

function usage(msg) {
  if (msg) console.error(msg);
  console.error('使い方: node scripts/run_e2e.mjs <edge|chrome> <組> [--sort date|name]');
  process.exit(2);
}

const args = process.argv.slice(2);
let which = null;
let set = null;
let sort = 'date';
for (let k = 0; k < args.length; k++) {
  if (args[k] === '--sort') sort = args[++k];
  else if (!which) which = args[k];
  else if (!set) set = args[k];
  else usage(`知らない引数：${args[k]}`);
}
if (which !== 'edge' && which !== 'chrome') usage('ブラウザは edge か chrome');
if (!set || !/^[A-Za-z0-9_-]+$/.test(set)) usage('組の名前が無いか、使えない字が有る');
if (sort !== 'date' && sort !== 'name') usage(`--sort は date か name：${sort}`);
const manifestPath = path.join(ROOT, 'fixtures', `${set}.manifest.json`);
if (!existsSync(manifestPath)) usage(`manifest が無い：fixtures/${set}.manifest.json`);
const manifest = JSON.parse(readFileSync(manifestPath, 'utf8'));

function verifyFixtures(when) {
  const r = spawnSync('python', ['-X', 'utf8', path.join('tools', 'make_fixtures.py'), '--verify', set], { cwd: ROOT, encoding: 'utf8' });
  const lines = (r.stdout || '').trim().split(/\r?\n/);
  console.log(`原本の照合（${when}）：${lines[lines.length - 1]}`);
  return r.status === 0;
}

const problems = [];
const t0 = Date.now();
if (!verifyFixtures('前')) problems.push('原本の照合（前）でちがいが有る');

const { vendorOk } = buildSite({ quiet: true });
if (!vendorOk) problems.push('dist の vendor の SHA-256 が元とちがう');

let onReport = null;
const reportArrived = new Promise((resolve) => { onReport = resolve; });
const serverLog = [];
const server = createServer({
  e2e: true,
  onReport: (name, body) => onReport(body),
  onRequest: (method, url) => serverLog.push(`${method} ${url}`),
});
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const port = server.address().port;
const origin = `http://127.0.0.1:${port}/`;

// 上限の時間＝1 分＋1 ファイル 3 秒（big は 1 ファイル 30 秒）
const timeoutMs = 60000 + manifest.files.length * (set === 'big' ? 30000 : 3000);
let browser = null;
let page = null;
let report = null;
let shot = null;
try {
  browser = await launch(which);
  console.log(`${browser.version}  PID ${browser.pid}  CDP ${browser.port}  ${origin}`);
  page = await openPage(browser);
  const url = `${origin}__e2e/harness.html?set=${encodeURIComponent(set)}&sort=${sort}&tag=${which}`;
  await page.goto(url);
  let timer = null;
  report = await Promise.race([
    reportArrived,
    new Promise((r) => { timer = setTimeout(() => r(null), timeoutMs); }),
  ]);
  clearTimeout(timer);
  if (!report) problems.push(`報告が ${Math.round(timeoutMs / 1000)} 秒のうちに来ない`);
  try {
    shot = await page.screenshot();
  } catch { /* 撮れなくても試験は続ける */ }
} catch (e) {
  problems.push(`ブラウザの操りで失敗：${e.message}`);
} finally {
  if (page) page.close();
  await stop(browser);
  await new Promise((r) => server.close(r));
}

mkdirSync(OUT_E2E, { recursive: true });
const stem = `${which}_${set}_${sort}`;
if (shot) writeFileSync(path.join(OUT_E2E, `${stem}.png`), shot);

if (!verifyFixtures('後')) problems.push('原本の照合（後）でちがいが有る');

// ---- 報告を読む ----
if (report && report.fatal) problems.push(`ページの中で止まった：${report.fatal.split('\n')[0]}`);
if (report && !report.fatal) {
  for (const m of report.mismatches) problems.push(m);
  for (const v of report.violations) problems.push(`CSP 違反：${v.directive} ${v.blocked} ${v.source}:${v.line}`);
  const wantOk = manifest.files.filter((f) => f.expect === 'ok').length;
  const inBooks = report.books.reduce((a, b) => a + b.photos.length, 0);
  if (inBooks !== wantOk) problems.push(`台帳に入った写真が ${inBooks} 枚（${wantOk} 枚のはず）`);
}
if (page) {
  const outside = page.outsideRequests(origin);
  for (const q of outside) problems.push(`外への要求：${q.method} ${q.url}`);
  for (const ex of page.exceptions) problems.push(`ページの例外：${ex.split('\n')[0]}`);
  // ブラウザが自分で取りに行く /favicon.ico の 404 だけは数えない（ページの作りと関係ない）
  const errs = page.logEntries.filter((l) => l.level === 'error' && !(l.url && l.url === `${origin}favicon.ico`));
  for (const le of errs) problems.push(`ブラウザのエラー：${le.source} ${le.text} ${le.url || ''}`);
}

// ---- 検査の道具が読む expect.json（tools/build_ledger.mjs と同じ形）----
if (report && !report.fatal) {
  const books = report.books.map((b) => ({
    file: b.file,
    startNumber: b.startNumber,
    pages: pageCount(b.photos.length),
    photos: b.photos.map((p, j) => ({
      number: b.startNumber + j,
      name: p.name,
      shownName: fitFileName(cleanText(p.name)),
      dateText: formatDateJa(p.dateTimeOriginal),
      label: p.label,
      imageW: p.imageW,
      imageH: p.imageH,
    })),
  }));
  for (const b of books) for (const p of b.photos) if (p.label == null) problems.push(`${b.file} ${p.number}: label が無い`);
  writeFileSync(path.join(OUT_E2E, `${stem}.expect.json`), JSON.stringify({ set, sort, books }, null, 2) + '\n');

  const t = report.timing;
  const f = (ms) => `${(ms / 1000).toFixed(1)} s`;
  console.log(`${report.counts.files} ファイル → 写し ${report.counts.photos}・開けない ${report.counts.failed}`);
  for (const x of report.failed) console.log(`  開けない：${x.relativePath}  ${x.reason}（${x.detail}）`);
  for (const b of report.books) console.log(`  out/e2e/${b.file}  ${b.photos.length} 枚  ${b.pages} ページ  ${b.bytes} B`);
  console.log(`時間：全体 ${f(t.totalMs)}（fetch ${f(t.fetchMs)}・写し ${f(t.processMs)}・台帳 ${t.ledgerMs.map(f).join(' / ')}）`
    + `・1 枚の平均 ${Math.round(t.perPhotoAvgMs)} ms・最大 ${Math.round(t.perPhotoMaxMs)} ms（${t.perPhotoMaxPath}）`);
  if (report.memory) console.log(`JS ヒープ（終わり）：${Math.round(report.memory.usedJSHeapSize / 1048576)} MB`);
  console.log(`CSP 違反 ${report.violations.length} 件・外への要求 ${page ? page.outsideRequests(origin).length : '?'} 件（要求 ${page ? page.requests.length : '?'} 件）`);
  console.log(`out/e2e/${stem}.expect.json`);
}
writeFileSync(path.join(OUT_E2E, `${stem}.requests.log`), serverLog.join('\n') + '\n');

console.log(`かかった時間（起こす・止める・照合を含む）：${((Date.now() - t0) / 1000).toFixed(1)} s`);
if (problems.length) {
  for (const p of problems) console.log(`外れ：${p}`);
  console.log(`外れ ${problems.length} 件`);
  process.exit(1);
}
console.log('外れ 0');
