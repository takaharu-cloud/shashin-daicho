// 通しの試験：本物のページ（dist/shashin-daicho/）を、置き先と同じ CSP の下で headless のブラウザに開かせ、
// 人と同じ口（<input type=file>・ボタン）から操る。試験の口（--e2e）は使わない。
// 使い方: node scripts/run_page.mjs <edge|chrome> <組> --assets <置き先のサイトの assets のフォルダ> [--shots]
//   1. build_site → サーバー（--e2e なし）→ ブラウザ → /shashin-daicho/ を開く（開いた直後の CSP 違反・エラー・外への要求・ExcelJS 未読み）
//   2. #input-files に組の全部のファイル → ready を待つ → 一覧・まとめ・開けなかった写真・並び順・番号・小さい絵を manifest と比べる
//   3. 名前の順 ↔ 撮影日時の順 → 保存のボタンを全部押す（out/page/<ブラウザ>_<組>/ に落とす）→ expect.json → tools/inspect_xlsx.py
//   4. 場合分け：中止・#input-folder（webkitRelativePath）で入れ替わる・300 枚を超える・写真が 0・最初からやり直す
//   5. 終わりの CSP 違反・エラー・外への要求・原本の照合（make_fixtures.py --verify）
// --shots の時、幅 1280 と 390 の画面を out/page/shots/ に撮る。外れが有れば 1 行ずつ出して、終了コード 1。
import { spawnSync } from 'node:child_process';
import { existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';
import { buildSite } from './build_site.mjs';
import { createServer } from './serve_csp.mjs';
import { launch, openPage, stop } from './lib/cdp.mjs';
import { sortPhotos } from '../src/sort.js';
import { splitBooks, pageCount, fitFileName, cleanText } from '../src/layout.js';
import { formatDateJa, formatDateTimeShort } from '../src/exif.js';
import { targetSize } from '../src/pipeline.js';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const OUT_PAGE = path.join(ROOT, 'out', 'page');
const SHOTS = path.join(OUT_PAGE, 'shots');
const WIDE = { w: 1280, h: 900 };
const NARROW = { w: 390, h: 844 };
const MAX_PHOTOS = 300;

function usage(msg) {
  if (msg) console.error(msg);
  console.error('使い方: node scripts/run_page.mjs <edge|chrome> <組> --assets <置き先のサイトの assets のフォルダ> [--shots]');
  process.exit(2);
}

// ---- 引数 ----
const args = process.argv.slice(2);
let which = null;
let set = null;
let assets = null;
let shots = false;
for (let k = 0; k < args.length; k++) {
  if (args[k] === '--assets') assets = args[++k];
  else if (args[k] === '--shots') shots = true;
  else if (!which) which = args[k];
  else if (!set) set = args[k];
  else usage(`知らない引数：${args[k]}`);
}
if (which !== 'edge' && which !== 'chrome') usage('ブラウザは edge か chrome');
if (!set || !/^[A-Za-z0-9_-]+$/.test(set)) usage('組の名前が無いか、使えない字が有る');
if (!assets) usage('--assets が無い');
try {
  if (!statSync(assets).isDirectory()) throw new Error();
} catch {
  usage(`--assets のフォルダが無い：${assets}`);
}
const readManifest = (s) => JSON.parse(readFileSync(path.join(ROOT, 'fixtures', `${s}.manifest.json`), 'utf8'));
if (!existsSync(path.join(ROOT, 'fixtures', `${set}.manifest.json`))) usage(`manifest が無い：fixtures/${set}.manifest.json`);
const manifest = readManifest(set);
const fixturePath = (s, rel) => path.join(ROOT, 'fixtures', s, ...rel.split('/'));
const baseName = (rel) => rel.split('/').pop();

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
const sec = (ms) => `${(ms / 1000).toFixed(1)} s`;
const problems = [];
const notes = [];
const bad = (s) => problems.push(s);

function verifyFixtures(s, when) {
  const r = spawnSync('python', ['-X', 'utf8', path.join('tools', 'make_fixtures.py'), '--verify', s], { cwd: ROOT, encoding: 'utf8' });
  const lines = (r.stdout || '').trim().split(/\r?\n/);
  console.log(`原本の照合（${s}・${when}）：${lines[lines.length - 1]}`);
  if (r.status !== 0) bad(`原本の照合（${s}・${when}）でちがいが有る`);
}

// ---- 画面の言葉（発注文の表と同じ。ページの側と別に書いて、食いちがいを見つける）----
const T = {
  total: (n) => `台帳に入る写真：${n}枚（${pageCount(n)}ページ）`,
  noDate: (n) => `撮影日時の記録が無い写真：${n}枚（撮影年月日は空欄になります）`,
  failed: (n) => `開けなかった写真：${n}件`,
  others: (n) => `写真ではないファイル：${n}件（読みませんでした）`,
  othersHead: (n) => `写真ではないファイル（${n}件）`,
  aborted: (n) => (n > 0 ? `中止しました。読めた${n}枚で、このまま台帳を作れます。` : '中止しました。'),
  noPhoto: '開ける写真（JPEG・PNG）が見つかりませんでした。',
  tooMany: (n) => `写真が${n}枚あります。いちどに読めるのは300枚までです。フォルダを分けて選んでください。`,
  saved: (name) => `保存しました：${name}。ブラウザのダウンロードの場所を見てください。`,
  saveOne: 'Excelの台帳を保存する',
  saveBook: (k, s, e) => `台帳${k}を保存する（${s}〜${e}枚目）`,
  heic: 'HEIC形式は開けません。JPEGに変えてから選んでください。',
  unsupported: 'この形式は開けません。開けるのはJPEGとPNGです。',
  empty: '中身が空です。',
  broken: '壊れていて開けません。',
};

/** manifest の 1 行が、画面の「開けなかった写真」で何と出るはずか。写真ではないファイルは null。 */
function failedText(f) {
  if (f.expect === 'ok' || f.expect === 'not-image') return null;
  const ext = (f.path.split('.').pop() || '').toLowerCase();
  if (f.expect === 'unsupported') return ext === 'heic' || ext === 'heif' ? T.heic : T.unsupported;
  return f.bytes === 0 ? T.empty : T.broken;
}

// ---- 前の照合・組み立て・サーバー ----
const t0 = Date.now();
verifyFixtures(set, '前');
const { vendorOk } = buildSite({ quiet: true });
if (!vendorOk) bad('dist の vendor の SHA-256 が元とちがう');
const serverLog = [];
const server = createServer({ assets, onRequest: (method, url) => serverLog.push(`${method} ${url}`) });
await new Promise((r) => server.listen(0, '127.0.0.1', r));
const origin = `http://127.0.0.1:${server.address().port}/`;
const pageUrl = `${origin}shashin-daicho/`;

const runDir = path.join(OUT_PAGE, `${which}_${set}`);
rmSync(runDir, { recursive: true, force: true });
mkdirSync(runDir, { recursive: true });
if (shots) mkdirSync(SHOTS, { recursive: true });

let browser = null;
let page = null;
const timing = {};
try {
  browser = await launch(which);
  console.log(`${browser.version}  PID ${browser.pid}  ${pageUrl}`);
  page = await openPage(browser);
  // CSP の違反を、ページのどのスクリプトより先に数え始める（試験の側から足す。ページのファイルは変えない）
  await page.addInitScript(`window.__csp = []; document.addEventListener('securitypolicyviolation', (e) => {
    window.__csp.push({ directive: e.violatedDirective, blocked: e.blockedURI, source: e.sourceFile, line: e.lineNumber });
  }, true);`);
  await page.setViewport(WIDE.w, WIDE.h);
  await page.allowDownloads(runDir);

  const ev = (expr) => page.evaluate(expr);
  const state = () => ev(`document.getElementById('tool').dataset.state`);
  const waitFor = async (expr, ms, label) => {
    const until = Date.now() + ms;
    while (Date.now() < until) {
      if (await ev(expr)) return true;
      await sleep(100);
    }
    bad(`${label} が ${Math.round(ms / 1000)} 秒のうちに来ない`);
    return false;
  };
  const click = (sel) => ev(`document.querySelector(${JSON.stringify(sel)}).click()`);

  /** 画面を撮る（幅 1280 と 390）。sel が有ればそこまで送ってから。 */
  async function shot(name, sel, opts = {}) {
    if (!shots) return;
    for (const vp of [WIDE, NARROW]) {
      await page.setViewport(vp.w, vp.h, vp === NARROW);
      await sleep(250);
      if (sel) {
        // ページは scroll-behavior: smooth なので、instant で送る
        await ev(`(() => { const e = document.querySelector(${JSON.stringify(sel)}); if (e) { e.scrollIntoView({ block: 'start', behavior: 'instant' }); window.scrollBy({ top: -${opts.above ?? 64}, behavior: 'instant' }); } })()`);
      } else {
        await ev(`window.scrollTo({ top: 0, behavior: 'instant' })`);
      }
      await sleep(350);
      writeFileSync(path.join(SHOTS, `${which}_${set}_${name}_${vp.w}.png`), await page.screenshot());
    }
    await page.setViewport(WIDE.w, WIDE.h);
    await sleep(150);
  }

  /** 画面の一覧とまとめを読む。 */
  const readPage = () => ev(`(() => {
    const q = (s) => document.querySelector(s);
    const texts = (s) => Array.from(document.querySelectorAll(s)).map((e) => e.textContent);
    const rows = Array.from(document.querySelectorAll('#list tbody tr')).map((tr) => {
      const c = tr.querySelector('canvas');
      return {
        no: tr.querySelector('.tool-no').textContent,
        name: tr.querySelector('.tool-file .tool-name').textContent,
        path: (tr.querySelector('.tool-file .tool-path') || {}).textContent || '',
        date: tr.querySelector('.tool-date').textContent,
        drawn: c ? c.dataset.drawn || '' : '',
        label: c ? c.getAttribute('aria-label') : '',
      };
    });
    return {
      state: q('#tool').dataset.state,
      error: q('#error').hidden ? null : q('#error').textContent,
      rows,
      summary: texts('#summary li'),
      failed: q('#failed').hidden ? [] : texts('#failed li .tool-name'),
      othersHead: q('#others').hidden ? null : q('#others summary').textContent,
      others: q('#others').hidden ? [] : texts('#others li .tool-name'),
      reviewHidden: q('#review').hidden,
      saveHidden: q('#save').hidden,
      readingHidden: q('#reading').hidden,
      status: q('#status').textContent,
      saveButtons: texts('#save-buttons button'),
    };
  })()`);

  // ===== 1. 開いた直後 =====
  await page.goto(pageUrl);
  await sleep(500);
  const atOpen = await ev(`({ csp: window.__csp.length, excel: typeof window.ExcelJS, title: document.title,
    scripts: Array.from(document.scripts).map((s) => s.getAttribute('src')) })`);
  if (atOpen.csp) bad(`開いた直後に CSP 違反 ${atOpen.csp} 件`);
  if (atOpen.excel !== 'undefined' || page.requests.some((q) => q.url.includes('exceljs'))) bad('開いた直後に ExcelJS の束を読んでいる');
  const errsAtOpen = page.logEntries.filter((l) => l.level === 'error').length + page.console.filter((c) => c.type === 'error').length + page.exceptions.length;
  if (errsAtOpen) bad(`開いた直後にエラー ${errsAtOpen} 件`);
  const outAtOpen = page.requests.filter((q) => !q.url.startsWith(origin)).length;
  if (outAtOpen) bad(`開いた直後に外への要求 ${outAtOpen} 件`);
  console.log(`開いた直後：要求 ${page.requests.length} 件・外 ${outAtOpen}・CSP 違反 ${atOpen.csp}・エラー ${errsAtOpen}・ExcelJS ${atOpen.excel === 'undefined' ? 'まだ読んでいない' : '読んでいる'}`);
  await shot('01_head', null);
  await shot('02_tool', '#tool');

  // ===== 2. ファイルを渡して読む =====
  const okFiles = manifest.files.filter((f) => f.expect === 'ok');
  const files = manifest.files.map((f) => fixturePath(set, f.path));
  const readLimitMs = 60000 + manifest.files.length * 3000;
  const tRead = Date.now();
  await page.setFileInputFiles('#input-files', files);
  let shotReading = false;
  {
    const until = Date.now() + readLimitMs;
    let st = null;
    while (Date.now() < until) {
      st = await state();
      if (st === 'ready' || st === 'idle') break;
      if (shots && !shotReading && st === 'reading') {
        const v = await ev(`document.getElementById('progress').value`);
        if (v >= 1) {
          shotReading = true;
          await shot('03_reading', '#tool');
        }
      }
      await sleep(100);
    }
    timing.readMs = Date.now() - tRead;
    if (st !== 'ready') bad(`読み終わらない（${st}・${sec(readLimitMs)} まで待った）`);
  }
  if (shots && !shotReading) notes.push('読んでいる途中の絵は撮れなかった（読み終わるのが早かった）');
  console.log(`読み終わるまで：${sec(timing.readMs)}（${manifest.files.length} ファイル）`);

  // ===== 3. 画面を manifest と比べる =====
  const items = okFiles.map((f) => ({ name: baseName(f.path), relativePath: '', dateTimeOriginal: f.dateTimeOriginal ?? null, f }));
  const want = {
    date: sortPhotos(items, 'date'),
    name: sortPhotos(items, 'name'),
  };
  const dateCell = (d) => (d ? formatDateTimeShort(d) : '記録なし');
  function checkOrder(rows, mode) {
    const w = want[mode];
    if (rows.length !== w.length) {
      bad(`${mode} の順：一覧の行が ${rows.length}（${w.length} のはず）`);
      return;
    }
    let miss = 0;
    rows.forEach((r, i) => {
      if (r.no !== String(i + 1)) { miss += 1; if (miss <= 5) bad(`${mode} の順：${i + 1} 行目の番号が ${r.no}`); }
      if (r.name !== w[i].name || r.date !== dateCell(w[i].dateTimeOriginal)) {
        miss += 1;
        if (miss <= 5) bad(`${mode} の順：${i + 1} 行目が ${r.name} ${r.date}（${w[i].name} ${dateCell(w[i].dateTimeOriginal)} のはず）`);
      }
    });
    if (miss > 5) bad(`${mode} の順：ほかに ${miss - 5} 件`);
  }
  const sameList = (a, b) => JSON.stringify([...a].sort()) === JSON.stringify([...b].sort());

  // 小さい絵が全部描き終わるのを待つ
  await waitFor(`Array.from(document.querySelectorAll('#list canvas')).every((c) => c.dataset.drawn)`, 30000 + okFiles.length * 200, '小さい絵');
  let view = await readPage();
  checkOrder(view.rows, 'date');
  const noDate = okFiles.filter((f) => !f.dateTimeOriginal).length;
  const failedWant = manifest.files.map(failedText).filter(Boolean);
  const failedFiles = manifest.files.filter((f) => f.expect !== 'ok' && f.expect !== 'not-image');
  const othersFiles = manifest.files.filter((f) => f.expect === 'not-image');
  const sumWant = [T.total(okFiles.length)];
  if (noDate) sumWant.push(T.noDate(noDate));
  if (failedFiles.length) sumWant.push(T.failed(failedFiles.length));
  if (othersFiles.length) sumWant.push(T.others(othersFiles.length));
  if (JSON.stringify(view.summary) !== JSON.stringify(sumWant)) bad(`まとめが ${JSON.stringify(view.summary)}（${JSON.stringify(sumWant)} のはず）`);
  const failedLines = failedFiles.map((f) => `${baseName(f.path)} — ${failedText(f)}`);
  if (!sameList(view.failed, failedLines)) bad(`開けなかった写真が ${JSON.stringify(view.failed)}（${JSON.stringify(failedLines)} のはず）`);
  if (failedWant.length !== failedFiles.length) bad('試験の側の数えちがい（開けなかった写真）');
  if (othersFiles.length) {
    if (view.othersHead !== T.othersHead(othersFiles.length)) bad(`写真ではないファイルの見出しが ${view.othersHead}`);
    if (!sameList(view.others, othersFiles.map((f) => baseName(f.path)))) bad(`写真ではないファイルが ${JSON.stringify(view.others)}`);
  } else if (view.othersHead !== null) bad('写真ではないファイルが無いのに出ている');
  if (view.error) bad(`読み終わった後にエラーが出ている：${view.error}`);
  const notDrawn = view.rows.filter((r) => r.drawn !== '1').length;
  if (notDrawn) bad(`小さい絵が描けていない行 ${notDrawn}`);
  if (view.rows.some((r, i) => r.label !== `${r.name} の小さい絵`)) bad('小さい絵の aria-label が決まりとちがう');
  // 小さい絵が空でない（8×8 の点の色が 2 つ以上ある）
  const blank = await ev(`Array.from(document.querySelectorAll('#list canvas')).filter((c) => {
    const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data;
    const seen = new Set();
    for (let y = 0; y < 8; y++) for (let x = 0; x < 8; x++) {
      const i = ((Math.floor((y + .5) * c.height / 8) * c.width) + Math.floor((x + .5) * c.width / 8)) * 4;
      seen.add(d[i] + ',' + d[i + 1] + ',' + d[i + 2] + ',' + d[i + 3]);
    }
    return seen.size < 2;
  }).length`);
  if (blank) bad(`小さい絵が空（1 色だけ）の行 ${blank}`);
  console.log(`一覧 ${view.rows.length} 行・まとめ ${view.summary.join(' ／ ')}・開けない ${view.failed.length}・写真でない ${view.others.length}`);

  // 幅 390 で横にはみ出さない
  await page.setViewport(NARROW.w, NARROW.h, true);
  await sleep(300);
  const overflow = await ev(`({ doc: document.documentElement.scrollWidth, body: document.body.scrollWidth, w: window.innerWidth,
    wide: Array.from(document.querySelectorAll('main *')).filter((e) => {
      if (e.closest('.toc ul') || e.closest('#list-wrap')) return false;
      const r = e.getBoundingClientRect();
      return r.width > 0 && (r.right > window.innerWidth + 1 || r.left < -1);
    }).slice(0, 5).map((e) => e.tagName + (e.id ? '#' + e.id : '') + '.' + e.className) })`);
  if (overflow.doc > overflow.w || overflow.body > overflow.w || overflow.wide.length) {
    bad(`幅 390 で横にはみ出す（文書 ${overflow.doc}・body ${overflow.body}・画面 ${overflow.w}・${overflow.wide.join(' ')}）`);
  }
  await page.setViewport(WIDE.w, WIDE.h);
  await sleep(200);

  await shot('04_list', '#review');
  if (failedFiles.length || othersFiles.length) {
    await ev(`document.getElementById('others').open = true`);
    await shot('05_failed', '#failed:not([hidden]), #others:not([hidden])', { above: 200 });
    await ev(`document.getElementById('others').open = false`);
  } else if (shots) notes.push('開けなかった写真が無い組なので、⑤ は撮っていない');

  // ===== 4. 並び順を切り替える =====
  await click('#sort-name');
  view = await readPage();
  checkOrder(view.rows, 'name');
  await click('#sort-date');
  view = await readPage();
  checkOrder(view.rows, 'date');

  // ===== 5. 保存のボタンを全部押す =====
  const books = splitBooks(okFiles.length);
  const btnWant = books.length === 1 ? [T.saveOne] : books.map((b, k) => T.saveBook(k + 1, b.start + 1, b.end));
  if (JSON.stringify(view.saveButtons) !== JSON.stringify(btnWant)) bad(`保存のボタンが ${JSON.stringify(view.saveButtons)}（${JSON.stringify(btnWant)} のはず）`);
  const listXlsx = () => readdirSync(runDir).filter((n) => n.toLowerCase().endsWith('.xlsx'));
  const savedNames = [];
  timing.saveMs = [];
  for (let k = 0; k < books.length; k++) {
    const before = new Set(listXlsx());
    const tSave = Date.now();
    await click(`#save-buttons button[data-book="${k + 1}"]`);
    let got = null;
    const until = Date.now() + 120000;
    let lastSize = -1;
    while (Date.now() < until) {
      const fresh = listXlsx().filter((n) => !before.has(n));
      const partial = readdirSync(runDir).some((n) => n.endsWith('.crdownload') || n.endsWith('.tmp'));
      if (fresh.length === 1 && !partial) {
        const size = statSync(path.join(runDir, fresh[0])).size;
        if (size > 0 && size === lastSize) { got = fresh[0]; break; }
        lastSize = size;
      }
      await sleep(150);
    }
    const ms = Date.now() - tSave;
    if (!got) {
      bad(`台帳 ${k + 1} のファイルが落ちてこない`);
      continue;
    }
    timing.saveMs.push(ms);
    savedNames.push(got);
    await waitFor(`document.getElementById('tool').dataset.state === 'ready'`, 10000, `台帳 ${k + 1} の後の ready`);
    const saved = await ev(`document.getElementById('saved').textContent`);
    if (saved !== T.saved(got)) bad(`保存の後の言葉が「${saved}」（「${T.saved(got)}」のはず）`);
    console.log(`保存 ${k + 1}：${got}  ${statSync(path.join(runDir, got)).size} B  ${sec(ms)}`);
  }
  const shape = books.length === 1 ? /^工事写真台帳_(\d{8}_\d{4})\.xlsx$/ : /^工事写真台帳_(\d{8}_\d{4})_(\d+)\.xlsx$/;
  const stamps = new Set();
  savedNames.forEach((n, k) => {
    const m = shape.exec(n);
    if (!m) bad(`保存の名前の形がちがう：${n}`);
    else {
      stamps.add(m[1]);
      if (books.length > 1 && m[2] !== String(k + 1)) bad(`保存の名前の番号がちがう：${n}（_${k + 1} のはず）`);
    }
  });
  if (stamps.size > 1) bad(`同じ回の冊の日時がそろっていない：${[...stamps].join(' ')}`);
  await shot('06_saved', '#save', { above: 160 });

  // ===== 6. expect.json を書いて、本物の Excel で確かめる =====
  // 画面の行を manifest に突き合わせる：名前（拡張子込み）と撮影日時の両方が同じ物を、manifest の順に 1 つずつ使う
  // （同じ名前が 2 つ以上ある時も、使った物は外していくので取りちがえない。並び順は上で sortPhotos と同じ事を確かめた）。
  if (savedNames.length === books.length) {
    const pool = [...okFiles];
    const matched = view.rows.map((r) => {
      const k = pool.findIndex((f) => baseName(f.path) === r.name && dateCell(f.dateTimeOriginal ?? null) === r.date);
      if (k < 0) { bad(`一覧の ${r.no} 行目（${r.name}）が manifest に無い`); return null; }
      return pool.splice(k, 1)[0];
    });
    if (matched.every(Boolean)) {
      const exp = {
        set, sort: 'date',
        books: books.map((b, k) => ({
          file: savedNames[k],
          startNumber: b.start + 1,
          pages: pageCount(b.end - b.start),
          photos: matched.slice(b.start, b.end).map((f, j) => {
            const name = baseName(f.path);
            const size = targetSize(f.upright.w, f.upright.h);
            return {
              number: b.start + 1 + j,
              name,
              shownName: fitFileName(cleanText(name)),
              dateText: formatDateJa(f.dateTimeOriginal ?? null),
              label: f.label,
              imageW: size.w,
              imageH: size.h,
            };
          }),
        })),
      };
      const expPath = path.join(runDir, `${which}_${set}.expect.json`);
      writeFileSync(expPath, JSON.stringify(exp, null, 2) + '\n');
      const tIns = Date.now();
      const r = spawnSync('python', ['-X', 'utf8', path.join('tools', 'inspect_xlsx.py'), expPath], { cwd: ROOT, encoding: 'utf8' });
      for (const line of (r.stdout || '').trim().split(/\r?\n/)) if (line) console.log(`  inspect: ${line}`);
      if (r.status !== 0) bad(`inspect_xlsx.py が ${r.status} で終わった ${(r.stderr || '').trim().split(/\r?\n/).pop() || ''}`);
      console.log(`inspect_xlsx：${sec(Date.now() - tIns)}`);
    }
  }

  // ===== 7. 場合分け =====
  // (a) 中止：mix60 を読み始めて 2 枚目で中止 → 読めた分で ready
  {
    const m = readManifest('mix60');
    await page.setFileInputFiles('#input-files', m.files.map((f) => fixturePath('mix60', f.path)));
    await waitFor(`document.getElementById('progress').value >= 2 || document.getElementById('tool').dataset.state !== 'reading'`, 60000, '中止：読み始め');
    await click('#cancel');
    await waitFor(`document.getElementById('tool').dataset.state !== 'reading'`, 30000, '中止の後');
    const v = await readPage();
    const allOk = m.files.filter((f) => f.expect === 'ok').length;
    if (v.rows.length >= allOk) bad(`中止：全部読めてしまった（${v.rows.length} 行）`);
    if (v.state !== (v.rows.length ? 'ready' : 'idle')) bad(`中止：状態が ${v.state}`);
    if (v.status !== T.aborted(v.rows.length)) bad(`中止：言葉が「${v.status}」（「${T.aborted(v.rows.length)}」のはず）`);
    if (v.rows.length && !v.saveButtons.length) bad('中止：保存のボタンが無い');
    const shown = await page.evaluate(`['cancel','progress'].filter((id) => { const e = document.getElementById(id); return e && e.getClientRects().length > 0; })`);
    if (shown.length) bad(`中止：止まった後も見えている：${shown.join('・')}`);
    console.log(`場合分け 中止：${v.rows.length} 枚で ${v.state}「${v.status}」`);
    await shot('08_aborted', '#tool');
  }
  // (b) #input-folder に組のフォルダの道を渡す（webkitRelativePath が入るか）→ 前の一覧と入れ替わる
  {
    const fset = set === 'bulk300' ? 'smoke' : set; // bulk300 はもう 1 回読むと重いので smoke で試す
    const m = readManifest(fset);
    const wantRows = m.files.filter((f) => f.expect === 'ok').length;
    let ok = false;
    try {
      await page.setFileInputFiles('#input-folder', [path.join(ROOT, 'fixtures', fset)]);
      await sleep(300);
      const st = await state();
      if (st === 'reading' || st === 'ready') {
        await waitFor(`document.getElementById('tool').dataset.state !== 'reading'`, 60000 + m.files.length * 3000, 'フォルダの読み');
        const v = await readPage();
        const withPath = v.rows.filter((r) => r.path).length;
        ok = v.rows.length === wantRows && withPath === wantRows && v.rows.every((r) => r.path === fset || r.path.startsWith(`${fset}/`));
        const sample = v.rows.slice(0, 2).map((r) => `${r.path}/${r.name}`).join('・');
        notes.push(`#input-folder（${fset}）：${ok ? 'できた' : 'できなかった'}＝${v.rows.length} 行（${wantRows} のはず）・道が入った行 ${withPath}（例 ${sample}）`);
        if (!ok) bad(`#input-folder に道を渡した結果が合わない（${v.rows.length} 行・道 ${withPath}）`);
      } else {
        notes.push(`#input-folder（${fset}）：できなかった＝渡しても読み始めない（状態 ${st}）。ファイルの口だけで通した`);
      }
    } catch (e) {
      notes.push(`#input-folder（${fset}）：できなかった＝${e.message}。ファイルの口だけで通した`);
    }
    if (!ok) {
      // 回り道：ファイルの口で入れ替わるかを確かめる
      await page.setFileInputFiles('#input-files', m.files.map((f) => fixturePath(fset, f.path)));
      await waitFor(`document.getElementById('tool').dataset.state === 'ready'`, 60000, '入れ替え');
      const v = await readPage();
      if (v.rows.length !== wantRows) bad(`もう一度選ぶと入れ替わる：${v.rows.length} 行（${wantRows} のはず）`);
    }
    const v = await readPage();
    if (v.status && !v.readingHidden) bad(`入れ替えの後に前の言葉が残っている：${v.status}`);
    console.log(`場合分け 入れ替え：${v.rows.length} 行（${wantRows} のはず）`);
  }
  // (c) 300 枚を超える：bulk300 に 1 枚足す → 読まずに知らせる
  {
    const m = readManifest('bulk300');
    const extra = readManifest('smoke').files.find((f) => f.expect === 'ok');
    const list = [...m.files.map((f) => fixturePath('bulk300', f.path)), fixturePath('smoke', extra.path)];
    const n = list.filter((p) => /\.(jpg|jpeg|jpe|png)$/i.test(p)).length;
    await page.setFileInputFiles('#input-files', list);
    await waitFor(`!document.getElementById('error').hidden`, 10000, '300 枚を超えた時のエラー');
    const v = await readPage();
    if (v.error !== T.tooMany(n)) bad(`300 枚超：言葉が「${v.error}」（「${T.tooMany(n)}」のはず）`);
    if (v.state !== 'idle' || !v.readingHidden || v.rows.length || !v.reviewHidden || !v.saveHidden) {
      bad(`300 枚超：読まずに止まっていない（${v.state}・読む欄 ${v.readingHidden ? '隠れ' : '出ている'}・行 ${v.rows.length}）`);
    }
    if (n <= MAX_PHOTOS) bad('試験の側の数えちがい（300 枚超）');
    console.log(`場合分け 300 枚超（${n} 枚）：「${v.error}」`);
    await shot('09_too_many', '#tool');
  }
  // (d) 写真が 0（memo.txt だけ）
  {
    await page.setFileInputFiles('#input-files', [fixturePath('mix60', 'memo.txt')]);
    await waitFor(`!document.getElementById('error').hidden && document.getElementById('tool').dataset.state === 'idle'`, 10000, '写真が 0 の時のエラー');
    const v = await readPage();
    if (v.error !== T.noPhoto) bad(`写真が 0：言葉が「${v.error}」`);
    if (v.rows.length || !v.saveHidden) bad('写真が 0：一覧か保存が出ている');
    if (v.othersHead !== T.othersHead(1)) bad(`写真が 0：写真ではないファイルの見出しが ${v.othersHead}`);
    console.log(`場合分け 写真が 0：「${v.error}」・${v.othersHead}`);
    await shot('10_no_photo', '#tool');
  }
  // (e) 最初からやり直す
  {
    const m = readManifest('smoke');
    await page.setFileInputFiles('#input-files', m.files.map((f) => fixturePath('smoke', f.path)));
    await waitFor(`document.getElementById('tool').dataset.state === 'ready'`, 60000, 'やり直しの前の読み');
    await click('#reset');
    const v = await readPage();
    if (v.state !== 'idle' || v.rows.length || !v.reviewHidden || !v.saveHidden || !v.readingHidden || v.error) {
      bad(`最初からやり直す：残っている（${v.state}・行 ${v.rows.length}・エラー ${v.error}）`);
    }
    console.log(`場合分け やり直し：${v.state}・行 ${v.rows.length}`);
  }

  await shot('07a_limit', '#limit');
  await shot('07b_faq', '#faq');
  await shot('07c_foot', '.site-foot', { above: 300 });

  // ===== 8. 終わり =====
  await sleep(300);
  const csp = await ev('window.__csp');
  for (const v of csp) bad(`CSP 違反：${v.directive} ${v.blocked} ${v.source}:${v.line}`);
  timing.heapMB = await ev('performance.memory ? Math.round(performance.memory.usedJSHeapSize / 1048576) : null');
} catch (e) {
  bad(`ブラウザの操りで失敗：${e.message}`);
} finally {
  if (page) page.close();
  await stop(browser);
  await new Promise((r) => server.close(r));
}

if (page) {
  // 保存は同じ origin の blob: から落とす（外への要求ではない）
  const outside = page.requests.filter((q) => !q.url.startsWith(origin) && !q.url.startsWith(`blob:${origin.slice(0, -1)}`));
  for (const q of outside) bad(`外への要求：${q.method} ${q.url}`);
  for (const ex of page.exceptions) bad(`ページの例外：${ex.split('\n')[0]}`);
  for (const c of page.console.filter((x) => x.type === 'error' || x.type === 'warning')) bad(`コンソール（${c.type}）：${c.text}`);
  const errs = page.logEntries.filter((l) => l.level === 'error' && !(l.url && l.url === `${origin}favicon.ico`));
  for (const le of errs) bad(`ブラウザのエラー：${le.source} ${le.text} ${le.url || ''}`);
  console.log(`要求 ${page.requests.length} 件・外 ${outside.length} 件・ブラウザのエラー ${errs.length} 件`);
}
writeFileSync(path.join(runDir, 'requests.log'), serverLog.join('\n') + '\n');
verifyFixtures(set, '後');
if (set !== 'smoke') verifyFixtures('smoke', '後');
if (set !== 'mix60') verifyFixtures('mix60', '後');
if (set !== 'bulk300') verifyFixtures('bulk300', '後');

for (const n of notes) console.log(`メモ：${n}`);
if (timing.readMs != null) {
  console.log(`時間：読み終わるまで ${sec(timing.readMs)}・保存 ${(timing.saveMs || []).map(sec).join(' / ') || '-'}`
    + `${timing.heapMB != null ? `・JS ヒープ（終わり）${timing.heapMB} MB` : ''}`);
}
console.log(`かかった時間（起こす・照合・検査を含む）：${sec(Date.now() - t0)}`);
if (problems.length) {
  for (const p of problems) console.log(`外れ：${p}`);
  console.log(`外れ ${problems.length} 件`);
  process.exit(1);
}
console.log('外れ 0');
