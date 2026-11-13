// 工事写真台帳を作るページの動き。写真はこのブラウザの中だけで扱い、外へ送らない。
// 置き先の CSP の下で動かす＝インラインのスクリプトとスタイル・style の書き換え・eval・外への通信を使わない。
// 見た目の切り替えは、クラスと hidden だけで行う。
import { processAll } from './lib/pipeline.js';
import { sortPhotos } from './lib/sort.js';
import { splitBooks, pageCount } from './lib/layout.js';
import { buildLedger } from './lib/ledger.js';
import { formatDateTimeShort } from './lib/exif.js';

const MAX_PHOTOS = 300; // いちどに読める写真の数（名前の拡張子で数える）
const PHOTO_NAME = /\.(jpg|jpeg|jpe|png)$/i;
const EXCELJS_SRC = './vendor/exceljs.bare.min.js';
const XLSX_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
const THUMB = { w: 96, h: 72, scale: 2 }; // 小さい絵（画面の px）と、描く時の倍率
const THUMB_BATCH = 4; // 小さい絵を何枚描くごとに間を置くか
const REVOKE_MS = 30000;

const MSG = {
  noPhoto: '開ける写真（JPEG・PNG）が見つかりませんでした。',
  tooMany: (n) => `写真が${n}枚あります。いちどに読めるのは${MAX_PHOTOS}枚までです。フォルダを分けて選んでください。`,
  noLib: '台帳を組み立てる部品を読めませんでした。ページを読み込み直してください。',
  noBuild: '台帳を作れませんでした。ページを読み込み直して、もう一度試してください。',
  oldBrowser: 'このブラウザでは動きません。新しいEdgeかChromeで開いてください。',
  reading: (d, t) => `写真を読んでいます（${d} / ${t}）`,
  aborted: (n) => (n > 0 ? `中止しました。読めた${n}枚で、このまま台帳を作れます。` : '中止しました。'),
  total: (n) => `台帳に入る写真：${n}枚（${pageCount(n)}ページ）`,
  noDate: (n) => `撮影日時の記録が無い写真：${n}枚（撮影年月日は空欄になります）`,
  failed: (n) => `開けなかった写真：${n}件`,
  others: (n) => `写真ではないファイル：${n}件（読みませんでした）`,
  othersHead: (n) => `写真ではないファイル（${n}件）`,
  saveOne: 'Excelの台帳を保存する',
  saveBook: (k, s, e) => `台帳${k}を保存する（${s}〜${e}枚目）`,
  building: '台帳を組み立てています',
  saved: (name) => `保存しました：${name}。ブラウザのダウンロードの場所を見てください。`,
  noDateCell: '記録なし',
};

/** 開けなかった理由の言葉。 */
function reasonText(r) {
  if (r.reason === 'heic') return 'HEIC形式は開けません。JPEGに変えてから選んでください。';
  if (r.reason === 'unsupported') return 'この形式は開けません。開けるのはJPEGとPNGです。';
  if (r.detail === 'empty') return '中身が空です。';
  if (r.detail === 'truncated') return '途中で切れています。';
  return '壊れていて開けません。';
}

const $ = (id) => document.getElementById(id);
const el = {
  tool: $('tool'),
  error: $('error'),
  announce: $('announce'),
  drop: $('drop'),
  pickFolder: $('pick-folder'),
  pickFiles: $('pick-files'),
  inputFolder: $('input-folder'),
  inputFiles: $('input-files'),
  reading: $('reading'),
  status: $('status'),
  progress: $('progress'),
  cancel: $('cancel'),
  review: $('review'),
  summary: $('summary'),
  sort: $('sort'),
  sortCap: $('sort-cap'),
  sortLock: $('sort-lock'),
  sortDate: $('sort-date'),
  sortName: $('sort-name'),
  listWrap: $('list-wrap'),
  tbody: document.querySelector('#list tbody'),
  failed: $('failed'),
  failedList: document.querySelector('#failed ul'),
  others: $('others'),
  othersHead: document.querySelector('#others summary'),
  othersList: document.querySelector('#others ul'),
  save: $('save'),
  saveButtons: $('save-buttons'),
  saved: $('saved'),
  reset: $('reset'),
};

// ---- いまの状態 ----
let generation = 0; // 選び直すたびに 1 増える（古い読みの結果を捨てるため）
let run = null; // { ctrl, done }＝読んでいる途中の物
let photos = []; // 読めた写真（読んだ順）
let ordered = []; // いまの並び順
let rows = new Map(); // 写真 → 一覧の行
let thumbQueue = [];
let thumbBusy = false;
let stamp = null; // 保存の名前の日時（この回で最初に保存を押した時）
let savedOnce = false; // この回で 1 冊でも保存した
let excelPromise = null;

/** 2 冊以上に分かれる回で 1 冊でも保存した後は、並び順を変えると冊どうしで写真が重なり・抜けるので、変えさせない。 */
function sortLocked() {
  return savedOnce && splitBooks(ordered.length).length > 1;
}

function setState(s) {
  el.tool.dataset.state = s;
  const building = s === 'building';
  const locked = sortLocked();
  for (const b of el.saveButtons.querySelectorAll('button')) b.disabled = building;
  el.reset.disabled = building;
  el.sort.disabled = locked;
  el.sortLock.hidden = !locked;
  el.sortDate.disabled = building || locked;
  el.sortName.disabled = building || locked;
  el.pickFolder.disabled = building;
  el.pickFiles.disabled = building;
}

function showError(text) {
  el.error.textContent = text;
  el.error.hidden = false;
}

function clearError() {
  el.error.textContent = '';
  el.error.hidden = true;
}

/** ExcelJS の束を、要る時に 1 回だけ読む。読めなければ次に呼ばれた時に読み直す。 */
function loadExcel() {
  if (window.ExcelJS) return Promise.resolve(window.ExcelJS);
  if (!excelPromise) {
    excelPromise = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = EXCELJS_SRC;
      s.addEventListener('load', () => (window.ExcelJS ? resolve(window.ExcelJS) : reject(new Error('ExcelJS が無い'))));
      s.addEventListener('error', () => {
        s.remove();
        reject(new Error('ExcelJS を読めない'));
      });
      document.head.append(s);
    });
    excelPromise.catch(() => { excelPromise = null; });
  }
  return excelPromise;
}

// ---- 前の結果を捨てる ----
function clearResult() {
  photos = [];
  ordered = [];
  rows = new Map();
  thumbQueue = [];
  stamp = null;
  savedOnce = false;
  el.tbody.replaceChildren();
  el.summary.replaceChildren();
  el.failedList.replaceChildren();
  el.othersList.replaceChildren();
  el.othersHead.textContent = '';
  el.saveButtons.replaceChildren();
  el.saved.textContent = '';
  el.failed.hidden = true;
  el.others.hidden = true;
  el.others.open = false;
  el.review.hidden = true;
  el.save.hidden = true;
  el.reading.hidden = true;
  el.status.textContent = '';
  el.announce.textContent = '';
  el.sortDate.checked = true;
}

// ---- 読む ----
async function startWith(items, extraFailed = []) {
  const my = ++generation;
  if (run) {
    run.ctrl.abort();
    try { await run.done; } catch { /* 古い読みの結果は使わない */ }
  }
  if (my !== generation) return;
  clearError();
  clearResult();

  const count = items.filter((it) => PHOTO_NAME.test(it.file.name)).length;
  if (count > MAX_PHOTOS) {
    showError(MSG.tooMany(count));
    setState('idle');
    return;
  }
  if (items.length === 0 && extraFailed.length === 0) {
    showError(MSG.noPhoto);
    setState('idle');
    return;
  }

  loadExcel().catch(() => { /* 保存を押した時に知らせる */ });
  const ctrl = new AbortController();
  const total = items.length;
  el.progress.max = Math.max(1, total);
  el.progress.value = 0;
  el.progress.hidden = false;
  el.cancel.hidden = false;
  el.status.textContent = MSG.reading(0, total);
  el.reading.hidden = false;
  setState('reading');

  const done = processAll(items, {
    signal: ctrl.signal,
    onProgress: (d, t) => {
      if (my !== generation) return;
      el.progress.value = d;
      el.status.textContent = MSG.reading(d, t);
    },
  });
  run = { ctrl, done };
  let res;
  try {
    res = await done;
  } finally {
    if (run && run.done === done) run = null;
  }
  if (my !== generation) return;
  showResult({ ...res, failed: [...res.failed, ...extraFailed] });
}

function showResult(res) {
  const failed = res.failed.filter((r) => r.reason !== 'not-image');
  const others = res.failed.filter((r) => r.reason === 'not-image');
  photos = res.photos;

  if (res.aborted) {
    el.status.textContent = MSG.aborted(photos.length);
    el.progress.hidden = true;
    el.cancel.hidden = true;
  } else {
    el.reading.hidden = true;
  }

  // まとめ
  const lines = [];
  if (photos.length > 0) lines.push(MSG.total(photos.length));
  const noDate = photos.filter((p) => !p.dateTimeOriginal).length;
  if (noDate > 0) lines.push(MSG.noDate(noDate));
  if (failed.length > 0) lines.push(MSG.failed(failed.length));
  if (others.length > 0) lines.push(MSG.others(others.length));
  el.summary.replaceChildren(...lines.map((t) => {
    const li = document.createElement('li');
    li.textContent = t;
    return li;
  }));

  // 開けなかった写真・写真ではないファイル
  el.failedList.replaceChildren(...failed.map((r) => nameItem(`${r.name} — ${reasonText(r)}`, r.relativePath)));
  el.failed.hidden = failed.length === 0;
  el.othersList.replaceChildren(...others.map((r) => nameItem(r.name, r.relativePath)));
  el.othersHead.textContent = MSG.othersHead(others.length);
  el.others.hidden = others.length === 0;

  const hasList = photos.length > 0;
  el.sort.hidden = !hasList;
  el.sortCap.hidden = !hasList;
  el.listWrap.hidden = !hasList;
  el.review.hidden = !hasList && failed.length === 0 && others.length === 0;

  if (!hasList) {
    if (!res.aborted) showError(MSG.noPhoto);
    setState('idle');
    return;
  }

  for (const p of photos) rows.set(p, makeRow(p));
  applySort(el.sortName.checked ? 'name' : 'date');
  makeSaveButtons();
  el.save.hidden = false;
  setState('ready');
  // 読み終わった事を、まとめの 1 行目と同じ文で読み上げに知らせる（フォーカスは動かさない）。中止の時は #status の言葉が知らせ。
  if (!res.aborted) el.announce.textContent = lines[0];
  thumbQueue = [...ordered];
  drawThumbs();
}

/** フォルダの道（ファイル名の前の部分）。無ければ ''。 */
function folderOf(relativePath) {
  if (typeof relativePath !== 'string') return '';
  const k = relativePath.lastIndexOf('/');
  return k > 0 ? relativePath.slice(0, k) : '';
}

function nameItem(text, relativePath) {
  const li = document.createElement('li');
  const name = document.createElement('span');
  name.className = 'tool-name';
  name.textContent = text;
  li.append(name);
  const folder = folderOf(relativePath);
  if (folder) {
    const sub = document.createElement('span');
    sub.className = 'tool-path';
    sub.textContent = folder;
    li.append(sub);
  }
  return li;
}

function makeRow(p) {
  const tr = document.createElement('tr');
  const no = document.createElement('td');
  no.className = 'tool-no';
  const pic = document.createElement('td');
  pic.className = 'tool-pic';
  const canvas = document.createElement('canvas');
  canvas.className = 'tool-thumb';
  canvas.width = THUMB.w * THUMB.scale;
  canvas.height = THUMB.h * THUMB.scale;
  canvas.setAttribute('role', 'img');
  canvas.setAttribute('aria-label', `${p.name} の小さい絵`);
  pic.append(canvas);
  const name = document.createElement('td');
  name.className = 'tool-file';
  const n = document.createElement('span');
  n.className = 'tool-name';
  n.textContent = p.name;
  name.append(n);
  const folder = folderOf(p.relativePath);
  if (folder) {
    const sub = document.createElement('span');
    sub.className = 'tool-path';
    sub.textContent = folder;
    name.append(sub);
  }
  const date = document.createElement('td');
  date.className = 'tool-date';
  date.textContent = p.dateTimeOriginal ? formatDateTimeShort(p.dateTimeOriginal) : MSG.noDateCell;
  if (!p.dateTimeOriginal) date.classList.add('tool-date-none');
  tr.append(no, pic, name, date);
  return tr;
}

/** 並べ替える（行は作り直さない）・番号を 1 から振り直す。 */
function applySort(mode) {
  ordered = sortPhotos(photos, mode);
  const trs = ordered.map((p) => rows.get(p));
  trs.forEach((tr, i) => { tr.firstChild.textContent = String(i + 1); });
  el.tbody.append(...trs);
}

// ---- 小さい絵（写しの JPEG から。数枚ごとに間を置いて、画面を止めない）----
async function drawThumbs() {
  if (thumbBusy) return;
  thumbBusy = true;
  const my = generation;
  try {
    let n = 0;
    while (thumbQueue.length > 0 && my === generation) {
      const p = thumbQueue.shift();
      const tr = rows.get(p);
      if (tr) await drawOne(p, tr.querySelector('canvas'));
      n += 1;
      if (n % THUMB_BATCH === 0) await new Promise((r) => setTimeout(r, 0));
    }
  } finally {
    thumbBusy = false;
    if (thumbQueue.length > 0 && my !== generation) drawThumbs();
  }
}

async function drawOne(p, canvas) {
  const cw = canvas.width;
  const ch = canvas.height;
  const s = Math.min(cw / p.image.width, ch / p.image.height);
  const w = Math.max(1, Math.round(p.image.width * s));
  const h = Math.max(1, Math.round(p.image.height * s));
  let bmp = null;
  try {
    bmp = await createImageBitmap(new Blob([p.image.bytes], { type: 'image/jpeg' }), {
      resizeWidth: w, resizeHeight: h, resizeQuality: 'medium',
    });
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#F2EEE4';
    ctx.fillRect(0, 0, cw, ch);
    ctx.drawImage(bmp, Math.round((cw - w) / 2), Math.round((ch - h) / 2), w, h);
    canvas.dataset.drawn = '1';
  } catch {
    // 写しは読めているので、小さい絵だけ描けなかった。一覧はそのまま使える
    canvas.dataset.drawn = 'no';
  } finally {
    if (bmp) bmp.close();
  }
}

// ---- 保存 ----
function makeSaveButtons() {
  const books = splitBooks(ordered.length);
  el.saveButtons.replaceChildren(...books.map(({ start, end }, k) => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'btn tool-btn-main save-book';
    b.dataset.book = String(k + 1);
    b.textContent = books.length === 1 ? MSG.saveOne : MSG.saveBook(k + 1, start + 1, end);
    return b;
  }));
}

const z2 = (n) => String(n).padStart(2, '0');
function makeStamp(d) {
  return `${d.getFullYear()}${z2(d.getMonth() + 1)}${z2(d.getDate())}_${z2(d.getHours())}${z2(d.getMinutes())}`;
}

async function save(k) {
  if (el.tool.dataset.state !== 'ready') return;
  clearError();
  setState('building');
  el.saved.textContent = MSG.building;
  try {
    let ExcelJS;
    try {
      ExcelJS = await loadExcel();
    } catch {
      el.saved.textContent = '';
      showError(MSG.noLib);
      return;
    }
    const books = splitBooks(ordered.length);
    const book = books[k];
    if (!book) throw new Error('冊が無い');
    if (!stamp) stamp = makeStamp(new Date());
    const fileName = books.length === 1 ? `工事写真台帳_${stamp}.xlsx` : `工事写真台帳_${stamp}_${k + 1}.xlsx`;
    const r = await buildLedger(ExcelJS, ordered.slice(book.start, book.end), { startNumber: book.start + 1 });
    if (!r.ok) {
      el.saved.textContent = '';
      showError(MSG.noBuild);
      return;
    }
    const url = URL.createObjectURL(new Blob([r.bytes], { type: XLSX_TYPE }));
    const a = document.createElement('a');
    a.href = url;
    a.download = fileName;
    a.hidden = true;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), REVOKE_MS);
    const wasLocked = sortLocked();
    savedOnce = true;
    el.saved.textContent = MSG.saved(fileName);
    // 並び順が押せなくなった、その時に 1 回だけ、#sort-lock と同じ文を読み上げにも知らせる
    if (!wasLocked && sortLocked()) el.announce.textContent = el.sortLock.textContent;
  } catch {
    el.saved.textContent = '';
    showError(MSG.noBuild);
  } finally {
    if (el.tool.dataset.state === 'building') setState('ready');
  }
}

// ---- 入口 ----
function fromInput(input) {
  const list = Array.from(input.files || []).map((file) => ({ file, relativePath: file.webkitRelativePath || '' }));
  input.value = ''; // 同じフォルダをもう一度選んでも読めるように
  return list;
}

/** 落とされた物を、drop の場で取り出す（entries が取れなければファイルの一覧）。 */
function takeDropped(dt) {
  const entries = Array.from(dt.items || [])
    .filter((it) => it.kind === 'file' && typeof it.webkitGetAsEntry === 'function')
    .map((it) => it.webkitGetAsEntry())
    .filter(Boolean);
  return { entries, files: entries.length === 0 ? Array.from(dt.files || []) : [] };
}

/** 落とされた物をたどる（フォルダは readEntries を空が返るまで繰り返す）。 */
async function fromDrop({ entries, files }) {
  if (entries.length === 0) return { items: files.map((file) => ({ file, relativePath: '' })), failed: [] };
  const items = [];
  const failed = [];
  const walk = async (entry) => {
    const rel = entry.fullPath.replace(/^\/+/, '');
    if (entry.isFile) {
      try {
        const file = await new Promise((resolve, reject) => entry.file(resolve, reject));
        items.push({ file, relativePath: rel });
      } catch {
        failed.push({ ok: false, name: entry.name, relativePath: rel, reason: 'broken', detail: 'read' });
      }
    } else if (entry.isDirectory) {
      const reader = entry.createReader();
      for (;;) {
        const batch = await new Promise((resolve, reject) => reader.readEntries(resolve, reject));
        if (batch.length === 0) break;
        for (const e of batch) await walk(e);
      }
    }
  };
  for (const e of entries) await walk(e);
  return { items, failed };
}

function supported() {
  return typeof createImageBitmap === 'function'
    && typeof AbortController === 'function'
    && typeof Blob === 'function' && typeof Blob.prototype.arrayBuffer === 'function'
    && (typeof OffscreenCanvas === 'function' || typeof HTMLCanvasElement.prototype.toBlob === 'function')
    && typeof URL.createObjectURL === 'function'
    && 'download' in HTMLAnchorElement.prototype
    && typeof Intl === 'object' && typeof Intl.Collator === 'function';
}

/** 思わぬ例外を受けた後、画面を落ち着いた状態に戻す。読めた写真が 1 枚でも手元に有れば ready、無ければ idle。 */
function settle() {
  const s = el.tool.dataset.state;
  if (s !== 'reading' && s !== 'building') return; // もう落ち着いている
  if (s === 'reading' && run) return; // 別の読みが進んでいる（古い読みの例外で止めない）
  el.reading.hidden = true;
  setState(photos.length > 0 ? 'ready' : 'idle');
}

/** 思わぬ例外は「台帳が作れない」に落とす（黙らない）。知らせた上で、「読んでいます」や「組み立てています」のまま残さない。 */
function guard(fn) {
  return (...args) => {
    Promise.resolve().then(() => fn(...args)).catch(() => {
      showError(MSG.noBuild);
      settle();
    });
  };
}

function init() {
  if (!supported()) {
    showError(MSG.oldBrowser);
    el.pickFolder.disabled = true;
    el.pickFiles.disabled = true;
    return;
  }
  el.pickFolder.addEventListener('click', () => el.inputFolder.click());
  el.pickFiles.addEventListener('click', () => el.inputFiles.click());
  el.inputFolder.addEventListener('change', guard(() => startWith(fromInput(el.inputFolder))));
  el.inputFiles.addEventListener('change', guard(() => startWith(fromInput(el.inputFiles))));

  el.drop.addEventListener('dragover', (e) => {
    e.preventDefault();
    if (e.dataTransfer) e.dataTransfer.dropEffect = 'copy';
    el.drop.classList.add('tool-drop-over');
  });
  el.drop.addEventListener('dragleave', (e) => {
    if (!el.drop.contains(e.relatedTarget)) el.drop.classList.remove('tool-drop-over');
  });
  el.drop.addEventListener('drop', (e) => {
    // 落とされた物は、この場（同期）でしか取り出せない
    e.preventDefault();
    el.drop.classList.remove('tool-drop-over');
    if (el.tool.dataset.state === 'building' || !e.dataTransfer) return;
    const dropped = takeDropped(e.dataTransfer);
    guard(async () => {
      const { items, failed } = await fromDrop(dropped);
      await startWith(items, failed);
    })();
  });
  // 道具の外に落としても、ブラウザが写真を開いてページを離れないようにする
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => e.preventDefault());

  el.cancel.addEventListener('click', () => { if (run) run.ctrl.abort(); });
  for (const r of [el.sortDate, el.sortName]) {
    r.addEventListener('change', () => {
      if (!r.checked || photos.length === 0 || sortLocked()) return;
      applySort(r.value);
      // 前の並び順で保存した知らせが、今の一覧の物に見えないように消す（ここに来るのは 1 冊の時だけ）
      el.saved.textContent = '';
    });
  }
  el.saveButtons.addEventListener('click', (e) => {
    const b = e.target.closest('button.save-book');
    if (b) guard(save)(Number(b.dataset.book) - 1);
  });
  el.reset.addEventListener('click', () => {
    generation += 1;
    if (run) run.ctrl.abort();
    clearError();
    clearResult();
    setState('idle');
  });
}

init();
