// 写真を 1 枚ずつ読んで、台帳に貼る写し（JPEG・長辺 1200）を作る流れ。
// 原本は読むだけ。classify と targetSize は DOM に触らない。processPhoto はブラウザの中だけで動く。
// 置き先の CSP の下で動かす＝<img> に blob:／data: を入れない・eval を使わない・外へ通信しない。
import { readPhotoInfo } from './exif.js';

const HEAD_BYTES = 131072; // 頭の 128KB だけ読んで種類・撮影日時・向きを見る
const PHOTO_EXT = new Set(['jpg', 'jpeg', 'jpe', 'png', 'heic', 'heif']);
// 写真の拡張子ではない、ほかの絵の拡張子（この道具では扱わない）
const OTHER_IMAGE_EXT = new Set([
  'gif', 'bmp', 'webp', 'tif', 'tiff', 'avif', 'jxl', 'jp2', 'ico', 'svg',
  'dng', 'cr2', 'cr3', 'nef', 'arw', 'raf', 'orf', 'rw2',
]);

function extOf(name) {
  if (typeof name !== 'string') return '';
  const dot = name.lastIndexOf('.');
  return dot >= 0 ? name.slice(dot + 1).toLowerCase() : '';
}

/**
 * 種類の見分け（DOM に触らない）。info は readPhotoInfo の返り値。
 * @returns {'photo'|'heic'|'broken'|'unsupported'|'not-image'}
 */
export function classify(name, size, info) {
  const kind = info && typeof info === 'object' ? info.kind : 'other';
  if (kind === 'jpeg' || kind === 'png') return 'photo';
  if (kind === 'heic') return 'heic';
  const ext = extOf(name);
  if (PHOTO_EXT.has(ext)) return 'broken'; // 0 バイト・頭が写真でない
  if (OTHER_IMAGE_EXT.has(ext)) return 'unsupported';
  return 'not-image';
}

const isPos = (v) => typeof v === 'number' && Number.isFinite(v) && v > 0;

/** 写しの大きさ（DOM に触らない）。長辺が maxLongSide を超える時だけ縮める。拡大しない。→ { w, h }・不正は null */
export function targetSize(w, h, maxLongSide = 1200) {
  if (!isPos(w) || !isPos(h) || !isPos(maxLongSide)) return null;
  const long = Math.max(w, h);
  if (long <= maxLongSide) return { w: Math.max(1, Math.round(w)), h: Math.max(1, Math.round(h)) };
  const s = maxLongSide / long;
  const side = (v) => (v === long ? Math.round(maxLongSide) : Math.max(1, Math.round(v * s)));
  return { w: side(w), h: side(h) };
}

const fail = (name, relativePath, reason, detail) => ({ ok: false, name, relativePath, reason, detail });

function makeCanvas(w, h) {
  if (typeof OffscreenCanvas === 'function') return new OffscreenCanvas(w, h);
  const c = document.createElement('canvas');
  c.width = w;
  c.height = h;
  return c;
}

async function canvasToJpeg(canvas, quality) {
  let blob;
  if (typeof canvas.convertToBlob === 'function') {
    blob = await canvas.convertToBlob({ type: 'image/jpeg', quality });
  } else {
    blob = await new Promise((resolve) => canvas.toBlob(resolve, 'image/jpeg', quality));
  }
  if (!blob) return null;
  return new Uint8Array(await blob.arrayBuffer());
}

/**
 * 1 枚を処理する（ブラウザの中だけ）。原本は読むだけ。例外を外へ投げない。
 * @param {File} file
 * @param {{ maxLongSide?: number, quality?: number, relativePath?: string }} [options] 既定 1200・0.8
 */
export async function processPhoto(file, options = {}) {
  const opt = options && typeof options === 'object' ? options : {};
  const name = file && typeof file.name === 'string' ? file.name : '';
  const relativePath = typeof opt.relativePath === 'string'
    ? opt.relativePath
    : (file && typeof file.webkitRelativePath === 'string' ? file.webkitRelativePath : '');
  const maxLongSide = isPos(opt.maxLongSide) ? opt.maxLongSide : 1200;
  const quality = typeof opt.quality === 'number' && opt.quality > 0 && opt.quality <= 1 ? opt.quality : 0.8;

  if (!file || typeof file.slice !== 'function') return fail(name, relativePath, 'not-image', 'no-file');

  let info;
  try {
    const size = file.size;
    const head = size === 0 ? new Uint8Array(0) : new Uint8Array(await file.slice(0, HEAD_BYTES).arrayBuffer());
    info = readPhotoInfo(head);
    const c = classify(name, size, info);
    if (c === 'heic') return fail(name, relativePath, 'heic', 'heic');
    if (c === 'unsupported') return fail(name, relativePath, 'unsupported', extOf(name));
    if (c === 'not-image') return fail(name, relativePath, 'not-image', extOf(name) || 'no-extension');
    if (c === 'broken') return fail(name, relativePath, 'broken', size === 0 ? 'empty' : 'header');
  } catch {
    return fail(name, relativePath, 'broken', 'read');
  }

  let bmp = null;
  let canvas = null;
  try {
    try {
      bmp = await createImageBitmap(file, { imageOrientation: 'from-image' });
    } catch {
      return fail(name, relativePath, 'broken', 'decode');
    }
    const size = targetSize(bmp.width, bmp.height, maxLongSide);
    if (!size) return fail(name, relativePath, 'broken', 'decode');
    let ctx = null;
    try {
      canvas = makeCanvas(size.w, size.h);
      ctx = canvas.getContext('2d');
    } catch {
      ctx = null;
    }
    if (!ctx) return fail(name, relativePath, 'broken', 'decode');
    ctx.fillStyle = '#ffffff'; // 透明な所（PNG）を白にしてから描く
    ctx.fillRect(0, 0, size.w, size.h);
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(bmp, 0, 0, size.w, size.h);
    bmp.close();
    bmp = null;
    const bytes = await canvasToJpeg(canvas, quality);
    if (!bytes || bytes.byteLength === 0) return fail(name, relativePath, 'broken', 'encode');
    return {
      ok: true,
      name,
      relativePath,
      dateTimeOriginal: info.dateTimeOriginal,
      orientation: info.orientation,
      image: { bytes, width: size.w, height: size.h, type: 'jpeg' },
    };
  } catch {
    return fail(name, relativePath, 'broken', 'encode');
  } finally {
    if (bmp) try { bmp.close(); } catch { /* 放すだけ */ }
    if (canvas) { canvas.width = 0; canvas.height = 0; }
  }
}

/**
 * 全部を 1 枚ずつ順に処理する（並列にしない）。止められたら、いまの 1 枚を終えてから止まる。
 * @param {Array<File | { file: File, relativePath: string }>} files
 * @param {{ onProgress?: (done: number, total: number, last: object) => void, signal?: AbortSignal,
 *           maxLongSide?: number, quality?: number }} [options]
 * @returns {Promise<{ photos: object[], failed: object[], aborted: boolean }>}
 */
export function processAll(files, options) {
  return processAllWith(processPhoto, files, options);
}

/** processAll の中身。1 枚の処理（processOne）を差し替えて試すための口（ページからは processAll を使う）。 */
export async function processAllWith(processOne, files, options = {}) {
  const opt = options && typeof options === 'object' ? options : {};
  const photos = [];
  const failed = [];
  if (!Array.isArray(files)) return { photos, failed, aborted: false };
  const total = files.length;
  for (let i = 0; i < total; i++) {
    if (opt.signal && opt.signal.aborted) return { photos, failed, aborted: true };
    const item = files[i];
    const wrapped = item && typeof item === 'object' && 'file' in item && !(typeof item.slice === 'function');
    const file = wrapped ? item.file : item;
    const relativePath = wrapped && typeof item.relativePath === 'string'
      ? item.relativePath
      : (file && typeof file.webkitRelativePath === 'string' ? file.webkitRelativePath : '');
    const name = file && typeof file.name === 'string' ? file.name : '';
    let r;
    if (!file || typeof file !== 'object') {
      r = fail(name, relativePath, 'not-image', 'no-file');
    } else {
      try {
        r = await processOne(file, { maxLongSide: opt.maxLongSide, quality: opt.quality, relativePath });
      } catch {
        r = null;
      }
      if (!r || typeof r !== 'object') r = fail(name, relativePath, 'broken', 'internal');
    }
    (r.ok ? photos : failed).push(r);
    if (typeof opt.onProgress === 'function') {
      try { opt.onProgress(i + 1, total, r); } catch { /* 画面の側の失敗で止めない */ }
    }
  }
  return { photos, failed, aborted: false };
}
