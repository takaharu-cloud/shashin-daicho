// ExcelJS で写真台帳（A4 縦・1 ページ 3 段）を組む。ブラウザでも Node でも同じコードで動かす
// （Buffer・fs・process を使わない）。枠とページ割りの数字は layout.js にある。
import {
  SHEET, fitImage, blockRows, blockLayout, imageAnchor, hasPageBreakAfter, pageCount, fitFileName, cleanText, textWidth,
} from './layout.js';
import { formatDateJa } from './exif.js';

const APP = 'shashin-daicho';
const TYPES = new Set(['jpeg', 'png']);
const thin = { style: 'thin', color: { argb: 'FF000000' } };
const BOX = { top: thin, left: thin, bottom: thin, right: thin };

// ---- 標準の字体（テーマ）----
// 列の幅は「標準の字体」の 1 字の幅で決まる。ExcelJS の既定のままだと日本語の Excel では
// ＭＳ Ｐゴシック 扱いになるので、テーマの minor／major の latin と Jpan を 游ゴシック に固定する。
// ExcelJS に公開の API が無く、`_themes` は ExcelJS の内部の名前（4.4.0 で確かめた）。
// 内部の名前に触るのはこの関数だけにする。テーマの元は束ごとに 1 回だけ作って使い回す。
const themeCache = new WeakMap();
async function applyThemeFont(ExcelJS, wb) {
  let p = themeCache.get(ExcelJS);
  if (!p) {
    p = (async () => {
      const t = new ExcelJS.Workbook();
      t.addWorksheet('a');
      const r = new ExcelJS.Workbook();
      await r.xlsx.load(await t.xlsx.writeBuffer());
      const face = SHEET.font.name;
      const xml = r._themes.theme1
        .replace(/(<a:(?:minor|major)Font>\s*<a:latin typeface=")[^"]*"/g, `$1${face}"`)
        .replace(/(<a:font script="Jpan" typeface=")[^"]*"/g, `$1${face}"`);
      if (!xml.includes(`typeface="${face}"`)) throw new Error('テーマの字体を置き換えられない');
      return xml;
    })();
    themeCache.set(ExcelJS, p);
    p.catch(() => themeCache.delete(ExcelJS)); // 失敗した物は使い回さない
  }
  wb._themes = { theme1: await p };
}

// ---- 頭の題 ----
// cleanText した後、印刷の幅（SHEET.titleMaxWidth）か titleMaxChars 字（… を含む）を超えたら切って … を付ける。
// 空になったら既定の題。& → && はこの後で行う（伸びても 255 字を超えない上限は SHEET の側で決めてある）。
function headerTitle(t) {
  const chars = Array.from(cleanText(t));
  if (chars.length === 0) return SHEET.defaultTitle;
  if (chars.length <= SHEET.titleMaxChars && textWidth(chars.join('')) <= SHEET.titleMaxWidth) return chars.join('');
  const ell = textWidth('…');
  let w = 0;
  let n = 0;
  while (n < chars.length && n < SHEET.titleMaxChars - 1 && w + textWidth(chars[n]) + ell <= SHEET.titleMaxWidth) {
    w += textWidth(chars[n++]);
  }
  return chars.slice(0, n).join('') + '…';
}

// ---- 入力の確かめ ----
const isPos = (v) => typeof v === 'number' && Number.isFinite(v) && v > 0;
function toBytes(b) {
  if (b instanceof Uint8Array) return b;
  if (b instanceof ArrayBuffer) return new Uint8Array(b);
  return null;
}
function checkEntry(e) {
  if (!e || typeof e !== 'object' || typeof e.name !== 'string') return null;
  const img = e.image;
  if (!img || typeof img !== 'object') return null;
  const bytes = toBytes(img.bytes);
  if (!bytes || bytes.byteLength === 0) return null;
  if (!isPos(img.width) || !isPos(img.height) || !TYPES.has(img.type)) return null;
  return { bytes, width: img.width, height: img.height, type: img.type };
}

/**
 * @param {object} ExcelJS 束の本体（ブラウザ＝グローバルの ExcelJS／Node＝tools/load_exceljs.mjs）
 * @param {Array<{ name: string, dateTimeOriginal: string|null,
 *                 image: { bytes: Uint8Array|ArrayBuffer, width: number, height: number, type: 'jpeg'|'png' } }>} entries
 *        並べ終えた順・1 ブック分
 * @param {{ startNumber?: number, title?: string, now?: Date, fields?: string[] }} [options]
 *        fields＝右の欄のキー（無ければ全部。知らないキーは捨てる）
 * @returns {Promise<{ ok: true, bytes: Uint8Array, pages: number }
 *                 | { ok: false, reason: 'empty'|'too-many'|'bad-entry'|'write-failed', index?: number, message?: string }>}
 * 例外を外へ投げない。
 */
export async function buildLedger(ExcelJS, entries, options = {}) {
  if (!Array.isArray(entries)) return { ok: false, reason: 'bad-entry', message: 'entries が配列でない' };
  if (entries.length === 0) return { ok: false, reason: 'empty' };
  if (entries.length > SHEET.maxPerBook) return { ok: false, reason: 'too-many' };
  const images = [];
  for (let i = 0; i < entries.length; i++) {
    const img = checkEntry(entries[i]);
    if (!img) return { ok: false, reason: 'bad-entry', index: i };
    images.push(img);
  }

  const opt = options && typeof options === 'object' ? options : {};
  const startNumber = Number.isInteger(opt.startNumber) && opt.startNumber >= 1 ? opt.startNumber : 1;
  const title = headerTitle(opt.title);
  const now = opt.now instanceof Date && !Number.isNaN(opt.now.getTime()) ? opt.now : new Date();

  try {
    const wb = new ExcelJS.Workbook();
    await applyThemeFont(ExcelJS, wb);
    wb.creator = APP;
    wb.lastModifiedBy = APP;
    wb.created = now;
    wb.modified = now;

    const n = entries.length;
    const ws = wb.addWorksheet(SHEET.sheetName, {
      pageSetup: { ...SHEET.pageSetup, margins: { ...SHEET.pageSetup.margins } },
      views: [{ showGridLines: false }],
    });
    ws.columns = SHEET.colPx.map((px) => ({ width: px / SHEET.mdwPx }));
    // 頭＝題（& は && にする。字の大きさの指定は付けない）／足＝ページ / 総ページ
    ws.headerFooter.oddHeader = '&C' + title.replace(/&/g, '&&');
    ws.headerFooter.oddFooter = '&C&P / &N';

    const font = { ...SHEET.font };
    const layout = blockLayout(opt.fields); // 選んだ欄（無い・配列でない時は全部）から、8 行の高さと行ごとの中身
    for (let i = 0; i < n; i++) {
      const e = entries[i];
      const { first, last, gap } = blockRows(i);
      const values = {
        number: startNumber + i,
        date: formatDateJa(e.dateTimeOriginal),
        file: fitFileName(cleanText(e.name)),
      };
      const rowOf = new Array(8); // 段の k 行目（0 始まり）を受け持つ欄
      for (const r of layout.rows) for (let s = 0; s < r.span; s++) rowOf[r.row + s] = r;
      layout.rowPt.forEach((pt, k) => {
        const row = ws.getRow(first + k);
        row.height = pt;
        const vTop = rowOf[k].key === 'note' || rowOf[k].key === 'blank' ? 'top' : 'middle';
        const c = row.getCell(3);
        c.font = font;
        c.border = BOX;
        c.alignment = { horizontal: 'distributed', vertical: vTop };
        const d = row.getCell(4);
        d.font = font;
        d.border = BOX;
        d.alignment = { horizontal: 'left', vertical: vTop, wrapText: true };
        row.getCell(1).border = BOX;
      });
      for (const r of layout.rows) {
        const top = first + r.row;
        if (r.key !== 'blank') {
          ws.getCell(top, 3).value = r.label;
          // 文字は文字列のまま入れる（ExcelJS は文字列を式にしない）。値の無い欄・空の欄は値を入れない。
          const v = values[r.key];
          if (v !== undefined && v !== '') ws.getCell(top, 4).value = v;
        }
        if (r.key === 'blank') ws.mergeCells(top, 3, last, 4);
        else if (r.span > 1) {
          ws.mergeCells(top, 3, top + r.span - 1, 3);
          ws.mergeCells(top, 4, top + r.span - 1, 4);
        }
      }
      ws.mergeCells(first, 1, last, 1);
      const gapRow = ws.getRow(gap);
      gapRow.height = SHEET.gapPt;
      if (hasPageBreakAfter(i, n)) gapRow.addPageBreak();

      // 絵＝tl は native（EMU・整数）、ext は px（小数のまま）。br・editAs は渡さない。
      const img = images[i];
      const fit = fitImage(img.width, img.height);
      const id = wb.addImage({ buffer: img.bytes, extension: img.type });
      ws.addImage(id, { tl: imageAnchor(i, fit, layout.rowPt), ext: { width: fit.w, height: fit.h } });
    }
    ws.pageSetup.printArea = `A1:D${blockRows(n - 1).gap}`;

    const out = await wb.xlsx.writeBuffer();
    const bytes = out instanceof Uint8Array
      ? new Uint8Array(out.buffer, out.byteOffset, out.byteLength)
      : new Uint8Array(out);
    return { ok: true, bytes, pages: pageCount(n) };
  } catch (err) {
    return { ok: false, reason: 'write-failed', message: String(err && err.message ? err.message : err) };
  }
}
