// 台帳の枠とページ割りの計算。DOM にも ExcelJS にも触らない。
// 数字は試作 1a（spike/1a/REPORT.md）で本物の Excel に開いて確かめた物。

const EMU_PER_PX = 9525;
const EMU_PER_PT = 12700;

const deepFreeze = (o) => {
  for (const v of Object.values(o)) if (v && typeof v === 'object') deepFreeze(v);
  return Object.freeze(o);
};

export const SHEET = deepFreeze({
  sheetName: '写真台帳',
  defaultTitle: '工事写真台帳',
  mdwPx: 8, // 游ゴシック の最大桁幅。列の幅（字数）＝ px ÷ 8
  colPx: [420, 8, 64, 166], // A 写真の枠／B すき間／C 見出し／D 値
  labels: ['写真番号', '撮影年月日', 'ファイル名', '工種', '種別', '細別', '撮影箇所', '備考'],
  rowPt: [18, 18, 45, 18, 18, 18, 18, 83.25], // 段の 1〜8 行目
  gapPt: 7.5, // 段の 9 行目（すき間）
  rowsPerBlock: 9,
  frame: { wPx: 420, hPx: 315 }, // A 列の 8 行を結合した枠（236.25pt＝315px）
  padPx: 12, // 印刷で行が約 2% 縮むので、枠の内側に取る余白
  imageBox: { wPx: 396, hPx: 291 },
  emuPerPx: EMU_PER_PX,
  emuPerPt: EMU_PER_PT,
  blocksPerPage: 3,
  maxPerBook: 150,
  font: { name: '游ゴシック', size: 9, charset: 128 },
  pageSetup: {
    paperSize: 9, orientation: 'portrait', horizontalCentered: true,
    fitToPage: true, fitToWidth: 1, fitToHeight: 0,
    margins: { left: 0.71, right: 0.39, top: 0.59, bottom: 0.59, header: 0.2, footer: 0.2 },
  },
  // ファイル名の欄（D 列・45pt＝3 行）に収まる幅の見積り（px・9pt の字）。1 行に約 157px 入り、折り返しで
  // 1 行ごとに字 1 つ分ほど余る。3 行＝約 440。段 3b で本物の Excel の PDF で確かめた（W 14 字・日本語 13 字が 1 行）。
  fileNameMaxWidth: 440,
  charWidthNarrow: 7, // U+0080 未満と半角カナの標準（小文字・数字ほか）
  // 太い ASCII は広い（段 3b・本物の Excel の PDF で字ごとに測った値を 9pt の px にした物）
  charWidthWider: { 11: 'MW', 10: '%@m', 9: '&+<=>DGHNOQUw~', 8: 'ABCKPRSTVXY' },
  charWidthWide: 12, // ほか
  extMaxChars: 10, // 最後の . から後ろが「. 込みで 10 字以内」（かつ . が先頭でない）時だけ拡張子として残す
  // 頭（ヘッダー）に入れる題の字数（Array.from で数える・… を含む）。Excel のヘッダーは 255 字までなので、
  // & が全部 && に伸びても '&C' を足して 255 を超えないよう安全側に取る（2＋2×99＋1＝201）。
  titleMaxChars: 100,
  // 頭は折り返されず、印刷の幅（A4 − 左右の余白＝約 516pt）を超えた分が切られる（段 3b・本物の Excel の PDF で見た：
  // 日本語は 1 字 11pt で 46 字まで）。字の幅は charWidth と同じ見積り（px・日本語＝12）で、516pt≒561 に余裕を見て 540。
  titleMaxWidth: 540,
});

const isPos = (v) => typeof v === 'number' && Number.isFinite(v) && v > 0;

/** 絵を縦横比を保って箱に収め、枠の真ん中に置く。→ { w, h, x, y }（px・枠の左上から）。不正は null。 */
export function fitImage(w, h) {
  if (!isPos(w) || !isPos(h)) return null;
  const { imageBox, frame } = SHEET;
  const s = Math.min(imageBox.wPx / w, imageBox.hPx / h);
  const fw = w * s;
  const fh = h * s;
  return { w: fw, h: fh, x: (frame.wPx - fw) / 2, y: (frame.hPx - fh) / 2 };
}

/** ブックの中の i 段目（0 始まり）の行番号（1 始まり）。 */
export function blockRows(i) {
  const first = i * SHEET.rowsPerBlock + 1;
  return { first, last: first + 7, gap: first + 8 };
}

/** 絵の左上を「A 列の何 EMU 目・どの行の何 EMU 目か」で返す（nativeRow は 0 始まり）。 */
export function imageAnchor(i, fit) {
  let yEmu = Math.round(fit.y * EMU_PER_PX);
  let rr = 0;
  while (rr < SHEET.rowPt.length - 1 && yEmu >= SHEET.rowPt[rr] * EMU_PER_PT) {
    yEmu -= SHEET.rowPt[rr] * EMU_PER_PT;
    rr++;
  }
  return {
    nativeCol: 0,
    nativeColOff: Math.round(fit.x * EMU_PER_PX),
    nativeRow: blockRows(i).first - 1 + rr,
    nativeRowOff: yEmu,
  };
}

/** i 段目の後ろで改ページするか（3 段ごと・最後の段には付けない）。 */
export function hasPageBreakAfter(i, n) {
  return (i + 1) % SHEET.blocksPerPage === 0 && i + 1 < n;
}

export function pageCount(n) {
  return Math.ceil(n / SHEET.blocksPerPage);
}

/** n 件を max 件ずつのブックに分ける。→ [{ start, end }]（0 始まり・end は含まない） */
export function splitBooks(n, max = SHEET.maxPerBook) {
  const out = [];
  for (let start = 0; start < n; start += max) out.push({ start, end: Math.min(n, start + max) });
  return out;
}

const WIDER = Object.entries(SHEET.charWidthWider).map(([w, set]) => [Number(w), set]);
const charWidth = (ch) => {
  const c = ch.codePointAt(0);
  if (c >= 0x80 && !(c >= 0xff61 && c <= 0xff9f)) return SHEET.charWidthWide;
  if (c < 0x80) for (const [w, set] of WIDER) if (set.includes(ch)) return w;
  return SHEET.charWidthNarrow;
};
const widthOf = (chars) => chars.reduce((a, ch) => a + charWidth(ch), 0);

/** 文字列の幅の見積り（px・9pt 換算・日本語＝12）。頭の題の切り方が使う。 */
export function textWidth(s) {
  return widthOf(Array.from(s));
}

// 欄の 1 行に入る幅（px）。本物の Excel の PDF で、W が 14 字・日本語が 13 字入り、1 字増やすと次の行へ行く事を見た。
const LINE_WIDTH = 157;
const MAX_LINES = 3;
const isWide = (ch) => charWidth(ch) === SHEET.charWidthWide;

/**
 * 名前が欄（幅 LINE_WIDTH の行）で何行になるかの見積り。折り返せる所＝空白の後ろ・「…」の後ろ・全角の字の前後。
 * （「…」の後ろで折り返す事は本物の Excel で見た。ハイフンの後ろで折るかは未確認なので、折れない物として数える。）
 * 1 行に収まらない単語は、新しい行から始めて字ごとに折り返す。
 */
export function nameLines(name) {
  const words = [];
  let cur = [];
  for (const ch of Array.from(name)) {
    cur.push(ch);
    if (ch === ' ' || ch === '…' || isWide(ch)) { words.push(cur); cur = []; }
  }
  if (cur.length) words.push(cur);
  let lines = 1;
  let x = 0;
  for (const w of words) {
    const trail = w[w.length - 1] === ' ' ? charWidth(' ') : 0;
    const ww = widthOf(w);
    if (x + ww - trail <= LINE_WIDTH) { x += ww; continue; }
    if (ww - trail <= LINE_WIDTH) { lines++; x = ww; continue; }
    if (x > 0) lines++;
    x = 0;
    for (const ch of w) {
      const cw = charWidth(ch);
      if (x + cw > LINE_WIDTH) { lines++; x = 0; }
      x += cw;
    }
  }
  return lines;
}

// 真ん中を「…」で詰めて、幅を max 以下にする（拡張子は残す）。
function squeeze(chars, max) {
  const dot = chars.lastIndexOf('.');
  const extLen = dot > 0 && chars.length - dot <= SHEET.extMaxChars ? chars.length - dot : 0;
  const ext = chars.slice(chars.length - extLen);
  const stem = chars.slice(0, chars.length - extLen);

  const rest = max - SHEET.charWidthWide - widthOf(ext); // 「…」と拡張子を除いた残り
  let headBudget = Math.ceil(rest / 2);
  let h = 0;
  while (h < stem.length && charWidth(stem[h]) <= headBudget) headBudget -= charWidth(stem[h++]);
  let tailBudget = rest - Math.ceil(rest / 2) + headBudget; // 頭で余った分は尾へ
  let t = stem.length;
  while (t > h && charWidth(stem[t - 1]) <= tailBudget) tailBudget -= charWidth(stem[--t]);

  return stem.slice(0, h).join('') + '…' + stem.slice(t).join('') + ext.join('');
}

/** 欄（3 行）に収まらない名前なら、真ん中を「…」で詰める（拡張子は残す）。収まる行数になるまで幅の上限を下げる。 */
export function fitFileName(name) {
  const chars = Array.from(name);
  if (widthOf(chars) <= SHEET.fileNameMaxWidth && nameLines(name) <= MAX_LINES) return name;
  let max = SHEET.fileNameMaxWidth;
  let r = squeeze(chars, max);
  while (nameLines(r) > MAX_LINES && max > 60) {
    max -= 6;
    r = squeeze(chars, max);
  }
  return r;
}

/**
 * 制御文字（U+0000〜U+001F・U+007F）と、対になっていないサロゲートと、XML に書けない U+FFFE・U+FFFF を
 * 取り除く。文字列でなければ ''。
 */
export function cleanText(s) {
  if (typeof s !== 'string') return '';
  return Array.from(s).filter((ch) => {
    const c = ch.codePointAt(0);
    return !(c <= 0x1f || c === 0x7f || (c >= 0xd800 && c <= 0xdfff) || c === 0xfffe || c === 0xffff);
  }).join('');
}
