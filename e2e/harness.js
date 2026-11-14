// 通しの試験のページ側（serve_csp.mjs --e2e の時だけ /__e2e/ で配る。本番のページではない）。
// 置き先と同じ CSP の下で、架空の写真を 1 つずつ fetch して File にし、processAll → sortPhotos → splitBooks →
// buildLedger まで通す。写しの縦横・四隅の色・番号の帯・撮影日時・向きを manifest と比べ、xlsx を /__save、結果を /__report へ送る。
// 普通のスクリプト（モジュールでない）にして、ExcelJS より先に CSP の違反を数え始める。
'use strict';

(() => {
  const violations = [];
  document.addEventListener('securitypolicyviolation', (e) => {
    violations.push({
      directive: e.violatedDirective, blocked: e.blockedURI, source: e.sourceFile, line: e.lineNumber, sample: e.sample,
    });
  });

  const params = new URLSearchParams(location.search);
  const SET = params.get('set') || '';
  const SORT = params.get('sort') === 'name' ? 'name' : 'date';
  const TAG = params.get('tag') || 'browser';
  const COLORS = { tl: [255, 0, 0], tr: [0, 255, 0], bl: [0, 0, 255], br: [255, 255, 0] };
  const CORNER_TOL = 40;

  const now = () => performance.now();
  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const status = (s) => { const el = document.getElementById('status'); if (el) el.textContent = s; };
  const encPath = (p) => p.split('/').map(encodeURIComponent).join('/');

  function expectedSize(up) {
    const long = Math.max(up.w, up.h);
    const s = long > 1200 ? 1200 / long : 1;
    return { w: up.w * s, h: up.h * s };
  }

  /** 写しを canvas に描いて、四隅の色と番号の帯を読む。 */
  async function readCopy(bytes) {
    const bmp = await createImageBitmap(new Blob([bytes], { type: 'image/jpeg' }));
    const w = bmp.width;
    const h = bmp.height;
    const c = new OffscreenCanvas(w, h);
    const ctx = c.getContext('2d');
    ctx.drawImage(bmp, 0, 0);
    bmp.close();
    const data = ctx.getImageData(0, 0, w, h).data;
    c.width = 0;
    c.height = 0;
    const px = (x, y) => {
      const i = (Math.round(y) * w + Math.round(x)) * 4;
      return [data[i], data[i + 1], data[i + 2]];
    };
    const ins = Math.max(2, Math.floor(Math.max(w, h) * 0.04));
    const corners = {
      tl: px(ins, ins), tr: px(w - 1 - ins, ins), bl: px(ins, h - 1 - ins), br: px(w - 1 - ins, h - 1 - ins),
    };
    const cornerBad = Object.entries(corners)
      .filter(([k, got]) => got.some((v, i) => Math.abs(v - COLORS[k][i]) > CORNER_TOL))
      .map(([k, got]) => `${k}=${got.join(',')}`);
    // 帯＝x 0.25W〜0.75W・y 0.90H〜0.97H を 12 等分。マスの真ん中の明るさが 128 より下なら 1。
    const y = 0.935 * h;
    const bits = [];
    for (let i = 0; i < 12; i++) {
      const [r, g, b] = px(0.25 * w + 0.5 * w * (i + 0.5) / 12, y);
      bits.push(0.299 * r + 0.587 * g + 0.114 * b < 128 ? 1 : 0);
    }
    let n = 0;
    for (const bit of bits.slice(1, 11)) n = n * 2 + bit;
    // 下の端の真ん中の色（途中で切れた JPEG が灰色で埋まっているかを見るため）
    const bottomMid = px(w / 2, h - 2);
    return { w, h, cornerBad, ends: bits[0] === 1 && bits[11] === 1, label: n, bottomMid };
  }

  async function post(url, body, type) {
    const r = await fetch(url, { method: 'POST', body, headers: { 'Content-Type': type } });
    if (!r.ok) throw new Error(`${url} → ${r.status}`);
    return r.text();
  }

  async function main() {
    const tAll = now();
    const lib = await import('/shashin-daicho/lib/pipeline.js');
    const { sortPhotos } = await import('/shashin-daicho/lib/sort.js');
    const { splitBooks } = await import('/shashin-daicho/lib/layout.js');
    const { buildLedger } = await import('/shashin-daicho/lib/ledger.js');
    if (typeof ExcelJS === 'undefined') throw new Error('ExcelJS が読めていない');

    status('manifest');
    const manifest = await (await fetch(`/__fixtures/${encodeURIComponent(SET)}.manifest.json`)).json();
    const byPath = new Map(manifest.files.map((f) => [f.path, f]));

    // 写真を 1 つずつ fetch して File にする（webkitRelativePath の代わりに relativePath を渡す）
    status('fetch');
    const tFetch = now();
    const items = [];
    for (const f of manifest.files) {
      const r = await fetch(`/__fixtures/${encodeURIComponent(SET)}/${encPath(f.path)}`);
      if (!r.ok) throw new Error(`fetch ${f.path} → ${r.status}`);
      const blob = await r.blob();
      items.push({ file: new File([blob], f.path.split('/').pop(), { lastModified: 0 }), relativePath: f.path });
    }
    const fetchMs = now() - tFetch;

    status('processAll');
    const per = [];
    let prev = now();
    const tProc = now();
    const res = await lib.processAll(items, {
      onProgress: (done, total, last) => {
        const t = now();
        per.push({ path: last.relativePath, ms: t - prev, ok: last.ok });
        prev = t;
        if (done % 10 === 0 || done === total) status(`processAll ${done}/${total}`);
      },
    });
    const processMs = now() - tProc;
    items.length = 0; // 原本の File を放す

    // ---- manifest と比べる ----
    status('check');
    const mismatches = [];
    const results = new Map();
    for (const r of [...res.photos, ...res.failed]) results.set(r.relativePath, r);
    const observed = [];
    for (const f of manifest.files) {
      const r = results.get(f.path);
      if (!r) { mismatches.push(`${f.path}: 結果が無い`); continue; }
      const want = f.expect;
      if (want === 'ok') {
        if (!r.ok) { mismatches.push(`${f.path}: 写しが出来ない（${r.reason}・${r.detail}）`); continue; }
        const es = expectedSize(f.upright);
        if (Math.abs(r.image.width - es.w) > 1 || Math.abs(r.image.height - es.h) > 1) {
          mismatches.push(`${f.path}: 写しの縦横 ${r.image.width}×${r.image.height}（${es.w.toFixed(1)}×${es.h.toFixed(1)} のはず）`);
        }
        const rc = await readCopy(r.image.bytes);
        if (rc.w !== r.image.width || rc.h !== r.image.height) mismatches.push(`${f.path}: 写しの中身の縦横が ${rc.w}×${rc.h}`);
        if (rc.cornerBad.length) mismatches.push(`${f.path}: 四隅の色がちがう ${rc.cornerBad.join(' ')}`);
        if (!rc.ends) mismatches.push(`${f.path}: 番号の帯の両端が黒でない`);
        else if (rc.label !== f.label) mismatches.push(`${f.path}: 番号の帯が ${rc.label}（${f.label} のはず）`);
        if (r.dateTimeOriginal !== (f.dateTimeOriginal ?? null)) {
          mismatches.push(`${f.path}: 撮影日時 ${r.dateTimeOriginal}（${f.dateTimeOriginal} のはず）`);
        }
        if (r.orientation !== (f.orientation ?? null)) mismatches.push(`${f.path}: 向き ${r.orientation}（${f.orientation} のはず）`);
        if (r.image.type !== 'jpeg' || r.image.bytes[0] !== 0xff || r.image.bytes[1] !== 0xd8) mismatches.push(`${f.path}: 写しが JPEG でない`);
      } else {
        const okReason = want === 'unsupported' ? ['heic', 'unsupported'] : [want];
        if (r.ok) {
          // 黙って台帳に入る形。どう開けたかを残す
          const rc = await readCopy(r.image.bytes);
          observed.push({ path: f.path, note: f.note, decoded: true, w: r.image.width, h: r.image.height, bottomMid: rc.bottomMid });
          mismatches.push(`${f.path}: 開けて写しが出来た（${want} のはず・${f.note}）`);
        } else {
          observed.push({ path: f.path, note: f.note, decoded: false, reason: r.reason, detail: r.detail });
          if (!okReason.includes(r.reason)) mismatches.push(`${f.path}: 理由 ${r.reason}（${okReason.join(' か ')} のはず）`);
        }
      }
    }
    for (const p of results.keys()) if (!byPath.has(p)) mismatches.push(`${p}: manifest に無い結果`);

    // ---- 並べて、ブックに分けて、台帳を作る ----
    status('ledger');
    const sorted = sortPhotos(res.photos, SORT);
    const books = [];
    const ledgerMs = [];
    for (const [k, { start, end }] of splitBooks(sorted.length).entries()) {
      const part = sorted.slice(start, end);
      const t = now();
      const lr = await buildLedger(ExcelJS, part, { startNumber: start + 1 });
      ledgerMs.push(now() - t);
      const file = `${TAG}_${SET}_${SORT}_${k + 1}.xlsx`;
      if (!lr.ok) {
        mismatches.push(`${file}: 台帳が作れない（${lr.reason}${lr.message ? ' ' + lr.message : ''}）`);
        continue;
      }
      await post(`/__save?name=${encodeURIComponent(file)}`, lr.bytes, 'application/octet-stream');
      books.push({
        file, startNumber: start + 1, pages: lr.pages, bytes: lr.bytes.length,
        photos: part.map((p) => ({
          name: p.name, relativePath: p.relativePath, dateTimeOriginal: p.dateTimeOriginal,
          label: byPath.has(p.relativePath) ? byPath.get(p.relativePath).label : null,
          imageW: p.image.width, imageH: p.image.height, copyBytes: p.image.bytes.length,
        })),
      });
    }

    await sleep(200); // 違反のイベントは後から来る
    const ms = per.map((p) => p.ms);
    const slow = per.reduce((a, b) => (b.ms > (a ? a.ms : -1) ? b : a), null);
    const report = {
      set: SET, sort: SORT, tag: TAG, userAgent: navigator.userAgent,
      counts: { files: manifest.files.length, photos: res.photos.length, failed: res.failed.length, aborted: res.aborted },
      mismatches,
      violations,
      observed,
      failed: res.failed.map((r) => ({ relativePath: r.relativePath, reason: r.reason, detail: r.detail })),
      books,
      timing: {
        totalMs: now() - tAll, fetchMs, processMs,
        perPhotoAvgMs: ms.length ? ms.reduce((a, b) => a + b, 0) / ms.length : 0,
        perPhotoMaxMs: slow ? slow.ms : 0, perPhotoMaxPath: slow ? slow.path : null,
        ledgerMs,
      },
      memory: performance.memory ? { usedJSHeapSize: performance.memory.usedJSHeapSize } : null,
    };
    await post(`/__report?name=${encodeURIComponent(`${TAG}_${SET}_${SORT}.report.json`)}`, JSON.stringify(report), 'application/json');
    status('done');
    window.__e2eDone = true;
  }

  window.addEventListener('load', () => {
    main().catch(async (err) => {
      status('fatal');
      try {
        await post(`/__report?name=${encodeURIComponent(`${TAG}_${SET}_${SORT}.report.json`)}`,
          JSON.stringify({ set: SET, sort: SORT, tag: TAG, fatal: String((err && err.stack) || err), violations }), 'application/json');
      } catch { /* 送れなければ run_e2e の時間切れで分かる */ }
    });
  });
})();
