// 置き先と同じ CSP を付けた、試験用のサーバー（127.0.0.1 だけで待つ）。
// 使い方: node scripts/serve_csp.mjs [--port 8733] [--e2e] [--assets <フォルダ>]
//   /shashin-daicho/ → dist/shashin-daicho/（先に node scripts/build_site.mjs）
//   --assets を渡した時だけ /assets/ をそのフォルダから読むだけで配る（渡さなければ 404）
//   --e2e の時だけ、試験の口を足す：/__e2e/ → e2e/・/__fixtures/ → fixtures/・
//     POST /__report?name=…（JSON を out/e2e/ に保存）・POST /__save?name=…（xlsx を out/e2e/ に保存）
//   試験の口は本番のページ（web/・src/）には入れない。
import http from 'node:http';
import { createReadStream, mkdirSync, statSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
export const OUT_E2E = path.join(ROOT, 'out', 'e2e');

// 置き先と同じ（一字そのまま。変えない）
export const CSP = "default-src 'self'; script-src 'self' https://static.cloudflareinsights.com; connect-src 'self' https://cloudflareinsights.com; img-src 'self'; style-src 'self'; base-uri 'self'; form-action 'none'; frame-ancestors 'none'; upgrade-insecure-requests";

const TYPES = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.json': 'application/json; charset=utf-8',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.png': 'image/png',
  '.svg': 'image/svg+xml',
  '.txt': 'text/plain; charset=utf-8',
  '.xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
};
const SAFE_NAME = /^[A-Za-z0-9_-][A-Za-z0-9_.-]{0,120}$/;
const MAX_BODY = 512 * 1024 * 1024;

function send(res, status, body, type = 'text/plain; charset=utf-8') {
  res.statusCode = status;
  res.setHeader('Content-Type', type);
  res.end(body);
}

/** URL の道（/ で始まる・前置きを外した残り）を、base の下の本当の道にする。外へ出る物・おかしな物は null。 */
function resolveUnder(base, rest) {
  let segs;
  try {
    segs = rest.split('/').map((s) => decodeURIComponent(s));
  } catch {
    return null;
  }
  segs = segs.filter((s) => s !== '');
  for (const s of segs) {
    if (s === '.' || s === '..' || /[\\:\0]/.test(s) || s.includes('/')) return null;
  }
  const p = path.resolve(base, ...segs);
  const rel = path.relative(base, p);
  if (rel.startsWith('..') || path.isAbsolute(rel)) return null;
  return p;
}

function serveFile(res, file, req) {
  let st;
  try {
    st = statSync(file);
    if (st.isDirectory()) {
      file = path.join(file, 'index.html');
      st = statSync(file);
    }
  } catch {
    return send(res, 404, 'not found');
  }
  if (!st.isFile()) return send(res, 404, 'not found');
  res.statusCode = 200;
  res.setHeader('Content-Type', TYPES[path.extname(file).toLowerCase()] || 'application/octet-stream');
  res.setHeader('Content-Length', st.size);
  if (req.method === 'HEAD') return res.end();
  createReadStream(file).on('error', () => res.destroy()).pipe(res);
}

function readBody(req, cb) {
  const chunks = [];
  let size = 0;
  req.on('data', (d) => {
    size += d.length;
    if (size > MAX_BODY) { req.destroy(); return; }
    chunks.push(d);
  });
  req.on('end', () => cb(Buffer.concat(chunks)));
}

/**
 * @param {{ e2e?: boolean, assets?: string|null, onReport?: (name: string, body: object) => void,
 *           onRequest?: (method: string, url: string) => void }} [opts]
 */
export function createServer(opts = {}) {
  const { e2e = false, assets = null, onReport = null, onRequest = null } = opts;
  const assetsDir = assets ? path.resolve(assets) : null;
  const mounts = [['/shashin-daicho/', path.join(ROOT, 'dist', 'shashin-daicho')]];
  if (assetsDir) mounts.push(['/assets/', assetsDir]);
  if (e2e) {
    mounts.push(['/__e2e/', path.join(ROOT, 'e2e')]);
    mounts.push(['/__fixtures/', path.join(ROOT, 'fixtures')]);
  }

  return http.createServer((req, res) => {
    // 全部の応答に付ける
    res.setHeader('Content-Security-Policy', CSP);
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('Cache-Control', 'no-store');
    if (onRequest) onRequest(req.method, req.url);

    let u;
    try {
      u = new URL(req.url, 'http://127.0.0.1');
    } catch {
      return send(res, 400, 'bad request');
    }
    const p = u.pathname;

    if (e2e && req.method === 'POST' && (p === '/__report' || p === '/__save')) {
      const name = u.searchParams.get('name') || (p === '/__report' ? 'report.json' : '');
      const want = p === '/__report' ? '.json' : '.xlsx';
      if (!SAFE_NAME.test(name) || name.includes('..') || !name.toLowerCase().endsWith(want)) {
        return send(res, 400, 'bad name');
      }
      return readBody(req, (buf) => {
        mkdirSync(OUT_E2E, { recursive: true });
        if (p === '/__report') {
          let obj;
          try {
            obj = JSON.parse(buf.toString('utf8'));
          } catch {
            return send(res, 400, 'bad json');
          }
          writeFileSync(path.join(OUT_E2E, name), JSON.stringify(obj, null, 1) + '\n');
          send(res, 200, 'ok');
          if (onReport) onReport(name, obj);
          return;
        }
        writeFileSync(path.join(OUT_E2E, name), buf);
        send(res, 200, String(buf.length));
      });
    }

    if (req.method !== 'GET' && req.method !== 'HEAD') return send(res, 405, 'method not allowed');
    for (const [prefix, base] of mounts) {
      if (p === prefix.slice(0, -1)) {
        res.statusCode = 301;
        res.setHeader('Location', prefix);
        return res.end();
      }
      if (p.startsWith(prefix)) {
        const file = resolveUnder(base, p.slice(prefix.length));
        if (!file) return send(res, 400, 'bad path');
        return serveFile(res, file, req);
      }
    }
    return send(res, 404, 'not found');
  });
}

if (process.argv[1] && path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  const args = process.argv.slice(2);
  let port = 8733;
  let e2e = false;
  let assets = null;
  for (let k = 0; k < args.length; k++) {
    if (args[k] === '--port') port = Number(args[++k]);
    else if (args[k] === '--e2e') e2e = true;
    else if (args[k] === '--assets') assets = args[++k];
    else {
      console.error(`知らない引数：${args[k]}\n使い方: node scripts/serve_csp.mjs [--port 8733] [--e2e] [--assets <フォルダ>]`);
      process.exit(2);
    }
  }
  if (!Number.isInteger(port) || port < 0 || port > 65535) {
    console.error('--port は 0〜65535');
    process.exit(2);
  }
  if (assets) {
    try {
      if (!statSync(assets).isDirectory()) throw new Error();
    } catch {
      console.error(`--assets のフォルダが無い：${assets}`);
      process.exit(2);
    }
  }
  const server = createServer({ e2e, assets });
  server.listen(port, '127.0.0.1', () => {
    console.log(`http://127.0.0.1:${server.address().port}/shashin-daicho/${e2e ? '  （--e2e：試験の口あり）' : ''}`);
  });
}
