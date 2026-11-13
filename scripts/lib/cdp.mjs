// ブラウザ（Edge・Chrome の headless）を CDP で操る小さな部品。部品を入れず、Node の WebSocket と http だけで書く。
// 起こす・つなぐ・ページを開く・式を評価する・コンソールと例外を集める・要求を数える・画面を撮る・自分が起こした PID を止める。
// 止めるのは自分が起こした PID だけ（名前で止めない）。
import { spawn, execFileSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import http from 'node:http';
import path from 'node:path';

export const BROWSERS = {
  edge: [
    'C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe',
    'C:\\Program Files\\Microsoft\\Edge\\Application\\msedge.exe',
  ],
  chrome: [
    'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
    'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
  ],
};

// 最初の実行の画面や、外への問い合わせを減らす引数
const QUIET_ARGS = [
  '--headless=new',
  '--no-first-run',
  '--no-default-browser-check',
  '--disable-background-networking',
  '--disable-component-update',
  '--disable-sync',
  '--disable-extensions',
  '--disable-default-apps',
  '--disable-domain-reliability',
  '--disable-client-side-phishing-detection',
  '--disable-features=Translate,OptimizationHints,MediaRouter,InterestFeedContentSuggestions,EdgeCollections,msEdgeShopping',
  '--metrics-recording-only',
  '--no-pings',
  '--password-store=basic',
  '--enable-precise-memory-info',
];

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

function getJson(port, p) {
  return new Promise((resolve, reject) => {
    http.get({ host: '127.0.0.1', port, path: p }, (res) => {
      let b = '';
      res.setEncoding('utf8');
      res.on('data', (d) => { b += d; });
      res.on('end', () => {
        try { resolve(JSON.parse(b)); } catch (e) { reject(e); }
      });
    }).on('error', reject);
  });
}

/** 1 本の WebSocket の上の CDP のつなぎ。 */
class Session {
  constructor(ws) {
    this.ws = ws;
    this.nextId = 1;
    this.pending = new Map();
    this.handlers = new Map();
    ws.addEventListener('message', (ev) => {
      const msg = JSON.parse(typeof ev.data === 'string' ? ev.data : Buffer.from(ev.data).toString('utf8'));
      if (msg.id && this.pending.has(msg.id)) {
        const { resolve, reject } = this.pending.get(msg.id);
        this.pending.delete(msg.id);
        if (msg.error) reject(new Error(`${msg.error.message}${msg.error.data ? ' ' + msg.error.data : ''}`));
        else resolve(msg.result);
      } else if (msg.method) {
        for (const fn of this.handlers.get(msg.method) || []) fn(msg.params);
      }
    });
    ws.addEventListener('close', () => {
      for (const { reject } of this.pending.values()) reject(new Error('CDP のつなぎが切れた'));
      this.pending.clear();
    });
  }

  send(method, params = {}) {
    const id = this.nextId++;
    return new Promise((resolve, reject) => {
      this.pending.set(id, { resolve, reject });
      this.ws.send(JSON.stringify({ id, method, params }));
    });
  }

  on(method, fn) {
    if (!this.handlers.has(method)) this.handlers.set(method, []);
    this.handlers.get(method).push(fn);
  }

  close() {
    try { this.ws.close(); } catch { /* 閉じるだけ */ }
  }
}

function connect(url) {
  return new Promise((resolve, reject) => {
    const ws = new WebSocket(url);
    ws.addEventListener('open', () => resolve(new Session(ws)), { once: true });
    ws.addEventListener('error', (e) => reject(new Error(`WebSocket につなげない：${e.message || url}`)), { once: true });
  });
}

/**
 * ブラウザを headless で起こす（--remote-debugging-port=0・この試験だけの一時プロファイル）。
 * @param {'edge'|'chrome'} which
 * @param {{ extraArgs?: string[], timeoutMs?: number }} [opts]
 */
export async function launch(which, opts = {}) {
  const exe = (BROWSERS[which] || []).find((p) => existsSync(p));
  if (!exe) throw new Error(`ブラウザが見つからない：${which}`);
  const profile = mkdtempSync(path.join(tmpdir(), `shashin-e2e-${which}-`));
  const args = [...QUIET_ARGS, `--user-data-dir=${profile}`, '--remote-debugging-port=0', ...(opts.extraArgs || []), 'about:blank'];
  const child = spawn(exe, args, { stdio: 'ignore', windowsHide: true });
  const pid = child.pid;
  let exited = false;
  child.on('exit', () => { exited = true; });

  // 起きたら、プロファイルの DevToolsActivePort に口の番号が書かれる
  const portFile = path.join(profile, 'DevToolsActivePort');
  const until = Date.now() + (opts.timeoutMs || 30000);
  let port = null;
  while (Date.now() < until && !exited) {
    try {
      const n = Number(readFileSync(portFile, 'utf8').split(/\r?\n/)[0]);
      if (n > 0) { port = n; break; }
    } catch { /* まだ */ }
    await sleep(100);
  }
  const browser = { which, exe, pid, profile, port, child, get exited() { return exited; } };
  if (!port) {
    await stop(browser);
    throw new Error(`ブラウザの CDP の口が開かない（${which}）`);
  }
  const version = await getJson(port, '/json/version');
  browser.version = version.Browser;
  return browser;
}

/**
 * 最初のページにつなぎ、コンソール・例外・要求を集め始める。
 * @returns {Promise<Page>}
 */
export async function openPage(browser) {
  let target = null;
  for (let k = 0; k < 50 && !target; k++) {
    const list = await getJson(browser.port, '/json/list');
    target = list.find((t) => t.type === 'page');
    if (!target) await sleep(100);
  }
  if (!target) throw new Error('ページの target が無い');
  const s = await connect(target.webSocketDebuggerUrl);
  const page = new Page(s);
  await page.init();
  return page;
}

export class Page {
  constructor(session) {
    this.s = session;
    this.console = [];
    this.exceptions = [];
    this.logEntries = [];
    this.requests = [];
  }

  async init() {
    const s = this.s;
    s.on('Runtime.consoleAPICalled', (p) => {
      this.console.push({ type: p.type, text: p.args.map((a) => (a.value !== undefined ? String(a.value) : a.description || a.type)).join(' ') });
    });
    s.on('Runtime.exceptionThrown', (p) => {
      const d = p.exceptionDetails;
      this.exceptions.push((d.exception && d.exception.description) || d.text);
    });
    s.on('Log.entryAdded', (p) => this.logEntries.push({ level: p.entry.level, source: p.entry.source, text: p.entry.text, url: p.entry.url }));
    s.on('Network.requestWillBeSent', (p) => this.requests.push({ url: p.request.url, method: p.request.method, type: p.type }));
    await s.send('Runtime.enable');
    await s.send('Log.enable');
    await s.send('Network.enable');
    await s.send('Page.enable');
  }

  /** 開いて、load を待つ。 */
  async goto(url, timeoutMs = 30000) {
    const loaded = new Promise((resolve) => this.s.on('Page.loadEventFired', resolve));
    const r = await this.s.send('Page.navigate', { url });
    if (r.errorText) throw new Error(`開けない：${r.errorText}`);
    await Promise.race([loaded, sleep(timeoutMs)]);
  }

  /** 式を評価して値を返す（Promise は待つ）。 */
  async evaluate(expression) {
    const r = await this.s.send('Runtime.evaluate', { expression, awaitPromise: true, returnByValue: true });
    if (r.exceptionDetails) throw new Error((r.exceptionDetails.exception && r.exceptionDetails.exception.description) || r.exceptionDetails.text);
    return r.result.value;
  }

  /** 画面を撮る → PNG のバイト列。 */
  async screenshot() {
    const r = await this.s.send('Page.captureScreenshot', { format: 'png' });
    return Buffer.from(r.data, 'base64');
  }

  /** CDP の呼び出しをそのまま送る。 */
  send(method, params = {}) {
    return this.s.send(method, params);
  }

  /** 次からの文書の読み込みごとに、ページのどのスクリプトより先に式を走らせる（試験の見張り用。ページのファイルは変えない）。 */
  async addInitScript(source) {
    await this.s.send('Page.addScriptToEvaluateOnNewDocument', { source });
  }

  /** selector の <input type=file> に、ディスクのファイル（またはフォルダ）の道を渡す。change が起きる。 */
  async setFileInputFiles(selector, files) {
    const r = await this.s.send('Runtime.evaluate', { expression: `document.querySelector(${JSON.stringify(selector)})`, returnByValue: false });
    if (!r.result || !r.result.objectId) throw new Error(`入力が見つからない：${selector}`);
    try {
      await this.s.send('DOM.setFileInputFiles', { files, objectId: r.result.objectId });
    } finally {
      await this.s.send('Runtime.releaseObject', { objectId: r.result.objectId }).catch(() => {});
    }
  }

  /** 画面の幅と高さを決める（mobile＝狭い幅の端末として扱う）。 */
  async setViewport(width, height, mobile = false) {
    await this.s.send('Emulation.setDeviceMetricsOverride', { width, height, deviceScaleFactor: 1, mobile });
  }

  /** ダウンロードを downloadPath に落とす（聞かずに保存）。 */
  async allowDownloads(downloadPath) {
    await this.s.send('Browser.setDownloadBehavior', { behavior: 'allow', downloadPath, eventsEnabled: true });
  }

  /** 集めた要求のうち、origin（例 http://127.0.0.1:8733/）で始まらない物。 */
  outsideRequests(origin) {
    return this.requests.filter((q) => !q.url.startsWith(origin));
  }

  close() {
    this.s.close();
  }
}

/** 自分が起こしたブラウザだけを止める（CDP の Browser.close → 残れば自分の PID の木を taskkill）。プロファイルを消す。 */
export async function stop(browser) {
  if (!browser) return;
  if (browser.port && !browser.exited) {
    try {
      const v = await getJson(browser.port, '/json/version');
      const s = await connect(v.webSocketDebuggerUrl);
      await Promise.race([s.send('Browser.close').catch(() => {}), sleep(3000)]);
      s.close();
    } catch { /* 次で止める */ }
  }
  for (let k = 0; k < 30 && !browser.exited; k++) await sleep(100);
  if (!browser.exited && browser.pid) {
    try { execFileSync('taskkill', ['/PID', String(browser.pid), '/T', '/F'], { stdio: 'ignore' }); } catch { /* もう無い */ }
    for (let k = 0; k < 30 && !browser.exited; k++) await sleep(100);
  }
  // 子のプロセスがプロファイルを放すまで少し待って消す
  for (let k = 0; k < 20; k++) {
    try {
      rmSync(browser.profile, { recursive: true, force: true });
      break;
    } catch {
      await sleep(250);
    }
  }
}
