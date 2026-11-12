// html-FileNest.js — 网页版本地文件管理器
// 项目：html-FileNest  版本：1.4.0.0
// 双击无参数 => 部署并启动；--serve => 仅服务；--uninstall => 卸载

const fs     = require('fs');
const path   = require('path');
const net    = require('net');
const http   = require('http');
const https  = require('https');
const urlLib = require('url');
const crypto = require('crypto');
const mammoth = require('mammoth');
const WordExtractor = require('word-extractor');
const XLSX = require('xlsx');
const JSZip = require('jszip');
const { spawn, spawnSync } = require('child_process');
const { pipeline } = require('stream');
const zlib = require('zlib');
const { promisify } = require('util');

const pipeAsync = promisify(pipeline);

process.on('uncaughtException', (err) => {
  console.error('[html-FileNest 防崩溃] ' + (err && err.stack || err));
});

const APP_NAME    = 'html-FileNest';
const SERVICE_VER = '1.4.2.2';
const HOST        = '127.0.0.1';
const INSTALL_DIR = 'C:\\html-FileNest';
const EXE_PATH    = path.join(INSTALL_DIR, 'html-FileNest.exe');
const CONFIG_PATH = path.join(INSTALL_DIR, 'config.json');
const PFX_PATH    = path.join(INSTALL_DIR, 'filenest.pfx');
const CER_PATH    = path.join(INSTALL_DIR, 'filenest.cer');
const VBS_PATH    = path.join(INSTALL_DIR, 'start.vbs');
const TASK_NAME   = 'html-FileNest';
const PFX_PASS    = 'htmlfilenest';

// 默认端口（避让原 htmlFileAgent 的 19000/19443）
const DEFAULT_PORTS = { http: 20000, https: 20443 };

// ============ 默认配置 ============
const DEFAULT_CONFIG = {
  ports: DEFAULT_PORTS,
  mode: 'whitelist',
  sources: {
    default: { type: 'local', tryAll: true, roots: ['*'] }
  },
  allowedExtensions: [],
  blockExecutables: true,
  blockedExtensions: [
    '.exe', '.bat', '.cmd', '.com', '.scr', '.ps1', '.vbs', '.vbe',
    '.js', '.jse', '.wsf', '.wsh', '.msi', '.cpl', '.hta', '.pif',
    '.reg', '.inf', '.lnk', '.url', '.dll'
  ],
  cacheDir: 'C:\\html-FileNest\\cache',
  cacheMaxAgeHours: 24,
  downloadTimeoutMs: 120000,
  allowedHosts: [],
  security: {
    requireLocalOrigin: true,
    authToken: ''
  }
};

// ============ 入口 ============
// 注意：入口分发放在文件末尾执行。若在此处直接调用 runServer()/doAuto()，
// 会先于后面的 let _cfgCache 声明执行 loadConfig()，触发暂时性死区（TDZ）异常。
// 具体分发逻辑见文件末尾。

// ============ 配置热加载 ============
let _cfgCache = null, _cfgMtime = 0;
function loadConfig() {
  try {
    const st = fs.statSync(CONFIG_PATH);
    if (_cfgCache && st.mtimeMs === _cfgMtime) return _cfgCache;
    const cfg = JSON.parse(fs.readFileSync(CONFIG_PATH, 'utf8'));
    _cfgCache = normalizeConfig(cfg);
    _cfgMtime = st.mtimeMs;
    return _cfgCache;
  } catch (e) {
    if (!_cfgCache) _cfgCache = normalizeConfig(DEFAULT_CONFIG);
    return _cfgCache;
  }
}

function normalizeConfig(cfg) {
  const c = JSON.parse(JSON.stringify(cfg));
  c.ports = c.ports || DEFAULT_PORTS;
  c.sources = c.sources || DEFAULT_CONFIG.sources;
  c.mode = c.mode || 'whitelist';
  c.allowedExtensions = Array.isArray(c.allowedExtensions) ? c.allowedExtensions : [];
  c.blockExecutables = c.blockExecutables !== false;
  c.blockedExtensions = Array.isArray(c.blockedExtensions)
    ? c.blockedExtensions : DEFAULT_CONFIG.blockedExtensions;
  c.cacheDir = c.cacheDir || DEFAULT_CONFIG.cacheDir;
  c.cacheMaxAgeHours = c.cacheMaxAgeHours != null ? c.cacheMaxAgeHours : 24;
  c.downloadTimeoutMs = c.downloadTimeoutMs || 120000;
  c.allowedHosts = c.allowedHosts || [];
  c.security = c.security || {};
  if (c.security.requireLocalOrigin === undefined) c.security.requireLocalOrigin = true;
  c.security.authToken = c.security.authToken || '';

  const expand = (s) => typeof s === 'string'
    ? s.replace(/\$\{([^}]+)\}/g, (_, k) => process.env[k] || '')
    : s;

  Object.keys(c.sources).forEach(k => {
    const src = c.sources[k];
    if (src.type === 'local' && !src.roots && src.root) {
      src.roots = [src.root]; delete src.root;
    }
    if (src.type === 'local' && (!src.roots || !src.roots.length)) {
      src.roots = ['*'];
    }
    if (src.roots) src.roots = src.roots.map(expand);
    if (src.base) src.base = expand(src.base);
    if (src.tryAll == null) src.tryAll = true;
  });
  c.cacheDir = expand(c.cacheDir);
  return c;
}

// ============ 部署 ============
async function doAuto() {
  const cfg = loadConfig();
  if (fs.existsSync(EXE_PATH) && taskExists(TASK_NAME) && await portOpen(cfg.ports.http)) {
    showMsg(APP_NAME + ' 已在运行',
      '· 文件管理器：http://127.0.0.1:' + cfg.ports.http + '/\n' +
      '· 健康检查：http://127.0.0.1:' + cfg.ports.http + '/ping',
      'Information');
    return;
  }
  if (!isAdmin()) { elevateSelf([]); return; }
  await install();
  showMsg(APP_NAME + ' 部署成功',
    '· 文件管理器：http://127.0.0.1:' + cfg.ports.http + '/\n' +
    '· HTTPS：https://127.0.0.1:' + cfg.ports.https + '/\n' +
    '· 配置：' + CONFIG_PATH + '\n\n' +
    '默认允许浏览全盘。若需限制范围，请编辑 config.json 的 sources.default.roots。\n\n' +
    '卸载：html-FileNest.exe --uninstall',
    'Information');
}

async function install() {
  if (!fs.existsSync(INSTALL_DIR)) fs.mkdirSync(INSTALL_DIR, { recursive: true });
  if (process.execPath.toLowerCase() !== EXE_PATH.toLowerCase()) {
    fs.copyFileSync(process.execPath, EXE_PATH);
  }
  // 把独立 UI 页面 index.html 复制到 INSTALL_DIR（源：exe 旁 / 脚本旁 / cwd）
  {
    const uiDst = path.join(INSTALL_DIR, 'index.html');
    let uiSrc = null;
    for (const cand of findUIDirCandidates()) {
      try {
        if (fs.existsSync(cand) && fs.statSync(cand).isFile()) { uiSrc = cand; break; }
      } catch (e) {}
    }
    if (uiSrc && path.resolve(uiSrc).toLowerCase() !== path.resolve(uiDst).toLowerCase()) {
      fs.copyFileSync(uiSrc, uiDst);
    } else if (!uiSrc && !fs.existsSync(uiDst)) {
      // 自动部署：从内嵌资源释放 index.html（构建时注入，用户可在此文件基础上自由修改）
      fs.writeFileSync(uiDst, UI_HTML_DEPLOY, 'utf8');
      log('index.html 已自动部署到 ' + uiDst);
    }
  }
  if (!fs.existsSync(CONFIG_PATH)) {
    fs.writeFileSync(CONFIG_PATH, JSON.stringify(DEFAULT_CONFIG, null, 2), 'utf8');
  }
  if (!fs.existsSync(VBS_PATH)) {
    fs.writeFileSync(VBS_PATH,
      'CreateObject("WScript.Shell").Run """' + EXE_PATH + '"" --serve", 0, False\r\n',
      'ascii');
  }
  if (!fs.existsSync(PFX_PATH)) makeCert();

  spawnSync('schtasks', ['/create', '/tn', TASK_NAME,
    '/tr', 'wscript.exe "' + VBS_PATH + '"',
    '/sc', 'onlogon', '/rl', 'highest', '/f'],
    { stdio: 'ignore', windowsHide: true });

  spawn('wscript.exe', [VBS_PATH],
    { detached: true, stdio: 'ignore', windowsHide: true }).unref();

  const cfg = loadConfig();
  for (let i = 0; i < 50; i++) {
    if (await portOpen(cfg.ports.http)) return;
    await sleep(300);
  }
  throw new Error('服务启动超时');
}

// ============ 卸载 ============
function doUninstall() {
  if (!isAdmin()) { elevateSelf(['--uninstall']); return; }
  spawnSync('schtasks', ['/end', '/tn', TASK_NAME], { stdio: 'ignore', windowsHide: true });
  spawnSync('schtasks', ['/delete', '/tn', TASK_NAME, '/f'], { stdio: 'ignore', windowsHide: true });
  const cfg = loadConfig();
  try {
    runPS("Get-NetTCPConnection -LocalPort " + cfg.ports.http + "," + cfg.ports.https +
      " -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique | " +
      "ForEach-Object { Stop-Process -Id $_ -Force -ErrorAction SilentlyContinue }");
  } catch (e) {}
  try {
    runPS("Get-ChildItem Cert:\\LocalMachine\\Root | Where-Object {$_.Subject -like '*html-FileNest*'} | " +
      "Remove-Item -Force -ErrorAction SilentlyContinue");
  } catch (e) {}
  [VBS_PATH, PFX_PATH, CER_PATH].forEach(f => {
    try { if (fs.existsSync(f)) fs.unlinkSync(f); } catch (e) {}
  });
  showMsg('卸载完成', 'html-FileNest.exe 本体仍在 ' + INSTALL_DIR, 'Information');
}

// ============ 安全校验 ============
function isLocalHostHeader(h) {
  if (!h) return false;
  const s = String(h).toLowerCase();
  const m = s.match(/^\[([^\]]+)\](?::\d+)?$/);
  if (m) return m[1] === '::1';
  const parts = s.split(':');
  return parts[0] === '127.0.0.1' || parts[0] === 'localhost';
}
function isLocalOrigin(o) {
  if (!o) return true;
  try {
    const u = new URL(o);
    const h = u.hostname.toLowerCase().replace(/^\[|\]$/g, '');
    return h === '127.0.0.1' || h === 'localhost' || h === '::1';
  } catch (e) { return false; }
}
// If Origin is present it must be local. Without Origin (e.g. cross-site <img>/<script>),
// fall back to Sec-Fetch-Site. Clients sending neither (curl, local scripts) are allowed.
function requestOriginAllowed(req) {
  const o = req.headers.origin;
  if (o) return isLocalOrigin(o);
  const site = req.headers['sec-fetch-site'];
  return !site || site === 'same-origin' || site === 'none';
}

// ============ 服务 ============
function runServer() {
  const cfg = loadConfig();
  http.createServer(handleRequest)
    .on('error', e => log('HTTP ' + cfg.ports.http + ' error: ' + (e.code || e.message)))
    .listen(cfg.ports.http, HOST, () => log('HTTP  http://' + HOST + ':' + cfg.ports.http));

  if (fs.existsSync(PFX_PATH)) {
    try {
      const pfx = fs.readFileSync(PFX_PATH);
      https.createServer({ pfx, passphrase: PFX_PASS }, handleRequest)
        .on('error', e => log('HTTPS ' + cfg.ports.https + ' error: ' + (e.code || e.message)))
        .listen(cfg.ports.https, HOST, () => log('HTTPS https://' + HOST + ':' + cfg.ports.https));
    } catch (e) {
      log('HTTPS disabled: ' + e.message);
    }
  }
  setTimeout(() => cleanupCache(false).catch(() => {}), 5000);
}

function handleRequest(req, res) {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'SAMEORIGIN');
  if (req.headers.origin && isLocalOrigin(req.headers.origin)) {
    res.setHeader('Access-Control-Allow-Origin', req.headers.origin);
    res.setHeader('Vary', 'Origin');
  }
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers',
    'Content-Type, X-Requested-With, X-Auth-Token, Access-Control-Request-Private-Network');
  res.setHeader('Access-Control-Allow-Private-Network', 'true');
  res.setHeader('Cache-Control', 'no-store');
  if (req.method === 'OPTIONS') { res.writeHead(204); return res.end(); }

  let parsed;
  try { parsed = urlLib.parse(req.url, true); }
  catch (e) { return sendJson(res, 400, { ok: false, error: 'bad url' }); }

  const p = parsed.pathname, q = parsed.query || {};
  const cfg = loadConfig();

  if (p === '/ping' || p === '/open' || p === '/exists' || p === '/reveal' ||
      p === '/open-folder' || p === '/cache/clear') {
    if (!isLocalHostHeader(req.headers.host)) {
      return sendJson(res, 403, { ok: false, error: 'invalid host header' });
    }
    const secLegacy = cfg.security || {};
    if (secLegacy.requireLocalOrigin !== false && !requestOriginAllowed(req)) {
      return sendJson(res, 403, { ok: false, error: 'origin not allowed' });
    }
  }

  if (p === '/ping') {
    return sendJson(res, 200, {
      ok: true, service: APP_NAME, name: APP_NAME, version: SERVICE_VER, pid: process.pid,
      ports: cfg.ports,
      sources: Object.keys(cfg.sources),
      mode: cfg.mode,
      allowAllExtensions: cfg.allowedExtensions.length === 0,
      allowedExtensions: cfg.allowedExtensions
    });
  }

  if (p === '/open' || p === '/exists' || p === '/reveal' || p === '/open-folder') {
    return handleFileOp(p, q, res, cfg).catch(err => {
      sendJson(res, 500, { ok: false, error: String(err && err.message || err) });
    });
  }

  if (p === '/cache/clear') {
    cleanupCache(true).then(() => sendJson(res, 200, { ok: true }))
      .catch(e => sendJson(res, 500, { ok: false, error: String(e) }));
    return;
  }

  if (p === '/' || p === '/index.html' || p === '/ui') {
    return serveUI(res);
  }

  if (p.startsWith('/api/')) {
    return handleApi(req, res, p, q, cfg);
  }

  sendJson(res, 404, { ok: false, error: 'not found' });
}

// ============ UI 页面加载（index.html 独立文件，带 mtime 缓存） ============
// 查找顺序：优先 INSTALL_DIR，其次 exe 同目录，最后 __dirname（开发运行场景）
function findUIDirCandidates() {
  return [
    path.join(path.dirname(process.execPath), 'index.html'), // exe 自身旁边
    path.join(__dirname, 'index.html'),                       // 脚本旁（开发模式）
    path.join(process.cwd(), 'index.html')                    // 当前目录兜底
  ];
}
function findUIFile() {
  const candidates = [
    path.join(INSTALL_DIR, 'index.html'), // 优先部署目录
    ...findUIDirCandidates()
  ];
  for (const f of candidates) {
    try {
      if (fs.existsSync(f) && fs.statSync(f).isFile()) return f;
    } catch (e) {}
  }
  return null;
}

// ★ 部署资源：install() 时自动释放 index.html 到 INSTALL_DIR（用户可自行修改）
const UI_HTML_DEPLOY = __EMBEDDED_UI_HTML__;

// ★ 第三方开源许可全文（嵌入 exe，--licenses 可随时查看）
const THIRD_PARTY_NOTICES = __EMBEDDED_NOTICES__;

let _uiCache = null, _uiMtime = 0, _uiPath = null;
function loadUI() {
  const f = findUIFile();
  if (!f) return null;
  try {
    const st = fs.statSync(f);
    if (_uiCache !== null && _uiPath === f && st.mtimeMs === _uiMtime) return _uiCache;
    _uiCache = fs.readFileSync(f, 'utf8');
    _uiMtime = st.mtimeMs;
    _uiPath = f;
    return _uiCache;
  } catch (e) {
    return _uiCache;
  }
}

function serveUI(res) {
  const body = loadUI();
  if (body == null) {
    return sendJson(res, 500, {
      ok: false,
      error: 'index.html not found (deployed at ' + INSTALL_DIR + '\\index.html)'
    });
  }
  res.writeHead(200, {
    'Content-Type': 'text/html; charset=utf-8',
    'Cache-Control': 'no-cache',
    'Content-Length': Buffer.byteLength(body)
  });
  res.end(body);
}

// ============ 文件管理器 API ============
async function handleApi(req, res, p, q, cfg) {
  if (!isLocalHostHeader(req.headers.host)) {
    return sendJson(res, 403, { ok: false, error: 'invalid host header' });
  }
  const sec = cfg.security || {};
  if (sec.requireLocalOrigin !== false && !requestOriginAllowed(req)) {
    return sendJson(res, 403, { ok: false, error: 'origin not allowed' });
  }
  if (sec.authToken && req.headers['x-auth-token'] !== sec.authToken) {
    return sendJson(res, 401, { ok: false, error: 'unauthorized' });
  }

  try {
    if (req.method === 'GET' || req.method === 'HEAD') {
      if (p === '/api/roots')  return listRoots(res, cfg);
      if (p === '/api/drives') return listDrives(res);
      if (p === '/api/ls')     return await listDir(res, q, cfg);
      if (p === '/api/stat')   return await statPath(res, q, cfg);
      if (p === '/api/raw')    return await streamFile(req, res, q, cfg);
      if (p === '/api/read')   return await readTextFile(res, q, cfg);
      if (p === '/api/office-preview') return await officePreview(res, q, cfg);
      if (p === '/api/archive-list') return await archiveList(res, q, cfg);
      if (p === '/api/archive-extract') return await archiveExtract(req, res, q, cfg);
      return sendJson(res, 404, { ok: false, error: 'not found' });
    }

    if (req.method === 'POST') {
      if (p === '/api/upload') return await uploadOp(req, res, q, cfg);
      const body = await readBody(req);
      const data = body ? JSON.parse(body) : {};
      if (p === '/api/mkdir')  return await mkdirOp(res, data, cfg);
      if (p === '/api/rename') return await renameOp(res, data, cfg);
      if (p === '/api/rm')     return await rmOp(res, data, cfg);
      if (p === '/api/mv')     return await mvOp(res, data, cfg);
      if (p === '/api/cp')     return await cpOp(res, data, cfg);
      if (p === '/api/write')  return await writeFileOp(res, data, cfg);
      if (p === '/api/open')   return await openOp(res, data, cfg);
      return sendJson(res, 404, { ok: false, error: 'not found' });
    }

    sendJson(res, 405, { ok: false, error: 'method not allowed' });
  } catch (e) {
    if (!res.headersSent) sendJson(res, 400, { ok: false, error: String(e.message || e) });
    else { try { res.end(); } catch (_) {} }
  }
}

// ---------- 辅助 ----------
function readBody(req, limit = 10 * 1024 * 1024) {
  return new Promise((resolve, reject) => {
    let size = 0;
    const chunks = [];
    req.on('data', c => {
      size += c.length;
      if (size > limit) { reject(new Error('body too large')); req.destroy(); return; }
      chunks.push(c);
    });
    req.on('end', () => resolve(Buffer.concat(chunks).toString('utf8')));
    req.on('error', reject);
  });
}

function resolveAllowedPath(pathStr, cfg) {
  const s = String(pathStr || '');
  if (!s) throw new Error('path required');
  if (/[\x00-\x1f]/.test(s)) throw new Error('invalid characters');
  let norm;
  try { norm = path.win32.normalize(s); }
  catch (e) { throw new Error('invalid path'); }
  if (!norm.startsWith('\\\\') && !/^[A-Za-z]:\\/.test(norm)) {
    throw new Error('must be absolute path');
  }
  for (const name of Object.keys(cfg.sources)) {
    const src = cfg.sources[name];
    if (src.type !== 'local') continue;
    if (matchAnyRoot(norm, src.roots)) return norm;
  }
  throw new Error('path not in any configured source');
}

function listRoots(res, cfg) {
  const roots = [];
  const seen = new Set();
  for (const [name, src] of Object.entries(cfg.sources)) {
    if (src.type !== 'local') continue;
    for (const r of src.roots) {
      if (!r) continue;
      if (r === '*' || r === '**') {
        if (!seen.has('*')) { seen.add('*'); roots.push({ source: name, path: '*', display: '所有位置' }); }
        continue;
      }
      let expanded = null;
      if (!/[?*]/.test(r)) {
        expanded = r.replace(/[\\\/]$/, '');
      } else if (/[\\\/]\*+$/.test(r)) {
        expanded = r.replace(/[\\\/]\*+$/, '');
      }
      if (!expanded) continue;
      const key = name + '|' + expanded.toLowerCase();
      if (seen.has(key)) continue;
      seen.add(key);
      roots.push({
        source: name,
        path: expanded,
        display: path.win32.basename(expanded) || expanded
      });
    }
  }
  sendJson(res, 200, { ok: true, roots });
}

function listDrives(res) {
  const drives = [];
  for (let c = 65; c <= 90; c++) {
    const d = String.fromCharCode(c) + ':\\';
    try { fs.accessSync(d); drives.push(d); } catch (e) {}
  }
  sendJson(res, 200, { ok: true, drives });
}

async function listDir(res, q, cfg) {
  const dir = resolveAllowedPath(q.path, cfg);
  const st = await fs.promises.stat(dir);
  if (!st.isDirectory()) throw new Error('not a directory');

  const ents = await fs.promises.readdir(dir, { withFileTypes: true });
  const items = [];
  for (const ent of ents) {
    const full = path.win32.join(dir, ent.name);
    let type = 'other', size = 0, mtime = 0;
    try {
      const s = await fs.promises.stat(full);
      type = s.isDirectory() ? 'dir' : s.isFile() ? 'file' : 'other';
      size = s.size;
      mtime = s.mtimeMs;
    } catch (e) {}
    items.push({
      name: ent.name,
      path: full,
      type,
      size,
      mtime,
      ext: (path.win32.extname(ent.name) || '').toLowerCase()
    });
  }
  items.sort((a, b) => {
    if (a.type === 'dir' && b.type !== 'dir') return -1;
    if (b.type === 'dir' && a.type !== 'dir') return 1;
    return a.name.localeCompare(b.name, 'zh', { numeric: true });
  });

  sendJson(res, 200, { ok: true, path: dir, items });
}

async function statPath(res, q, cfg) {
  const file = resolveAllowedPath(q.path, cfg);
  const s = await fs.promises.stat(file);
  sendJson(res, 200, {
    ok: true, path: file,
    type: s.isDirectory() ? 'dir' : s.isFile() ? 'file' : 'other',
    size: s.size, mtime: s.mtimeMs, ctime: s.ctimeMs,
    ext: (path.win32.extname(file) || '').toLowerCase()
  });
}

// —— 文本解码：识别 BOM（UTF-8 / UTF-16 LE / UTF-16 BE），修复 UTF-16 文本被误判为 binary ——
function utf16beToString(buf) {
  // 手动交换字节序后按 utf16le 解码（不依赖 ICU 的 utf-16be 支持）
  const swapped = Buffer.allocUnsafe(buf.length - (buf.length % 2));
  for (let i = 0; i + 1 < buf.length; i += 2) {
    swapped[i] = buf[i + 1];
    swapped[i + 1] = buf[i];
  }
  return swapped.toString('utf16le');
}

// 合理性检查：解码结果中控制字符（除常见空白）占比过高 => 不是文本
function looksLikeText(str) {
  if (!str.length) return false;
  let bad = 0;
  for (let i = 0; i < str.length; i++) {
    const c = str.charCodeAt(i);
    if (c === 9 || c === 10 || c === 13 || c === 11 || c === 12) continue; // \t \n \r \v \f
    if (c === 0 || c < 32 || c === 127 || c === 0xFFFE || c === 0xFFFF) bad++;
  }
  return bad / str.length < 0.05;
}

function decodeTextBuffer(buf) {
  // 1) BOM 检测
  if (buf.length >= 3 && buf[0] === 0xEF && buf[1] === 0xBB && buf[2] === 0xBF) {
    return { encoding: 'utf-8', content: buf.slice(3).toString('utf8') };
  }
  if (buf.length >= 2 && buf[0] === 0xFF && buf[1] === 0xFE) {
    // FF FE 00 00 是 UTF-32 LE BOM，不支持，按二进制处理
    if (buf.length >= 4 && buf[2] === 0x00 && buf[3] === 0x00) throw new Error('binary file');
    return { encoding: 'utf-16le', content: buf.slice(2).toString('utf16le') };
  }
  if (buf.length >= 2 && buf[0] === 0xFE && buf[1] === 0xFF) {
    return { encoding: 'utf-16be', content: utf16beToString(buf.slice(2)) };
  }

  // 2) 无 BOM：扫描前 4096 字节的 null 分布
  const n = Math.min(buf.length, 4096);
  let nulls = 0, evenNulls = 0, oddNulls = 0;
  for (let i = 0; i < n; i++) {
    if (buf[i] === 0) {
      nulls++;
      if (i % 2 === 0) evenNulls++; else oddNulls++;
    }
  }
  // 完全没有 null 字节 => 按 UTF-8 解码，若结果含替换字符(U+FFFD)则回退 GB18030
  if (nulls === 0) {
    const utf8 = buf.toString('utf8');
    if (!utf8.includes('\uFFFD')) return { encoding: 'utf-8', content: utf8 };
    // UTF-8 解码失败（含替换字符）→ 尝试 GB18030（兼容 GBK/GB2312）
    try {
      const gb = new TextDecoder('gb18030', { fatal: false }).decode(buf);
      if (looksLikeText(gb) && (gb.match(/\uFFFD/g) || []).length < (utf8.match(/\uFFFD/g) || []).length) {
        return { encoding: 'gb18030', content: gb };
      }
    } catch (_) {}
    return { encoding: 'utf-8', content: utf8 };
  }

  // 有 null 字节：UTF-16 文本的 ASCII 字符后跟 0x00，先按 UTF-16 启发式解码再验证
  // 奇数位 null 多 => UTF-16 LE；偶数位 null 多 => UTF-16 BE
  const tryDecode = (enc) => {
    if (enc === 'utf-16le') {
      const b = buf.slice(0, buf.length - (buf.length % 2));
      return b.toString('utf16le');
    }
    return utf16beToString(buf);
  };
  const candidates = oddNulls >= evenNulls
    ? ['utf-16le', 'utf-16be']
    : ['utf-16be', 'utf-16le'];
  for (const enc of candidates) {
    const str = tryDecode(enc);
    if (str.length && looksLikeText(str)) return { encoding: enc, content: str };
  }
  // UTF-16 也解不出 → 最后试 GB18030（GBK 混合少量 null 字节的情况）
  try {
    const gb = new TextDecoder('gb18030', { fatal: false }).decode(buf);
    if (looksLikeText(gb)) return { encoding: 'gb18030', content: gb };
  } catch (_) {}
  // 确实无法解码为合理文本 => 才报 binary
  throw new Error('binary file');
}

async function readTextFile(res, q, cfg) {
  const file = resolveAllowedPath(q.path, cfg);
  const extCheck = checkExtension(file, cfg);
  if (!extCheck.ok) return sendJson(res, 403, { ok: false, error: extCheck.reason });
  const s = await fs.promises.stat(file);
  if (!s.isFile()) throw new Error('not a file');
  const MAX = 5 * 1024 * 1024;
  if (s.size > MAX) throw new Error('file too large (max ' + (MAX / 1048576) + 'MB)');
  const buf = await fs.promises.readFile(file);
  const decoded = decodeTextBuffer(buf);
  sendJson(res, 200, {
    ok: true, path: file, content: decoded.content, encoding: decoded.encoding
  });
}

// —— PSD 缩略图提取：解析 Image Resources 区段的 Resource ID 1033/1036 ——
function extractPsdThumbnail(buf) {
  if (buf.length < 26) return null;
  // PSD 签名 '8BPS'
  if (buf[0] !== 0x38 || buf[1] !== 0x42 || buf[2] !== 0x50 || buf[3] !== 0x53) return null;
  let offset = 26;
  // 跳过 Color Mode Data
  if (offset + 4 > buf.length) return null;
  const colorModeLen = buf.readUInt32BE(offset);
  offset += 4 + colorModeLen;
  // 读 Image Resources 区段
  if (offset + 4 > buf.length) return null;
  const resourcesLen = buf.readUInt32BE(offset);
  offset += 4;
  const resourcesEnd = Math.min(offset + resourcesLen, buf.length);
  while (offset + 12 < resourcesEnd) {
    // '8BIM' 签名
    if (buf[offset] !== 0x38 || buf[offset + 1] !== 0x42 || buf[offset + 2] !== 0x49 || buf[offset + 3] !== 0x4D) break;
    offset += 4;
    const resourceId = buf.readUInt16BE(offset);
    offset += 2;
    // Pascal string name（1 字节长度 + 名称，补偶数）
    const nameLen = buf[offset];
    offset += 1 + nameLen;
    if ((nameLen + 1) % 2 !== 0) offset++;
    // 数据大小
    if (offset + 4 > buf.length) break;
    const dataSize = buf.readUInt32BE(offset);
    offset += 4;
    // Resource ID 0x0409 (1033) 或 0x040C (1036) = 缩略图
    if (resourceId === 0x0409 || resourceId === 0x040C) {
      if (offset + 24 > buf.length) break;
      const format = buf.readUInt32BE(offset); // 0=JPEG, 1=raw RGB
      if (format === 0) {
        // 跳过 format(4)+width(4)+height(4)+widthBytes(4)+totalSize(4)+compressedSize(4)=24
        const jpegEnd = Math.min(offset + dataSize, buf.length);
        return buf.slice(offset + 24, jpegEnd);
      }
    }
    offset += dataSize;
    if (dataSize % 2 !== 0) offset++;
  }
  return null;
}

// —— PSD 复合图像提取：解析尾部 Image Data 区段（8-bit RGB/RGBA, raw/RLE）——
function extractPsdComposite(buf) {
  if (buf.length < 30) return null;
  if (buf[0] !== 0x38 || buf[1] !== 0x42 || buf[2] !== 0x50 || buf[3] !== 0x53) return null;
  let offset = 26;
  // 跳过 Color Mode Data
  const cmLen = buf.readUInt32BE(offset); offset += 4 + cmLen;
  // 跳过 Image Resources
  const irLen = buf.readUInt32BE(offset); offset += 4 + irLen;
  // 跳过 Layer & Mask Info
  const lmLen = buf.readUInt32BE(offset); offset += 4 + lmLen;
  // Image Data 区段
  if (offset + 2 > buf.length) return null;
  const compression = buf.readUInt16BE(offset); offset += 2;
  const height = buf.readUInt32BE(14);  // header: height at offset 14
  const width = buf.readUInt32BE(18);   // header: width at offset 18
  const channels = buf.readUInt16BE(12); // header: channels at offset 12
  const depth = buf.readUInt16BE(22);    // header: depth at offset 22
  if (depth !== 8 && depth !== 16) return null;
  const bytesPerSample = depth / 8;
  if (width === 0 || height === 0 || width > 30000 || height > 30000) return null;
  const pxCount = width * height;
  const planeSize = pxCount * bytesPerSample;

  // 解码为 RGBA
  const rgba = Buffer.alloc(pxCount * 4);

  if (compression === 0) {
    // Raw: 按通道顺序存储（R平面, G平面, B平面, A平面）
    for (let ch = 0; ch < Math.min(channels, 4); ch++) {
      const chStart = offset + ch * planeSize;
      for (let i = 0; i < pxCount; i++) {
        const byteOff = chStart + i * bytesPerSample;
        if (byteOff < buf.length) {
          rgba[i * 4 + ch] = depth === 16 ? buf[byteOff] : buf[byteOff]; // 16-bit 取高字节
        }
      }
    }
  } else if (compression === 1) {
    // RLE (PackBits): 先读行长度表，再解压数据
    const rowsPerChannel = height;
    const rowCount = rowsPerChannel * channels;
    const countsStart = offset;
    // 读每行压缩长度
    const rowLengths = [];
    for (let r = 0; r < rowCount; r++) {
      rowLengths.push(buf.readUInt16BE(countsStart + r * 2));
    }
    let dataStart = countsStart + rowCount * 2;
    // 按通道解压每行
    for (let ch = 0; ch < Math.min(channels, 4); ch++) {
      let dataOff = dataStart;
      for (let r = 0; r < rowsPerChannel; r++) {
        const rowLen = rowLengths[ch * rowsPerChannel + r];
        // PackBits 解压一行
        const row = Buffer.alloc(width);
        let ri = 0;
        let end = Math.min(dataOff + rowLen, buf.length);
        while (dataOff < end && ri < width) {
          const n = buf.readInt8(dataOff++);
          if (n >= 0) {
            // literal: n+1 bytes
            const count = n + 1;
            for (let k = 0; k < count && ri < width && dataOff < end; k++) {
              row[ri++] = buf[dataOff++];
            }
          } else if (n !== -128) {
            // repeat: 1-n times
            const count = 1 - n;
            const val = dataOff < end ? buf[dataOff++] : 0;
            for (let k = 0; k < count && ri < width; k++) {
              row[ri++] = val;
            }
          }
          // n === -128: no-op
        }
        // 写入 RGBA
        for (let x = 0; x < width; x++) {
          rgba[(r * width + x) * 4 + ch] = row[x];
        }
        dataOff = dataStart + rowLengths.slice(0, ch * rowsPerChannel + r + 1).reduce((a, b) => a + b, 0);
      }
    }
  } else if (compression === 2 || compression === 3) {
    // ZIP (deflate): 直接解压整个 Image Data，再按通道平面处理
    const zlib = require('zlib');
    let decompressed;
    try {
      decompressed = zlib.inflateRawSync(buf.slice(offset));
    } catch (e1) {
      try { decompressed = zlib.inflateSync(buf.slice(offset)); } catch (e2) { return null; }
    }
    if (decompressed.length < pxCount * Math.min(channels, 4)) return null;
    // compression 3 = ZIP + prediction（每行每字节是前一字节的增量）
    if (compression === 3) {
      for (let ch = 0; ch < Math.min(channels, 4); ch++) {
        for (let r = 0; r < height; r++) {
          const rowStart = ch * pxCount + r * width;
          for (let x = 1; x < width; x++) {
            decompressed[rowStart + x] = (decompressed[rowStart + x - 1] + decompressed[rowStart + x]) & 0xFF;
          }
        }
      }
    }
    for (let ch = 0; ch < Math.min(channels, 4); ch++) {
      for (let i = 0; i < pxCount; i++) {
        rgba[i * 4 + ch] = decompressed[ch * planeSize + i * bytesPerSample];
      }
    }
  } else {
    return null;
  }

  // 确保 Alpha 通道存在
  if (channels < 4) {
    for (let i = 0; i < pxCount; i++) rgba[i * 4 + 3] = 255;
  }
  return { width, height, rgba };
}

// —— ZIP 压缩包浏览：列出内容 ——
// ============ 多格式压缩包支持 ============
// 文件名编码解码：优先 UTF-8，有替换字符则回退 GB18030（中文 Windows 常见编码）
function decodeArchiveName(buf) {
  if (!buf || buf.length === 0) return '';
  const bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  const utf8 = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  const bad = (utf8.match(/\uFFFD/g) || []).length;
  if (bad === 0) return utf8;
  try {
    return new TextDecoder('gb18030', { fatal: false }).decode(bytes);
  } catch (_) { return utf8; }
}

function detectArchiveFormat(filePath) {
  const ext = path.extname(filePath).toLowerCase().replace('.', '');
  if (ext === 'zip' || ext === 'jar') return 'zip';
  if (ext === 'tar') return 'tar';
  if (ext === 'gz' || ext === 'tgz') return 'gz';
  if (ext === 'rar' || ext === '7z' || ext === 'bz2' || ext === 'xz') return 'sevenzip';
  return null;
}

// TAR 解析器：解析 512 字节文件头（USTAR + GNU 扩展）
function parseTarEntries(buf) {
  const entries = [];
  let offset = 0;
  let longName = null; // GNU 长文件名缓冲
  let paxName = null;  // PaxHeaders 真实文件名缓冲
  while (offset + 512 <= buf.length) {
    const header = buf.slice(offset, offset + 512);
    let allZero = true;
    for (let i = 0; i < 512; i++) { if (header[i] !== 0) { allZero = false; break; } }
    if (allZero) break;
    // 文件名 (0-99)：编码自适应
    let name = decodeArchiveName(header.slice(0, 100)).replace(/\0.*$/, '');
    // USTAR 前缀 (345-499) 用于长路径
    const prefix = decodeArchiveName(header.slice(345, 500)).replace(/\0.*$/, '');
    if (prefix) name = prefix + '/' + name;
    // 文件大小 (124-135, 八进制)
    const sizeStr = header.slice(124, 136).toString('ascii').replace(/\0.*$/, '').trim();
    const size = parseInt(sizeStr, 8) || 0;
    // 类型标志 (156)：'5'=目录, '0'/'\0'=普通文件, 'L'=GNU长文件名, 'K'=GNU长链接名, 'x'/'g'=PaxHeaders
    const typeFlag = String.fromCharCode(header[156]);
    const dataStart = offset + 512;
    // GNU 长文件名（type 'L'）：文件名存储在数据区，应用到下一个条目
    if (typeFlag === 'L') {
      longName = decodeArchiveName(buf.slice(dataStart, dataStart + size)).replace(/\0.*$/, '');
      offset = dataStart + Math.ceil(size / 512) * 512;
      continue;
    }
    // PaxHeaders（type 'x'/'g'）：解析 path= 真实文件名，应用到下一个条目
    if (typeFlag === 'x' || typeFlag === 'g') {
      const data = buf.slice(dataStart, dataStart + size);
      const text = decodeArchiveName(data);
      // PaxHeaders 格式: "LENGTH key=value\n" 每行
      for (const line of text.split('\n')) {
        const eqIdx = line.indexOf('=');
        if (eqIdx < 0) continue;
        const spaceIdx = line.indexOf(' ');
        if (spaceIdx < 0 || spaceIdx > eqIdx) continue;
        const key = line.slice(spaceIdx + 1, eqIdx);
        if (key === 'path') {
          paxName = line.slice(eqIdx + 1).replace(/\r$/, '');
        }
      }
      offset = dataStart + Math.ceil(size / 512) * 512;
      continue;
    }
    // 应用 GNU 长文件名或 PaxHeaders 真实文件名
    if (longName) { name = longName; longName = null; }
    if (paxName) { name = paxName; paxName = null; }
    const isDir = typeFlag === '5' || name.endsWith('/');
    if (name !== '.' && name !== './' && name !== '' && !name.startsWith('PaxHeaders.')) {
      entries.push({ name: name, size: isDir ? 0 : size, type: isDir ? 'dir' : 'file', dataOffset: dataStart });
    }
    offset = dataStart + Math.ceil(size / 512) * 512;
  }
  return entries;
}

function extractTarEntry(buf, entryPath) {
  const norm = entryPath.replace(/\\/g, '/').replace(/^\/+/, '');
  const entries = parseTarEntries(buf);
  const target = entries.find(e => e.name === norm && e.type === 'file');
  if (!target) throw new Error('entry not found: ' + entryPath);
  return buf.slice(target.dataOffset, target.dataOffset + target.size);
}

// 7z CLI 集成（RAR / 7Z / 其他 7z 支持的格式）
function get7zExe() {
  const known = ['C:\\Program Files\\7-Zip\\7z.exe', 'C:\\Program Files (x86)\\7-Zip\\7z.exe'];
  for (const p of known) { try { if (fs.existsSync(p)) return p; } catch (_) {} }
  return '7z';
}

function run7z(args) {
  return new Promise((resolve, reject) => {
    const exe = get7zExe();
    const proc = spawn(exe, args, { windowsHide: true });
    const chunks = [], errChunks = [];
    proc.stdout.on('data', d => chunks.push(d));
    proc.stderr.on('data', d => errChunks.push(d));
    proc.on('close', code => {
      // 7z exit code 1 = warning (非致命)
      if (code !== 0 && code !== 1) return reject(new Error('7z exit ' + code + ': ' + Buffer.concat(errChunks).toString('utf-8')));
      resolve(Buffer.concat(chunks));
    });
    proc.on('error', reject);
  });
}

async function list7zEntries(filePath) {
  const out = await run7z(['l', '-slt', filePath]);
  // 7z 输出在中文 Windows 下可能是 GBK/GB18030 编码
  const text = decodeArchiveName(out);
  const entries = [];
  // ---------- 只在头部和条目之间出现一次；条目之间用空行分隔
  const sepIdx = text.indexOf('----------');
  const afterSep = sepIdx >= 0 ? text.slice(sepIdx + 10) : text;
  // 按空行分隔条目（兼容 \r\n 和 \n）
  const sections = afterSep.split(/\r?\n\s*\r?\n/);
  for (const section of sections) {
    const lines = section.split(/\r?\n/);
    let p = '', size = 0, isDir = false;
    for (const line of lines) {
      if (line.startsWith('Path = ')) p = line.slice(7);
      else if (line.startsWith('Size = ')) size = parseInt(line.slice(7)) || 0;
      else if (line.startsWith('Attributes = ')) isDir = line.slice(13, 14) === 'D';
    }
    if (p) entries.push({ name: p, size: isDir ? 0 : size, type: isDir ? 'dir' : 'file' });
  }
  return entries;
}

async function extract7zEntry(filePath, entryPath) {
  return await run7z(['x', filePath, '-so', entryPath]);
}

// 统一提取接口：按格式分发
async function extractEntryFromArchive(filePath, entryPath) {
  const format = detectArchiveFormat(filePath);
  if (!format) throw new Error('unsupported archive format');
  if (format === 'zip') {
    const buf = await fs.promises.readFile(filePath);
    const zip = await JSZip.loadAsync(buf);
    const target = zip.file(entryPath);
    if (!target) throw new Error('entry not found: ' + entryPath);
    return await target.async('nodebuffer');
  }
  if (format === 'tar') {
    const buf = await fs.promises.readFile(filePath);
    return extractTarEntry(buf, entryPath);
  }
  if (format === 'gz') {
    const buf = await fs.promises.readFile(filePath);
    const decompressed = zlib.gunzipSync(buf);
    if (decompressed.length >= 262 && decompressed.slice(257, 262).toString('ascii') === 'ustar') {
      return extractTarEntry(decompressed, entryPath);
    }
    return decompressed; // 单文件 GZ
  }
  if (format === 'sevenzip') {
    return await extract7zEntry(filePath, entryPath);
  }
  throw new Error('unsupported archive format');
}

// 平铺条目 → 虚拟目录树（共用逻辑）
function buildArchiveItems(rawEntries, subdir) {
  const items = [];
  const seen = {};
  for (const entry of rawEntries) {
    const norm = entry.name.replace(/\\/g, '/');
    if (subdir && !norm.startsWith(subdir + '/')) continue;
    const rel = subdir ? norm.slice(subdir.length + 1) : norm;
    const seg = rel.split('/')[0];
    if (!seg || seg === '.' || seg === '..' || seen[seg]) continue;
    seen[seg] = true;
    const isDir = entry.type === 'dir' || rel.endsWith('/') || rel.indexOf('/') > 0;
    const fullPath = subdir ? subdir + '/' + seg : seg;
    items.push({ name: seg, type: isDir ? 'dir' : 'file', entry: isDir ? fullPath + '/' : fullPath, size: isDir ? 0 : (entry.size || 0) });
  }
  items.sort((a, b) => {
    if (a.type !== b.type) return a.type === 'dir' ? -1 : 1;
    return a.name.localeCompare(b.name);
  });
  return items;
}

async function archiveList(res, q, cfg) {
  const file = resolveAllowedPath(q.path, cfg);
  const s = await fs.promises.stat(file);
  if (!s.isFile()) throw new Error('not a file');
  const format = detectArchiveFormat(file);
  if (!format) throw new Error('unsupported archive format');
  const subdir = (q.subdir || '').replace(/\\/g, '/').replace(/^\/+|\/+$/g, '');
  let rawEntries = [];
  if (format === 'zip') {
    const buf = await fs.promises.readFile(file);
    const zip = await JSZip.loadAsync(buf);
    for (const name of Object.keys(zip.files)) {
      const entry = zip.files[name];
      const norm = name.replace(/\\/g, '/');
      rawEntries.push({ name: norm, size: entry._data ? entry._data.uncompressedSize : 0, type: (entry.dir || norm.endsWith('/')) ? 'dir' : 'file' });
    }
  } else if (format === 'tar') {
    const buf = await fs.promises.readFile(file);
    rawEntries = parseTarEntries(buf).map(e => ({ name: e.name, size: e.size, type: e.type }));
  } else if (format === 'gz') {
    const buf = await fs.promises.readFile(file);
    let decompressed;
    try { decompressed = zlib.gunzipSync(buf); } catch (e) { throw new Error('gzip decompress failed: ' + e.message); }
    if (decompressed.length >= 262 && decompressed.slice(257, 262).toString('ascii') === 'ustar') {
      rawEntries = parseTarEntries(decompressed).map(e => ({ name: e.name, size: e.size, type: e.type }));
    } else {
      const baseName = path.basename(file).replace(/\.t?gz$/i, '') || 'content';
      rawEntries = [{ name: baseName, size: decompressed.length, type: 'file' }];
    }
  } else if (format === 'sevenzip') {
    rawEntries = await list7zEntries(file);
  }
  const items = buildArchiveItems(rawEntries, subdir);
  sendJson(res, 200, { ok: true, subdir: subdir, items });
}

// —— 多格式压缩包提取单个文件 ——
async function archiveExtract(req, res, q, cfg) {
  const file = resolveAllowedPath(q.path, cfg);
  const entry = String(q.entry || '');
  if (!entry) throw new Error('missing entry');
  for (const p of [entry, file]) {
    const extCheck = checkExtension(p, cfg);
    if (!extCheck.ok) return sendJson(res, 403, { ok: false, error: extCheck.reason });
  }
  const data = await extractEntryFromArchive(file, entry);
  // 猜测 MIME
  const ext = path.extname(entry).toLowerCase();
  const mimeMap = {
    '.txt': 'text/plain', '.html': 'text/html', '.css': 'text/css',
    '.js': 'application/javascript', '.json': 'application/json',
    '.png': 'image/png', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.jpe': 'image/jpeg', '.jfif': 'image/jpeg',
    '.gif': 'image/gif', '.svg': 'image/svg+xml', '.pdf': 'application/pdf',
    '.bmp': 'image/bmp', '.ico': 'image/x-icon', '.webp': 'image/webp', '.apng': 'image/apng',
    '.tif': 'image/tiff', '.tiff': 'image/tiff', '.avif': 'image/avif',
    '.mp4': 'video/mp4', '.mp3': 'audio/mpeg', '.wav': 'audio/wav',
  };
  const mime = mimeMap[ext] || 'application/octet-stream';
  res.writeHead(200, {
    'Content-Type': mime,
    'Content-Length': data.length,
    'Content-Disposition': 'inline; filename="' + path.basename(entry).replace(/[^\x20-\x7e]/g, '_') + '"; filename*=UTF-8\'\'' + encodeURIComponent(path.basename(entry))
  });
  res.end(data);
}
// —— PPT (OLE2) 文本提取：扫描二进制中的 UTF-16LE / ASCII 文本串 ——
function extractPptText(buf) {
  const strings = [];
  const MIN_LEN = 4;

  // 扫描 UTF-16LE 文本（中文、英文混排）
  let text = '';
  for (let i = 0; i < buf.length - 1; i += 2) {
    const code = buf.readUInt16LE(i);
    const isPrintable =
      (code >= 0x20 && code <= 0x7e) ||   // ASCII
      (code >= 0xa0 && code <= 0xff) ||   // Latin-1
      (code >= 0x3000 && code <= 0x9fff) || // CJK
      (code >= 0xff00 && code <= 0xffef) || // Fullwidth
      code === 0x0a || code === 0x0d;
    if (isPrintable) {
      text += String.fromCharCode(code);
    } else {
      if (text.length >= MIN_LEN) strings.push(text.trim());
      text = '';
    }
  }
  if (text.length >= MIN_LEN) strings.push(text.trim());

  // 去重 + 过滤噪声（纯数字/符号、单行重复）
  const seen = new Set();
  const filtered = strings.filter(s => {
    if (!s || s.length < MIN_LEN) return false;
    // 过滤纯二进制噪声：要求至少 50% 可读字符
    let readable = 0;
    for (const ch of s) {
      const c = ch.charCodeAt(0);
      if ((c >= 0x30 && c <= 0x39) || (c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a) || (c >= 0x4e00 && c <= 0x9fff)) readable++;
    }
    if (readable / s.length < 0.5) return false;
    if (seen.has(s)) return false;
    seen.add(s);
    return true;
  });

  return filtered.length > 0 ? [filtered.join('\n')] : ['(未能提取文本内容，请用系统程序打开)'];
}

async function officePreview(res, q, cfg) {
  // 支持 :: 虚拟路径（压缩包内文件）
  let file, buf;
  const zipSep = String(q.path || '').indexOf('::');
  if (zipSep > 0) {
    const archivePath = q.path.substring(0, zipSep);
    const entry = q.path.substring(zipSep + 2);
    file = resolveAllowedPath(archivePath, cfg);
    buf = await extractEntryFromArchive(file, entry);
    if (!buf) throw new Error('entry not found: ' + entry);
  } else {
    file = resolveAllowedPath(q.path, cfg);
    const s = await fs.promises.stat(file);
    if (!s.isFile()) throw new Error('not a file');
    buf = await fs.promises.readFile(file);
  }
  const ext = path.extname(zipSep > 0 ? q.path.substring(zipSep + 2) : file).toLowerCase();

  if (ext === '.docx') {
    const r = await mammoth.convertToHtml({ buffer: buf });
    return sendJson(res, 200, { ok: true, type: 'html', html: r.value, ext: 'docx' });
  }

  if (ext === '.doc') {
    try {
      const extractor = new WordExtractor();
      const doc = await extractor.extract(buf);
      const text = doc.getBody();
      const escaped = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      const html = '<pre style="white-space:pre-wrap;font-family:sans-serif;line-height:1.8;max-width:100%">' + escaped + '</pre>';
      return sendJson(res, 200, { ok: true, type: 'html', html, ext: 'doc' });
    } catch (e) {
      return sendJson(res, 200, { ok: false, error: 'DOC parse error: ' + e.message });
    }
  }

  if (ext === '.xlsx' || ext === '.xls') {
    const wb = XLSX.read(buf, { type: 'buffer' });
    const sheets = {};
    for (const name of wb.SheetNames) {
      sheets[name] = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, defval: '' });
    }
    return sendJson(res, 200, { ok: true, type: 'xlsx', sheets, names: wb.SheetNames, ext: 'xlsx' });
  }

  if (ext === '.pptx') {
    const zip = await JSZip.loadAsync(buf);
    const slideFiles = Object.keys(zip.files)
      .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
      .sort((a, b) => {
        const na = parseInt(a.match(/slide(\d+)/)[1], 10);
        const nb = parseInt(b.match(/slide(\d+)/)[1], 10);
        return na - nb;
      });
    const slides = [];
    for (const sf of slideFiles) {
      const xml = await zip.files[sf].async('string');
      const texts = [];
      const re = /<a:t[^>]*>([^<]*)<\/a:t>/g;
      let m;
      while ((m = re.exec(xml)) !== null) {
        if (m[1].trim()) texts.push(m[1]);
      }
      slides.push(texts.join(' '));
    }
    return sendJson(res, 200, { ok: true, type: 'pptx', slides, ext: 'pptx' });
  }

  if (ext === '.ppt') {
    try {
      // PPT (OLE2) 无现成 JS 解析库，从二进制扫描提取文本
      const slides = extractPptText(buf);
      return sendJson(res, 200, { ok: true, type: 'pptx', slides, ext: 'ppt' });
    } catch (e) {
      return sendJson(res, 200, { ok: false, error: 'PPT parse error: ' + e.message });
    }
  }

  // —— AI 文件：PDF 兼容型 → PDF 预览；旧版 PostScript → 提示系统打开 ——
  if (ext === '.ai') {
    if (buf.length >= 5 && buf[0] === 0x25 && buf[1] === 0x50 && buf[2] === 0x44 && buf[3] === 0x46 && buf[4] === 0x2D) {
      // %PDF- 头：PDF 兼容 AI，返回 pdf 标记让前端用 iframe 预览
      return sendJson(res, 200, { ok: true, type: 'pdf', ext });
    }
    if (buf.length >= 2 && buf[0] === 0x25 && buf[1] === 0x21) {
      // %!PS-Adobe 头：旧版 PostScript AI，无法网页预览
      return sendJson(res, 200, { ok: true, type: 'fallback', message: 'PostScript 格式 AI 文件，需用 Illustrator 打开', ext });
    }
    // 未知格式：尝试按 PSD 处理（某些 AI 可能内嵌 PSD 数据）
  }

  // —— PSD / PSB：提取内嵌 JPEG 缩略图 → 复合图像兜底 ——
  if (ext === '.psd' || ext === '.psb') {
    const thumb = extractPsdThumbnail(buf);
    if (thumb) {
      return sendJson(res, 200, { ok: true, type: 'image', dataUrl: 'data:image/jpeg;base64,' + thumb.toString('base64'), ext });
    }
    // 缩略图提取失败 → 尝试提取复合图像
    const composite = extractPsdComposite(buf);
    if (composite) {
      return sendJson(res, 200, { ok: true, type: 'raw-image', width: composite.width, height: composite.height, data: composite.rgba.toString('base64'), ext });
    }
    throw new Error('cannot extract image from PSD (compression=' + buf.readUInt16BE(38) + ' depth=' + buf.readUInt16BE(24) + ' ch=' + buf.readUInt16BE(22) + ' w=' + buf.readUInt32BE(18) + ' h=' + buf.readUInt32BE(14) + ' len=' + buf.length + ')');
  }

  // —— AI 文件：PDF 兼容检测 ——
  if (ext === '.ai') {
    // 现代 AI 文件 (CS+) 是 PDF 兼容格式
    if (buf.length > 5 && buf.slice(0, 5).toString('ascii') === '%PDF-') {
      return sendJson(res, 200, { ok: true, type: 'pdf', ext: 'ai' });
    }
    // 旧版 PostScript AI → 尝试提取内嵌图像
    throw new Error('legacy PostScript AI file (not PDF-compatible), use system handler');
  }

  throw new Error('unsupported office format: ' + ext);
}

function streamFile(req, res, q, cfg) {
  return new Promise(async (resolve, reject) => {
    try {
      const file = resolveAllowedPath(q.path, cfg);
      const extCheck = checkExtension(file, cfg);
      if (!extCheck.ok) {
        sendJson(res, 403, { ok: false, error: extCheck.reason });
        return resolve();
      }
      const st = await fs.promises.stat(file);
      if (!st.isFile()) throw new Error('not a file');

      const ext  = path.win32.extname(file).toLowerCase().slice(1);
      const mime = getMime(ext);
      const dl   = q.download === '1' || q.download === 'true';
      const fname = path.win32.basename(file);

      const headers = {
        'Content-Type': mime,
        'Accept-Ranges': 'bytes',
        'Cache-Control': 'no-cache'
      };
      if (dl) {
        headers['Content-Disposition'] = "attachment; filename*=UTF-8''" + encodeURIComponent(fname);
      } else {
        const asciiName = fname.replace(/[^\x20-\x7e]/g, '_');
        headers['Content-Disposition'] = "inline; filename=\"" + asciiName + "\"; filename*=UTF-8''" + encodeURIComponent(fname);
      }

      const range = req.headers.range;
      if (range) {
        const m = /^bytes=(\d*)-(\d*)$/.exec(range);
        if (m) {
          let start = m[1] ? parseInt(m[1], 10) : 0;
          let end   = m[2] ? parseInt(m[2], 10) : st.size - 1;
          if (isNaN(start)) start = 0;
          if (isNaN(end) || end >= st.size) end = st.size - 1;
          if (start > end || start >= st.size) {
            res.writeHead(416, { 'Content-Range': 'bytes */' + st.size });
            res.end();
            return resolve();
          }
          headers['Content-Range']  = 'bytes ' + start + '-' + end + '/' + st.size;
          headers['Content-Length'] = end - start + 1;
          res.writeHead(206, headers);
          const rs = fs.createReadStream(file, { start, end });
          rs.on('error', reject);
          rs.on('end', resolve);
          res.on('close', () => { rs.destroy(); resolve(); }); // 客户端中断时销毁流，避免挂住连接
          rs.pipe(res);
          return;
        }
      }

      headers['Content-Length'] = st.size;
      res.writeHead(200, headers);
      if (req.method === 'HEAD') { res.end(); return resolve(); }
      const rs = fs.createReadStream(file);
      rs.on('error', reject);
      rs.on('end', resolve);
      res.on('close', () => { rs.destroy(); resolve(); }); // 客户端中断时销毁流，避免挂住连接
      rs.pipe(res);
    } catch (e) { reject(e); }
  });
}

function getMime(ext) {
  const m = {
    txt: 'text/plain; charset=utf-8',
    html: 'text/html; charset=utf-8', htm: 'text/html; charset=utf-8',
    css: 'text/css; charset=utf-8',
    js: 'application/javascript; charset=utf-8',
    mjs: 'application/javascript; charset=utf-8',
    json: 'application/json; charset=utf-8',
    xml: 'application/xml; charset=utf-8',
    svg: 'image/svg+xml',
    md: 'text/markdown; charset=utf-8',
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', jpe: 'image/jpeg', jfif: 'image/jpeg',
    apng: 'image/apng', tif: 'image/tiff', tiff: 'image/tiff', avif: 'image/avif',
    heic: 'image/heic', heif: 'image/heif', dib: 'image/bmp', jp2: 'image/jp2',
    gif: 'image/gif', webp: 'image/webp', bmp: 'image/bmp', ico: 'image/x-icon',
    mp4: 'video/mp4', webm: 'video/webm', mkv: 'video/x-matroska',
    mov: 'video/quicktime', avi: 'video/x-msvideo',
    mp3: 'audio/mpeg', wav: 'audio/wav', flac: 'audio/flac',
    m4a: 'audio/mp4', ogg: 'audio/ogg', aac: 'audio/aac',
    pdf: 'application/pdf',
    zip: 'application/zip', '7z': 'application/x-7z-compressed',
    rar: 'application/vnd.rar', gz: 'application/gzip', tar: 'application/x-tar'
  };
  return m[ext] || 'application/octet-stream';
}

async function mkdirOp(res, data, cfg) {
  const dir = resolveAllowedPath(data.path, cfg);
  const name = String(data.name || '').trim();
  if (!name || /[\\\/:*?"<>|]/.test(name) || name === '.' || name === '..') {
    throw new Error('invalid folder name');
  }
  const full = path.win32.join(dir, name);
  await fs.promises.mkdir(full);
  sendJson(res, 200, { ok: true, path: full });
}

async function renameOp(res, data, cfg) {
  const src = resolveAllowedPath(data.path, cfg);
  const name = String(data.newName || '').trim();
  if (!name || /[\\\/:*?"<>|]/.test(name) || name === '.' || name === '..') {
    throw new Error('invalid name');
  }
  const dst = path.win32.join(path.win32.dirname(src), name);
  if (dst === src) return sendJson(res, 200, { ok: true, path: dst });
  const st = await fs.promises.stat(src);
  if (!st.isDirectory()) {
    const extCheck = checkExtension(dst, cfg);
    if (!extCheck.ok) return sendJson(res, 403, { ok: false, error: extCheck.reason });
  }
  await fs.promises.rename(src, dst);
  sendJson(res, 200, { ok: true, path: dst });
}

async function rmOp(res, data, cfg) {
  const target = resolveAllowedPath(data.path, cfg);
  const st = await fs.promises.stat(target);
  if (st.isDirectory()) {
    if (!data.recursive) throw new Error('directory requires recursive=true');
    await fs.promises.rm(target, { recursive: true, force: false });
  } else {
    await fs.promises.unlink(target);
  }
  sendJson(res, 200, { ok: true });
}

async function mvOp(res, data, cfg) {
  const src = resolveAllowedPath(data.from, cfg);
  const dst = resolveAllowedPath(data.to, cfg);
  try {
    await fs.promises.rename(src, dst);
  } catch (e) {
    // 跨盘移动（EXDEV）rename 会失败，降级为 复制 + 删除
    if (e && e.code === 'EXDEV') {
      await copyRecursive(src, dst);
      await fs.promises.rm(src, { recursive: true, force: true });
    } else {
      throw e;
    }
  }
  sendJson(res, 200, { ok: true });
}

// 递归复制文件 / 目录（POST /api/cp {from, to}）
async function cpOp(res, data, cfg) {
  const src = resolveAllowedPath(data.from, cfg);
  const dst = resolveAllowedPath(data.to, cfg);
  if (path.win32.normalize(src).toLowerCase() === path.win32.normalize(dst).toLowerCase()) {
    throw new Error('source and destination are the same');
  }
  await copyRecursive(src, dst);
  sendJson(res, 200, { ok: true });
}

async function copyRecursive(src, dst) {
  const st = await fs.promises.stat(src);
  if (st.isDirectory()) {
    // 防止把目录复制进自身内部形成死循环
    const srcPrefix = path.win32.normalize(src).toLowerCase() + path.sep;
    if (path.win32.normalize(dst).toLowerCase().startsWith(srcPrefix)) {
      throw new Error('cannot copy a directory into itself');
    }
    await fs.promises.mkdir(dst, { recursive: true });
    const ents = await fs.promises.readdir(src, { withFileTypes: true });
    for (const ent of ents) {
      await copyRecursive(path.win32.join(src, ent.name), path.win32.join(dst, ent.name));
    }
  } else if (st.isFile()) {
    await fs.promises.copyFile(src, dst);
  }
}

async function writeFileOp(res, data, cfg) {
  const file = resolveAllowedPath(data.path, cfg);
  const content = String(data.content || '');
  const extCheck = checkExtension(file, cfg);
  if (!extCheck.ok) return sendJson(res, 403, { ok: false, error: extCheck.reason });
  await fs.promises.writeFile(file, content, 'utf8');
  sendJson(res, 200, { ok: true, path: file });
}

// —— 用系统默认程序打开文件/文件夹 ——
// explorer.exe 参数数组传递，不经任何 shell 解释（防命令注入；勿改回 cmd /c start）
function openWithSystem(targetPath) {
  const child = spawn('explorer.exe', [targetPath], {
    windowsHide: true, detached: true, stdio: 'ignore'
  });
  child.on('error', () => {
    try {
      spawn('rundll32.exe', ['url.dll,FileProtocolHandler', targetPath], {
        windowsHide: true, detached: true, stdio: 'ignore'
      }).unref();
    } catch (_) {}
  });
  child.unref();
}

// Windows 非法文件名字符（含控制字符）：出现即异常，拒绝零误伤。
// 合法元字符（& % ^ !）无需拦截——explorer.exe 参数数组传递无 shell 解释路径。
function hasUnsafeName(p) {
  return /["<>|*?\u0000-\u001f]/.test(String(p || ''));
}

async function openOp(res, data, cfg) {
  const extCheck = checkExtension(data.path || '', cfg);
  if (!extCheck.ok) return sendJson(res, 403, { ok: false, error: extCheck.reason });
  if (hasUnsafeName(data.path || '')) return sendJson(res, 400, { ok: false, error: 'invalid characters in path' });
  // 压缩包虚拟路径（path.ext::entry）→ 提取到临时文件后打开
  const zipSep = (data.path || '').indexOf('::');
  if (zipSep > 0) {
    const archivePath = data.path.substring(0, zipSep);
    const entry = data.path.substring(zipSep + 2);
    const entryExtCheck = checkExtension(entry, cfg);
    if (!entryExtCheck.ok) return sendJson(res, 403, { ok: false, error: entryExtCheck.reason });
    if (hasUnsafeName(entry)) return sendJson(res, 400, { ok: false, error: 'invalid characters in entry' });
    const archiveFull = resolveAllowedPath(archivePath, cfg);
    if (!fs.existsSync(archiveFull)) throw new Error('archive not found');
    const content = await extractEntryFromArchive(archiveFull, entry);
    const os = await import('os');
    const tmpDir = path.join(os.tmpdir(), 'html-FileNest-open');
    await fs.promises.mkdir(tmpDir, { recursive: true });
    const base = path.basename(entry);
    const tmpFile = path.join(tmpDir, Date.now() + '_' + base);
    await fs.promises.writeFile(tmpFile, content);
    openWithSystem(tmpFile);
    return sendJson(res, 200, { ok: true, path: tmpFile, extracted: true });
  }
  const target = resolveAllowedPath(data.path, cfg);
  if (!fs.existsSync(target)) throw new Error('path not found');
  openWithSystem(target);
  sendJson(res, 200, { ok: true, path: target });
}

async function uploadOp(req, res, q, cfg) {
  const dir = resolveAllowedPath(q.path, cfg);
  const name = String(q.name || '').trim();
  if (!name || /[\\\/:*?"<>|]/.test(name) || name === '.' || name === '..' || /[. ]$/.test(name)) {
    throw new Error('invalid name');
  }
  const full = resolveAllowedPath(path.win32.join(dir, name), cfg);
  const extCheck = checkExtension(full, cfg);
  if (!extCheck.ok) return sendJson(res, 403, { ok: false, error: extCheck.reason });
  const overwrite = q.overwrite === '1';
  if (!overwrite && fs.existsSync(full)) throw new Error('file exists: ' + name);

  const tmp = full + '.upload.' + process.pid + '.' + Date.now();
  await new Promise((resolve, reject) => {
    const ws = fs.createWriteStream(tmp);
    let done = false;
    const finish = (err) => {
      if (done) return; done = true;
      try { ws.close(); } catch (_) {}
      if (err) { try { fs.unlinkSync(tmp); } catch (_) {} reject(err); }
      else resolve();
    };
    req.on('error', finish);
    ws.on('error', finish);
    ws.on('finish', () => finish());
    req.pipe(ws);
  });

  try { if (fs.existsSync(full)) fs.unlinkSync(full); } catch (_) {}
  fs.renameSync(tmp, full);
  sendJson(res, 200, { ok: true, path: full, size: fs.statSync(full).size });
}

// ============ 文件打开 / 定位（保留原 /open 语义） ============
async function handleFileOp(op, q, res, cfg) {
  let target;
  try { target = await resolveTarget(q, cfg); }
  catch (e) { return sendJson(res, 403, { ok: false, error: e.message }); }
  if (!target) return sendJson(res, 403, { ok: false, error: 'path not allowed' });

  const extCheck = checkExtension(target.displayPath || target.localPath || '', cfg);
  if (!extCheck.ok) {
    return sendJson(res, 403, { ok: false, error: extCheck.reason });
  }

  let localPath;
  try { localPath = await obtainLocal(target, cfg); }
  catch (e) { return sendJson(res, 404, { ok: false, error: e.message }); }

  if (op === '/exists') return sendJson(res, 200, { ok: true, exists: true, path: target.displayPath });

  if (op === '/open') {
    let replied = false;
    const reply = (data) => {
      if (replied) return;
      replied = true;
      sendJson(res, data.ok ? 200 : 500, data);
    };

    try {
      const child = spawn('explorer.exe', [localPath], {
        windowsHide: true, detached: true, stdio: 'ignore'
      });
      child.on('error', (err) => {
        log('explorer 失败，降级 rundll32: ' + err.message);
        try {
          const child2 = spawn('rundll32.exe', ['url.dll,FileProtocolHandler', localPath], {
            windowsHide: true, detached: true, stdio: 'ignore'
          });
          child2.on('error', () => reply({ ok: false, error: '系统拒绝执行打开命令' }));
          child2.unref();
          reply({ ok: true, path: target.displayPath, local: localPath, method: 'rundll32-fallback' });
        } catch (e2) {
          reply({ ok: false, error: '系统拒绝执行' });
        }
      });
      child.unref();
      setTimeout(() => {
        reply({ ok: true, path: target.displayPath, local: localPath, method: 'explorer' });
      }, 100);
    } catch (e) {
      reply({ ok: false, error: 'open failed: ' + e.message });
    }
    return;
  }

  if (op === '/reveal') {
    try {
      spawn('explorer.exe', ['/select,', localPath],
        { windowsHide: false, detached: true, stdio: 'ignore' }).unref();
      return sendJson(res, 200, { ok: true });
    } catch (e) {
      return sendJson(res, 500, { ok: false, error: 'reveal failed: ' + e.message });
    }
  }

  if (op === '/open-folder') {
    try {
      const folderPath = path.win32.dirname(localPath);
      spawn('explorer.exe', [folderPath],
        { windowsHide: false, detached: true, stdio: 'ignore' }).unref();
      return sendJson(res, 200, { ok: true, folder: folderPath });
    } catch (e) {
      return sendJson(res, 500, { ok: false, error: 'open folder failed: ' + e.message });
    }
  }
}

function checkExtension(filePath, cfg) {
  // 去掉尾部点/空格再取扩展名，防 "evil.exe." / "evil.exe " 绕过（Windows 会忽略尾部点空格）
  const normalized = String(filePath).replace(/[. ]+$/, '');
  const ext = path.win32.extname(normalized).toLowerCase();
  if (cfg.blockExecutables && ext && cfg.blockedExtensions.indexOf(ext) >= 0) {
    return { ok: false, reason: 'executable blocked: ' + ext };
  }
  if (cfg.allowedExtensions.length > 0) {
    if (!ext) return { ok: false, reason: 'file has no extension' };
    if (cfg.allowedExtensions.indexOf(ext) < 0) {
      return { ok: false, reason: 'extension not in whitelist: ' + ext };
    }
  }
  return { ok: true };
}

// ============ Resolver（含根目录自身匹配修复） ============
function wildcardToRegex(pattern) {
  const esc = pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&');
  return new RegExp('^' + esc.replace(/\*/g, '.*').replace(/\?/g, '.') + '$', 'i');
}

function matchAnyRoot(absPath, roots) {
  const norm = path.win32.normalize(absPath);
  const nl = norm.toLowerCase();
  for (const r of roots) {
    if (!r) continue;
    const rl = r.toLowerCase();

    if (!/[?*]/.test(r)) {
      const rClean = rl.replace(/[\\\/]$/, '');
      const nClean = nl.replace(/[\\\/]$/, '');
      if (nClean === rClean) return true;
      if (nClean.startsWith(rClean + '\\')) return true;
      continue;
    }

    if (rl === '*' || rl === '**') return true;

    if (rl.endsWith('\\*') || rl.endsWith('/*')) {
      const prefix = rl.slice(0, -2);
      if (nl === prefix) return true;
      if (nl.startsWith(prefix + '\\') || nl.startsWith(prefix + '/')) return true;
    }

    const pat = rl.endsWith('*') ? rl : rl + '*';
    if (wildcardToRegex(pat).test(norm)) return true;
  }
  return false;
}

async function resolveTarget(q, cfg) {
  if (q.source) {
    const src = cfg.sources[q.source];
    if (!src) throw new Error('unknown source: ' + q.source);
    const rel = String(q.path || '').replace(/^[\\\/]+/, '');
    if (!rel) throw new Error('path required');
    if (/[\x00-\x1f]/.test(rel)) throw new Error('invalid characters');

    if (src.type === 'local') {
      const candidates = [];
      for (const root of src.roots) {
        if (root === '*' || root === '**') {
          let abs;
          try { abs = path.win32.normalize(rel); } catch (e) { continue; }
          if (!abs.startsWith('\\\\') && !/^[A-Za-z]:\\/.test(abs)) continue;
          candidates.push(abs);
        } else {
          const base = root.replace(/[*?]/g, '');
          const abs = path.win32.join(base, rel);
          if (!matchAnyRoot(abs, [root])) continue;
          candidates.push(abs);
        }
      }
      if (!candidates.length) throw new Error('path escapes all roots');
      const chosen = src.tryAll === false
        ? candidates[0]
        : candidates.find(p => { try { return fs.existsSync(p); } catch (e) { return false; } }) || candidates[0];
      return { type: 'local', displayPath: chosen, localPath: chosen, source: q.source };
    }
    if (src.type === 'http') {
      const base = src.base.endsWith('/') ? src.base : src.base + '/';
      const safeRel = rel.split(/[\\\/]/).map(encodeURIComponent).join('/');
      const url = base + safeRel;
      checkHost(url, cfg);
      return { type: 'http', displayPath: url, url, source: q.source, auth: src.auth };
    }
    throw new Error('unsupported source type');
  }

  const abs = String(q.path || '');
  if (!abs) throw new Error('path required');
  if (/[\x00-\x1f]/.test(abs)) throw new Error('invalid characters');
  let norm;
  try { norm = path.win32.normalize(abs); } catch (e) { throw new Error('invalid path'); }
  if (!norm.startsWith('\\\\') && !/^[A-Za-z]:\\/.test(norm)) {
    throw new Error('must be absolute path');
  }

  for (const name of Object.keys(cfg.sources)) {
    const src = cfg.sources[name];
    if (src.type !== 'local') continue;
    if (matchAnyRoot(norm, src.roots)) {
      return { type: 'local', displayPath: norm, localPath: norm, source: name };
    }
  }
  throw new Error('path not in any configured source');
}

function checkHost(urlStr, cfg) {
  if (!cfg.allowedHosts || !cfg.allowedHosts.length) return;
  const u = new URL(urlStr);
  if (cfg.allowedHosts.indexOf(u.hostname.toLowerCase()) < 0) {
    throw new Error('host not allowed: ' + u.hostname);
  }
}

// ============ Fetcher ============
async function obtainLocal(target, cfg) {
  if (target.type === 'local') {
    if (!fs.existsSync(target.localPath)) {
      throw new Error('file not found or no access: ' + target.displayPath);
    }
    return target.localPath;
  }
  if (target.type === 'http') return await downloadToCache(target, cfg);
  throw new Error('unsupported target type');
}

async function downloadToCache(target, cfg) {
  const cacheDir = cfg.cacheDir;
  if (!fs.existsSync(cacheDir)) fs.mkdirSync(cacheDir, { recursive: true });

  const urlPath = new URL(target.url).pathname;
  const ext = path.win32.extname(urlPath) || '.dat';
  const hash = crypto.createHash('sha1')
    .update(target.url + (target.auth ? target.auth.user : ''))
    .digest('hex').slice(0, 16);
  const fpath = path.join(cacheDir, hash + ext);

  if (fs.existsSync(fpath)) {
    const st = fs.statSync(fpath);
    if ((Date.now() - st.mtimeMs) / 3600000 < cfg.cacheMaxAgeHours) {
      try { fs.utimesSync(fpath, new Date(), new Date()); } catch (e) {}
      return fpath;
    }
  }

  const tmp = fpath + '.part.' + process.pid;
  try {
    await httpDownload(target.url, tmp, {
      timeoutMs: cfg.downloadTimeoutMs,
      auth: target.auth
    });
    if (fs.existsSync(fpath)) fs.unlinkSync(fpath);
    fs.renameSync(tmp, fpath);
    return fpath;
  } catch (e) {
    try { if (fs.existsSync(tmp)) fs.unlinkSync(tmp); } catch (_) {}
    throw new Error('download failed: ' + e.message);
  }
}

function httpDownload(urlStr, destPath, opts) {
  return new Promise((resolve, reject) => {
    const u = new URL(urlStr);
    const lib = u.protocol === 'https:' ? https : http;
    const headers = {};
    if (opts.auth && opts.auth.user) {
      headers['Authorization'] = 'Basic ' +
        Buffer.from(opts.auth.user + ':' + (opts.auth.pass || '')).toString('base64');
    }
    const req = lib.get(u, { headers, timeout: opts.timeoutMs }, res => {
      if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
        const next = new URL(res.headers.location, u).toString();
        res.resume();
        return httpDownload(next, destPath, opts).then(resolve, reject);
      }
      if (res.statusCode !== 200) {
        res.resume();
        return reject(new Error('HTTP ' + res.statusCode));
      }
      pipeAsync(res, fs.createWriteStream(destPath)).then(resolve, reject);
    });
    req.on('timeout', () => req.destroy(new Error('timeout')));
    req.on('error', reject);
  });
}

async function cleanupCache(force) {
  const cfg = loadConfig();
  if (!fs.existsSync(cfg.cacheDir)) return;
  const maxMs = (cfg.cacheMaxAgeHours || 24) * 3600000;
  const now = Date.now();
  try {
    for (const f of fs.readdirSync(cfg.cacheDir)) {
      const full = path.join(cfg.cacheDir, f);
      try {
        const st = fs.statSync(full);
        if (st.isFile() && (force || now - st.mtimeMs > maxMs)) fs.unlinkSync(full);
      } catch (e) {}
    }
  } catch (e) {}
}

// ============ 工具 ============
function sendJson(res, code, obj) {
  const body = JSON.stringify(obj);
  res.writeHead(code, {
    'Content-Type': 'application/json; charset=utf-8',
    'Content-Length': Buffer.byteLength(body)
  });
  res.end(body);
}
function isAdmin() {
  return spawnSync('net', ['session'], { stdio: 'ignore', windowsHide: true }).status === 0;
}
function elevateSelf(args) {
  const exe = process.execPath.replace(/'/g, "''");
  let ps = "Start-Process -FilePath '" + exe + "'";
  if (args && args.length) {
    ps += ' -ArgumentList ' + args.map(a => "'" + a.replace(/'/g, "''") + "'").join(',');
  }
  ps += ' -Verb RunAs';
  try {
    spawnSync('powershell', ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-WindowStyle', 'Hidden',
      '-EncodedCommand', Buffer.from(ps, 'utf16le').toString('base64')],
      { stdio: 'ignore', windowsHide: true });
  } catch (e) {}
  process.exit(0);
}
function makeCert() {
  const ps = `
$ErrorActionPreference = 'Stop'
try {
  $c = New-SelfSignedCertificate -Subject 'CN=html-FileNest Local' -DnsName 'localhost','127.0.0.1' ` +
    `-CertStoreLocation Cert:\\LocalMachine\\My -NotAfter (Get-Date).AddYears(10) ` +
    `-KeyExportPolicy Exportable -KeyLength 2048
  $p = ConvertTo-SecureString -String '${PFX_PASS}' -Force -AsPlainText
  Export-PfxCertificate -Cert $c -FilePath '${PFX_PATH}' -Password $p | Out-Null
  $c | Export-Certificate -FilePath '${CER_PATH}' -Type CERT | Out-Null
  Import-Certificate -FilePath '${CER_PATH}' -CertStoreLocation Cert:\\LocalMachine\\Root | Out-Null
  Write-Output 'OK'
} catch { Write-Output ('ERR:' + $_.Exception.Message) }
`;
  const r = runPS(ps);
  if (!/OK/.test(r.stdout || '')) log('cert warning: ' + (r.stdout || r.stderr || ''));
}
function runPS(script) {
  return spawnSync('powershell',
    ['-NoProfile', '-ExecutionPolicy', 'Bypass', '-EncodedCommand',
     Buffer.from(script, 'utf16le').toString('base64')],
    { encoding: 'utf8', windowsHide: true });
}
function showMsg(title, msg, icon) {
  const ps = `
Add-Type -AssemblyName PresentationFramework
[System.Windows.MessageBox]::Show(
  '${msg.replace(/'/g, "''")}',
  '${title.replace(/'/g, "''")}',
  'OK',
  '${icon || 'Information'}'
) | Out-Null
`;
  try { runPS(ps); } catch (e) {}
}
function taskExists(name) {
  return spawnSync('schtasks', ['/query', '/tn', name],
    { stdio: 'ignore', windowsHide: true }).status === 0;
}
function portOpen(port) {
  return new Promise(resolve => {
    const s = net.connect({ host: HOST, port });
    let done = false;
    const fin = ok => { if (!done) { done = true; s.destroy(); resolve(ok); } };
    s.once('connect', () => fin(true));
    s.once('error', () => fin(false));
    s.setTimeout(800, () => fin(false));
  });
}
function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }
function log(s) { try { console.log('[html-FileNest] ' + s); } catch (e) {} }

// ============ UI 页面 ============
// UI 已抽离为独立文件 index.html（部署于 INSTALL_DIR 或与本文件同目录），
// 由 serveUI() 在运行时读取并对外提供，不再内嵌于本 JS 中。

// ============ 入口分发（放文件末尾，确保所有 let 初始化完成后再执行） ============
const argv = process.argv.slice(2);
if (argv.includes('--licenses'))       console.log(THIRD_PARTY_NOTICES);
else if (argv.includes('--serve'))     runServer();
else if (argv.includes('--uninstall')) doUninstall();
else                                   doAuto().catch(e => {
  showMsg(APP_NAME + ' 部署失败', String(e && e.message || e), 'Error');
  process.exit(1);
});
