// html-FileNest.js — 网页版本地文件管理器
// 项目：html-FileNest  版本：1.4.0.0
// 双击无参数 => 部署并启动；--serve => 仅服务；--uninstall => 卸载

const fs     = require('fs');
const os     = require('os');
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
const { Jimp } = require('jimp');
const { spawn, spawnSync } = require('child_process');
const { pipeline } = require('stream');
const zlib = require('zlib');
const { promisify } = require('util');

const pipeAsync = promisify(pipeline);

process.on('uncaughtException', (err) => {
  console.error('[html-FileNest 防崩溃] ' + (err && err.stack || err));
});

// fs 线程池扩容（须在任何 fs I/O 之前设置）：默认 4 线程被慢盘/UNC 占满会拖垮目录
// 枚举（2026-11-15 卡顿根因之一）；start.vbs 侧同设，覆盖所有启动方式
if (!process.env.UV_THREADPOOL_SIZE) process.env.UV_THREADPOOL_SIZE = '64';

const APP_NAME    = 'html-FileNest';
const SERVICE_VER = '2.2.0.1';
const HOST        = '127.0.0.1';
const INSTALL_DIR = process.env.FIN_INSTALL_DIR || 'C:\\html-FileNest';
// 打包 exe 自身所在目录（便携运行时可能与 INSTALL_DIR 不同）；从源码运行时为 null
const SELF_DIR = /html-FileNest\.exe$/i.test(process.execPath) ? path.dirname(process.execPath) : null;
const EXE_PATH    = path.join(INSTALL_DIR, 'html-FileNest.exe');
const CONFIG_PATH = path.join(INSTALL_DIR, 'config.json');
const PFX_PATH    = path.join(INSTALL_DIR, 'filenest.pfx');
const CER_PATH    = path.join(INSTALL_DIR, 'filenest.cer');
const VBS_PATH    = path.join(INSTALL_DIR, 'start.vbs');
// 沙箱锚点连带作用域（2026-11-18）：FIN_INSTALL_DIR 下任务名/证书主题加后缀，沙箱安装
// 与真实安装完全隔离——修复"产品启动清理 bug"（沙箱清理曾误删默认安装的任务/证书/文件）
const SANDBOX     = !!process.env.FIN_INSTALL_DIR;
const TASK_NAME   = SANDBOX ? 'html-FileNest-fin-test' : 'html-FileNest';
const CERT_CN     = SANDBOX ? 'html-FileNest FinTest' : 'html-FileNest';
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
    '.reg', '.inf', '.url', '.dll'
  ],
  cacheDir: 'C:\\html-FileNest\\cache',
  cacheMaxAgeHours: 24,
  downloadTimeoutMs: 120000,
  allowedHosts: [],
  deniedPaths: [],
  security: {
    requireLocalOrigin: true,
    allowedOrigins: [],
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
  if (!Array.isArray(c.security.allowedOrigins)) c.security.allowedOrigins = [];
  c.security.authToken = c.security.authToken || '';
  // —— 路径黑名单（防篡改）：内置安装目录/自身目录始终生效（配置可追加、不可移除） ——
  const deny = [];
  const pushDeny = (p) => {
    if (!p) return;
    try {
      const n = path.win32.normalize(String(p)).replace(/[\\\/]+$/, '').toLowerCase();
      if (n && n !== '\\' && deny.indexOf(n) < 0) deny.push(n);
    } catch (_) {}
  };
  pushDeny(INSTALL_DIR);
  pushDeny(SELF_DIR);
  for (const p of (Array.isArray(c.deniedPaths) ? c.deniedPaths : [])) pushDeny(p);
  c.deniedPaths = deny;

  // ${ENV} 展开加白名单（#23）：仅预定义安全变量可展开，未知变量按空串处理，防环境变量值泄露
  const ENV_WHITELIST = new Set([
    'USERPROFILE', 'USERNAME', 'COMPUTERNAME', 'TEMP', 'TMP', 'LOCALAPPDATA', 'APPDATA',
    'PUBLIC', 'HOMEDRIVE', 'HOMEPATH', 'SYSTEMROOT', 'WINDIR', 'PROGRAMFILES',
    'PROGRAMFILES(X86)', 'COMMONPROGRAMFILES', 'COMMONPROGRAMFILES(X86)',
    'ALLUSERSPROFILE', 'PROGRAMDATA'
  ]);
  const expand = (s) => typeof s === 'string'
    ? s.replace(/\$\{([^}]+)\}/g, (_, k) =>
        (ENV_WHITELIST.has(String(k).toUpperCase()) ? (process.env[k] || '') : ''))
    : s;

  Object.keys(c.sources).forEach(k => {
    const src = c.sources[k];
    if (src.type === 'local' && !src.roots && src.root) {
      src.roots = [src.root]; delete src.root;
    }
    if (src.type === 'local' && (!src.roots || !src.roots.length)) {
      src.roots = ['*'];
    }
    if (src.roots) {
      // roots 展开后必须为绝对路径（盘符或 UNC 网络路径）或通配 '*'（#23）：拒绝相对路径，防路径注入
      // UNC 不拒绝（杨总 2026-11-13 指示，NAS 共享盘场景需要）
      src.roots = src.roots.map(expand).filter(r =>
        typeof r === 'string' && (r === '*' || /^[A-Za-z]:[\\/]/.test(r) || /^[\\/]{2}[^\\/]/.test(r)));
      if (!src.roots.length) src.roots = [];
    }
    if (src.base) src.base = expand(src.base);
    if (src.tryAll == null) src.tryAll = true;
  });
  c.cacheDir = expand(c.cacheDir);
  return c;
}

// ============ 部署 ============
// 删除文件/目录，带重试（杀进程后文件句柄可能延迟释放；一次性整目录盲删抛错会半途而废留下残留）
async function removeWithRetry(p, isDir) {
  for (let i = 0; i < 5; i++) {
    try {
      if (!fs.existsSync(p)) return true;
      if (isDir) fs.rmSync(p, { recursive: true, force: true });
      else fs.rmSync(p, { force: true });
      return true;
    } catch (e) {}
    await sleep(300);
  }
  return false;
}

// 强卸载（2026-11-18 自特例版引入，按用户要求收窄）：安装/升级前清旧版程序
// （进程+计划任务+安装目录程序文件），只针对 html-FileNest 本产品；
// 用户数据一律保留——config.json/favorites.json/recycle_bin/recycle_meta.json/
// upload-journal.json/index.html/index.html.bak/filenest.pfx/filenest.cer/file-index.db*，
// 升级后原配置、原证书信任与原索引库自动继承（2.2.0.0 起 SQLite 索引升级不重建）。
// 证书信任链升级时一并继承（不清证书库、保留 filenest.pfx/cer，浏览器免重信任），
// 显式卸载 --uninstall 才清证书。
// FIN_INSTALL_DIR 沙箱作用域：候选目录只认 INSTALL_DIR、进程按 exe 路径定点杀、
// 任务用沙箱名（TASK_NAME），绝不触碰默认安装（清理 bug 修复）。
async function uninstallProduct(name, knownDir) {
  // 0) 收集安装目录候选（必须在删计划任务之前提取，删了任务线索就没了）：
  //    已知目录 + 计划任务 XML 里含产品名的绝对路径所在目录 + 约定目录 C:\<name>
  const candidates = [];
  if (SANDBOX) {
    candidates.push(INSTALL_DIR);
  } else {
    if (knownDir) candidates.push(knownDir);
    try {
      const r = spawnSync('schtasks', ['/query', '/tn', TASK_NAME, '/xml'],
        { encoding: 'utf8', windowsHide: true });
      const all = ((r.stdout || '').match(/[A-Za-z]:\\[^<"\r\n]*/g) || []);
      for (const p of all) {
        if (p.toLowerCase().includes(name.toLowerCase())) {
          const isFile = /\.(exe|vbs|cmd|bat|ps1)$/i.test(p);
          candidates.push(isFile ? path.win32.dirname(p) : p);
        }
      }
    } catch (e) {}
    candidates.push('C:\\' + name);
  }
  // 1) 杀进程并等真正退出（句柄释放后才能删文件）。正常=按进程名杀（排除自身，
  //    安装器可能与产品同名）；沙箱=按 exe 路径定点杀，不伤真实安装的服务
  const selfPid = process.pid;
  const psQ = s => "'" + String(s).replace(/'/g, "''") + "'";
  const killQ = SANDBOX
    ? "Get-CimInstance Win32_Process -Filter \"Name='" + name + ".exe'\" | " +
      "Where-Object { $_.ProcessId -ne " + selfPid + " -and $_.ExecutablePath -like " + psQ(INSTALL_DIR + '\\*') + " } | " +
      "ForEach-Object { Stop-Process -Id $_.ProcessId -Force -ErrorAction SilentlyContinue }"
    : "Get-Process -Name " + psQ(name) + " -ErrorAction SilentlyContinue | " +
      "Where-Object { $_.Id -ne " + selfPid + " } | " +
      "ForEach-Object { Stop-Process -Id $_.Id -Force -ErrorAction SilentlyContinue }";
  const aliveQ = SANDBOX
    ? "Write-Output (@(Get-CimInstance Win32_Process -Filter \"Name='" + name + ".exe'\" | " +
      "Where-Object { $_.ProcessId -ne " + selfPid + " -and $_.ExecutablePath -like " + psQ(INSTALL_DIR + '\\*') + " }).Count)"
    : "Write-Output (@(Get-Process -Name " + psQ(name) + " -ErrorAction SilentlyContinue | " +
      "Where-Object { $_.Id -ne " + selfPid + " }).Count)";
  try { runPS(killQ); } catch (e) {}
  for (let i = 0; i < 10; i++) {
    let alive = '';
    try { alive = runPS(aliveQ).stdout || ''; } catch (e) { break; }
    if (!/[1-9]/.test(alive.trim())) break;
    await sleep(300);
  }
  // 2) 删计划任务（TASK_NAME；沙箱名与真实安装隔离）
  spawnSync('schtasks', ['/end', '/tn', TASK_NAME], { stdio: 'ignore', windowsHide: true });
  spawnSync('schtasks', ['/delete', '/tn', TASK_NAME, '/f'], { stdio: 'ignore', windowsHide: true });
  // 3) 清安装目录程序文件（逐候选；发现型候选的目录名必须等于产品名才删，防误删）
  //    不整目录盲删——盲删会把用户配置一锅端；被占用文件抛错时又半途而废留下残留
  const keep = ['config.json', 'favorites.json', 'recycle_bin', 'recycle_meta.json', 'upload-journal.json',
    'index.html', 'index.html.bak', 'filenest.pfx', 'filenest.cer',
    'file-index.db', 'file-index.db-wal', 'file-index.db-shm'];
  for (const dir of candidates) {
    try {
      // 已知目录（knownDir/INSTALL_DIR）由调用方明确指定不做目录名校验；
      // 发现型候选（任务 XML/约定目录）必须目录名==产品名，防误删
      const explicit = dir === knownDir || dir === INSTALL_DIR;
      if (!explicit && path.win32.basename(dir).toLowerCase() !== name.toLowerCase()) continue;
      if (!fs.existsSync(dir)) continue;
      for (const f of fs.readdirSync(dir)) {
        if (keep.indexOf(f.toLowerCase()) >= 0) continue;
        const full = path.join(dir, f);
        const isDir = fs.statSync(full).isDirectory();
        if (!(await removeWithRetry(full, isDir))) log(name + ' cleanup locked, left: ' + full);
      }
    } catch (e) {}
  }
  log(name + ' cleanup done');
}

async function doAuto() {
  const cfg = loadConfig();
  // 安装/升级=全新替换：先强卸载旧版再安装（2026-11-18 自特例版引入）。
  // 原"已在运行"拦截会在服务运行时直接 return，升级永远轮不到（特例版 2026-11-16 教训）
  if (!isAdmin()) { elevateSelf([]); return; }
  // 只卸载 html-FileNest 程序文件，用户配置/证书/索引库一律保留（见 uninstallProduct 的 keep 白名单）
  await uninstallProduct('html-FileNest', INSTALL_DIR);
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
    const uiSamePath = uiSrc && path.resolve(uiSrc).toLowerCase() === path.resolve(uiDst).toLowerCase();
    // 安装/升级必须刷新 UI（修复：旧 index.html 残留 + 装新版本不覆盖，导致界面永不更新）。
    // 从部署位置直接启动（uiSrc === uiDst）时不动现有文件，保留用户个性化修改。
    if (!uiSamePath) {
      const uiNew = uiSrc ? fs.readFileSync(uiSrc, 'utf8') : UI_HTML_DEPLOY;
      const uiOld = fs.existsSync(uiDst) ? fs.readFileSync(uiDst, 'utf8') : null;
      if (uiOld !== uiNew) {
        if (uiOld !== null) {
          // 旧 UI 先备份为 index.html.bak（用户个性化修改可找回）
          try { fs.writeFileSync(uiDst + '.bak', uiOld, 'utf8'); } catch (e) {}
        }
        fs.writeFileSync(uiDst, uiNew, 'utf8');
        log('index.html 已部署到 ' + uiDst + (uiOld !== null ? '（旧文件备份为 index.html.bak）' : ''));
      }
    }
  }
  if (!fs.existsSync(CONFIG_PATH)) {
    fs.writeFileSync(CONFIG_PATH, JSON.stringify(DEFAULT_CONFIG, null, 2), 'utf8');
  }
  // UV_THREADPOOL_SIZE=64：fs 线程池默认 4 线程，被慢盘/死 UNC 占满时目录枚举整体卡死
  // （2026-11-15 卡顿根因之一）；内容有变时重写 start.vbs，旧部署升级后自动换新。
  // VBS 引号：字符串内嵌引号用 ""（且不闭合外层字符串）；2026-11-15 曾误用 """ 在此
  // 处提前闭合字符串导致 exe 路径裸露、800A0401"语句未结束"，cscript 实测后修正
  const vbsBody = 'CreateObject("WScript.Shell").Run "cmd /c set UV_THREADPOOL_SIZE=64 && ""' + EXE_PATH + '"" --serve", 0, False\r\n';
  try {
    if (!fs.existsSync(VBS_PATH) || fs.readFileSync(VBS_PATH, 'ascii') !== vbsBody) {
      fs.writeFileSync(VBS_PATH, vbsBody, 'ascii');
    }
  } catch (e) {
    if (!fs.existsSync(VBS_PATH)) fs.writeFileSync(VBS_PATH, vbsBody, 'ascii');
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
    runPS("Get-ChildItem Cert:\\LocalMachine\\Root | Where-Object {$_.Subject -like '*" + CERT_CN + "*'} | " +
      "Remove-Item -Force -ErrorAction SilentlyContinue");
  } catch (e) {}
  [VBS_PATH, PFX_PATH, CER_PATH].forEach(f => {
    try { if (fs.existsSync(f)) fs.unlinkSync(f); } catch (e) {}
  });
  // 界面文件随卸载清理（含升级备份 index.html.bak），避免残留旧 UI 毒害下次安装
  [path.join(INSTALL_DIR, 'index.html'), path.join(INSTALL_DIR, 'index.html.bak')].forEach(f => {
    try { if (fs.existsSync(f)) fs.unlinkSync(f); } catch (e) {}
  });
  showMsg('卸载完成', 'html-FileNest.exe 本体仍在 ' + INSTALL_DIR + '，index.html 已清理', 'Information');
}

// ============ 服务停止/重启（客户机运维入口） ============
// 计划任务经 wscript start.vbs 拉起后 wscript 即退出，服务进程脱离任务管控
// （schtasks /end 摸不到它）。停止/重启动作全部委托给一次性 VBS 外力执行：
// WMI 按"同一 exe 路径 + 命令行带 --serve"精确停止，不误伤其他安装、不依赖 PID 文件；
// 不用 PowerShell 做关键路径（客户机环境差异大，启动不稳会把重启流程做丢）。
function serviceControlVbs(exePath, alsoRelaunch) {
  const exe = String(exePath).replace(/"/g, '""');
  const L = [];
  L.push('On Error Resume Next');
  L.push('Set wmi = GetObject("winmgmts:\\\\.\\root\\cimv2")');
  L.push('Function SvcAlive()');
  L.push('  SvcAlive = False');
  L.push("  For Each q In wmi.ExecQuery(\"SELECT CommandLine, ExecutablePath FROM Win32_Process WHERE Name = 'html-FileNest.exe'\")");
  L.push('    If Not q.ExecutablePath Is Nothing Then');
  L.push('      If LCase(q.ExecutablePath) = LCase("' + exe + '") And InStr(q.CommandLine, "--serve") > 0 Then SvcAlive = True');
  L.push('    End If');
  L.push('  Next');
  L.push('End Function');
  if (alsoRelaunch) L.push('WScript.Sleep 2000');  // 等 HTTP 响应发完再动手
  L.push("For Each p In wmi.ExecQuery(\"SELECT ProcessId, CommandLine, ExecutablePath FROM Win32_Process WHERE Name = 'html-FileNest.exe'\")");
  L.push('  If Not p.ExecutablePath Is Nothing Then');
  L.push('    If LCase(p.ExecutablePath) = LCase("' + exe + '") And InStr(p.CommandLine, "--serve") > 0 Then p.Terminate');
  L.push('  End If');
  L.push('Next');
  if (alsoRelaunch) {
    L.push('WScript.Sleep 2000');  // 等端口释放（Terminate 后句柄释放可能慢，1 秒不够）
    L.push('For i = 1 To 5');
    L.push('  CreateObject("WScript.Shell").Run """' + exe + '"" --serve", 0, False');
    L.push('  WScript.Sleep 3000');
    L.push('  If SvcAlive() Then Exit For');
    L.push('  WScript.Sleep 2000');
    L.push('Next');
    L.push('If Not SvcAlive() Then');
    L.push('  Set fso = CreateObject("Scripting.FileSystemObject")');
    L.push('  Set f = fso.CreateTextFile(fso.GetParentFolderName("' + exe + '") & "\\restart-failed.log", True)');
    L.push('  f.WriteLine CStr(Now) & " relaunch failed: ' + exe + '"');
    L.push('  f.Close');
    L.push('End If');
  }
  return L.join('\r\n') + '\r\n';
}

function runServiceControl(alsoRelaunch) {
  const vbs = path.join(process.env.TEMP || process.env.TMP || '.',
    'html-FileNest-service-' + (alsoRelaunch ? 'restart' : 'stop') + '-' + process.pid + '.vbs');
  // UTF-16 LE + BOM 写入：wscript 按 BOM 解码，exe 路径含中文/特殊字符才不会被 ascii 打碎成 '?'
  fs.writeFileSync(vbs, Buffer.concat([
    Buffer.from([0xFF, 0xFE]),
    Buffer.from(serviceControlVbs(process.execPath, alsoRelaunch), 'utf16le')
  ]));
  if (alsoRelaunch) {
    // 外力脚本自带"等响应->杀旧->等端口->拉起->验证重试"序列，调用方立即返回即可
    spawn('wscript.exe', [vbs], { detached: true, stdio: 'ignore', windowsHide: true }).unref();
  } else {
    spawnSync('cscript.exe', ['//Nologo', vbs], { stdio: 'ignore', windowsHide: true });
  }
}

function doStop() {
  runServiceControl(false);
  console.log('html-FileNest: service stopped');
}

function doRestart() {
  runServiceControl(true);
  console.log('html-FileNest: service restarting');
}

// UI 入口（POST /api/service-restart）：与 --restart 同一条路径
function serviceRestartOp(res) {
  runServiceControl(true);
  sendJson(res, 200, { ok: true, restarting: true });
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
function normalizeOrigin(o) {
  try { return new URL(o).origin; } catch (e) { return ''; }
}
// 应用主页恒放行，不参与白名单：本机（127.0.0.1 / localhost / ::1）+ 配置端口（cfg.ports，
// 默认 20000/20443）。端口跟随 config.json 动态判定（热加载，改端口后自家 UI 照常恒放行）；
// 另有 Origin==Host 同源判定兜底（见 requestOriginAllowed）。
function isSelfOrigin(o) {
  try {
    const u = new URL(o);
    const h = u.hostname.toLowerCase().replace(/^\[|\]$/g, '');
    if (h !== '127.0.0.1' && h !== 'localhost' && h !== '::1') return false;
    const ports = loadConfig().ports || DEFAULT_PORTS;
    const httpPort = String(ports.http || DEFAULT_PORTS.http);
    const httpsPort = String(ports.https || DEFAULT_PORTS.https);
    // URL 显式端口缺省时（http://h:80、https://h:443）u.port 为空串，归一化后再比
    const p = u.port || (u.protocol === 'https:' ? '443' : '80');
    return (u.protocol === 'http:' && p === httpPort) ||
           (u.protocol === 'https:' && p === httpsPort);
  } catch (e) { return false; }
}
// 调用网页白名单：主页恒放行；只对其他站点生效，空名单默认全禁止；
// 条目可写完整 Origin（http://192.168.1.10）或域名（192.168.1.10）；Origin: null 一律拒。
// 无 Origin 的请求（curl、本地脚本）与 v2.0.1.0 一致放行；浏览器 Sec-Fetch-Site: cross-site 拒绝。
function requestOriginAllowed(req) {
  const o = req.headers.origin;
  if (o) {
    if (isSelfOrigin(o)) return true;
    // 自源补充判定：Origin 与请求 Host 一致 = 同源 UI 自调用。
    // 客户改端口后 Origin 端口不再是固定默认值，光认 20000/20443 会把自家 UI 的请求拒掉；
    // Host 由浏览器按目标 URL 生成（禁止头），跨站伪造不了，安全面不受影响。
    try {
      if (new URL(o).host.toLowerCase() === String(req.headers.host || '').toLowerCase()) return true;
    } catch (e) {}
    const t = normalizeOrigin(o);
    if (!t) return false;
    const list = (loadConfig().security || {}).allowedOrigins;
    if (!Array.isArray(list) || list.length === 0) return false;
    let host = '';
    try { host = new URL(t).hostname.toLowerCase().replace(/^\[|\]$/g, ''); } catch (e) {}
    return list.some(e => {
      const s = String(e || '').trim().toLowerCase();
      return !!s && (normalizeOrigin(s) === t || s === host);
    });
  }
  const site = req.headers['sec-fetch-site'];
  return !site || site === 'same-origin' || site === 'none';
}

// ============ 服务 ============
function runServer() {
  const cfg = loadConfig();
  // 启动即后台预热共享枚举（等 PowerShell 冷启动），用户点盘符页时已缓存命中（2026-11-16）
  setTimeout(() => { listNetShares().catch(() => {}); }, 500);
  // ★ HTTP 监听失败必须退出：进程活着但端口不通的"半活"状态比死进程更糟——
  //   死进程能被计划任务/重启脚本拉起，半活僵尸会让服务永远隐身（重启后"再也进不去"）
  http.createServer(handleRequest)
    .on('error', e => {
      log('HTTP ' + cfg.ports.http + ' error: ' + (e.code || e.message) + ' -> exit');
      process.exit(1);
    })
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
  // 后台全局扫描建索引库（2026-11-17）：延时 3s 让服务先就绪；低速低优先级，
  // 页面进入任一非盘符/非根主机目录即永久停止（stopGlobalScan）
  setTimeout(() => { startGlobalScan(cfg).catch(() => {}); }, 3000);
  // 临时文件兜底清扫（#11/#12）：启动后延时首扫，此后每 6h
  setTimeout(() => { try { sweepStaleTemp(); } catch (_) {} }, 7000);
  setInterval(() => { try { sweepStaleTemp(); } catch (_) {} }, 6 * 3600000).unref();
}

function handleRequest(req, res) {
  res.setHeader('X-Content-Type-Options', 'nosniff');
  res.setHeader('X-Frame-Options', 'SAMEORIGIN');
  if (req.headers.origin && requestOriginAllowed(req)) {
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
      sendJson(res, err && err.status || 500, { ok: false, error: String(err && err.message || err) });
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
      if (p === '/api/desktop') return await desktopInfo(res, cfg);
      if (p === '/api/ls')     return await listDir(req, res, q, cfg);
      if (p === '/api/newfile-types') return await newfileTypesOp(res);
      if (p === '/api/favorites') return await favoritesListOp(res);
      if (p === '/api/search')  return await searchOp(req, res, q, cfg);
      if (p === '/api/recycle') return await recycleListOp(res);
      if (p === '/api/stat')   return await statPath(res, q, cfg);
      if (p === '/api/raw')    return await streamFile(req, res, q, cfg);
      if (p === '/api/read')   return await readTextFile(res, q, cfg);
      if (p === '/api/office-preview') return await officePreview(res, q, cfg);
      if (p === '/api/thumb') return await thumbFile(res, q, cfg);
      if (p === '/api/shortcut-info') return await shortcutInfoOp(res, q, cfg);
      if (p === '/api/archive-list') return await archiveList(res, q, cfg);
      if (p === '/api/archive-extract') return await archiveExtract(req, res, q, cfg);
      return sendJson(res, 404, { ok: false, error: 'not found' });
    }

    if (req.method === 'POST') {
      if (p === '/api/upload') return await uploadOp(req, res, q, cfg);
      if (p === '/api/service-restart') return serviceRestartOp(res);
      // /api/write 放宽到 50MB（前端文本编辑器保存大文件，#21）；其余保持 10MB
      const body = await readBody(req, p === '/api/write' ? 50 * 1024 * 1024 : 10 * 1024 * 1024);
      const data = body ? JSON.parse(body) : {};
      if (p === '/api/mkdir')  return await mkdirOp(res, data, cfg);
      if (p === '/api/mkfile') return await mkfileOp(res, data, cfg);
      if (p === '/api/stat-batch') return await statBatchOp(req, res, data, cfg);
      if (p === '/api/favorites') return await favoritesAddOp(res, data, cfg);
      if (p === '/api/unfavorite') return await unfavoriteOp(res, data);
      if (p === '/api/rename') return await renameOp(res, data, cfg);
      if (p === '/api/rm')     return await rmOp(res, data, cfg);
      if (p === '/api/extract-all')       return await extractAllOp(res, data, cfg);
      if (p === '/api/props')             return await propsOp(res, data, cfg);
      if (p === '/api/zip-create')        return await zipCreateOp(res, data, cfg);
      if (p === '/api/recycle-restore')   return await recycleRestoreOp(res, data, cfg);
      if (p === '/api/recycle-restore-all') return await recycleRestoreAllOp(res);
      if (p === '/api/recycle-purge')     return await recyclePurgeOp(res, data, cfg);
      if (p === '/api/recycle-empty')     return await recycleEmptyOp(res, data, cfg);
      if (p === '/api/mv')     return await mvOp(res, data, cfg);
      if (p === '/api/cp')     return await cpOp(res, data, cfg);
      if (p === '/api/write')  return await writeFileOp(res, data, cfg);
      if (p === '/api/open')   return await openOp(res, data, cfg);
      return sendJson(res, 404, { ok: false, error: 'not found' });
    }

    sendJson(res, 405, { ok: false, error: 'method not allowed' });
  } catch (e) {
    if (!res.headersSent) sendJson(res, e && e.status || 400, { ok: false, error: String(e.message || e) });
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

// 绝对路径归一。UNC（\\host[\...]）必须手工归一：
// path.win32.normalize 会把只有主机名的 \\host 错误吞成 \host，导致共享盘主机根入口
// 过不了绝对路径校验（打不开）；两级以上 \\host\share 不受影响（2026-11-13 修）
function normalizeAbs(s) {
  const m = s.match(/^[\\/]{2}(.*)$/);
  if (m) {
    const parts = m[1].split(/[\\/]+/).filter(Boolean);
    if (!parts.length) throw new Error('invalid path');
    let tail = parts.slice(1).join('\\');
    if (tail) {
      tail = path.win32.normalize(tail).replace(/[\\/]+$/, '');
      if (tail === '.' || tail === '') tail = '';
      else if (tail === '..' || tail.startsWith('..\\')) throw new Error('invalid path');
    }
    return '\\\\' + parts[0] + (tail ? '\\' + tail : '');
  }
  return path.win32.normalize(s);
}

// UNC 主机根（\\host，无共享段）
function isBareHost(p) {
  return /^\\\\[^\\]+$/.test(p);
}

// 控制台输出解码：net view 等系统命令按 OEM 代码页输出（中文 Windows = GBK），
// 直接当 UTF-8 解码会乱码（2026-11-13 修：共享名乱码导致子目录打不开）
function oemLabel() {
  try {
    const { dlopen, FFIType } = require('bun:ffi');
    const k32 = dlopen('kernel32.dll', { GetOEMCP: { returns: FFIType.u32, args: [] } });
    const cp = k32.symbols.GetOEMCP();
    const map = {
      936: 'gbk', 950: 'big5', 932: 'shift_jis', 949: 'euc-kr',
      1252: 'windows-1252', 437: 'windows-1252', 850: 'windows-1252', 65001: 'utf-8'
    };
    return map[cp] || 'gbk';
  } catch (e) {
    return 'gbk';
  }
}

function decodeConsole(buf) {
  // 纯 ASCII 直接解；非 ASCII 按 OEM 代码页解（GBK/Big5/Shift_JIS 均与 ASCII 兼容）
  let ascii = true;
  for (let i = 0; i < buf.length; i++) {
    if (buf[i] >= 0x80) { ascii = false; break; }
  }
  if (ascii) return buf.toString('latin1');
  try { return new TextDecoder(oemLabel()).decode(buf); }
  catch (e) { return new TextDecoder('gbk').decode(buf); }
}

// 列出 UNC 主机上的可见共享（net view = Windows 共享枚举；隐藏共享 $ 不列出）
async function listShares(hostPath) {
  let stdout = '';
  let killed = false;
  let failed = false;
  try {
    const proc = Bun.spawn(['net', 'view', hostPath], { stdout: 'pipe', stderr: 'pipe' });
    const killTimer = setTimeout(() => { killed = true; try { proc.kill(); } catch (e) {} }, 8000);
    stdout = decodeConsole(Buffer.from(await new Response(proc.stdout).arrayBuffer()));
    await proc.exited;
    clearTimeout(killTimer);
  } catch (e) {
    failed = true;
  }
  const lines = String(stdout).split(/\r?\n/);
  const dash = lines.findIndex((l) => /^\s*-{3,}\s*$/.test(l));
  const shares = [];
  if (dash >= 0) {
    for (let i = dash + 1; i < lines.length; i++) {
      const t = lines[i].trim();
      if (!t) continue;
      if (/^(命令|The command|发生系统错误|System error)/.test(t)) break;
      const m = t.match(/^(\S+)/);
      if (m) shares.push(m[1]);
    }
    return shares;
  }
  // 无表体：空列表原样返回；超时/系统错误报出来
  const joined = lines.join(' ').trim();
  if (failed || killed || /发生系统错误|System error|找不到网络路径|network path/i.test(joined)) {
    throw new Error('无法列出主机共享: ' + joined.slice(0, 120));
  }
  return shares;
}

// —— 路径黑名单（防篡改）：命中即拒绝 ——
// realPath：批量列表场景传入已解析的真实路径（调用方缓存父目录真实路径、拼接条目名），
// 免逐条目 realpathSync 系统调用（万级目录曾因此每条 0.12ms，2026-11-15 修）；
// 不传则回退 realpathSync（junction/符号链接指向黑名单目录时同样拒绝）
function isDenied(absNorm, cfg, realPath) {
  const list = cfg.deniedPaths || [];
  if (!list.length) return false;
  const nkey = (p) => path.win32.normalize(String(p)).replace(/[\\\/]+$/, '').toLowerCase();
  const nl = nkey(absNorm);
  for (const d of list) {
    if (d && (nl === d || nl.startsWith(d + '\\'))) return true;
  }
  // 真实路径兜底：junction/符号链接指向黑名单目录时同样拒绝
  try {
    const rl = realPath ? nkey(realPath) : nkey(fs.realpathSync(absNorm));
    for (const d of list) {
      if (d && (rl === d || rl.startsWith(d + '\\'))) return true;
    }
  } catch (_) {}
  return false;
}

// —— 回收站例外（只读）：回收站子树仅供"读"类接口直读（预览/下载/文本/Office/压缩包条目），
//    写入/删除/列目录/搜索仍受黑名单拒绝，保持"不可篡改/不可浏览"语义 ——
function isRecycleChild(absNorm) {
  try {
    const nkey = (p) => path.win32.normalize(String(p)).replace(/[\\\/]+$/, '').toLowerCase();
    const base = nkey(fs.realpathSync(RECYCLE_DIR));
    const rl = nkey(fs.realpathSync(absNorm));
    return rl === base || rl.startsWith(base + '\\');
  } catch (_) { return false; }
}

// —— 坏网络路径快速失败缓存：不可达 UNC 主机首次慢失败（SMB 超时 20s+）后，
//    5 分钟内对该主机的新请求直接快速拒绝，不反复踩网络超时把 fs 线程池堵死
//    （2026-11-16 服务端"塞车"修复，源自 2026-11-15 卡顿三刀验证版） ——
const _badHost = new Map();
function checkBadHost(norm) {
  if (!String(norm).startsWith('\\\\')) return;
  const host = String(norm).slice(2).split('\\')[0].toLowerCase();
  const until = _badHost.get(host);
  if (until) {
    if (Date.now() < until) {
      const e = new Error('network path temporarily unreachable');
      e.status = 400;
      throw e;
    }
    _badHost.delete(host);
  }
}
function markBadHost(absNorm) {
  if (!String(absNorm).startsWith('\\\\')) return;
  const host = String(absNorm).slice(2).split('\\')[0].toLowerCase();
  _badHost.set(host, Date.now() + 300000);
}

// —— 网络盘形态等价（2026-11-18）：映射盘符（Z:）↔ UNC（\\server\share）双向换算 —— //
// SMB 快捷方式目标与 roots 常以不同形态记录同一共享（一边映射盘、一边 UNC），
// 跨形态匹配不上会把真实存在的目标误报"找不到来源"；按当前会话磁盘映射换算形态重试。
// 映射查询缓存 15s（仅在失配兜底路径触发，spawnSync 百 ms 级可接受）。
let _driveMap = { at: 0, list: [] };
function getDriveMappingsSync() {
  if (Date.now() - _driveMap.at < 15000) return _driveMap.list;
  const list = [];
  try {
    const r = Bun.spawnSync(['powershell', '-NoProfile', '-NonInteractive', '-Command',
      "Get-SmbMapping | ForEach-Object { Write-Output ($_.LocalPath + '|' + $_.RemotePath) }"]);
    const out = String(r.stdout || '');
    for (const line of out.split(/\r?\n/)) {
      const i = line.indexOf('|');
      if (i <= 0) continue;
      const local = line.slice(0, i).trim().replace(/[\\\/]+$/, '').toLowerCase();
      const remote = line.slice(i + 1).trim().replace(/[\\\/]+$/, '').toLowerCase();
      if (/^[a-z]:$/.test(local) && remote.startsWith('\\\\')) list.push({ local, remote });
    }
  } catch (e) {}
  _driveMap = { at: Date.now(), list };
  return list;
}

// 路径的等价形态列表（自身形态除外）：盘符形态 ↔ UNC 形态
function mapPathForms(p, maps) {
  const s = String(p || '');
  const out = [];
  const m = s.match(/^([a-zA-Z]):(.*)$/);
  if (m) {
    const hit = maps.find(x => x.local === (m[1] + ':').toLowerCase());
    if (hit) out.push(hit.remote + m[2]);
    return out;
  }
  const nl = s.toLowerCase();
  for (const x of maps) {
    if (nl === x.remote) out.push(x.local);
    else if (nl.startsWith(x.remote + '\\')) out.push(x.local + s.slice(x.remote.length));
  }
  return out;
}

function resolveAllowedPath(pathStr, cfg, opts) {
  const s = String(pathStr || '');
  if (!s) throw new Error('path required');
  if (/[\x00-\x1f]/.test(s)) throw new Error('invalid characters');
  let norm;
  try { norm = normalizeAbs(s); }
  catch (e) { throw new Error('invalid path'); }
  if (!norm.startsWith('\\\\') && !/^[A-Za-z]:\\/.test(norm)) {
    throw new Error('must be absolute path');
  }
  checkBadHost(norm);   // 坏 UNC 主机 5 分钟内快速拒绝（2026-11-16）
  if (isDenied(norm, cfg) && !(opts && opts.trashRead && isRecycleChild(norm))) {
    const e = new Error('path is denied by blacklist');
    e.status = 403;
    throw e;
  }
  for (const name of Object.keys(cfg.sources)) {
    const src = cfg.sources[name];
    if (src.type !== 'local') continue;
    if (matchAnyRoot(norm, src.roots)) return norm;
  }
  // —— 形态/短名兜底（2026-11-18）：网络盘快捷方式目标与 roots 常为同一共享的不同形态
  //    （映射盘 vs UNC）或 8.3 短名——换算形态 + realpath 解短名后重试，仍不匹配才拒绝 ——
  for (const alt of mapPathForms(norm, getDriveMappingsSync())) {
    for (const name of Object.keys(cfg.sources)) {
      const src = cfg.sources[name];
      if (src.type !== 'local') continue;
      if (matchAnyRoot(alt, src.roots)) return alt;
    }
  }
  try {
    const real = normalizeAbs(fs.realpathSync.native(norm));
    if (real.toLowerCase() !== norm.toLowerCase()) {
      for (const name of Object.keys(cfg.sources)) {
        const src = cfg.sources[name];
        if (src.type !== 'local') continue;
        if (matchAnyRoot(real, src.roots)) return real;
      }
    }
  } catch (e) {}
  const rootPeek = [];
  for (const name of Object.keys(cfg.sources)) {
    const src = cfg.sources[name];
    if (src.type === 'local') for (const r of src.roots) rootPeek.push(r);
  }
  throw new Error('path not in any configured source: ' + norm
    + ' [roots: ' + rootPeek.slice(0, 3).join(' | ') + ']');
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
      if (isDenied(expanded, cfg)) continue;
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
  listNetShares().then(net => sendJson(res, 200, { ok: true, drives, net }))
    .catch(() => sendJson(res, 200, { ok: true, drives, net: [] }));
}

// —— 网络共享盘：映射网络驱动器 + 连接过的 UNC 共享 + 活跃 SMB 连接（用户机器上的远程共享） ——
let netCache = null, netCacheAt = 0;
async function listNetShares() {
  const now = Date.now();
  if (netCache && now - netCacheAt < 600000) return netCache;   // 10 分钟缓存
  const ps = '$out=@(); ' +
    'foreach ($m in @(Get-SmbGlobalMapping -ErrorAction SilentlyContinue; Get-SmbMapping -ErrorAction SilentlyContinue)) { ' +
    'if ($m.RemotePath) { $out += [PSCustomObject]@{ path=$m.RemotePath; letter=$m.LocalPath } } }; ' +
    'foreach ($c in @(Get-SmbConnection -ErrorAction SilentlyContinue)) { ' +
    'if ($c.ShareName) { $out += [PSCustomObject]@{ path=(\"\\\\\" + $c.ServerName + \"\\\" + $c.ShareName); letter=$null } } }; ' +
    '$rk=[Microsoft.Win32.Registry]::CurrentUser.OpenSubKey(\"Network\"); ' +
    'if ($rk) { foreach ($n in $rk.GetSubKeyNames()) { $rp=$rk.OpenSubKey($n).GetValue(\"RemotePath\"); ' +
    'if ($rp) { $out += [PSCustomObject]@{ path=$rp; letter=($n + \":\") } } }; $rk.Close() }; ' +
    'foreach ($s in (Get-ChildItem Registry::HKEY_USERS | Where-Object { $_.PSChildName -match \"^S-1-5-21\" } | ForEach-Object { $_.PSChildName })) { ' +
    '$mp=[Microsoft.Win32.Registry]::Users.OpenSubKey($s + \"\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\MountPoints2\"); ' +
    'if ($mp) { foreach ($n in $mp.GetSubKeyNames()) { if ($n.StartsWith(\"##\") -and -not $n.StartsWith(\"###\")) { ' +
    '$out += [PSCustomObject]@{ path=(\"\\\\\" + $n.Substring(2).Replace(\"#\",\"\\\")); letter=$null } } }; $mp.Close() } }; ' +
    '($out | Sort-Object path -Unique) | ConvertTo-Json -Compress';
  try {
    const proc = Bun.spawn(['powershell', '-NoProfile', '-NonInteractive', '-Command', ps],
      { stdout: 'pipe', stderr: 'ignore' });
    // 15s 强杀兜底：断开的 SMB 映射/慢 NAS 会让 Get-SmbConnection 阻塞等网络超时，
    // 无超时的 PowerShell 枚举能把 /api/drives 拖到分钟级
    const killTimer = setTimeout(() => { try { proc.kill(); } catch (e) {} }, 15000);
    const out = await new Response(proc.stdout).text();
    await proc.exited;
    clearTimeout(killTimer);
    let arr = [];
    try { const j = JSON.parse(out.trim() || '[]'); arr = Array.isArray(j) ? j : [j]; } catch { }
    // —— 通用归一规则（产品级，不针对任何具体主机/机器） ——
    // 所有网络共享条目一律收敛到主机根：任何 \\host\share\... 深度的路径
    // （映射盘符 / UNC 记录 / 活跃 SMB 连接，任何主机、任何深度）统一归一为 \\host，
    // 与 C:\ D:\ 同级呈现；映射盘符保留盘符作为入口，点击仍进盘根（2026-11-13 杨总指示）
    const toHostRoot = (p) => {
      const s = String(p);
      if (s.startsWith('\\\\')) {
        const host = s.slice(2).split('\\')[0];
        if (host) return '\\\\' + host;
      }
      return s;
    };
  netCache = arr.filter(x => x && x.path && /^\\\\/.test(x.path))
    .map(x => ({ path: toHostRoot(x.path), letter: x.letter || null }))
    .filter((x, i, a) => a.findIndex(y => (y.letter || '') === (x.letter || '') && y.path.toLowerCase() === x.path.toLowerCase()) === i);
  } catch (e) {
    netCache = netCache || [];   // 失败/超时也缓存（10 分钟内不重试），防反复卡（2026-11-16）
  }
  netCacheAt = now;
  return netCache;
}

// —— 桌面集成（所有位置页的桌面模块）——
function getDesktopDir() {
  // 首选注册表 User Shell Folders\Desktop（兼容桌面重定向到 OneDrive 等场景），回退 %USERPROFILE%\Desktop
  try {
    const q = spawnSync('reg', ['query',
      'HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Explorer\\User Shell Folders', '/v', 'Desktop'],
      { encoding: 'utf8', windowsHide: true });
    if (q.status === 0) {
      const m = String(q.stdout).match(/Desktop\s+REG_(?:EXPAND_)?SZ\s+(.+)/i);
      let p = m && m[1] ? m[1].trim() : '';
      if (p) {
        p = p.replace(/%([^%]+)%/g, (_, k) => process.env[k] || ('%' + k + '%'));
        if (fs.existsSync(p)) return p;
      }
    }
  } catch (_) {}
  return path.join(process.env.USERPROFILE || os.homedir(), 'Desktop');
}

async function desktopInfo(res, cfg) {
  const dir = getDesktopDir();
  // 受 roots 约束校验：受限配置下桌面不可见时，前端隐藏/降级显示桌面模块
  resolveAllowedPath(dir, cfg);
  const st = await fs.promises.stat(dir);
  if (!st.isDirectory()) throw new Error('desktop not found');
  const ents = await fs.promises.readdir(dir, { withFileTypes: true });
  const items = [];
  for (const ent of ents) {
    const full = path.win32.join(dir, ent.name);
    if (isDenied(full, cfg)) continue;
    let type = 'other', size = 0, mtime = 0;
    try {
      const s = await fs.promises.stat(full);
      type = s.isDirectory() ? 'dir' : s.isFile() ? 'file' : 'other';
      size = s.size;
      mtime = s.mtimeMs;
    } catch (_) {}
    items.push({
      name: ent.name, path: full, type: type, size: size, mtime: mtime,
      ext: path.win32.extname(ent.name).toLowerCase().slice(1)
    });
  }
  sendJson(res, 200, { ok: true, path: dir, items: items });
}

// ============ 文件元数据索引（SQLite，2026-11-17）============
// 服务启动即后台全局扫描建库（位置/文件名/类型/大小/修改时间）；
// 页面进入任一非盘符/非根主机目录 = 用户真实连接 → 全局扫描立即停止且永不再启（用户优先，
// 宁可库不完整也绝不抢用户 IO）；未扫完的目录靠进入时的局部扫描逐目录补。
// /api/ls 优先查库秒回（fromDb:true），前端随后以 ?fresh=1 触发该目录局部扫描校准
// （增删改=该目录行全量替换入库）并刷新界面；库内没有的冷目录直接走实时扫描兜底。
// 索引任何异常（损坏/被锁/环境异常）整体降级为实时扫描（等同旧行为），绝不阻塞请求。
const INDEX_DB_PATH = path.join(INSTALL_DIR, 'file-index.db');
let _idxDb = null, _idxDead = false;
function indexDb() {
  if (_idxDead) return null;
  if (_idxDb) return _idxDb;
  try {
    const { Database } = require('bun:sqlite');
    const db = new Database(INDEX_DB_PATH, { create: true });
    db.exec('PRAGMA journal_mode=WAL; PRAGMA busy_timeout=5000; PRAGMA synchronous=NORMAL;');
    db.exec(
      'CREATE TABLE IF NOT EXISTS dirs(path TEXT PRIMARY KEY COLLATE NOCASE, scanned_at INTEGER);' +
      'CREATE TABLE IF NOT EXISTS files(' +
      ' path TEXT PRIMARY KEY COLLATE NOCASE, parent TEXT NOT NULL COLLATE NOCASE,' +
      ' name TEXT NOT NULL, type TEXT NOT NULL, size INTEGER NOT NULL, mtime INTEGER NOT NULL, ext TEXT NOT NULL);' +
      'CREATE INDEX IF NOT EXISTS idx_files_parent ON files(parent);');
    _idxDb = db;
    return db;
  } catch (e) {
    _idxDead = true;
    log('index db disabled: ' + (e && e.message || e));
    return null;
  }
}

// 全局扫描停止标志（一次性、永不再启）
let _scanStopped = false;
function stopGlobalScan() {
  if (_scanStopped) return;
  _scanStopped = true;
  log('global scan stopped by user connection');
}

// 页面进入任一非盘符/非根主机目录 = 用户真实连接（2026-11-17 定）：
// 盘符根（C:\）与裸主机根（\\host）不触发，进入任一更深层目录即触发停止
function isRootLevel(dir) {
  return isBareHost(dir) || /^[A-Za-z]:[\\\/]?$/.test(dir);
}

// 局部校准写库：增删改 = 该目录行全量替换（单事务原子），dirs 记"已扫过"标记
function writeDirIndex(dir, items) {
  const db = indexDb();
  if (!db) return;
  const now = Date.now();
  db.transaction(() => {
    db.query('DELETE FROM files WHERE parent = ?').run(dir);
    const ins = db.query('INSERT OR REPLACE INTO files(path,parent,name,type,size,mtime,ext) VALUES(?,?,?,?,?,?,?)');
    for (const it of items) ins.run(it.path, dir, it.name, it.type, it.size || 0, it.mtime || 0, it.ext || '');
    db.query('INSERT OR REPLACE INTO dirs(path,scanned_at) VALUES(?,?)').run(dir, now);
  })();
}

// 查库：{scanned, items}；scanned=false = 冷目录（走实时扫描兜底，绝不空屏）
function queryDirIndex(dir) {
  const db = indexDb();
  if (!db) return { scanned: false, items: [] };
  const d = db.query('SELECT path FROM dirs WHERE path = ?').get(dir);
  if (!d) return { scanned: false, items: [] };
  const rows = db.query('SELECT path,name,type,size,mtime,ext FROM files WHERE parent = ?').all(dir);
  return {
    scanned: true,
    items: rows.map(r => ({ name: r.name, path: r.path, type: r.type, size: r.size, mtime: r.mtime, ext: r.ext }))
  };
}

// —— 写操作库同步（2026-11-17）：页面删除/新建/移动/复制等写操作逐行同步库，
// 消除"写后读到旧数据"的幽灵窗口；目录级影响（删/改名/移动目录、解压、批量恢复）
// 用 indexInvalidate 子树失效兜底（下次进入实时校准自愈）。任何异常只记日志，
// 绝不影响写操作本身——库坏了写操作照样成功（实时扫描永远兜底）。
function indexRowUpsert(parent, name, type, size, mtime) {
  const db = indexDb(); if (!db) return;
  try {
    const full = path.win32.join(parent, name);
    db.query('INSERT OR REPLACE INTO files(path,parent,name,type,size,mtime,ext) VALUES(?,?,?,?,?,?,?)')
      .run(full, parent, name, type, size || 0, mtime || 0, (path.win32.extname(name) || '').toLowerCase());
  } catch (e) { log('index upsert skip: ' + (e && e.message || e)); }
}
function indexInvalidate(full) {
  // 子树失效：清自己（含父目录里的行）与整个子树的"已扫过"标记和条目行；
  // substr 前缀匹配防 LIKE 通配符注入，COLLATE NOCASE 对齐路径大小写
  const db = indexDb(); if (!db) return;
  try {
    const pre = String(full) + '\\';
    db.query('DELETE FROM dirs WHERE path = ? COLLATE NOCASE OR substr(path, 1, ?) = ? COLLATE NOCASE').run(full, pre.length, pre);
    db.query('DELETE FROM files WHERE path = ? COLLATE NOCASE OR substr(path, 1, ?) = ? COLLATE NOCASE').run(full, pre.length, pre);
  } catch (e) { log('index invalidate skip: ' + (e && e.message || e)); }
}
function indexMove(src, dst) {
  // 移动/改名：旧路径子树失效 + 新路径子树失效（防旧残留）+ 新行按磁盘现状写入
  indexInvalidate(src);
  indexInvalidate(dst);
  try {
    const st = fs.statSync(dst);
    indexRowUpsert(path.win32.dirname(dst), path.win32.basename(dst), st.isDirectory() ? 'dir' : 'file', st.size, st.mtimeMs);
  } catch (e) {}
}
function indexWrite(full) {
  // 新建/上传/复制/恢复落成后：按磁盘现状写自己一行；目录先做子树失效（防旧残留）
  try {
    const st = fs.statSync(full);
    const isDir = st.isDirectory();
    if (isDir) indexInvalidate(full);
    indexRowUpsert(path.win32.dirname(full), path.win32.basename(full), isDir ? 'dir' : 'file', st.size, st.mtimeMs);
  } catch (e) {}
}

// 源头定序（2026-11-16）：目录优先固定（不随方向翻转），语义与前端 sortItems 一致；
// 库快照与实时扫描共用同一比较函数，两条路径排序结果逐项一致
function sortDirItems(items, sortKey, sortAsc) {
  items.sort((a, b) => {
    if (a.type === 'dir' && b.type !== 'dir') return -1;
    if (b.type === 'dir' && a.type !== 'dir') return 1;
    let r = 0;
    if (sortKey === 'size') r = (a.size || 0) - (b.size || 0);
    else if (sortKey === 'mtime') r = (a.mtime || 0) - (b.mtime || 0);
    else if (sortKey === 'type') r = (a.ext || '').localeCompare(b.ext || '');
    if (r === 0) r = a.name.localeCompare(b.name, 'zh', { numeric: true });
    return sortAsc ? r : -r;
  });
}

function expandScanRoots(cfg) {
  const roots = ((cfg.sources || {}).default || {}).roots || [];
  const out = [];
  for (const r of roots) {
    if (r === '*') {
      for (let c = 65; c <= 90; c++) {
        const d = String.fromCharCode(c) + ':\\';
        try { fs.accessSync(d); out.push(d); } catch (e) {}
      }
    } else out.push(r);
  }
  return out;
}

// 全局扫描建库：目录串行 + 目录内 8 路并发 stat（给前台请求留足 UV 线程池余量，杜绝塞车），
// 每 256 条目检查停止标志；符号链接/junction 目录不入队（防链接环）；无权限/坏盘跳过不入库
// （冷路径会报真实错误）；条目 key 统一由 path.win32.join 生成，与 listDir 同源保证库命中
async function startGlobalScan(cfg) {
  if (!indexDb()) return;
  const queue = expandScanRoots(cfg);
  if (!queue.length) return;
  log('global scan start: ' + queue.join(', '));
  let ndir = 0, nfile = 0;
  while (queue.length) {
    if (_scanStopped) { log('global scan aborted: ' + ndir + ' dirs, ' + nfile + ' files'); return; }
    const dir = queue.shift();
    let items = [];
    try {
      const ents = await fs.promises.readdir(dir, { withFileTypes: true });
      let _n = 0;
      for (const ent of ents) {
        if ((++_n & 255) === 0 && _scanStopped) { log('global scan aborted: ' + ndir + ' dirs'); return; }
        const full = path.win32.join(dir, ent.name);
        const link = ent.isSymbolicLink();
        if (isDenied(full, cfg)) continue;
        items.push({
          name: ent.name, path: full,
          type: ent.isDirectory() ? 'dir' : ent.isFile() ? 'file' : 'other',
          size: 0, mtime: 0,
          ext: (path.win32.extname(ent.name) || '').toLowerCase()
        });
        if (!link && ent.isDirectory()) queue.push(full);
      }
    } catch (e) { continue; }
    // 8 路消费式并发 stat 补齐 size/mtime/type（快者多干，无批栅栏）
    let idx = 0;
    const workers = [];
    for (let w = 0; w < 8 && w < items.length; w++) {
      workers.push((async () => {
        while (idx < items.length) {
          if (_scanStopped) return;
          const it = items[idx++];
          try {
            const s = await fs.promises.stat(it.path);
            it.type = s.isDirectory() ? 'dir' : s.isFile() ? 'file' : 'other';
            it.size = s.size;
            it.mtime = s.mtimeMs;
          } catch (e) {}
        }
      })());
    }
    await Promise.all(workers);
    if (_scanStopped) { log('global scan aborted: ' + ndir + ' dirs'); return; }
    try { writeDirIndex(dir, items); } catch (e) {}
    ndir++; nfile += items.length;
    if ((ndir % 200) === 0) log('global scan: ' + ndir + ' dirs, ' + nfile + ' files');
    await sleep(0);   // 让出事件循环：前台请求永远优先
  }
  log('global scan done: ' + ndir + ' dirs, ' + nfile + ' files');
}


async function listDir(req, res, q, cfg) {
  const dir = resolveAllowedPath(q.path, cfg);
  // 页面进入任一非盘符/非根主机目录 = 用户真实连接：全局扫描立即永久停止（用户优先）
  if (!isRootLevel(dir)) stopGlobalScan();
  // 客户端已断开（AbortController）则放弃剩余处理：否则旧请求跑完才放行新
  // 请求，快速进/返回时服务端排队塞车（2026-11-15）
  const dead = () => !!(req && req.signal && req.signal.aborted);
  // 一次性读取 + 源头定序（2026-11-16）：按用户设定的排序方式一次读全量（readdir+stat）
  // 并在服务端排好序再返回，页面拿到后按序流式加载；不再两阶段补齐（stat-batch）
  const sortKey = (q.sort || '').toLowerCase();      // name|size|mtime|type
  const sortAsc = q.asc !== '0';
  // fast 骨架无条件生效（2026-11-18 两阶段）：即使 sort=size/mtime 也先出骨架
  // （元数据空、同值按名并列序），补齐请求全量 stat 后按用户排序整体替换——
  // 首屏速度优先于骨架期的临时顺序（默认排序就是 mtime，忽略 fast 会让两阶段失效）
  const fast = q.fast === '1' || q.fast === 'true';
  // UNC 主机根（\\host）不能用 readdir 列目录——Windows 共享枚举是另一套 API，
  // 用 net view 列出主机上的可见共享，以目录形式呈现（2026-11-13）
  if (isBareHost(dir)) {
    const names = (await listShares(dir)).filter(n => !isDenied(dir + '\\' + n, cfg));
    const items = names.map((n) => ({
      name: n, path: dir + '\\' + n, type: 'dir', size: 0, mtime: 0, ext: ''
    }));
    items.sort((a, b) => sortAsc
      ? a.name.localeCompare(b.name, 'zh', { numeric: true })
      : b.name.localeCompare(a.name, 'zh', { numeric: true }));
    sendJson(res, 200, { ok: true, path: dir, fast, items });
    return;
  }
  // 快路径（2026-11-17）：库内已扫过的目录查库秒回（fromDb），前端随后 ?fresh=1 局部校准；
  // 冷目录/库禁用/任何异常一律落回实时扫描兜底（=旧行为），绝不空屏、绝不塞车
  if (q.fresh !== '1' && !isBareHost(dir)) {
    try {
      const hit = queryDirIndex(dir);
      if (hit.scanned) {
        const items = hit.items;
        sortDirItems(items, sortKey, sortAsc);
        sendJson(res, 200, { ok: true, path: dir, fromDb: true, items });
        return;
      }
    } catch (e) {}
  }
  const st = await fs.promises.stat(dir);
  if (!st.isDirectory()) throw new Error('not a directory');

  const ents = await fs.promises.readdir(dir, { withFileTypes: true });
  const items = [];
  // realDir：父目录真实路径每目录解析一次；非符号链接条目的真实路径 = 父真实路径 + 条目名，
  // 免逐条目 realpathSync（万级目录曾因此每条 0.12ms，2026-11-15 修）
  let realDir = '';
  if ((cfg.deniedPaths || []).length) {
    try { realDir = path.win32.normalize(fs.realpathSync(dir)).replace(/[\\\/]+$/, ''); } catch (e) {}
  }
  let _n = 0;
  for (const ent of ents) {
    if ((++_n & 255) === 0 && dead()) return;   // 每 256 条检查一次客户端是否已放弃
    const full = path.win32.join(dir, ent.name);
    const link = ent.isSymbolicLink();
    if (isDenied(full, cfg, link || !realDir ? undefined : realDir + '\\' + ent.name)) continue;
    let type = 'other', size = 0, mtime = 0;
    if (fast) {
      // Dirent 自带类型（零系统调用）；符号链接/junction 需 stat 跟随判断
      if (link) {
        try {
          const s = await fs.promises.stat(full);
          type = s.isDirectory() ? 'dir' : s.isFile() ? 'file' : 'other';
        } catch (e) { markBadHost(full); }
      } else {
        type = ent.isDirectory() ? 'dir' : ent.isFile() ? 'file' : 'other';
      }
    }
    items.push({
      name: ent.name,
      path: full,
      type,
      size: fast ? null : size,
      mtime: fast ? null : mtime,
      ext: (path.win32.extname(ent.name) || '').toLowerCase()
    });
  }
  if (!fast) {
    // 消费式并发 stat（64 worker）：旧版逐条 await 是 N 次串行系统调用，万级目录耗时
    // 随文件数线性膨胀（2026-11-15 卡顿主因）；并发批的批栅栏会被慢条目拖住，改
    // 无栅栏消费式，快者多干
    let idx = 0;
    const workers = [];
    for (let w = 0; w < 64 && w < items.length; w++) {
      workers.push((async () => {
        while (idx < items.length) {
          const it = items[idx++];
          try {
            const s = await fs.promises.stat(it.path);
            it.type = s.isDirectory() ? 'dir' : s.isFile() ? 'file' : 'other';
            it.size = s.size;
            it.mtime = s.mtimeMs;
          } catch (e) { markBadHost(it.path); }
        }
      })());
    }
    await Promise.all(workers);
  }
  // 源头定序（2026-11-16）：目录优先固定（不随方向翻转），语义与前端 sortItems 一致
  sortDirItems(items, sortKey, sortAsc);

  // 局部校准入库（2026-11-17）：增删改 = 该目录行全量替换，同步落库后才响应——
  // "响应即已入库"语义原子；单目录批量事务毫秒级，不构成塞车。
  // fast 骨架（size/mtime 未 stat）不落库——防占位值污染索引（2026-11-18），
  // 补齐阶段的全量请求会把完整数据落库
  if (!fast) { try { writeDirIndex(dir, items); } catch (e) {} }
  sendJson(res, 200, { ok: true, path: dir, fast, items });
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

// 批量 stat：两阶段加载的第二阶段——前端分批补齐 fast 列表的 size/mtime。
// 单批上限 2000；逐路径过 resolveAllowedPath + isDenied 同门禁，只读。
// 真实路径兜底与逐条 realpathSync 语义等价：parent realpath 每目录缓存一次，
// lstat 判符号链接、仅链接条目逐条 realpath（2026-11-15 提速，防万条逐条系统调用）
async function statBatchOp(req, res, data, cfg) {
  const paths = Array.isArray(data && data.paths) ? data.paths.slice(0, 2000) : [];
  const list = cfg.deniedPaths || [];
  const nkey = (x) => path.win32.normalize(String(x)).replace(/[\\\/]+$/, '').toLowerCase();
  const hit = (rl) => {
    for (const d of list) { if (d && (rl === d || rl.startsWith(d + '\\'))) return true; }
    return false;
  };
  const realCache = new Map();
  const out = await Promise.all(paths.map(async (p) => {
    // 断开检测用 socket 语义：Bun 里 POST body 读完后 req.signal 即 abort，不能当断开用
    if (res.writableEnded || (req.socket && req.socket.destroyed)) return null;
    try {
      const full = resolveAllowedPath(p, cfg);
      if (hit(nkey(full))) return null;
      let s;
      if (list.length) {
        const parent = path.win32.dirname(full);
        let realParent = realCache.get(parent);
        if (realParent === undefined) {
          try { realParent = nkey(fs.realpathSync(parent)); } catch (e) { realParent = ''; }
          realCache.set(parent, realParent);
        }
        const lst = await fs.promises.lstat(full);
        if (lst.isSymbolicLink()) {
          let rl;
          try { rl = nkey(fs.realpathSync(full)); } catch (e) { markBadHost(full); return null; }
          if (hit(rl)) return null;
          s = await fs.promises.stat(full);        // 链接条目跟随取目标真值
        } else {
          if (realParent && hit(realParent + '\\' + nkey(path.win32.basename(full)))) return null;
          s = lst;                                  // 非链接：lstat 即目标元数据
        }
      } else {
        s = await fs.promises.stat(full);
      }
      return {
        path: full,
        type: s.isDirectory() ? 'dir' : s.isFile() ? 'file' : 'other',
        size: s.size,
        mtime: s.mtimeMs
      };
    } catch (e) { markBadHost(p); return null; }
  }));
  sendJson(res, 200, { ok: true, items: out.filter(Boolean) });
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
  const file = resolveAllowedPath(q.path, cfg, { trashRead: true });
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
  let bytes = buf instanceof Uint8Array ? buf : new Uint8Array(buf);
  // 去掉尾部 NUL 填充（TAR 头按定长字段填充），避免稀释极短名的判断（#10）
  let end = bytes.length;
  while (end > 0 && bytes[end - 1] === 0) end--;
  if (end === 0) return '';
  bytes = bytes.slice(0, end);
  const utf8 = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  // 极短名（<=5 字节）：GB18030 双字节序列可能恰好是合法 UTF-8（如 GBK "啊" = B0A1），
  // 两边都解、取"常见字符占比"更高的一侧（#10）；平手取 UTF-8（维持原行为）
  if (bytes.length <= 5) {
    const decode = (enc) => { try { return new TextDecoder(enc, { fatal: false }).decode(bytes); } catch (_) { return null; } };
    const gb = decode('gb18030');
    if (gb && gb !== utf8) {
      const plausible = (s) => {
        const cs = Array.from(s);
        if (!cs.length) return 0;
        let ok = 0;
        for (const ch of cs) {
          const c = ch.codePointAt(0);
          const good = (c >= 0x20 && c <= 0x7e) ||
            (c >= 0xa0 && c <= 0xff && c !== 0xd7 && c !== 0xf7) ||
            (c >= 0x100 && c <= 0x24f) || (c >= 0x370 && c <= 0x4ff) ||
            (c >= 0x3000 && c <= 0x303f) || (c >= 0x3040 && c <= 0x30ff) ||
            (c >= 0x4e00 && c <= 0x9fff) || (c >= 0xac00 && c <= 0xd7af) ||
            (c >= 0xff00 && c <= 0xffef);
          if (good) ok++;
        }
        return ok / cs.length;
      };
      if (plausible(gb) > plausible(utf8)) return gb;
    }
    return utf8;
  }
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
function parseTarEntries(buf, stopAt) {
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
      // 命中目标条目即停止解析（stopAt，#16）：避免大 TAR 全量扫描
      if (stopAt && name === stopAt && !isDir) return entries;
    }
    offset = dataStart + Math.ceil(size / 512) * 512;
  }
  return entries;
}

function extractTarEntry(buf, entryPath) {
  const norm = entryPath.replace(/\\/g, '/').replace(/^\/+/, '');
  const entries = parseTarEntries(buf, norm);
  const target = entries.find(e => e.name === norm && e.type === 'file');
  if (!target) throw new Error('entry not found: ' + entryPath);
  // 越界校验：声明 size 超出实际缓冲区时 Buffer.slice 会静默截断 → 无声数据损坏，显式拒绝
  if (target.dataOffset + target.size > buf.length) throw new Error('corrupted archive: entry size out of range');
  return buf.slice(target.dataOffset, target.dataOffset + target.size);
}

// 7z CLI 集成（RAR / 7Z / 其他 7z 支持的格式）
let _7zExeCache;
function get7zExe() {
  if (_7zExeCache) return _7zExeCache;
  const known = ['C:\\Program Files\\7-Zip\\7z.exe', 'C:\\Program Files (x86)\\7-Zip\\7z.exe'];
  for (const p of known) { try { if (fs.existsSync(p)) return (_7zExeCache = p); } catch (_) {} }
  // 自定义安装路径：查 PATH（#18），找不到给出明确错误而非静默 spawn ENOENT
  try {
    const which = spawnSync('where', ['7z'], { encoding: 'utf8', windowsHide: true });
    if (which.status === 0) {
      const hit = String(which.stdout || '').split(/\r?\n/).map(s => s.trim()).filter(Boolean)[0];
      if (hit) return (_7zExeCache = hit);
    }
  } catch (_) {}
  throw new Error('未找到 7z.exe，请安装 7-Zip 或将其加入 PATH');
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
  const out = await run7z(['l', '-slt', '-y', filePath]);
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
  // -spd: 禁用通配符匹配（entryPath 含 * ? 时按字面处理）; -y: 假设全部提示为是; --: 终止选项解析（防 entryPath 以 - 开头被当选项）
  const out = await run7z(['x', filePath, '-so', '-spd', '-y', '--', entryPath]);
  if (!out.length) {
    // 空输出：区分"真实空文件"与"条目不存在"（-spd 后通配符不再匹配会静默给出空输出）
    const entries = await list7zEntries(filePath);
    const norm = entryPath.replace(/\\/g, '/').replace(/^\/+/, '');
    if (!entries.some(e => e.name.replace(/\\/g, '/') === norm && e.type === 'file')) {
      throw new Error('entry not found: ' + entryPath);
    }
  }
  return out;
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
    // 单文件 GZ：虚拟条目名为去除 .gz/.tgz 后缀的文件名，须精确匹配（否则会把整个解压内容错发给任意请求）
    const baseName = path.basename(filePath).replace(/\.t?gz$/i, '') || 'content';
    if (entryPath.replace(/\\/g, '/') !== baseName) throw new Error('entry not found: ' + entryPath);
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
    return a.name.localeCompare(b.name, 'zh', { numeric: true });
  });
  return items;
}

async function archiveList(res, q, cfg) {
  const file = resolveAllowedPath(q.path, cfg, { trashRead: true });
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
  const file = resolveAllowedPath(q.path, cfg, { trashRead: true });
  const entry = String(q.entry || '');
  if (!entry) throw new Error('missing entry');
  for (const p of [entry, file]) {
    const extCheck = checkExtension(p, cfg);
    if (!extCheck.ok) return sendJson(res, 403, { ok: false, error: extCheck.reason });
  }
  const data = await extractEntryFromArchive(file, entry);
  // 猜测 MIME：复用 getMime 全量表（#15，此前独立简表缺 20+ 种类型）
  const mime = getMime(path.extname(entry).toLowerCase().slice(1));
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

// —— 文件缩略图 /api/thumb（2026-11-16）：文件列表图标用低清小图（64px 长边）——
// PSD/PSB 复用内嵌 JPEG 提取（资源区段在文件头部，只读头 32MB）；
// AI 尽力提取（PDF 兼容型扫内嵌 JPEG / DOS EPS 头取 TIFF 预览转 JPEG）；
// 图片走 GDI+ 缩放（PowerShell System.Drawing）存 cacheDir/thumbs（键含 mtime+size）；
// svg 原样回传（浏览器自缩放）。提不出缩略图回 404，前端保留原图标。
const THUMB_EDGE = 64;

function readHead(file, n) {
  const fd = fs.openSync(file, 'r');
  try {
    const buf = Buffer.alloc(n);
    const got = fs.readSync(fd, buf, 0, n, 0);
    return buf.subarray(0, got);
  } finally { try { fs.closeSync(fd); } catch (e) {} }
}

function readTail(file, n, size) {
  const fd = fs.openSync(file, 'r');
  try {
    const buf = Buffer.alloc(n);
    const got = fs.readSync(fd, buf, 0, n, Math.max(0, size - n));
    return buf.subarray(0, got);
  } finally { try { fs.closeSync(fd); } catch (e) {} }
}

// 扫描内嵌 JPEG（SOI..EOI）；小于 1KB 的碎片（图标噪声）跳过继续找
function scanFirstJpeg(buf) {
  const sig = Buffer.from([0xFF, 0xD8, 0xFF]);
  let p = 0;
  while (p < buf.length) {
    const soi = buf.indexOf(sig, p);
    if (soi < 0) return null;
    let end = -1;
    for (let i = soi + 3; i + 1 < buf.length; i++) {
      if (buf[i] === 0xFF && buf[i + 1] === 0xD9) { end = i + 2; break; }
    }
    if (end < 0) return null;
    if (end - soi >= 1024) return buf.subarray(soi, end);
    p = end;
  }
  return null;
}

// DOS EPS 二进制头：magic(4)=C5 D3 D6 C6 + 5 组 uint32 偏移/长度 + uint16 校验，
// TIFF 预览偏移/长度在第 20~27 字节；字节序双向验证（TIFF 魔数 II*/MM* 对得上才算）
function extractEpsTiff(head) {
  if (head.length < 30) return null;
  if (!(head[0] === 0xC5 && head[1] === 0xD3 && head[2] === 0xD6 && head[3] === 0xC6)) return null;
  for (const be of [true, false]) {
    const off = be ? head.readUInt32BE(20) : head.readUInt32LE(20);
    const len = be ? head.readUInt32BE(24) : head.readUInt32LE(24);
    if (off <= 0 || len <= 24 || off + len > head.length) continue;
    const sig = head.subarray(off, off + 4);
    const tiff = (sig[0] === 0x49 && sig[1] === 0x49 && sig[2] === 0x2A && sig[3] === 0x00) ||
                 (sig[0] === 0x4D && sig[1] === 0x4D && sig[2] === 0x00 && sig[3] === 0x2A);
    if (tiff) return head.subarray(off, off + len);
  }
  return null;
}

// PSD raw RGB 缩略图（8BIM 1033/1036，format 1）：BGR 底行在前、行宽 widthBytes
function extractPsdRawThumb(buf) {
  if (buf.length < 30) return null;
  if (buf[0] !== 0x38 || buf[1] !== 0x42 || buf[2] !== 0x50 || buf[3] !== 0x53) return null;
  let offset = 26;
  try {
    const cmLen = buf.readUInt32BE(offset); offset += 4 + cmLen;
    const irLen = buf.readUInt32BE(offset);
    const end = Math.min(offset + 4 + irLen, buf.length);
    offset += 4;
    while (offset + 12 < end) {
      if (buf[offset] !== 0x38 || buf[offset + 1] !== 0x42 || buf[offset + 2] !== 0x49 || buf[offset + 3] !== 0x4D) break;
      offset += 4;
      const resourceId = buf.readUInt16BE(offset); offset += 2;
      const nameLen = buf[offset];
      offset += 1 + nameLen;
      if ((nameLen + 1) % 2 !== 0) offset++;
      if (offset + 4 > end) break;
      const dataSize = buf.readUInt32BE(offset); offset += 4;
      if ((resourceId === 0x0409 || resourceId === 0x040C) && offset + 24 <= end) {
        const format = buf.readUInt32BE(offset);
        if (format === 1) {
          const w = buf.readUInt32BE(offset + 4), h = buf.readUInt32BE(offset + 8);
          const wbytes = buf.readUInt32BE(offset + 12), total = buf.readUInt32BE(offset + 16);
          if (w > 0 && h > 0 && w <= 4096 && h <= 4096 && wbytes >= w * 3 && wbytes * h <= 64 * 1048576) {
            // 像素区必须裁在资源载荷边界内：total 字段可能虚大，越界续读=跨资源垃圾噪点
            const payloadEnd = Math.min(offset + dataSize, end);
            const pixEnd = Math.min(offset + 24 + Math.min(total || wbytes * h, wbytes * h), payloadEnd);
            const pix = buf.subarray(offset + 24, pixEnd);
            // 载荷不足整幅 = 不是真 raw（可能是带前缀的 JPEG 变体），弃用防噪点
            if (pix.length >= wbytes * h) return { w, h, wbytes, pix };
          }
        }
      }
      offset += dataSize;
      if (dataSize % 2 !== 0) offset++;
    }
  } catch (e) {}
  return null;
}

// PSD 1033/1036 资源载荷内 JPEG 扫描（缩略图专用变体兜底）：
// 真实文件变体实证：fmt=1 但载荷实为 JPEG（4 字节前缀 + FFD8..FFD9）。
// 只给 /api/thumb 管线用；预览窗格（officePreview）走 extractPsdThumbnail 原逻辑，严禁混用。
function extractPsdThumbVariant(buf) {
  if (buf.length < 26) return null;
  if (buf[0] !== 0x38 || buf[1] !== 0x42 || buf[2] !== 0x50 || buf[3] !== 0x53) return null;
  let offset = 26;
  if (offset + 4 > buf.length) return null;
  const colorModeLen = buf.readUInt32BE(offset);
  offset += 4 + colorModeLen;
  if (offset + 4 > buf.length) return null;
  const resourcesLen = buf.readUInt32BE(offset);
  offset += 4;
  const resourcesEnd = Math.min(offset + resourcesLen, buf.length);
  while (offset + 12 < resourcesEnd) {
    if (buf[offset] !== 0x38 || buf[offset + 1] !== 0x42 || buf[offset + 2] !== 0x49 || buf[offset + 3] !== 0x4D) break;
    offset += 4;
    const resourceId = buf.readUInt16BE(offset);
    offset += 2;
    const nameLen = buf[offset];
    offset += 1 + nameLen;
    if ((nameLen + 1) % 2 !== 0) offset++;
    if (offset + 4 > buf.length) break;
    const dataSize = buf.readUInt32BE(offset);
    offset += 4;
    if (resourceId === 0x0409 || resourceId === 0x040C) {
      const payloadEnd = Math.min(offset + dataSize, buf.length);
      const found = scanFirstJpeg(buf.subarray(offset, payloadEnd));
      if (found) return found;
    }
    offset += dataSize;
    if (dataSize % 2 !== 0) offset++;
  }
  return null;
}

// RGBA（顶行在前）→ 24bpp BMP（底行在前、行 4 字节对齐），给 GDI+ 当中间格式
function rgbaToBmp(rgba, w, h) {
  const rowRaw = w * 3;
  const pad = (4 - rowRaw % 4) % 4;
  const rowSz = rowRaw + pad;
  const dataSz = rowSz * h;
  const bmp = Buffer.alloc(54 + dataSz);
  bmp.write('BM', 0);
  bmp.writeUInt32LE(54 + dataSz, 2);
  bmp.writeUInt32LE(54, 10);
  bmp.writeUInt32LE(40, 14);
  bmp.writeInt32LE(w, 18);
  bmp.writeInt32LE(h, 22);
  bmp.writeUInt16LE(1, 26);
  bmp.writeUInt16LE(24, 28);
  bmp.writeUInt32LE(dataSz, 34);
  for (let y = 0; y < h; y++) {
    const srcY = h - 1 - y;
    let o = 54 + y * rowSz;
    for (let x = 0; x < w; x++) {
      const s = (srcY * w + x) * 4;
      bmp[o++] = rgba[s + 2]; bmp[o++] = rgba[s + 1]; bmp[o++] = rgba[s];
    }
  }
  return bmp;
}

// raw RGB 缩略图（BGR 底行在前）→ RGBA（顶行在前）
function rawThumbToRgba(raw) {
  const w = raw.w, h = raw.h, wbytes = raw.wbytes, pix = raw.pix;
  const rgba = Buffer.alloc(w * h * 4);
  for (let y = 0; y < h; y++) {
    const src = (h - 1 - y) * wbytes;
    for (let x = 0; x < w; x++) {
      const s = src + x * 3, d = (y * w + x) * 4;
      rgba[d] = pix[s + 2]; rgba[d + 1] = pix[s + 1]; rgba[d + 2] = pix[s]; rgba[d + 3] = 255;
    }
  }
  return { rgba: rgba, w: w, h: h };
}

// raw RGB 缩略图（BGR 底行在前）→ BMP（备用桥接）
function rawThumbToBmp(raw) {
  const r = rawThumbToRgba(raw);
  return rgbaToBmp(r.rgba, r.w, r.h);
}

// 64px 长边 JPEG 缩略图（q75、不放大）。jimp 纯 JS 解码缩放——无 native 依赖、
// 完整打包进 exe（sharp 的 .node 走动态加载路径，bun build --compile 嵌不进）；
// 成功写 outFile 返回 true。长尾格式（ICO 等）退回 PowerShell GDI+ 兜底。
async function gdiThumb(srcFile, outFile) {
  try {
    const im = await Jimp.fromBuffer(fs.readFileSync(srcFile), JIMP_DECODE_OPTS);
    await writeThumbJpeg(im, outFile);
    return true;
  } catch (e) {
    return gdiFallback(srcFile, outFile);
  }
}

// raw 像素直喂（PSD raw RGB / 合成图桥接，免 BMP 转码）；RGBA 顶行在前
async function gdiThumbRgba(rgba, w, h, outFile) {
  try {
    await writeThumbJpeg(new Jimp({ data: Buffer.from(rgba), width: w, height: h }), outFile);
    return true;
  } catch (e) { return false; }
}

// 解码选项：jpeg-js 有 decompression-bomb 内存/分辨率上限，按本机服务端放开
const JIMP_DECODE_OPTS = { 'image/jpeg': { maxMemoryUsageInMB: 1024, maxResolutionInMP: 100 } };

// 等比缩放到 THUMB_EDGE 内（不放大）→ JPEG q75 写盘
async function writeThumbJpeg(im, outFile) {
  const w = im.width, h = im.height;
  if (w > THUMB_EDGE || h > THUMB_EDGE) {
    const nw = w >= h ? THUMB_EDGE : Math.max(1, Math.round(w * THUMB_EDGE / h));
    const nh = w >= h ? Math.max(1, Math.round(h * THUMB_EDGE / w)) : THUMB_EDGE;
    im.resize({ w: nw, h: nh });
  }
  const buf = await im.getBuffer('image/jpeg', { quality: 75 });
  fs.writeFileSync(outFile, buf);
}

// PowerShell GDI+ 兜底（sharp 不支持的输入格式）。EncodedCommand 传脚本（路径引号/特殊
// 字符零注入），异步子进程 + 15s 强杀（Bun.spawn，同 listShares 模式）；成功写 outFile 返回 true
async function gdiFallback(srcFile, outFile) {
  const s = String(srcFile).replace(/'/g, "''");
  const o = String(outFile).replace(/'/g, "''");
  const ps = [
    'Add-Type -AssemblyName System.Drawing',
    "$src = New-Object System.Drawing.Bitmap('" + s + "')",
    '$max = ' + THUMB_EDGE,
    '$w = $src.Width; $h = $src.Height',
    'if ($w -lt 1 -or $h -lt 1) { exit 1 }',
    'if ($w -le $max -and $h -le $max) { $nw = $w; $nh = $h } elseif ($w -ge $h) { $nw = $max; $nh = [Math]::Max(1, [int]($h * $max / $w)) } else { $nh = $max; $nw = [Math]::Max(1, [int]($w * $max / $h)) }',
    '$dst = New-Object System.Drawing.Bitmap($nw, $nh)',
    '$g = [System.Drawing.Graphics]::FromImage($dst)',
    '$g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic',
    '$g.DrawImage($src, 0, 0, $nw, $nh)',
    '$g.Dispose(); $src.Dispose()',
    "$enc = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }",
    '$ep = New-Object System.Drawing.Imaging.EncoderParameters(1)',
    '$ep.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter([System.Drawing.Imaging.Encoder]::Quality, [long]75)',
    "$dst.Save('" + o + "', $enc, $ep)",
    '$dst.Dispose()'
  ].join('\n');
  try {
    const proc = Bun.spawn(['powershell', '-NoProfile', '-NonInteractive', '-EncodedCommand',
      Buffer.from(ps, 'utf16le').toString('base64')], { stdout: 'ignore', stderr: 'ignore' });
    const killTimer = setTimeout(() => { try { proc.kill(); } catch (e) {} }, 15000);
    const code = await proc.exited;
    clearTimeout(killTimer);
    return code === 0 && fs.existsSync(outFile);
  } catch (e) { return false; }
}

function sendImageFile(res, file, mime) {
  const buf = fs.readFileSync(file);
  res.writeHead(200, {
    'Content-Type': mime,
    'Content-Length': buf.length,
    'Cache-Control': 'private, max-age=86400'
  });
  res.end(buf);
}

async function thumbFile(res, q, cfg) {
  const file = resolveAllowedPath(q.path, cfg, {});
  // 快捷方式图标分支（2026-11-18）：.lnk 不在黑名单时继承目标原图标；在黑名单 → 全面封禁走统一拦截
  if (isLnkPath(file)) return await shortcutIconOp(res, file, cfg);   // 封禁语义由其内部 selfCheck 统一给出
  const chk = checkExtension(file, cfg);
  if (!chk.ok) return sendJson(res, 403, { ok: false, error: chk.reason });
  const st = await fs.promises.stat(file);
  if (!st.isFile()) throw new Error('not a file');
  const ext = path.extname(file).toLowerCase();
  const MIME = { '.svg': 'image/svg+xml', '.webp': 'image/webp', '.avif': 'image/avif', '.png': 'image/png', '.gif': 'image/gif', '.jpg': 'image/jpeg', '.jpeg': 'image/jpeg', '.bmp': 'image/bmp', '.ico': 'image/x-icon' };

  // svg 原样回传（浏览器自缩放），限 2MB
  if (ext === '.svg') {
    if (st.size > 2 * 1048576) return sendJson(res, 404, { ok: false, error: 'too large' });
    return sendImageFile(res, file, MIME[ext]);
  }

  const tdir = path.join(cfg.cacheDir, 'thumbs');
  if (!fs.existsSync(tdir)) fs.mkdirSync(tdir, { recursive: true });
  const key = crypto.createHash('sha1').update(file + '|' + st.mtimeMs + '|' + st.size).digest('hex').slice(0, 20);
  const outFile = path.join(tdir, key + '.jpg');
  if (fs.existsSync(outFile)) return sendImageFile(res, outFile, 'image/jpeg');

  const head = readHead(file, ext === '.ai' ? 4 * 1048576 : 64 * 1048576);

  // PSD/PSB 三级兜底：内嵌 JPEG（1033/1036 format 0）→ raw RGB 缩略图（format 1）→ 合成图
  if (ext === '.psd' || ext === '.psb') {
    const thumb = extractPsdThumbnail(head);
    if (thumb && thumb.length >= 16 && thumb[0] === 0xFF && thumb[1] === 0xD8) {
      fs.writeFileSync(outFile, thumb);
      return sendImageFile(res, outFile, 'image/jpeg');
    }
    const raw = extractPsdRawThumb(head);
    if (raw) {
      const rr = rawThumbToRgba(raw);
      const ok = await gdiThumbRgba(rr.rgba, rr.w, rr.h, outFile);
      if (ok) return sendImageFile(res, outFile, 'image/jpeg');
    }
    // 变体兜底（缩略图专用）：fmt=1 但载荷实为 JPEG（4 字节前缀 + FFD8..FFD9），魔数扫描
    const vjpg = extractPsdThumbVariant(head);
    if (vjpg) {
      fs.writeFileSync(outFile, vjpg);
      return sendImageFile(res, outFile, 'image/jpeg');
    }
    // 合成图兜底（无缩略图资源的 PSD）：仅小文件、且合成数据是 raw/RLE（zip 压缩解析器不支持，防全黑假图）
    if (st.size <= 64 * 1048576) {
      try {
        let off = 26;
        const cmL = head.readUInt32BE(off); off += 4 + cmL;
        const irL = head.readUInt32BE(off); off += 4 + irL;
        const lmL = head.readUInt32BE(off); off += 4 + lmL;
        const comp0 = head.readUInt16BE(off);
        if (comp0 <= 1) {
          const comp = extractPsdComposite(head);
          if (comp && comp.width > 0 && comp.height > 0 && comp.width * comp.height <= 20000000) {
            const ok = await gdiThumbRgba(comp.rgba, comp.width, comp.height, outFile);
            if (ok) return sendImageFile(res, outFile, 'image/jpeg');
          }
        }
      } catch (e) {}
    }
    return sendJson(res, 404, { ok: false, error: 'no embedded thumbnail' });
  }

  // AI 三级兜底：DOS EPS 头 TIFF 预览 → PDF 兼容段内嵌 JPEG → 全文扫 JPEG
  // 现代 AI（%!PS 文本头）的 PDF 兼容段贴在文件尾部，所以扫描分头/尾两段
  if (ext === '.ai') {
    const tif = extractEpsTiff(head);
    if (tif && tif.length < 20 * 1048576) {
      const tifTmp = path.join(tdir, key + '.tif');
      fs.writeFileSync(tifTmp, tif);
      const ok = await gdiThumb(tifTmp, outFile);
      try { fs.unlinkSync(tifTmp); } catch (e) {}
      if (ok) return sendImageFile(res, outFile, 'image/jpeg');
    }
    let scanBufs;
    if (st.size <= 16 * 1048576) scanBufs = [readHead(file, st.size)];
    else scanBufs = [readTail(file, 12 * 1048576, st.size), head];   // 尾部先扫（PDF 段通常贴尾）
    for (const b of scanBufs) {
      const jpg = scanFirstJpeg(b);
      if (jpg) { fs.writeFileSync(outFile, jpg); return sendImageFile(res, outFile, 'image/jpeg'); }
    }
    return sendJson(res, 404, { ok: false, error: 'no embedded preview' });
  }

  // 图片：GDI+ 缩放（200MB 以上不缩）；GDI+ 认不了但浏览器能显示的格式小文件原样回传
  if (st.size <= 200 * 1048576 && await gdiThumb(file, outFile)) {
    return sendImageFile(res, outFile, 'image/jpeg');
  }
  if (MIME[ext] && st.size < 2 * 1048576) return sendImageFile(res, file, MIME[ext]);
  return sendJson(res, 404, { ok: false, error: 'no thumbnail' });
}

async function officePreview(res, q, cfg) {
  // 支持 :: 虚拟路径（压缩包内文件）
  let file, buf;
  const zipSep = String(q.path || '').indexOf('::');
  if (zipSep > 0) {
    const archivePath = q.path.substring(0, zipSep);
    const entry = q.path.substring(zipSep + 2);
    file = resolveAllowedPath(archivePath, cfg, { trashRead: true });
    buf = await extractEntryFromArchive(file, entry);
    if (!buf) throw new Error('entry not found: ' + entry);
  } else {
    file = resolveAllowedPath(q.path, cfg, { trashRead: true });
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
      const file = resolveAllowedPath(q.path, cfg, { trashRead: true });
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
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    doc: 'application/msword', xls: 'application/vnd.ms-excel', ppt: 'application/vnd.ms-powerpoint',
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
  indexWrite(full);   // 库同步
  sendJson(res, 200, { ok: true, path: full });
}

async function mkfileOp(res, data, cfg) {
  const dir = resolveAllowedPath(data.path, cfg);
  const name = String(data.name || '').trim();
  if (!name || /[\\\/:*?"<>|]/.test(name) || name === '.' || name === '..' || /[. ]$/.test(name)) {
    throw new Error('invalid file name');
  }
  const full = resolveAllowedPath(path.win32.join(dir, name), cfg);
  // 扩展名黑名单前置（2026-11-18 定正）：.lnk 在黑名单 → 新建快捷方式同样封禁，跟随配置放行/封禁
  const extCheck = checkExtension(full, cfg);
  if (!extCheck.ok) return sendJson(res, 403, { ok: false, error: extCheck.reason });
  // —— 新建菜单"快捷方式"专用流：仅允许 .lnk 且带用户指定目标 ——
  const shortcutTarget = data.shortcutTarget ? String(data.shortcutTarget).trim() : '';
  if (shortcutTarget) {
    if (path.win32.extname(name).toLowerCase() !== '.lnk') return sendJson(res, 400, { ok: false, error: 'shortcut must be .lnk' });
    if (shortcutTarget.length > 2000 || /[\x00-\x1f]/.test(shortcutTarget)) return sendJson(res, 400, { ok: false, error: 'invalid shortcut target' });
    if (fs.existsSync(full)) return sendJson(res, 409, { ok: false, error: 'file exists' });
    await createShortcutLnk(full, shortcutTarget);
    indexWrite(full);   // 库同步
    return sendJson(res, 200, { ok: true, path: full });
  }
  // —— 与 Windows 同源：按 ShellNew 注册的创建方式生成新文件（模板拷贝/初始数据/创建命令/空文件），绝不覆盖已存在文件 ——
  const types = await getNewFileTypes();
  const ext = path.win32.extname(name).toLowerCase();
  const t = types.find(x => x.ext === ext) || null;
  if (t && t.kind === 'template' && t.template) {
    const tpl = expandEnvVars(t.template);
    if (fs.existsSync(tpl)) {
      await fs.promises.copyFile(tpl, full, fs.constants.COPYFILE_EXCL);
    } else {
      await fs.promises.writeFile(full, '', { flag: 'wx' });
    }
  } else if (t && t.kind === 'data' && t.data) {
    await fs.promises.writeFile(full, Buffer.from(t.data, 'base64'), { flag: 'wx' });
  } else if (t && t.kind === 'command' && t.command) {
    if (fs.existsSync(full)) throw new Error('file exists');
    await createViaShellNewCommand(t.command, full);
  } else {
    await fs.promises.writeFile(full, '', { flag: 'wx' });   // 已存在则报错，绝不覆盖
  }
  indexWrite(full);   // 库同步
  sendJson(res, 200, { ok: true, path: full });
}

// —— ShellNew 新建文件辅助 ——
function expandEnvVars(s) {
  return String(s).replace(/%([^%]+)%/g, (m, k) => {
    const v = process.env[k] != null ? process.env[k] : process.env[k.toUpperCase()];
    return v != null ? v : m;
  });
}

function parseCommandLine(line) {
  // CommandLineToArgvW 风格简单分词（双引号包裹成单参数，引号本身剥离）
  const args = []; let cur = ''; let inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') { inQ = !inQ; continue; }
    if (!inQ && (ch === ' ' || ch === '\t')) { if (cur !== '') { args.push(cur); cur = ''; } continue; }
    cur += ch;
  }
  if (cur !== '') args.push(cur);
  return args;
}

async function createViaShellNewCommand(commandLine, fullPath) {
  // 与 Windows 行为一致：执行注册表里的创建命令，%1 = 新文件路径。
  // 直接按 argv 启动目标程序（不经 cmd shell），%1 只作独立参数传入，杜绝命令注入面。
  try {
    const argv = parseCommandLine(expandEnvVars(commandLine))
      .map(a => /%1/i.test(a) ? a.replace(/%1/gi, fullPath) : a);
    if (argv.length && argv[0]) {
      Bun.spawn([argv[0], ...argv.slice(1)], { stdin: 'ignore', stdout: 'ignore', stderr: 'ignore' });
    }
  } catch (e) { /* 启动失败走兜底 */ }
  for (let i = 0; i < 30; i++) {   // 等命令把文件建出来（最多 3 秒）
    if (fs.existsSync(fullPath)) return;
    await new Promise(r => setTimeout(r, 100));
  }
  await fs.promises.writeFile(fullPath, '', { flag: 'wx' }).catch(() => {});   // 兜底：空文件
}

async function createShortcutLnk(fullPath, target) {
  // WScript.Shell COM 生成真正的 .lnk；脚本经 -EncodedCommand 传递（UTF-16LE base64），
  // 目标字符串只做 PS 单引号转义，不经过任何 shell 解析，杜绝命令注入面。
  const esc = (s) => String(s).replace(/'/g, "''");
  const script = "$ErrorActionPreference='Stop'\n"
    + "$w=New-Object -ComObject WScript.Shell\n"
    + "$sc=$w.CreateShortcut('" + esc(fullPath) + "')\n"
    + "$sc.TargetPath='" + esc(target) + "'\n"
    + "$sc.Save()";
  const encoded = Buffer.from(script, 'utf16le').toString('base64');
  const proc = Bun.spawn(['powershell', '-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
    { stdout: 'ignore', stderr: 'ignore' });
  await proc.exited;
  if (!fs.existsSync(fullPath)) throw new Error('shortcut create failed');
}

// —— 快捷方式追溯（2026-11-18）：.lnk 不再一刀切拒绝，解析真实目标按目标类型/策略处理 —— //
function isLnkPath(p) {
  return path.win32.extname(String(p || '').replace(/[. ]+$/, '')).toLowerCase() === '.lnk';
}

// .lnk 二进制兜底解析（2026-11-18 v2，客户案例驱动）：部分网络盘/NAS/资源管理器生成的快捷
// 方式 COM 读回 TargetPath 为空（目标只存 ItemIDList/非常规块）；解析 LinkInfo/环境变量块的
// 结构化路径，再扫内嵌 UNC/盘符字符串救回。候选按【可靠度】排序而非“最长获胜”：
//   结构化 LinkInfo 拼接/ExpString(Unicode) > UTF-16LE 扫描（中文路径真身）> ANSI 扫描 >
//   latin1 容错扫描（GBK 断串 mojibake，垫底）；与 .lnk 同名的目标加权（共享空间.lnk
//   → ...\共享空间）。长垃圾不再抢戏（客户症状：中文 ANSI 断串 + 长垃圾赢真路径）。
function parseLnkCandidates(lnkPath) {
  const out = [];
  const push = (s, rank) => {
    const t = String(s || '').trim().replace(/\//g, '\\');
    if (t.length < 4 || t.length > 520) return;
    if (!/^[\\/]{2}/.test(t) && !/^[a-z]:[\\\/]/i.test(t) && !/^[a-z][a-z0-9+.-]*:/i.test(t)) return;  // UNC/盘符/URL
    out.push({ t, rank });
  };
  try {
    const buf = fs.readFileSync(lnkPath);
    if (buf.length < 76 || buf.readUInt32LE(0) !== 0x4C) return [];
    const flags = buf.readUInt32LE(20);
    const readAnsiZ = (off) => {
      if (!off || off < 0 || off >= buf.length) return '';
      const end = buf.indexOf(0, off);
      return buf.slice(off, end < 0 ? buf.length : end).toString('latin1');
    };
    const readUtf16Z = (off) => {
      if (!off || off < 0 || off + 2 > buf.length) return '';
      let end = off;
      while (end + 1 < buf.length && !(buf[end] === 0 && buf[end + 1] === 0)) end += 2;
      return buf.subarray(off, end).toString('utf16le');
    };
    const join2 = (base, suf) => {
      const b2 = String(base || '').trim();
      const s2 = String(suf || '').trim();
      if (!b2 || !s2 || s2 === '.') return b2;
      if (/^[a-z]:[\\\/]/i.test(s2)) return s2;   // suffix 自身是完整路径（非规范 .lnk）→ 直接取
      return b2.replace(/[\\/]+$/, '') + '\\' + s2.replace(/^[\\/]+/, '');
    };
    let pos = 76;
    if (flags & 0x01) pos += 2 + buf.readUInt16LE(pos);   // LinkTargetIDList（PIDL 不重建，跳过）
    if ((flags & 0x02) && pos + 28 <= buf.length) {       // HasLinkInfo：顺序块（IDList 存在时不在 76）
      try {
        const liSize = buf.readUInt32LE(pos);
        const liEnd = Math.min(pos + (liSize || buf.length), buf.length);
        const hdrSize = buf.readUInt32LE(pos + 4);
        const liFlags = buf.readUInt32LE(pos + 8);
        const baseOff = buf.readUInt32LE(pos + 0x10);
        const netOff = buf.readUInt32LE(pos + 0x14);
        const sufOff = buf.readUInt32LE(pos + 0x18);
        const uni = hdrSize >= 0x24;
        const baseOffU = uni ? buf.readUInt32LE(pos + 0x1c) : 0;
        const sufOffU = uni ? buf.readUInt32LE(pos + 0x20) : 0;
        const rel = (off) => (off > 0 && pos + off < liEnd ? pos + off : 0);
        const suffix = (sufOffU && rel(sufOffU) ? readUtf16Z(rel(sufOffU)) : '')
          || readAnsiZ(rel(sufOff)) || '';
        if (netOff && rel(netOff)) {
          const cn = rel(netOff);
          const nnOffU = cn + 0x0c < liEnd ? buf.readUInt32LE(cn + 0x0c) : 0;
          const nnOff = cn + 0x08 < liEnd ? buf.readUInt32LE(cn + 0x08) : 0;
          const netName = (nnOffU && cn + nnOffU < liEnd ? readUtf16Z(cn + nnOffU) : '')
            || readAnsiZ(nnOff ? cn + nnOff : 0) || '';
          if (netName) push(join2(netName, suffix), 900);
        }
        if ((liFlags & 0x01) && (baseOffU || baseOff)) {
          const base = (baseOffU && rel(baseOffU) ? readUtf16Z(rel(baseOffU)) : '')
            || readAnsiZ(rel(baseOff)) || '';
          if (base) push(join2(base, suffix), 850);
        }
      } catch {}
    }
    // StringData（Name/RelativePath/WorkingDir/Args/Icon 按 flags 存在位，定长跳过找 ExtraData）
    try {
      for (const bit of [0x04, 0x08, 0x10, 0x20, 0x40]) {
        if (!(flags & bit)) continue;
        if (flags & 0x80) {                       // IsUnicode：UTF-16 双 0 结尾
          let e = pos;
          while (e + 1 < buf.length && !(buf[e] === 0 && buf[e + 1] === 0)) e += 2;
          pos = e + 2;
        } else {
          const e = buf.indexOf(0, pos);
          pos = e < 0 ? buf.length : e + 1;
        }
      }
    } catch {}
    // ExtraData：EnvironmentVariableDataBlock（0xA0000001）——完整目标的 Unicode/ANSI 双份
    try {
      let p = pos;
      while (p + 8 <= buf.length) {
        const sz = buf.readUInt32LE(p);
        if (!sz || sz < 8 || p + sz > buf.length) break;
        if (buf.readUInt32LE(p + 4) === 0xa0000001) {
          push(readUtf16Z(p + 268), 950);    // TargetUnicode（260 ANSI + 8 头之后）
          push(readAnsiZ(p + 8), 800);       // TargetAnsi（GBK 断串风险，靠 Unicode 优先）
          break;
        }
        p += sz;
      }
    } catch {}
    // 字符串扫描兜底：UTF-16LE 先（中文路径真身），ANSI 后
    const rx = [
      /\\\\[A-Za-z0-9_.$-]+(?:\\[^\\/:*?"<>|\x00\r\n]+)+/g,
      /[A-Za-z]:\\(?:[^\\/:*?"<>|\x00\r\n]+\\)*[^\\/:*?"<>|\x00\r\n]+/g
    ];
    for (const [txt, rank] of [[buf.toString('utf16le'), 600], [buf.toString('latin1'), 100]]) {
      for (const re of rx) {
        re.lastIndex = 0;
        let m;
        while ((m = re.exec(txt))) push(m[0], rank);
      }
    }
  } catch { return []; }
  const stem = path.win32.basename(lnkPath).replace(/[. ]+$/, '').replace(/\.lnk$/i, '').toLowerCase();
  const seen = new Set();
  return out
    .map((c) => ({ ...c, score: c.rank + (stem && c.t.toLowerCase().endsWith(stem) ? 2000 : 0) }))
    .sort((a, b) => b.score - a.score || b.t.length - a.t.length)
    .filter((c) => (seen.has(c.t.toLowerCase()) ? false : (seen.add(c.t.toLowerCase()), true)));
}

function parseLnkTarget(lnkPath) {
  return (parseLnkCandidates(lnkPath)[0] || {}).t || '';
}

// WScript.Shell COM 解析 .lnk 目标（与 createShortcutLnk 同款 -EncodedCommand 零注入）；
// 目标含环境变量则展开；解析失败/空目标返回 ''；URL 目标原样返回（由调用方识别）
async function resolveShortcutTarget(lnkPath) {
  const esc = (s) => String(s).replace(/'/g, "''");
  // WScript.Shell 读 TargetPath；URL 目标 WScript 读回为空 → Shell.Application GetLink 备援
  const dir = path.win32.dirname(lnkPath);
  const base = path.win32.basename(lnkPath);
  const script = "$ErrorActionPreference='Stop'\n"
    + "$w=New-Object -ComObject WScript.Shell\n"
    + "$sc=$w.CreateShortcut('" + esc(lnkPath) + "')\n"
    + "$t=$sc.TargetPath\n"
    + "if(-not $t){ try { $sh=New-Object -ComObject Shell.Application\n"
    + "  $t=$sh.Namespace('" + esc(dir) + "').ParseName('" + esc(base) + "').GetLink().Target.Path } catch { $t='' } }\n"
    + "if($t -match '^[a-z]:[\\\\/]' -or $t -like '\\\\*'){ $t=[Environment]::ExpandEnvironmentVariables($t) }\n"
    + "Write-Output $t";
  const encoded = Buffer.from(script, 'utf16le').toString('base64');
  let out = '';
  try {
    const proc = Bun.spawn(['powershell', '-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
      { stdout: 'pipe', stderr: 'ignore' });
    out = await new Response(proc.stdout).text();
    await proc.exited;
  } catch {}
  let comT = out.split(/\r?\n/).map(s => s.trim()).filter(Boolean).pop() || '';
  if (!comT) comT = await comShortcutTargetCscript(lnkPath);   // PowerShell 失败/被禁 → cscript 备援
  // 二进制结构化候选优先、COM 兜底（2026-11-18 客户案例）：无 LinkInfo 的非常规 .lnk 里
  // WSH 会吐截断值（如 \\host\share 根，恰好存在会误导），ExpString/LinkInfo 的完整路径证据更强；
  // 候选已按可靠度排序（ExpString Unicode > LinkInfo 拼接 > UTF-16 扫描 > ANSI 扫描）+ .lnk 同名加权
  const cands = parseLnkCandidates(lnkPath);
  // 截断值剔除（2026-11-18 客户案例双向验证）：证据度低（同）的候选若是证据度高（同）候选
  // （或其映射形态/真实路径）的路径前缀 → 是残缺截断值（如 LinkInfo 只给到共享根、COM 只给
  // 到 \\host\share），弃；反之截断值证据更高（ExpString 完整路径）则保——防长垃圾/根值抢戏
  const keySet = (p) => {
    const ks = new Set();
    const add = (x) => { const v = String(x || '').toLowerCase().replace(/[\\/]+$/, ''); if (v) ks.add(v); };
    add(p);
    try { for (const f of mapPathForms(p, getDriveMappingsSync())) add(f); } catch {}
    try { add(fs.realpathSync.native(p)); } catch {}
    return ks;
  };
  const all = [];
  if (comT) all.push({ t: comT, rank: 920 });   // COM：结构化但非常规 .lnk 会吐截断值
  for (const c of cands) all.push(c);
  const keys = all.map((c) => keySet(c.t));
  const keep = all.filter((a, i) => !all.some((b, j) => j !== i && b.rank >= a.rank
    && [...keys[j]].some((k) => [...keys[i]].some((m) => m.length < k.length && k.startsWith(m + '\\')))));
  const t0 = (keep[0] || all[0] || {}).t || '';
  let t = String(t0 || '').trim();
  if (t) t = t.replace(/%([A-Za-z_][A-Za-z0-9_]*)%/g, (m, k) => (process.env[k] != null ? process.env[k] : m));
  if (t && !/^[a-z][a-z0-9+.-]*:/i.test(t) && !/^[a-z]:[\\\/]/i.test(t) && !/^[\\/]{2}/.test(t)) {
    t = path.win32.resolve(dir, t);   // 相对目标 → 相对 .lnk 所在目录定绝对
  }
  return t;
}

// PowerShell 整体失败/被禁（执行策略/AppLocker）时的 COM 备援：cscript + WScript.Shell
// （临时 .vbs 传参，零注入；用完即删）
async function comShortcutTargetCscript(lnkPath) {
  try {
    const vf = path.join(os.tmpdir(), 'fin_lnk_' + crypto.randomBytes(6).toString('hex') + '.vbs');
    fs.writeFileSync(vf, 'Set w=CreateObject("WScript.Shell")\nSet s=w.CreateShortcut(WScript.Arguments(0))\nWScript.Echo s.TargetPath\n');
    const proc = Bun.spawn(['cscript', '//Nologo', vf, lnkPath], { stdout: 'pipe', stderr: 'ignore' });
    const out = await new Response(proc.stdout).text();
    await proc.exited;
    try { fs.unlinkSync(vf); } catch {}
    return out.split(/\r?\n/).map(s => s.trim()).filter(Boolean).pop() || '';
  } catch { return ''; }
}

// 快捷方式打开计划：解析目标 → 过安全边界（roots+路径黑名单+扩展名黑名单按目标判定）→
// 返回 { kind:'dir'|'file'|'url', target }；不通过/目标失效抛 Error（原因即拦截理由）。
// 目标在黑名单（路径或扩展名）→ 拦截，杜绝用 .lnk 绕过黑名单；lnk 链目标仍是 .lnk →
// 命中扩展名黑名单拒绝（不做链式递归，防环）。
async function shortcutOpenPlan(lnkPath, cfg) {
  let target = await resolveShortcutTarget(lnkPath);
  // 链式快捷方式（.lnk→.lnk）逐级溯源到最终目标（上限 4 级防环）
  for (let hop = 0; hop < 4 && target && isLnkPath(target); hop++) {
    target = await resolveShortcutTarget(resolveAllowedPath(target, cfg));
  }
  if (!target || isLnkPath(target)) {
    throw new Error('shortcut target not found' + (target ? ' (lnk chain unresolved: ' + target + ')' : ' (no readable target in lnk)'));
  }
  if (/^[a-z][a-z0-9+.-]*:/i.test(target) && !/^[a-z]:[\\\/]/i.test(target)) {
    return { kind: 'url', target };   // URL 目标（http/https/mailto 等）→ 系统默认处理
  }
  const abs = resolveAllowedPath(target, cfg);   // roots + 黑名单 + 符号链接真实路径
  let st = await fs.promises.stat(abs).catch(() => null);
  let openPath = abs;
  if (!st) {
    // 形态兜底（2026-11-18）：映射盘在本会话不可见时改走 UNC 形态（反之亦然）
    for (const alt of mapPathForms(abs, getDriveMappingsSync())) {
      st = await fs.promises.stat(alt).catch(() => null);
      if (st) { openPath = alt; break; }
    }
  }
  if (!st) {
    // 网络计算机/共享根目标（2026-11-18）：\\server 或 \\server\share 这类服务器/共享根
    // 无法用 fs.stat 存在性判断（Windows 里 WSD"网络设备"快捷方式目标就是服务器根，
    // 资源管理器可正常打开）→ 交系统打开（与 Windows 双击行为一致），不报 not found。
    const uncRoot = /^\\\\[^\\\/]+(\\[^\\\/]+)?$/.test(abs.replace(/\//g, '\\').replace(/\\+$/, ''));
    if (uncRoot) return { kind: 'native', target: abs };
    throw new Error('shortcut target not found: ' + abs + ' (stat failed)');
  }
  if (st.isDirectory()) return { kind: 'dir', target: openPath };
  const extCheck = checkExtension(openPath, cfg);
  if (!extCheck.ok) throw new Error(extCheck.reason);   // 目标扩展名在黑名单（如 .exe）→ 拦截
  return { kind: 'file', target: openPath };
}

// 快捷方式信息（UI 双击 .lnk 前问询）：文件夹目标→前端进目录，文件目标→前端走 /open
async function shortcutInfoOp(res, q, cfg) {
  const file = resolveAllowedPath(q.path, cfg, {});
  if (!isLnkPath(file)) return sendJson(res, 400, { ok: false, error: 'not a shortcut' });
  // .lnk 在黑名单 → 全面封禁（与打开/图标/创建同步）
  const selfCheck = checkExtension(file, cfg);
  if (!selfCheck.ok) return sendJson(res, 403, { ok: false, error: selfCheck.reason });
  let plan;
  try { plan = await shortcutOpenPlan(file, cfg); }
  catch (e) { return sendJson(res, 403, { ok: false, error: e.message }); }
  return sendJson(res, 200, { ok: true, kind: plan.kind, target: plan.target });
}

// 快捷方式图标：继承目标在资源管理器显示的原图标（SHGetFileInfo 解析 .lnk 取目标图标 →
// PNG 32px；文件夹/URL/失效目标全兼容），存 cacheDir/thumbs；失败回 404 前端保留默认图标
async function shortcutIconOp(res, lnkPath, cfg) {
  const st = await fs.promises.stat(lnkPath).catch(() => null);
  if (!st || !st.isFile()) return sendJson(res, 404, { ok: false, error: 'not found' });
  // .lnk 在黑名单 → 全面封禁（图标与打开/预览/创建同步）
  const selfCheck = checkExtension(lnkPath, cfg);
  if (!selfCheck.ok) return sendJson(res, 403, { ok: false, error: selfCheck.reason });
  const tdir = path.join(cfg.cacheDir, 'thumbs');
  if (!fs.existsSync(tdir)) fs.mkdirSync(tdir, { recursive: true });
  const key = crypto.createHash('sha1')
    .update('lnkico|' + lnkPath + '|' + st.mtimeMs + '|' + st.size).digest('hex').slice(0, 20);
  const outFile = path.join(tdir, key + '.png');
  if (fs.existsSync(outFile)) return sendImageFile(res, outFile, 'image/png');
  const esc = (s) => String(s).replace(/'/g, "''");
  const script = "$ErrorActionPreference='Stop'\n"
    + "Add-Type -AssemblyName System.Drawing\n"
    + "Add-Type -TypeDefinition 'using System;using System.Runtime.InteropServices;public class ShellIco{[StructLayout(LayoutKind.Sequential,CharSet=CharSet.Unicode)]public struct SHFILEINFO{public IntPtr hIcon;public int iIcon;public uint dwAttributes;[MarshalAs(UnmanagedType.ByValTStr,SizeConst=260)]public string szDisplayName;[MarshalAs(UnmanagedType.ByValTStr,SizeConst=80)]public string szTypeName;}[DllImport(\"shell32.dll\",CharSet=CharSet.Unicode)]public static extern IntPtr SHGetFileInfo(string pszPath,uint dwFileAttributes,ref SHFILEINFO psfi,uint cbFileInfo,uint uFlags);[DllImport(\"user32.dll\")]public static extern bool DestroyIcon(IntPtr h);}'\n"
    + "$sfi=New-Object ShellIco+SHFILEINFO\n"
    + "$sz=[uint32][System.Runtime.InteropServices.Marshal]::SizeOf($sfi)\n"
    + "$ptr=[ShellIco]::SHGetFileInfo('" + esc(lnkPath) + "',[uint32]0,[ref]$sfi,$sz,[uint32]0x100)\n"
    + "if($ptr -eq [IntPtr]::Zero -or $sfi.hIcon -eq [IntPtr]::Zero){ throw 'no icon' }\n"
    + "$ico=[System.Drawing.Icon]::FromHandle($sfi.hIcon)\n"
    + "$bmp=$ico.ToBitmap()\n"
    + "$bmp.Save('" + esc(outFile) + "',[System.Drawing.Imaging.ImageFormat]::Png)\n"
    + "$bmp.Dispose()\n"
    + "[ShellIco]::DestroyIcon($sfi.hIcon) | Out-Null";
  const encoded = Buffer.from(script, 'utf16le').toString('base64');
  const proc = Bun.spawn(['powershell', '-NoProfile', '-NonInteractive', '-EncodedCommand', encoded],
    { stdout: 'ignore', stderr: 'ignore' });
  await proc.exited;
  if (fs.existsSync(outFile)) return sendImageFile(res, outFile, 'image/png');
  return sendJson(res, 404, { ok: false, error: 'no icon' });
}

// —— Windows 右键"新建"菜单的文件类型（注册表 ShellNew，与桌面右键菜单同源） ——
// v2.1.0.1：补全 Windows 实际查找的全部注册位置——
//   ① .ext\ShellNew 顶层；② .ext\<ProgID>\ShellNew 嵌套层（Office/WPS/压缩文件夹的标准注册位）；
//   ③ 32 位视图（WOW6432Node，WPS 等 32 位程序的注册落点）；④ 所有已加载用户的 Software\Classes；
//   ⑤ Explorer 新建菜单缓存兜底（菜单有、注册表缺的类型不缺项）。
// 并按 NullFile/FileName/Data/Command 注册的创建方式生成初始内容（旧版一律写空文件）。
let newFileTypesCache = null;
async function getNewFileTypes() {
  if (!newFileTypesCache) newFileTypesCache = await queryShellNewTypes();
  return newFileTypesCache;
}

const SHELLNEW_QUERY_PS = String.raw`$ErrorActionPreference = 'SilentlyContinue'
if (-not ('SnNew2.SH' -as [type])) {
  Add-Type -Namespace SnNew2 -Name SH -MemberDefinition @'
[DllImport("shlwapi.dll", CharSet = CharSet.Unicode)] public static extern int SHLoadIndirectString(string s, System.Text.StringBuilder b, uint c, IntPtr r);
[DllImport("shell32.dll", CharSet = CharSet.Unicode)] public static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, out SHFILEINFO psfi, uint cbFileInfo, uint uFlags);
[DllImport("user32.dll")] public static extern bool DestroyIcon(IntPtr hIcon);
public struct SHFILEINFO { public IntPtr hIcon; public int iIcon; public uint dwAttributes; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)] public string szDisplayName; [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)] public string szTypeName; }
'@
}
if (-not ('System.Drawing.Icon' -as [type])) { Add-Type -AssemblyName System.Drawing | Out-Null }

function Res3([string]$s) {
  if (-not $s) { return '' }
  if ($s.StartsWith('@')) {
    try {
      $sb = New-Object System.Text.StringBuilder 512
      if (([SnNew2.SH]::SHLoadIndirectString($s, $sb, 512, [IntPtr]::Zero)) -ge 0 -and $sb.Length -gt 0) { return $sb.ToString() }
    } catch { }
    return ''
  }
  return $s
}
function Strip-Accel3([string]$s) {
  if (-not $s) { return '' }
  return ($s -replace '\s*\(\&[^\)]*\)', '').Trim()
}
function Get-IconB64b([string]$ext) {
  try {
    $info = New-Object SnNew2.SH+SHFILEINFO
    $flags = [uint32]0x110
    $h = [SnNew2.SH]::SHGetFileInfo(('x' + $ext), [uint32]0x80, [ref]$info, [uint32]([System.Runtime.InteropServices.Marshal]::SizeOf($info)), $flags)
    if ($h -eq [IntPtr]::Zero -or $info.hIcon -eq [IntPtr]::Zero) { return '' }
    $ic = [System.Drawing.Icon]::FromHandle($info.hIcon)
    $bmp = $ic.ToBitmap()
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $b64 = [Convert]::ToBase64String($ms.ToArray())
    $bmp.Dispose(); $ic.Dispose()
    [void][SnNew2.SH]::DestroyIcon($info.hIcon)
    return $b64
  } catch { return '' }
}

$items = New-Object System.Collections.Generic.List[object]
$extDef = @{}        # ext -> 默认 ProgID（按 HKCU > 已加载用户 > HKLM 优先级，先到先得）
$extPer = @{}        # ext -> PerceivedType（Windows 对伪注册的门禁：text/document/image 才进菜单）
$extUc   = @{}       # ext -> UserChoice ProgId 列表（文件类型默认打开方式 = Windows 菜单取名的第一来源）
$cacheSet = @{}      # Explorer 新建菜单缓存（Windows 菜单自己维护的清单）

function Add-SnEntry($cl, $sn, $ext, $progId, $prio) {
  $vFile = $sn.GetValue('FileName')
  $vData = $sn.GetValue('Data')
  $vNull = $sn.GetValue('NullFile')
  $vCmd  = $sn.GetValue('Command')
  $vHand = $sn.GetValue('Handler')
  $kind = 'none'; $template = $null; $data = $null; $command = $null
  if ($vFile) {
    $kind = 'template'; $template = [string]$vFile
  } elseif ($null -ne $vData) {
    $kind = 'data'
    if ($vData -is [byte[]]) { $data = [Convert]::ToBase64String($vData) }
    else { $data = [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes([string]$vData)) }
  } elseif ($null -ne $vNull) {
    $kind = 'null'
  } elseif ($vHand) {
    $kind = 'handler'
  } elseif ($vCmd) {
    $kind = 'command'; $command = [string]$vCmd
  }
  $itemName = Res3 ([string]$sn.GetValue('ItemName'))
  $menuText = Strip-Accel3 (Res3 ([string]$sn.GetValue('MenuText')))
  $pd = ''
  if ($progId) {
    $pk = $cl.OpenSubKey($progId)
    if (-not $pk) { $pk = $cl.OpenSubKey($ext + '\' + $progId) }
    if ($pk) { $pd = Res3 ([string]$pk.GetValue('')); $pk.Close() }
  }
  $items.Add([PSCustomObject]@{
    ext = $ext; progId = $progId; prio = $prio; kind = $kind
    template = $template; data = $data; command = $command
    itemName = $itemName; menuText = $menuText; progDefault = $pd
    hasHandler = [bool]$vHand; hasNullFile = ($null -ne $vNull)
  })
}

# 根优先级：CurrentUser > 已加载用户 > LocalMachine（HKCR 合并语义：用户覆盖机器）
$hives = New-Object System.Collections.Generic.List[object]
$hives.Add(@{ h = 'CurrentUser';  p = 'Software\Classes' })
foreach ($s in [Microsoft.Win32.Registry]::Users.GetSubKeyNames()) {
  if ($s.StartsWith('S-1-5-21-')) { $hives.Add(@{ h = 'Users'; p = ($s + '\Software\Classes') }) }
}
$hives.Add(@{ h = 'LocalMachine'; p = 'Software\Classes' })
$roots = New-Object System.Collections.Generic.List[object]
foreach ($hv in $hives) {
  foreach ($view in @([Microsoft.Win32.RegistryView]::Registry64, [Microsoft.Win32.RegistryView]::Registry32)) {
    try {
      $hive = [Microsoft.Win32.RegistryHive]$hv.h
      $bk = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, $view)
      $cl = $bk.OpenSubKey($hv.p)
      if ($cl) { $roots.Add($cl) }
    } catch { }
  }
}

# UserChoice（默认打开方式）+ Explorer 新建菜单缓存，两样都是 Windows 菜单自己读的
foreach ($hv in $hives) {
  try {
    $hive = [Microsoft.Win32.RegistryHive]$hv.h
    $bk = [Microsoft.Win32.RegistryKey]::OpenBaseKey($hive, [Microsoft.Win32.RegistryView]::Registry64)
    $fx = $bk.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts')
    if ($fx) {
      foreach ($n in $fx.GetSubKeyNames()) {
        if (-not $n.StartsWith('.')) { continue }
        $uc = $fx.OpenSubKey($n + '\UserChoice')
        if ($uc) {
          $pid2 = [string]$uc.GetValue('ProgId')
          $uc.Close()
          if ($pid2) {
            if (-not $extUc.ContainsKey($n)) { $extUc[$n] = New-Object System.Collections.Generic.List[object] }
            if (-not $extUc[$n].Contains($pid2)) { $extUc[$n].Add($pid2) }
          }
        }
      }
      $fx.Close()
    }
    $cs = $bk.OpenSubKey('Software\Microsoft\Windows\CurrentVersion\Explorer\Discardable\PostSetup\ShellNew')
    if ($cs) {
      $cache = $cs.GetValue('Classes')
      $cs.Close()
      foreach ($c in @($cache)) {
        if ($c -and $c.StartsWith('.')) { $cacheSet[$c.ToLower()] = $true }
      }
    }
    $bk.Close()
  } catch { }
}

# 第一遍：扩展名默认 ProgID（按根优先级）；第二遍：枚举 ShellNew（顶层/嵌套/深嵌套/ProgID 根层）
foreach ($cl in $roots) {
  foreach ($n in $cl.GetSubKeyNames()) {
    if (-not $n.StartsWith('.')) { continue }
    $ek = $cl.OpenSubKey($n)
    if (-not $ek) { continue }
    $prog = [string]$ek.GetValue('')
    if ($prog -and (-not $extDef.ContainsKey($n))) { $extDef[$n] = $prog }
    $per = [string]$ek.GetValue('PerceivedType')
    if ($per -and (-not $extPer.ContainsKey($n))) { $extPer[$n] = $per }
    $sn = $ek.OpenSubKey('ShellNew')
    if ($sn) { Add-SnEntry $cl $sn $n $prog 0; $sn.Close() }
    foreach ($sub in $ek.GetSubKeyNames()) {
      if ($sub -eq 'ShellNew') { continue }
      $sn2 = $ek.OpenSubKey($sub + '\ShellNew')
      if ($sn2) { Add-SnEntry $cl $sn2 $n $sub 1; $sn2.Close() }
      $sk2 = $ek.OpenSubKey($sub)
      if ($sk2) {
        foreach ($sub2 in $sk2.GetSubKeyNames()) {
          if ($sub2 -eq 'ShellNew') { continue }
          $sn3 = $ek.OpenSubKey($sub + '\' + $sub2 + '\ShellNew')
          if ($sn3) { Add-SnEntry $cl $sn3 $n ($sub + '\' + $sub2) 2; $sn3.Close() }
        }
        $sk2.Close()
      }
    }
    $ek.Close()
  }
}
foreach ($cl in $roots) {
  foreach ($n in $cl.GetSubKeyNames()) {
    if ($n.StartsWith('.')) { continue }
    $sn = $cl.OpenSubKey($n + '\ShellNew')
    if (-not $sn) { continue }
    $sn.Close()
    foreach ($e in $extDef.Keys) {
      if ($extDef[$e] -eq $n) { $sn2 = $cl.OpenSubKey($n + '\ShellNew'); if ($sn2) { Add-SnEntry $cl $sn2 $e $n 0; $sn2.Close() } }
    }
  }
}

# 每扩展名取最优注册：UserChoice/默认 ProgID 优先 > 创建方式 > 顶层优先
$best = @{}
$scoreMap = @{ template = 5; data = 4; null = 3; handler = 2; none = 1; command = 0 }
foreach ($it in $items) {
  $k = $it.ext.ToLower()
  $def = $extDef[$it.ext]
  $uc = $extUc[$it.ext]
  $defRank = 0
  if ($def -and ($it.progId -ieq $def)) { $defRank = 1 }
  if ($uc -and (($uc | Where-Object { $_ -ieq $it.progId }).Count -gt 0)) { $defRank = 2 }
  $rank = $defRank * 1000 + $scoreMap[$it.kind] * 10 + (2 - $it.prio)
  if ((-not $best.ContainsKey($k)) -or ($rank -gt $best[$k].rank)) { $best[$k] = @{ rank = $rank; it = $it } }
}

function Get-FriendlyOf($cl, $ext, $progId) {
  if (-not $progId) { return '' }
  $pk = $cl.OpenSubKey($progId)
  if (-not $pk) { $pk = $cl.OpenSubKey($ext + '\' + $progId) }
  if (-not $pk) { return '' }
  $ftn = [string]$pk.GetValue('FriendlyTypeName')
  if (-not $ftn) { $ftn = [string]$pk.GetValue('') }
  $pk.Close()
  return (Res3 $ftn)
}

$out = New-Object System.Collections.Generic.List[object]
$allExt = New-Object System.Collections.Generic.HashSet[string]
foreach ($k in $best.Keys) { [void]$allExt.Add($k) }
foreach ($k in $cacheSet.Keys) { [void]$allExt.Add($k) }
foreach ($k in ($allExt | Sort-Object)) {
  $it = $null
  if ($best.ContainsKey($k)) { $it = $best[$k].it }
  $def = $extDef[$k]
  if (-not $def) {
    foreach ($e2 in $extDef.Keys) { if ($e2.ToLower() -eq $k) { $def = $extDef[$e2]; break } }
  }
  $ucList = $null
  foreach ($e2 in $extUc.Keys) { if ($e2.ToLower() -eq $k) { $ucList = $extUc[$e2]; break } }
  # —— Windows 菜单过滤规律（两台机器逐项实证） ——
  if ((-not $def) -and (-not $ucList)) { continue }                                   # ① 无默认 ProgID 不进菜单（.contact）
  if ($it -and $it.kind -eq 'command') { continue }                                   # ② 仅 Command 创建方式不进菜单
  if ($it -and $it.hasHandler -and $it.hasNullFile -and (-not $it.itemName) -and (-not $it.menuText)) { continue }  # ③ Handler+NullFile 无标签不进菜单（.library-ms）
  $perceived = ''
  foreach ($e2 in $extPer.Keys) { if ($e2.ToLower() -eq $k) { $perceived = $extPer[$e2]; break } }
  if ($it -and ($it.kind -eq 'none' -or $it.kind -eq 'handler')) {                    # ④ 伪注册（无创建值/仅 Handler）：
    if ($perceived -notin @('text', 'document', 'image')) { continue }                #    仅 文本/文档/图像 感知类型进菜单
  }                                                                                  #    （文本文档 text、BMP 图像 image 进；压缩 compressed、FontCreator 无 -> 不进）
  # —— 名称链条（与 Windows 菜单逐字一致）——
  # UserChoice 默认打开方式 -> 扩展名默认 ProgID -> 菜单文字 MenuText -> 所选注册 ProgID -> ItemName -> 扩展名
  $name = ''
  if ($ucList) { foreach ($u in $ucList) { if (-not $name) { $name = Get-FriendlyOf ([Microsoft.Win32.Registry]::ClassesRoot) $k $u } } }
  if (-not $name) { $name = Get-FriendlyOf ([Microsoft.Win32.Registry]::ClassesRoot) $k $def }
  if (-not $name -and $it) { $name = $it.menuText }
  if (-not $name -and $it) { $name = $it.progDefault }
  if (-not $name -and $it) { $name = $it.itemName }
  if (-not $name) { $name = $k }
  $kindOut = 'null'; $template = $null; $data = $null; $command = $null
  if ($it) {
    $kindOut = $it.kind
    if ($kindOut -eq 'none') { $kindOut = 'null' }
    $template = $it.template; $data = $it.data; $command = $it.command
  }
  $out.Add([PSCustomObject]@{
    ext = $k; name = $name; kind = $kindOut
    template = $template; data = $data; command = $command
    icon = (Get-IconB64b $k)
  })
}
$out | ConvertTo-Json -Compress
`;

async function queryShellNewTypes() {
  const proc = Bun.spawn(['powershell', '-NoProfile', '-NonInteractive', '-Command', SHELLNEW_QUERY_PS],
    { stdout: 'pipe', stderr: 'ignore' });
  const out = await new Response(proc.stdout).text();
  await proc.exited;
  let arr = [];
  try { const j = JSON.parse(out.trim() || '[]'); arr = Array.isArray(j) ? j : [j]; } catch { }
  return arr.filter(x => x && x.ext && x.name)
    .map(x => ({
      ext: String(x.ext).toLowerCase(),
      name: String(x.name || x.ext),
      kind: x.kind || 'null',
      template: x.template || null,
      data: x.data || null,
      command: x.command || null,
      icon: typeof x.icon === 'string' && /^[A-Za-z0-9+/=]+$/.test(x.icon) ? x.icon : '',
    }))
    .sort((a, b) => a.ext.localeCompare(b.ext));
}

async function newfileTypesOp(res) {
  const types = await getNewFileTypes();
  sendJson(res, 200, { ok: true, types: types.map(t => ({ ext: t.ext, name: t.name, icon: t.icon || '' })) });
}

// —— 我的收藏（持久化到安装目录 favorites.json） ——
function favoritesFile() { return path.join(INSTALL_DIR, 'favorites.json'); }
function loadFavorites() {
  try { return JSON.parse(fs.readFileSync(favoritesFile(), 'utf8')) || []; } catch { return []; }
}
function saveFavorites(list) {
  fs.writeFileSync(favoritesFile(), JSON.stringify(list, null, 2));
}
async function favoritesListOp(res) {
  sendJson(res, 200, { ok: true, favorites: loadFavorites() });
}
async function favoritesAddOp(res, data, cfg) {
  const p = resolveAllowedPath(data.path, cfg);
  const list = loadFavorites();
  if (!list.some(f => f.path === p)) {
    list.unshift({ path: p, name: path.win32.basename(p) || p, type: data.type === 'dir' ? 'dir' : 'file', addedAt: Date.now() });
    saveFavorites(list);
  }
  sendJson(res, 200, { ok: true, favorites: list });
}
async function unfavoriteOp(res, data) {
  const raw = String(data.path || '');
  const list = loadFavorites().filter(f => f.path !== raw);
  saveFavorites(list);
  sendJson(res, 200, { ok: true, favorites: list });
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
  indexMove(src, dst);   // 库同步：旧行移除+新行写入，子树失效防幽灵
  sendJson(res, 200, { ok: true, path: dst });
}

// ============ 回收站（应用内，删除默认进回收站可恢复） ============
const RECYCLE_DIR  = path.join(INSTALL_DIR, 'recycle_bin');
const RECYCLE_META = path.join(INSTALL_DIR, 'recycle_meta.json');

async function recycleMetaList() {
  try { return JSON.parse(await fs.promises.readFile(RECYCLE_META, 'utf8')) || []; } catch (e) { return []; }
}
async function recycleSave(list) {
  await fs.promises.mkdir(RECYCLE_DIR, { recursive: true });
  await fs.promises.writeFile(RECYCLE_META, JSON.stringify(list, null, 2));
}
async function recycleMove(target, st) {
  const id = Date.now().toString(36) + Math.random().toString(36).slice(2, 7);
  const name = path.win32.basename(target);
  const binPath = path.join(RECYCLE_DIR, id + '__' + name);
  await fs.promises.mkdir(RECYCLE_DIR, { recursive: true });
  try {
    await fs.promises.rename(target, binPath);
  } catch (e) {
    if (e && e.code === 'EXDEV') {
      await copyRecursive(target, binPath);
      await fs.promises.rm(target, { recursive: true, force: true });
    } else throw e;
  }
  const list = await recycleMetaList();
  list.unshift({ id, name, originalPath: target, binPath, isDir: st.isDirectory(), size: st.size, deletedAt: Date.now() });
  await recycleSave(list);
  return id;
}

async function rmOp(res, data, cfg) {
  const target = resolveAllowedPath(data.path, cfg);
  const st = await fs.promises.stat(target);
  if (data.purge) {
    // 彻底删除（回收站专用入口）
    await fs.promises.rm(target, { recursive: true, force: true });
    indexInvalidate(target);   // 库同步：子树清扫
    return sendJson(res, 200, { ok: true, purged: true });
  }
  // 默认：移入回收站（可恢复），不再直接删除
  const id = await recycleMove(target, st);
  indexInvalidate(target);   // 库同步：子树清扫
  sendJson(res, 200, { ok: true, recycled: id });
}

async function recycleListOp(res) {
  sendJson(res, 200, { ok: true, items: await recycleMetaList() });
}

async function recycleRestoreOne(info) {
  let dest = info.originalPath;
  if (fs.existsSync(dest)) {
    // 原位置已被占用：加后缀避免覆盖
    const ext = path.win32.extname(dest), base = dest.slice(0, dest.length - ext.length);
    dest = base + ' (restored)' + ext;
    let n = 2;
    while (fs.existsSync(dest)) dest = base + ' (restored' + n++ + ')' + ext;
  }
  await fs.promises.mkdir(path.win32.dirname(dest), { recursive: true });
  try {
    await fs.promises.rename(info.binPath, dest);
  } catch (e) {
    if (e && e.code === 'EXDEV') {
      await copyRecursive(info.binPath, dest);
      await fs.promises.rm(info.binPath, { recursive: true, force: true });
    } else throw e;
  }
  return dest;
}

async function recycleRestoreOp(res, data, cfg) {
  const list = await recycleMetaList();
  const info = list.find(x => x.id === data.id);
  if (!info) throw new Error('recycle item not found');
  const dest = await recycleRestoreOne(info);
  indexWrite(dest);   // 库同步
  await recycleSave(list.filter(x => x.id !== info.id));
  sendJson(res, 200, { ok: true, restored: dest });
}

async function recycleRestoreAllOp(res) {
  const list = await recycleMetaList();
  const remain = [];
  let count = 0;
  for (const info of list) {
    try { const d = await recycleRestoreOne(info); indexWrite(d); count++; } catch (e) { remain.push(info); }
  }
  await recycleSave(remain);
  sendJson(res, 200, { ok: true, count, failed: remain.length });
}

async function recyclePurgeOp(res, data, cfg) {
  const list = await recycleMetaList();
  const info = list.find(x => x.id === data.id);
  if (!info) throw new Error('recycle item not found');
  await fs.promises.rm(info.binPath, { recursive: true, force: true });
  await recycleSave(list.filter(x => x.id !== info.id));
  sendJson(res, 200, { ok: true });
}

async function recycleEmptyOp(res, data, cfg) {
  const list = await recycleMetaList();
  for (const info of list) {
    try { await fs.promises.rm(info.binPath, { recursive: true, force: true }); } catch (e) {}
  }
  await recycleSave([]);
  sendJson(res, 200, { ok: true, count: list.length });
}

// ============ 递归搜索（文件名包含匹配，上限 500 条） ============
async function searchOp(req, res, q, cfg) {
  const root = resolveAllowedPath(String(q.path || ''), cfg);
  const kw = String(q.q || '').trim().toLowerCase();
  if (!kw) return sendJson(res, 200, { ok: true, results: [] });
  const results = [];
  const LIMIT = 500;
  async function walk(dir) {
    if (results.length >= LIMIT) return;
    if (req && req.signal && req.signal.aborted) return;   // 客户端已放弃
    let entries;
    try { entries = await fs.promises.readdir(dir, { withFileTypes: true }); } catch (e) { return; }
    // 父目录真实路径每层解析一次，免逐条目 realpathSync（同 listDir，2026-11-15 修）
    let realDir = '';
    if ((cfg.deniedPaths || []).length) {
      try { realDir = path.win32.normalize(fs.realpathSync(dir)).replace(/[\\\/]+$/, ''); } catch (e) {}
    }
    for (const ent of entries) {
      if (results.length >= LIMIT) return;
      if (req && req.signal && req.signal.aborted) return;
      if (ent.isSymbolicLink()) continue;   // 防链接环
      const full = path.win32.join(dir, ent.name);
      if (isDenied(full, cfg, realDir ? realDir + '\\' + ent.name : undefined)) continue;
      if (ent.isDirectory()) {
        if (ent.name.toLowerCase().includes(kw)) {
          results.push({ name: path.win32.relative(root, full), path: full, type: 'dir', size: 0, mtime: 0, ext: '' });
        }
        await walk(full);
      } else if (ent.isFile()) {
        if (ent.name.toLowerCase().includes(kw)) {
          let size = 0, mtime = 0;
          try { const st = await fs.promises.stat(full); size = st.size; mtime = Math.floor(st.mtimeMs); } catch (e) { markBadHost(full); }
          results.push({ name: path.win32.relative(root, full), path: full, type: 'file', size, mtime, ext: path.win32.extname(ent.name).toLowerCase() });
        }
      }
    }
  }
  await walk(root);
  results.sort((a, b) => a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1);
  sendJson(res, 200, { ok: true, results, truncated: results.length >= LIMIT });
}

// ============ 属性（大小/时间/文件哈希） ============
async function propsOp(res, data, cfg) {
  const target = resolveAllowedPath(data.path, cfg);
  const st = await fs.promises.stat(target);
  const out = {
    name: path.win32.basename(target), path: target,
    type: st.isDirectory() ? 'dir' : 'file',
    size: st.size,
    created: Math.floor(st.birthtimeMs), modified: Math.floor(st.mtimeMs)
  };
  if (st.isDirectory()) {
    try { out.count = (await fs.promises.readdir(target)).length; } catch (e) { out.count = -1; }
  } else {
    const md5 = crypto.createHash('md5'), sha = crypto.createHash('sha256');
    await new Promise((resolve, reject) => {
      const s = fs.createReadStream(target);
      s.on('data', c => { md5.update(c); sha.update(c); });
      s.on('end', resolve);
      s.on('error', reject);
    });
    out.md5 = md5.digest('hex');
    out.sha256 = sha.digest('hex');
  }
  sendJson(res, 200, { ok: true, props: out });
}

// ============ 打包下载（多选打成 zip 返回） ============
async function zipCreateOp(res, data, cfg) {
  const paths = Array.isArray(data.paths) ? data.paths : [];
  if (!paths.length || paths.length > 500) throw new Error('invalid paths');
  const zip = new JSZip();
  for (const p of paths) {
    const src = resolveAllowedPath(p, cfg);
    await addPathToZip(zip, src, path.win32.basename(src));
  }
  const buf = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' });
  const fname = 'download-' + Date.now().toString(36) + '.zip';
  res.writeHead(200, {
    'Content-Type': 'application/zip',
    'Content-Disposition': "attachment; filename*=UTF-8''" + encodeURIComponent(fname),
    'Content-Length': buf.length
  });
  res.end(buf);
}
async function addPathToZip(zip, src, rel) {
  const st = await fs.promises.stat(src);
  if (st.isDirectory()) {
    for (const ent of await fs.promises.readdir(src, { withFileTypes: true })) {
      await addPathToZip(zip, path.win32.join(src, ent.name), rel + '/' + ent.name);
    }
  } else {
    zip.file(rel.replace(/\\/g, '/'), await fs.promises.readFile(src));
  }
}

// ============ 整包解压（解压到压缩包所在文件夹，已存在文件跳过，绝不覆盖） ============
async function extractAllOp(res, data, cfg) {
  const src = resolveAllowedPath(data.path, cfg);
  const fmt = detectArchiveFormat(src);
  if (!fmt) throw new Error('not a supported archive');
  const dest = path.win32.resolve(path.win32.dirname(src));
  let written = 0, skipped = 0;
  async function writeEntry(relName, dataBuf) {
    // Zip Slip 防护：条目路径必须落在目标文件夹内
    // 修：dest 为盘符根时结尾本就带分隔符，再拼 sep 使前缀多一个分隔符，startsWith 永假，误杀一切条目；
    //     同时去掉条目名开头分隔符并拒绝 '..' 段与盘符前缀
    const name = String(relName).replace(/^[\\/]+/, '');
    if (!name || /^[a-zA-Z]:/.test(name) || /(^|[\\/])\.\.([\\/]|$)/.test(name)) throw new Error('unsafe entry path: ' + relName);
    const target = path.win32.resolve(dest, name.replace(/[\\/]+/g, path.win32.sep));
    const base = dest.endsWith(path.win32.sep) ? dest : dest + path.win32.sep;
    if (!target.toLowerCase().startsWith(base.toLowerCase())) throw new Error('unsafe entry path: ' + relName);
    if (fs.existsSync(target)) { skipped++; return; }
    await fs.promises.mkdir(path.win32.dirname(target), { recursive: true });
    await fs.promises.writeFile(target, dataBuf);
    written++;
  }
  async function ensureDir(relName) {
    const name = String(relName).replace(/^[\\/]+/, '');
    if (!name || /^[a-zA-Z]:/.test(name) || /(^|[\\/])\.\.([\\/]|$)/.test(name)) return;
    const target = path.win32.resolve(dest, name.replace(/[\\/]+/g, path.win32.sep));
    const base = dest.endsWith(path.win32.sep) ? dest : dest + path.win32.sep;
    if (!target.toLowerCase().startsWith(base.toLowerCase())) return;
    if (!fs.existsSync(target)) await fs.promises.mkdir(target, { recursive: true });
  }
  if (fmt === 'zip') {
    const zip = await JSZip.loadAsync(await fs.promises.readFile(src));
    for (const name of Object.keys(zip.files)) {
      const ent = zip.files[name];
      if (ent.dir) { await ensureDir(name); continue; }
      await writeEntry(name, await ent.async('nodebuffer'));
    }
  } else if (fmt === 'tar' || fmt === 'gz') {
    let data = await fs.promises.readFile(src);
    if (fmt === 'gz') {
      data = zlib.gunzipSync(data);
      // .tar.gz → 按 TAR 解；纯 .gz 单文件 → 解出单个文件
      const entries = parseTarEntries(data);
      const sane = entries.length > 0 && entries.every(e => !/[\x00-\x1f]/.test(e.name) && e.dataOffset + e.size <= data.length);
      if (!sane) {
        await writeEntry(path.win32.basename(src).replace(/\.gz$/i, ''), data);
        return sendJson(res, 200, { ok: true, written, skipped });
      }
      for (const e of entries) {
        if (e.type === 'dir') { await ensureDir(e.name); continue; }
        await writeEntry(e.name, data.slice(e.dataOffset, e.dataOffset + e.size));
      }
    } else {
      for (const e of parseTarEntries(data)) {
        if (e.type === 'dir') { await ensureDir(e.name); continue; }
        if (e.dataOffset + e.size > data.length) continue;   // 越界条目跳过（#9 同款防护）
        await writeEntry(e.name, data.slice(e.dataOffset, e.dataOffset + e.size));
      }
    }
  } else {
    // rar / 7z / bz2 / xz：走 7z.exe，先解到临时目录再合并（已存在跳过，绝不覆盖）
    const sevenZip = get7zExe();
    const tmp = path.join(os.tmpdir(), 'filenest-extract-' + Date.now().toString(36));
    await fs.promises.mkdir(tmp, { recursive: true });
    try {
      const r = spawnSync(sevenZip, ['x', '-y', '-o' + tmp, src], { windowsHide: true, timeout: 600000 });
      if (r.status !== 0) throw new Error('7z extract failed');
      async function merge(dir) {
        for (const ent of await fs.promises.readdir(dir, { withFileTypes: true })) {
          const s = path.win32.join(dir, ent.name);
          if (ent.isDirectory()) await merge(s);
          else await writeEntry(path.win32.relative(tmp, s), await fs.promises.readFile(s));
        }
      }
      await merge(tmp);
    } finally {
      await fs.promises.rm(tmp, { recursive: true, force: true });
    }
  }
  indexInvalidate(dest);   // 库同步：解压到压缩包所在文件夹，该目录子树失效兜底（下次进实时校准）
  sendJson(res, 200, { ok: true, written, skipped });
}

async function mvOp(res, data, cfg) {
  const src = resolveAllowedPath(data.from, cfg);
  const dst = resolveAllowedPath(data.to, cfg);
  const st = await fs.promises.stat(src);
  if (!st.isDirectory()) {
    const extCheck = checkExtension(dst, cfg);
    if (!extCheck.ok) return sendJson(res, 403, { ok: false, error: extCheck.reason });
  }
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
  indexMove(src, dst);   // 库同步（含 EXDEV 复制+删除降级路径）
  sendJson(res, 200, { ok: true });
}

// 递归复制文件 / 目录（POST /api/cp {from, to}）
async function cpOp(res, data, cfg) {
  const src = resolveAllowedPath(data.from, cfg);
  const dst = resolveAllowedPath(data.to, cfg);
  if (path.win32.normalize(src).toLowerCase() === path.win32.normalize(dst).toLowerCase()) {
    throw new Error('source and destination are the same');
  }
  const st = await fs.promises.stat(src);
  if (!st.isDirectory()) {
    const extCheck = checkExtension(dst, cfg);
    if (!extCheck.ok) return sendJson(res, 403, { ok: false, error: extCheck.reason });
  }
  await copyRecursive(src, dst);
  indexWrite(dst);   // 库同步
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
  // 扩展名黑名单统一前置（2026-11-18 定正）：.lnk 在黑名单 → 全面封禁；不在 → 走下方追溯
  const extCheck = checkExtension(data.path || '', cfg);
  if (!extCheck.ok) return sendJson(res, 403, { ok: false, error: extCheck.reason });
  if (hasUnsafeName(data.path || '')) return sendJson(res, 400, { ok: false, error: 'invalid characters in path' });
  // 快捷方式追溯：.lnk 不在黑名单时解析真实目标按目标类型/策略打开（目标照常过黑名单）
  if (isLnkPath(data.path) && String(data.path || '').indexOf('::') < 0) {
    let plan;
    try { plan = await shortcutOpenPlan(data.path, cfg); }
    catch (e) { return sendJson(res, 403, { ok: false, error: e.message }); }
    if (plan.kind === 'url' || plan.kind === 'native') {
      openWithSystem(plan.target);
      return sendJson(res, 200, { ok: true, path: data.path, target: plan.target, shortcut: true });
    }
    data = Object.assign({}, data, { path: plan.target });
  }
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
    // 压缩包内 .lnk 同样追溯（2026-11-18）：杜绝 zip 快捷方式绕过目标黑名单
    let openTmp = tmpFile;
    if (isLnkPath(entry)) {
      let plan;
      try { plan = await shortcutOpenPlan(tmpFile, cfg); }
      catch (e) { return sendJson(res, 403, { ok: false, error: e.message }); }
      if (plan.kind === 'url' || plan.kind === 'native') {
        openWithSystem(plan.target);
        return sendJson(res, 200, { ok: true, path: data.path, target: plan.target, shortcut: true, extracted: true });
      }
      openTmp = plan.target;
    }
    openWithSystem(openTmp);
    // 打开副本延迟清理（#12）：1h 后尝试删除，被查看器占用则退避重试，24h 清扫兜底
    const tryDelete = (attempt) => {
      setTimeout(() => {
        fs.promises.unlink(tmpFile).catch(() => { if (attempt < 3) tryDelete(attempt + 1); });
      }, 3600000);
    };
    tryDelete(0);
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
  _activeUploadTmps.add(tmp);
  journalUploadDir(dir);
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
  indexWrite(full);   // 库同步
  _activeUploadTmps.delete(tmp);
  sendJson(res, 200, { ok: true, path: full, size: fs.statSync(full).size });
}

// ============ 文件打开 / 定位（保留原 /open 语义） ============
async function handleFileOp(op, q, res, cfg) {
  let target;
  try { target = await resolveTarget(q, cfg); }
  catch (e) { return sendJson(res, 403, { ok: false, error: e.message }); }
  if (!target) return sendJson(res, 403, { ok: false, error: 'path not allowed' });

  // 扩展名黑名单统一前置（2026-11-18 定正）：.lnk 在黑名单 → 全面封禁；不在 → /open 走追溯
  const extCheck = checkExtension(target.displayPath || target.localPath || '', cfg);
  if (!extCheck.ok) {
    return sendJson(res, 403, { ok: false, error: extCheck.reason });
  }

  let localPath;
  try { localPath = await obtainLocal(target, cfg); }
  catch (e) { return sendJson(res, 404, { ok: false, error: e.message }); }

  if (op === '/exists') return sendJson(res, 200, { ok: true, exists: true, path: target.displayPath });

  if (op === '/open') {
    // 快捷方式追溯：.lnk 不在黑名单时解析真实目标按目标策略处理（目标照常过黑名单）
    if (isLnkPath(localPath)) {
      let plan;
      try { plan = await shortcutOpenPlan(localPath, cfg); }
      catch (e) { return sendJson(res, 403, { ok: false, error: e.message }); }
      localPath = plan.target;
    }
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
      // 路径含逗号时 explorer 的 /select, 切分会误解析 → 退化为打开所在文件夹（#20）
      const sel = localPath.includes(',')
        ? [path.win32.dirname(localPath)]
        : ['/select,', localPath];
      spawn('explorer.exe', sel,
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
  // normalizeAbs 而非 path.win32.normalize：后者会把 UNC 主机根 \\host 吞成 \host（2026-11-18 修）
  let norm;
  try { norm = normalizeAbs(absPath); } catch (e) { return false; }
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
  try { norm = normalizeAbs(abs); } catch (e) { throw new Error('invalid path'); }
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
  // 递归清理缓存文件（含 thumbs 等子目录），空目录顺手回收；写入端自建目录，删了可再生。
  // lstat 不跟随符号链接、限深 8，保证绝不越出 cacheDir 真实子树。
  const sweep = (dir, depth) => {
    let entries;
    try { entries = fs.readdirSync(dir); } catch (e) { return; }
    for (const f of entries) {
      const full = path.join(dir, f);
      try {
        const st = fs.lstatSync(full);
        if (st.isSymbolicLink()) continue;
        if (st.isDirectory()) {
          if (depth < 8) sweep(full, depth + 1);
          try { if (fs.readdirSync(full).length === 0) fs.rmdirSync(full); } catch (e) {}
        } else if (st.isFile() && (force || now - st.mtimeMs > maxMs)) {
          fs.unlinkSync(full);
        }
      } catch (e) {}
    }
  };
  sweep(cfg.cacheDir, 0);
}

// —— 临时文件兜底回收（#11/#12）——
// 上传残留（.upload.<pid>.<ts>）：进程崩溃/断电后残留目标目录，靠目录日志 + 24h 阈值清扫；
// 压缩包"用系统程序打开"副本（%TEMP%\html-FileNest-open）：同样 24h 阈值清扫。
const _activeUploadTmps = new Set();
const UPLOAD_JOURNAL = path.join(INSTALL_DIR, 'upload-journal.json');
function journalUploadDir(dir) {
  try {
    let arr = [];
    try { arr = JSON.parse(fs.readFileSync(UPLOAD_JOURNAL, 'utf8')) || []; } catch (_) {}
    if (!Array.isArray(arr)) arr = [];
    if (arr.indexOf(dir) < 0) {
      arr.push(dir);
      if (arr.length > 50) arr = arr.slice(-50);
      fs.writeFileSync(UPLOAD_JOURNAL, JSON.stringify(arr));
    }
  } catch (_) {}
}
function sweepStaleTemp() {
  const now = Date.now();
  const maxMs = 24 * 3600000;
  const cleanDir = (dir, re) => {
    try {
      for (const f of fs.readdirSync(dir)) {
        if (re && !re.test(f)) continue;
        const full = path.join(dir, f);
        try {
          const st = fs.statSync(full);
          if (st.isFile() && now - st.mtimeMs > maxMs) fs.unlinkSync(full);
        } catch (_) {}
      }
    } catch (_) {}
  };
  // 上传残留：只扫日志记录过的目录，避免全盘扫描
  try {
    let arr = [];
    try { arr = JSON.parse(fs.readFileSync(UPLOAD_JOURNAL, 'utf8')) || []; } catch (_) {}
    if (Array.isArray(arr)) for (const dir of arr) cleanDir(dir, /\.upload\.\d+\.\d+$/);
  } catch (_) {}
  // 压缩包打开副本
  cleanDir(path.join(os.tmpdir(), 'html-FileNest-open'), null);
}
// 进程退出时尽力清理进行中的上传临时文件（崩溃/断电路径由 journal 清扫兜底）
process.on('exit', () => {
  for (const t of _activeUploadTmps) { try { fs.unlinkSync(t); } catch (_) {} }
});

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
  $c = New-SelfSignedCertificate -Subject 'CN=${CERT_CN} Local' -DnsName 'localhost','127.0.0.1' ` +
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
else if (argv.includes('--stop'))      doStop();
else if (argv.includes('--restart'))   doRestart();
else if (argv.includes('--uninstall')) doUninstall();
else                                   doAuto().catch(e => {
  showMsg(APP_NAME + ' 部署失败', String(e && e.message || e), 'Error');
  process.exit(1);
});
