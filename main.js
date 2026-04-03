/**
 * XylemView Pro — Electron Main Process
 * All file system operations, config management, and business logic live here.
 * The renderer (index.html) communicates via IPC through preload.js.
 *
 * SAFETY POLICY:
 *   - This program NEVER deletes or overwrites files.
 *   - Link creation uses fs.writeFile with wx flag (exclusive create).
 *   - File paste checks existence before every copy.
 *   - Folder creation uses mkdir (errors if exists).
 */

const { app, BrowserWindow, ipcMain, Tray, Menu, shell, globalShortcut, nativeImage, dialog, clipboard, nativeTheme } = require('electron');
const path = require('path');
const fs = require('fs');
const fsp = fs.promises;
const { execSync, exec, spawn } = require('child_process');
const { promisify } = require('util');
const execAsync = promisify(exec);
const os = require('os');

// ─── Config defaults ────────────────────────────────────────────────────
const IS_MAC = process.platform === 'darwin';
const IS_WIN = process.platform === 'win32';

if (IS_WIN) app.setAppUserModelId('com.xylem.xylemview-pro');

const DEFAULTS = {
  order_path:      IS_WIN ? 'L:\\Group\\orders' : path.join(os.homedir(), 'XylemTest/orders'),
  archive_path:    IS_WIN ? 'L:\\Group\\Orders-Archive' : path.join(os.homedir(), 'XylemTest/Orders-Archive'),
  drawing_path:    IS_WIN ? 'L:\\Drawings\\ACADDWGS' : path.join(os.homedir(), 'XylemTest/ACADDWGS'),
  ds_path:         IS_WIN ? 'L:\\drawings\\DSGNSTDS\\DS' : path.join(os.homedir(), 'XylemTest/DSGNSTDS'),
  dwg_viewer:      IS_WIN ? 'E:\\ITTVIEW\\DWGSee\\DWGSee.exe' : '',
  autocad_exe:     '',  // Auto-detected on first use
  contingency_exe: IS_WIN ? 'E:\\Contingency\\Contingency.exe' : '',
  theme: 'auto',
  minimize_to_tray: true,
  default_dwg_action: 'viewer',
  preferred_format: 'dwg',  // 'dwg' or 'pdf'
  recent_orders: [],  // [{order, line, label, ts}]
  recent_drawings: [],  // [{query, ts}]
  recent_opened_drawings: [],  // [{name, query, ts}]
  launch_on_startup: true,
  seen_changelog_ver: '',
};

// Auto-detect the newest AutoCAD installation
function detectAutoCAD() {
  if (!IS_WIN) return null;
  const base = 'C:\\Program Files\\Autodesk';
  try {
    if (!fs.existsSync(base)) return null;
    const dirs = fs.readdirSync(base)
      .filter(d => d.toLowerCase().startsWith('autocad') && fs.existsSync(path.join(base, d, 'acad.exe')))
      .sort()  // alphabetical = newest year last
      .reverse();
    if (dirs.length > 0) return path.join(base, dirs[0], 'acad.exe');
  } catch (e) {}
  return null;
}

function findAccoreConsole() {
  if (!IS_WIN) return null;
  const base = 'C:\\Program Files\\Autodesk';
  try {
    if (!fs.existsSync(base)) return null;
    const dirs = fs.readdirSync(base)
      .filter(d => d.toLowerCase().startsWith('dwg trueview'))
      .sort()
      .reverse(); // newest first
    for (const d of dirs) {
      const candidate = path.join(base, d, 'accoreconsole.exe');
      if (fs.existsSync(candidate)) return candidate;
    }
  } catch (e) {}
  return null;
}

function findOdaConverter() {
  if (!IS_WIN) return null;
  const base = 'C:\\Program Files\\ODA';
  try {
    if (!fs.existsSync(base)) return null;
    const dirs = fs.readdirSync(base).filter(d => d.toLowerCase().startsWith('odafileconverter')).sort().reverse();
    for (const d of dirs) {
      const candidate = path.join(base, d, 'ODAFileConverter.exe');
      if (fs.existsSync(candidate)) return candidate;
    }
  } catch (e) {}
  return null;
}

// Convert a single DWG to DXF using ODA File Converter
async function odaConvertDwgToDxf(dwgPath, outDir) {
  const oda = findOdaConverter();
  if (!oda) return { ok: false, msg: 'ODA File Converter not found' };
  const inputDir = path.dirname(dwgPath);
  const fileName = path.basename(dwgPath);
  const outputDir = outDir || inputDir;
  // ODA operates on folders, use filter param to target single file
  return new Promise((resolve) => {
    const proc = spawn(oda, [inputDir, outputDir, 'ACAD2018', 'DXF', '0', '1', fileName],
      { stdio: 'pipe', timeout: 60000, windowsHide: true });
    proc.on('close', () => {
      const dxfName = fileName.replace(/\.dwg$/i, '.dxf');
      const dxfPath = path.join(outputDir, dxfName);
      if (fs.existsSync(dxfPath)) {
        resolve({ ok: true, dxfPath, msg: `Converted to ${dxfName}` });
      } else {
        resolve({ ok: false, msg: 'ODA conversion produced no output' });
      }
    });
    proc.on('error', (e) => {
      resolve({ ok: false, msg: 'Could not start ODA: ' + e.message });
    });
  });
}

const CFG_VERSION = 1;

// Resolved shared data directory — set after loadConfig determines drive mappings
let SHARED_DATA_DIR = IS_WIN ? 'E:\\XylemView\\XylemView Pro' : path.join(os.homedir(), 'XylemTest');
function resolveSharedDataDir() {
  if (!IS_WIN) return;
  const eDrive = 'E:\\XylemView\\XylemView Pro';
  if (fs.existsSync(eDrive)) { SHARED_DATA_DIR = eDrive; return; }
  const unc = '\\\\01ckfp02-1\\Apps\\XylemView\\XylemView Pro';
  if (fs.existsSync(unc)) { SHARED_DATA_DIR = unc; return; }
  SHARED_DATA_DIR = eDrive; // fallback to E: even if missing
}

// ─── Config management ──────────────────────────────────────────────────
let config = { ...DEFAULTS };

function configDir() {
  return path.join(app.getPath('appData'), 'XylemViewPro');
}
function configPath() {
  return path.join(configDir(), 'config.json');
}

// UNC fallbacks — used when drive letters don't map to the expected server
const UNC_DEFAULTS = {
  order_path:      '\\\\01ckfp02-1\\vol1\\Group\\orders',
  archive_path:    '\\\\01ckfp02-1\\vol1\\Group\\Orders-Archive',
  drawing_path:    '\\\\01ckfp02-1\\vol1\\Drawings\\ACADDWGS',
  ds_path:         '\\\\01ckfp02-1\\vol1\\drawings\\DSGNSTDS\\DS',
  dwg_viewer:      '\\\\01ckfp02-1\\Apps\\ITTVIEW\\DWGSee\\DWGSee.exe',
  contingency_exe: '\\\\01ckfp02-1\\Apps\\Contingency\\Contingency.exe',
};

function checkDriveMapping(letter, expectedUnc) {
  if (!IS_WIN) return true;
  try {
    const out = execSync(`net use ${letter}:`, { encoding: 'utf-8', timeout: 3000 });
    const m = out.match(/Remote name\s+(.+)/i);
    if (m) return m[1].trim().toLowerCase() === expectedUnc.toLowerCase();
  } catch (e) {}
  return false;
}

function loadConfig() {
  const isFirstRun = !fs.existsSync(configPath());
  try {
    if (!isFirstRun) {
      const d = JSON.parse(fs.readFileSync(configPath(), 'utf-8'));
      for (const k of Object.keys(DEFAULTS)) {
        if (k in d) config[k] = d[k];
      }
    }
  } catch (e) { console.error('Config load error:', e); }

  // First install: check if L: and E: map to the expected servers
  if (isFirstRun && IS_WIN) {
    const lOk = checkDriveMapping('L', '\\\\01ckfp02-1\\vol1');
    const eOk = checkDriveMapping('E', '\\\\01ckfp02-1\\Apps');
    if (!lOk) {
      config.order_path   = UNC_DEFAULTS.order_path;
      config.archive_path = UNC_DEFAULTS.archive_path;
      config.drawing_path = UNC_DEFAULTS.drawing_path;
      config.ds_path      = UNC_DEFAULTS.ds_path;
    }
    if (!eOk) {
      config.dwg_viewer      = UNC_DEFAULTS.dwg_viewer;
      config.contingency_exe = UNC_DEFAULTS.contingency_exe;
    }
    saveConfig();
  }
}

function saveConfig() {
  try {
    fs.mkdirSync(configDir(), { recursive: true });
    fs.writeFileSync(configPath(), JSON.stringify({ _v: CFG_VERSION, ...config }, null, 2));
  } catch (e) { console.error('Config save error:', e); }
}

// ─── Drawing name parsing ───────────────────────────────────────────────
const DWG_RE = /^\[?(\d{6}[A-Za-z\d]\d{5})r([0-9A-Za-z]+)\.(dwg|dxf|plt|pdf|bak)(\]?)$/i;
const TL_RE  = /^\[?(TL[A-Za-z0-9]+)r([0-9A-Za-z]+)\.(dwg|dxf|plt|pdf|bak)(\]?)$/i;
const DWG_PDF_RE = /r[0-9][0-9A-Za-z]\.pdf\]?$/i;

function formatDrawingName(filename) {
  let m = DWG_RE.exec(filename);
  if (m) {
    const d = m[1], r = m[2], f = d[0];
    let fmt;
    if ('123'.includes(f))
      fmt = `${d[0]}-${d.slice(1,4)}-${d[4]}-${d.slice(5,7)}-${d.slice(7,10)}-${d.slice(10,12)}`;
    else if ('045'.includes(f))
      fmt = `${d[0]}-${d.slice(1,4)}-${d.slice(4,6)}-${d.slice(6,9)}-${d.slice(9,12)}`;
    else return null;
    const rev = (r.replace(/^0+/, '') || r.slice(-1)).toUpperCase();
    return { name: fmt.toUpperCase(), rev };
  }
  m = TL_RE.exec(filename);
  if (m) {
    const rev = (m[2].replace(/^0+/, '') || m[2].slice(-1)).toUpperCase();
    return { name: m[1].toUpperCase(), rev };
  }
  return null;
}

function isDrawingFile(name, ext) {
  if (['DWG','DXF','PLT','MVW','DWG]'].includes(ext)) return true;
  if (['PDF','PDF]'].includes(ext) && DWG_PDF_RE.test(name)) return true;
  return false;
}

function getExt(name) {
  const dot = name.lastIndexOf('.');
  return dot !== -1 ? name.slice(dot + 1).toUpperCase() : '';
}

function getFileTag(name) {
  const ext = getExt(name);
  const isLink = ext.endsWith(']');
  const base = ext.replace(']', '');
  if (isLink) {
    if (['DWG','DXF','PLT','BAK','MVW'].includes(base)) return 'dwg';
    if (base === 'PDF') return 'pdf';
    return 'link';
  }
  if (isDrawingFile(name, ext)) {
    if (['DWG','DXF','PLT','MVW'].includes(ext)) return 'dwg';
    if (ext === 'PDF') return 'pdf';
  }
  return 'normal';
}

async function buildFileItem(filepath) {
  const name = path.basename(filepath);
  const ext = getExt(name);
  const isDwg = isDrawingFile(name, ext);
  const fmt = isDwg ? formatDrawingName(name) : null;
  const tag = getFileTag(name);
  const isLink = ext.endsWith(']');
  const nameLower = name.toLowerCase();

  let isCnt = false;
  let cntPriority = 99;
  if (ext === 'PDF' && nameLower.includes('contingency')) { isCnt = true; cntPriority = 1; }
  else if (ext === 'PDF' && nameLower.includes('cntgy')) { isCnt = true; cntPriority = 2; }
  else if (ext === 'CNT]') { isCnt = true; cntPriority = 3; }
  else if (ext === 'PDF' && /^\d{6}-\d{1,2}-rev\d/i.test(name)) { isCnt = true; cntPriority = 4; }

  let size = 0, mtime = null;
  try {
    const st = await fsp.stat(filepath);
    size = st.size;
    mtime = st.mtime.toISOString();
  } catch (e) {}

  let displayName = name;
  let revDisplay = '';
  if (fmt) {
    displayName = fmt.name;
    revDisplay = fmt.rev;
  }

  return {
    filepath, name, ext, tag, isLink, isCnt, cntPriority,
    isDrawing: isDwg,
    displayName, revDisplay, size, mtime,
    shownName: fmt ? `${fmt.name} ʀ${fmt.rev}` : name,
  };
}

// ─── Folder resolution ──────────────────────────────────────────────────
async function isDir(p) { try { return (await fsp.stat(p)).isDirectory(); } catch { return false; } }
async function isFile(p) { try { return (await fsp.stat(p)).isFile(); } catch { return false; } }
async function dirExists(p) { return isDir(p); }
async function readdirSafe(p) { try { return await fsp.readdir(p); } catch { return []; } }

async function networkReachable() {
  try { await fsp.access(config.order_path); return true; } catch { /* fall through */ }
  try { await fsp.access(config.drawing_path); return true; } catch { /* fall through */ }
  return false;
}

async function resolveOrderFolder(order, line) {
  const folderName = order.padStart(6, '0') + line.padStart(2, '0');

  // 1. Primary
  const p = path.join(config.order_path, folderName);
  if (await isDir(p)) return { path: p, src: 'orders' };

  // 2. Archive expected
  const prefix = order.padStart(6, '0').slice(0, 2) + '0';
  const a = path.join(config.archive_path, prefix, folderName);
  if (await isDir(a)) return { path: a, src: 'archive' };

  return null;
}

// Deep archive search — broadened scan of all archive subdirectories
async function resolveOrderFolderDeep(order, line) {
  const folderName = order.padStart(6, '0') + line.padStart(2, '0');
  if (await dirExists(config.archive_path)) {
    try {
      for (const sub of await readdirSafe(config.archive_path)) {
        const candidate = path.join(config.archive_path, sub, folderName);
        if (await isDir(candidate))
          return { path: candidate, src: `archive/${sub}` };
      }
    } catch (e) {}
  }
  return null;
}

// ─── Drawing search ─────────────────────────────────────────────────────
const DIGIT_TO_FOLDER = { '0':'0D','1':'1D','2':'2D','3':'3D','4':'4d','5':'5d' };
const DIGIT_TO_DS     = { '0':'0D','1':'1D','2':'2D','3':'3D','4':'4D','5':'5D' };

function parseDrawingInput(raw) {
  let s = raw.trim();
  const h = s.indexOf('#');
  if (h !== -1) s = s.slice(0, h);
  s = s.replace(/[-\s]/g, '');
  s = s.replace(/\.(dwg|dxf|plt|bak|pdf)\]?$/i, '');
  s = s.replace(/r[0-9A-Za-z]+$/i, '');
  return s || null;
}

function buildWildcardRegex(pattern) {
  // Convert wildcard pattern: * and % = multi-char, _ = single-char
  const escaped = pattern.replace(/[.+^${}()|[\]\\]/g, '\\$&');
  const re = escaped.replace(/[*%]/g, '.*').replace(/_/g, '.');
  return new RegExp('^' + re + 'r[0-9a-z]', 'i');
}

async function scanFolder(folder, base, results, includeBak, wildcard) {
  if (!(await dirExists(folder))) return;
  try {
    const baseLower = base.toLowerCase();
    const wcRegex = wildcard ? buildWildcardRegex(baseLower) : null;
    let entries;
    try { entries = await fsp.readdir(folder, { withFileTypes: true }); } catch { return; }
    for (const ent of entries) {
      if (!ent.isFile()) continue;
      const lower = ent.name.toLowerCase();
      const validExt = lower.endsWith('.dwg') || lower.endsWith('.pdf') || (includeBak && lower.endsWith('.bak'));
      if (!validExt) continue;
      const nameNoExt = lower.replace(/\.(dwg|pdf|bak)$/i, '');
      const isFullMatch = !wildcard && nameNoExt.startsWith(baseLower + 'r');
      const isWildcardMatch = wildcard && wcRegex && wcRegex.test(nameNoExt);
      if (isFullMatch || isWildcardMatch) {
        results.push(await buildFileItem(path.join(folder, ent.name)));
      }
    }
  } catch (e) {}
}

async function scanMatchingSubfolders(parentDir, fixedPrefix, searchPattern, results, wildcard) {
  if (!(await dirExists(parentDir))) return;
  try {
    const entries = await fsp.readdir(parentDir, { withFileTypes: true });
    const promises = [];
    for (const ent of entries) {
      if (!ent.isDirectory()) continue;
      if (ent.name.toLowerCase().startsWith(fixedPrefix)) {
        promises.push(scanFolder(path.join(parentDir, ent.name), searchPattern, results, true, wildcard));
      }
    }
    await Promise.all(promises);
  } catch {}
}

async function searchDrawings(baseNum, wildcard, wcPattern) {
  const isTL = baseNum.toUpperCase().startsWith('TL');
  const acadResults = [], dsResults = [];

  const searchPattern = wcPattern || baseNum;
  if (isTL) {
    await scanFolder(path.join(config.drawing_path, 'TL'), searchPattern, acadResults, true, wildcard);
  } else if (wildcard && wcPattern) {
    // When wildcards exist, resolve folders from the original pattern (not stripped base)
    const upperPattern = wcPattern.toUpperCase();
    let fixedLen = 0;
    for (let i = 0; i < Math.min(4, upperPattern.length); i++) {
      if (/[*%_]/.test(upperPattern[i])) break;
      fixedLen = i + 1;
    }
    if (fixedLen >= 4) {
      // No wildcards in first 4 chars — direct folder lookup using pattern prefix
      const pc = upperPattern.slice(0, 4);
      const f = pc[0];
      const sub = DIGIT_TO_FOLDER[f];
      const ds = DIGIT_TO_DS[f];
      await Promise.all([
        sub ? scanFolder(path.join(config.drawing_path, sub, pc), searchPattern, acadResults, true, wildcard) : null,
        ds  ? scanFolder(path.join(config.ds_path, ds, pc), searchPattern, dsResults, true, wildcard) : null,
      ]);
    } else {
      // Wildcards in first 4 chars — enumerate matching subfolders
      const fixedPrefix = upperPattern.slice(0, fixedLen).toLowerCase();
      const digits = fixedLen === 0 ? Object.keys(DIGIT_TO_FOLDER) : [upperPattern[0]];
      const promises = [];
      for (const d of digits) {
        const sub = DIGIT_TO_FOLDER[d];
        const dsSub = DIGIT_TO_DS[d];
        if (sub) promises.push(scanMatchingSubfolders(path.join(config.drawing_path, sub), fixedPrefix, searchPattern, acadResults, wildcard));
        if (dsSub) promises.push(scanMatchingSubfolders(path.join(config.ds_path, dsSub), fixedPrefix, searchPattern, dsResults, wildcard));
      }
      await Promise.all(promises);
    }
  } else {
    const f = baseNum[0] || '';
    const pc = baseNum.slice(0, 4);
    const sub = DIGIT_TO_FOLDER[f];
    const ds = DIGIT_TO_DS[f];
    await Promise.all([
      sub ? scanFolder(path.join(config.drawing_path, sub, pc), searchPattern, acadResults, true, wildcard) : null,
      ds  ? scanFolder(path.join(config.ds_path, ds, pc), searchPattern, dsResults, true, wildcard) : null,
    ]);
  }

  acadResults.forEach(r => { r.location = 'ACADDWGS'; });
  dsResults.forEach(r => { r.location = 'DSGNSTDS'; });

  const allResults = [...acadResults, ...dsResults];

  // Parse rev from each result for grouping
  allResults.forEach(r => {
    const m = DWG_RE.exec(r.name) || TL_RE.exec(r.name);
    r._rev = m ? m[2] : '';
    r._revSort = m ? m[2].padStart(6, '0') : '';
    r._baseNum = m ? m[1].toUpperCase() : r.name.toUpperCase();
    r._ext = r.ext.toUpperCase().replace(']', '');
    r.owner = '';  // Owner loaded lazily on demand
  });

  // Sort: drawing number asc, then rev desc, then DWG before PDF before BAK
  const extOrder = { DWG: 0, PDF: 1, BAK: 2 };
  allResults.sort((a, b) => {
    if (a._baseNum !== b._baseNum) return a._baseNum < b._baseNum ? -1 : 1;
    if (a._revSort !== b._revSort) return b._revSort < a._revSort ? -1 : 1;
    return (extOrder[a._ext] ?? 9) - (extOrder[b._ext] ?? 9);
  });

  return { acad: acadResults, ds: dsResults, all: allResults };
}

// ─── Link creation (SAFE) ───────────────────────────────────────────────
async function createDrawingLink(sourceFilepath, orderFolder, createFolder) {
  const sourceName = path.basename(sourceFilepath);
  const sourceExt = getExt(sourceName);
  const isAlreadyLink = sourceExt.endsWith(']');

  let linkFilename;
  if (isAlreadyLink) {
    linkFilename = sourceName;
  } else {
    const linkExt = ['DWG','DXF','PLT','BAK','MVW'].includes(sourceExt) ? 'dwg]' : 'pdf]';
    let baseName = sourceName;
    if (baseName.startsWith('[')) baseName = baseName.slice(1);
    const dot = baseName.lastIndexOf('.');
    if (dot !== -1) baseName = baseName.slice(0, dot);
    linkFilename = `[${baseName}.${linkExt}`;
  }
  const linkPath = path.join(orderFolder, linkFilename);

  if (await isFile(linkPath)) return { ok: false, msg: `Link already exists:\n${linkFilename}` };

  if (!(await dirExists(orderFolder))) {
    if (createFolder) {
      try { await fsp.mkdir(orderFolder, { recursive: false }); }
      catch (e) {
        if (e.code !== 'EEXIST') return { ok: false, msg: `Could not create folder:\n${e.message}` };
      }
    } else {
      return { ok: false, msg: `Order folder not found:\n${orderFolder}` };
    }
  }

  try {
    if (isAlreadyLink) {
      await fsp.copyFile(sourceFilepath, linkPath, fs.constants.COPYFILE_EXCL);
    } else {
      await fsp.writeFile(linkPath, sourceFilepath, { flag: 'wx' });
    }
    return { ok: true, msg: `Created:\n${linkFilename}\n\nin ${orderFolder}`, folderCreated: createFolder };
  } catch (e) {
    if (e.code === 'EEXIST') return { ok: false, msg: `Link already exists:\n${linkFilename}` };
    return { ok: false, msg: `Error: ${e.message}` };
  }
}

// ─── Birthday check ─────────────────────────────────────────────────────
function checkBirthdays() {
  const candidates = [
    path.join(SHARED_DATA_DIR, 'birthdays.csv'),
    path.join(path.dirname(process.execPath), 'birthdays.csv'),
    path.join(app.isPackaged ? process.resourcesPath : __dirname, 'birthdays.csv'),
    path.join(process.cwd(), 'birthdays.csv'),
    path.join(configDir(), 'birthdays.csv'),
  ];

  let file = null;
  for (const c of candidates) {
    if (fs.existsSync(c)) { file = c; break; }
  }
  if (!file) return null;

  const today = new Date();
  const weekday = today.getDay(); // 0=Sun, 5=Fri, 6=Sat
  const datesToCheck = new Set();
  datesToCheck.add(`${today.getMonth()+1}-${today.getDate()}`);

  if (weekday === 5) { // Friday: check Sat+Sun
    const sat = new Date(today); sat.setDate(sat.getDate() + 1);
    const sun = new Date(today); sun.setDate(sun.getDate() + 2);
    datesToCheck.add(`${sat.getMonth()+1}-${sat.getDate()}`);
    datesToCheck.add(`${sun.getMonth()+1}-${sun.getDate()}`);
  } else if (weekday === 1) { // Monday: check prev Sat+Sun
    const sat = new Date(today); sat.setDate(sat.getDate() - 2);
    const sun = new Date(today); sun.setDate(sun.getDate() - 1);
    datesToCheck.add(`${sat.getMonth()+1}-${sat.getDate()}`);
    datesToCheck.add(`${sun.getMonth()+1}-${sun.getDate()}`);
  }

  const names = [];
  try {
    const lines = fs.readFileSync(file, 'utf-8').split('\n');
    for (const line of lines) {
      const trimmed = line.trim();
      if (!trimmed || trimmed.startsWith('#') || trimmed.toLowerCase().startsWith('name')) continue;
      const parts = trimmed.split(',').map(s => s.trim());
      if (parts.length >= 3) {
        const [name, month, day] = [parts[0], parseInt(parts[1]), parseInt(parts[2])];
        if (!isNaN(month) && !isNaN(day) && datesToCheck.has(`${month}-${day}`)) {
          names.push(name);
        }
      }
    }
  } catch (e) {}

  if (!names.length) return null;
  if (names.length === 1) return `🎂  Happy birthday, ${names[0]}!`;
  return `🎂  Happy birthday to ${names.slice(0, -1).join(', ')} and ${names[names.length - 1]}!`;
}

// ─── Read link target ───────────────────────────────────────────────────
async function readLinkTarget(filepath) {
  try {
    const content = await fsp.readFile(filepath, 'utf-8');
    for (const line of content.split('\n')) {
      const trimmed = line.trim().replace(/^["']|["']$/g, '');
      if (trimmed) return trimmed;
    }
  } catch (e) {}
  return null;
}

// ─── Windows system theme detection ─────────────────────────────────────
function detectSystemDark() {
  if (!IS_WIN) return false;
  try {
    const result = execSync(
      'reg query "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize" /v AppsUseLightTheme',
      { encoding: 'utf-8', timeout: 2000 }
    );
    return result.includes('0x0');
  } catch (e) { return false; }
}

// ─── Login item (launch on startup) ─────────────────────────────────────
function applyLoginItemSettings() {
  if (!app.isPackaged) return; // skip in dev
  app.setLoginItemSettings({ openAtLogin: config.launch_on_startup !== false });
}

// ─── Main window ────────────────────────────────────────────────────────
let mainWindow = null;
let tray = null;
let isQuitting = false;

function createWindow() {
  const winOpts = {
    width: 490, height: 600, minWidth: 360, minHeight: 440,
    frame: false,
    maximizable: false,
    fullscreenable: false,
    titleBarStyle: 'hidden',
    titleBarOverlay: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    icon: path.join(__dirname, 'icons', 'icon.png'),
    show: false,
  };

  // Windows 11: Acrylic backdrop
  if (IS_WIN) {
    winOpts.backgroundColor = '#00000000';
    winOpts.backgroundMaterial = 'acrylic';
  } else {
    winOpts.vibrancy = 'under-window';
    winOpts.visualEffectState = 'active';
    winOpts.backgroundColor = '#00000000';
  }

  mainWindow = new BrowserWindow(winOpts);

  mainWindow.loadFile('index.html');

  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  // Track when the window last had focus — used by tray click to decide show vs hide
  mainWindow.on('focus', () => { _lastFocusedAt = Date.now(); });
  mainWindow.on('blur',  () => { _lastFocusedAt = Date.now(); });

  mainWindow.on('close', (e) => {
    if (!isQuitting && config.minimize_to_tray) {
      e.preventDefault();
      mainWindow.hide();
    }
  });
}

let _lastFocusedAt = 0;

function _trayShow() {
  mainWindow.setOpacity(0);
  mainWindow.show();
  mainWindow.focus();
  setTimeout(() => mainWindow.setOpacity(1), 60);
}

function createTray() {
  if (tray) return;
  let trayIcon;
  try {
    const iconPath = path.join(__dirname, 'icons', 'icon.png');
    trayIcon = nativeImage.createFromPath(iconPath).resize({ width: 16, height: 16 });
    if (trayIcon.isEmpty()) throw new Error('empty');
  } catch (e) {
    trayIcon = nativeImage.createFromDataURL(
      'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAaklEQVQ4y2NgGAXEACYo+z8Q/yfGEGaoAYxQNjYDmKBsbAYwIYkxYDEAq8twGcCEJMaApAYnA3AagNMFOA3AywCsBuAzAKsBhAxgQhJjQFKDkwHoXoCTAfgMwGoAPgOwGkDIAJwuGAWkAQAE/TYR0IflWgAAAABJRU5ErkJggg=='
    );
  }
  tray = new Tray(trayIcon);
  tray.setToolTip('XylemView Pro');
  tray.setContextMenu(Menu.buildFromTemplate([
    { label: 'Show XylemView Pro', click: () => _trayShow() },
    { label: 'Quit', click: () => { isQuitting = true; app.quit(); } },
  ]));
  // Single click: if hidden → show; if was just focused (user clicked tray from app) → hide; else → focus
  tray.on('click', () => {
    if (!mainWindow.isVisible()) {
      _trayShow();
    } else if (Date.now() - _lastFocusedAt < 300) {
      mainWindow.hide();
    } else {
      mainWindow.show();
      mainWindow.focus();
    }
  });
}

// ─── App lifecycle ──────────────────────────────────────────────────────
// Single-instance lock — prevent multiple tray icons
const gotSingleLock = app.requestSingleInstanceLock();
if (!gotSingleLock) {
  app.quit();
} else {
  app.on('second-instance', () => {
    if (mainWindow) {
      if (mainWindow.isMinimized()) mainWindow.restore();
      mainWindow.show();
      mainWindow.focus();
    }
  });
}

app.whenReady().then(() => {
  if (!gotSingleLock) return;
  loadConfig();
  resolveSharedDataDir();
  // Auto-detect latest AutoCAD if not configured
  if (!config.autocad_exe) {
    const detected = detectAutoCAD();
    if (detected) { config.autocad_exe = detected; saveConfig(); }
  }
  applyLoginItemSettings();
  // Log version + user + ODA status to shared network file (fire-and-forget)
  try {
    const logPath = path.join(SHARED_DATA_DIR, 'version-log.csv');
    const username = os.userInfo().username;
    const ts = new Date().toISOString();
    const hasOda = !!findOdaConverter();
    const hasAcad = !!findPdfAccore();
    const line = `${ts},${username},${CURRENT_VERSION},${app.isPackaged ? 'installed' : 'dev'},oda=${hasOda},acad=${hasAcad}\n`;
    fs.appendFile(logPath, line, () => {});
  } catch (e) {}
  createWindow();
  nativeTheme.on('updated', () => { if (win) win.webContents.send('system-theme-changed'); });
  createTray();  // Always show tray icon

  // Global shortcut for quick search (Ctrl+Shift+X)
  globalShortcut.register('CommandOrControl+Shift+X', () => {
    if (mainWindow) {
      if (mainWindow.isVisible()) {
        mainWindow.focus();
      } else {
        mainWindow.show();
        mainWindow.focus();
      }
    }
  });
});

app.on('before-quit', () => { isQuitting = true; });
ipcMain.handle('quit-app', () => { isQuitting = true; app.quit(); });
app.on('window-all-closed', () => { if (!IS_MAC) app.quit(); });
app.on('activate', () => { if (!mainWindow) createWindow(); else mainWindow.show(); });

// ─── Error logging to shared network drive ─────────────────────────────
function logError(source, error) {
  try {
    const logPath = path.join(SHARED_DATA_DIR, 'error-log.csv');
    const username = os.userInfo().username;
    const ts = new Date().toISOString();
    const msg = String(error && error.stack || error || '').replace(/[\r\n]+/g, ' ').slice(0, 500);
    const line = `${ts},${username},${CURRENT_VERSION},${source},"${msg.replace(/"/g, '""')}"\n`;
    fs.appendFile(logPath, line, () => {});
  } catch (e) {}
}

process.on('uncaughtException', (e) => { logError('main-uncaught', e); });
process.on('unhandledRejection', (e) => { logError('main-rejection', e); });

ipcMain.handle('log-renderer-error', (_, msg) => { logError('renderer', msg); });

// ─── IPC Handlers ───────────────────────────────────────────────────────
ipcMain.handle('notify-chat', (_, title, body) => {
  if (tray) tray.displayBalloon({ iconType: 'info', title, content: body });
});

ipcMain.handle('is-window-focused', () => mainWindow ? mainWindow.isFocused() : false);

ipcMain.handle('get-config', () => ({ ...config }));
ipcMain.handle('get-downloads-path', () => app.getPath('downloads'));

ipcMain.handle('save-config', (_, newCfg) => {
  Object.assign(config, newCfg);
  saveConfig();
  applyLoginItemSettings();
  return true;
});

ipcMain.handle('check-network', () => networkReachable());
ipcMain.handle('get-path-defaults', () => ({
  order_path: DEFAULTS.order_path,
  archive_path: DEFAULTS.archive_path,
  drawing_path: DEFAULTS.drawing_path,
  ds_path: DEFAULTS.ds_path,
}));
ipcMain.handle('get-system-dark', () => detectSystemDark());
ipcMain.handle('get-birthdays', () => checkBirthdays());
ipcMain.handle('get-platform', () => process.platform);

ipcMain.handle('search-order', async (_, order, line) => {
  const result = await resolveOrderFolder(order, line);
  if (!result) {
    const expected = path.join(config.order_path, order.padStart(6, '0') + line.padStart(2, '0'));
    const offline = !(await networkReachable());
    // Find sibling lines even though this specific line doesn't exist
    const orderPad = order.padStart(6, '0');
    const lineSet = new Set();
    try {
      for (const name of await readdirSafe(config.order_path)) {
        if (name.startsWith(orderPad) && name.length === 8 && await isDir(path.join(config.order_path, name)))
          lineSet.add(name.slice(6, 8));
      }
    } catch {}
    if (lineSet.size === 0 && config.archive_path) {
      try {
        const prefix = orderPad.slice(0, 2) + '0';
        for (const name of await readdirSafe(path.join(config.archive_path, prefix))) {
          if (name.startsWith(orderPad) && name.length === 8 && await isDir(path.join(config.archive_path, prefix, name)))
            lineSet.add(name.slice(6, 8));
        }
      } catch {}
    }
    // Exclude the line we just failed to find
    lineSet.delete(line.padStart(2, '0'));
    const siblingLines = [...lineSet].sort();
    return { found: false, expected, offline, siblingLines };
  }

  // List files
  const files = [];
  const cntFiles = [];
  try {
    const entries = await readdirSafe(result.path);
    for (const name of entries) {
      const fp = path.join(result.path, name);
      if (name === '$' || name.startsWith('~$')) continue;  // Skip lock files
      const extLc = name.slice(name.lastIndexOf('.') + 1).toLowerCase();
      if (extLc === 'dwl' || extLc === 'dwl2') continue;  // Skip AutoCAD lock files
      if (await isFile(fp)) {
        const fi = await buildFileItem(fp);
        files.push(fi);
        if (fi.isCnt) cntFiles.push(fi);
      }
    }
  } catch (e) {}

  // Sort contingency by priority
  cntFiles.sort((a, b) => a.cntPriority - b.cntPriority);

  // Mark order-line matching DWG/PDF files as order drawings (bold + colored, sorted to top)
  const orderLine = order.padStart(6, '0') + line.padStart(2, '0');
  const orderDwgRe = new RegExp('^\\[?' + orderLine + '(_P\\d{2,3})?\\.(?:dwg|pdf)\\]?$', 'i');
  for (const f of files) {
    const baseExt = f.ext.replace(']', '').toUpperCase();
    if ((baseExt === 'DWG' || baseExt === 'PDF') && orderDwgRe.test(f.name)) {
      f.isOrderDrawing = true;
      f.isDrawing = true;
      f.tag = baseExt === 'DWG' ? 'dwg' : 'pdf';
      f.displayName = f.name.replace(/\.\w+\]?$/, '');
    }
  }

  // Sort order drawings by preferred format, then name
  const pref = config.preferred_format === 'pdf' ? 'PDF' : 'DWG';
  const orderDwgSort = (a, b) => {
    const ae = a.ext.replace(']', '').toUpperCase(), be = b.ext.replace(']', '').toUpperCase();
    if (ae === pref && be !== pref) return -1;
    if (be === pref && ae !== pref) return 1;
    return a.name.localeCompare(b.name);
  };

  // Sort: for 00 folders, all alphabetical by name; otherwise drawings first (desc), others asc
  let sortedFiles;
  if (line.padStart(2, '0') === '00') {
    const od = files.filter(f => f.isOrderDrawing).sort(orderDwgSort);
    const rest = files.filter(f => !f.isOrderDrawing).sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));
    sortedFiles = [...od, ...rest];
  } else {
    const od = files.filter(f => f.isOrderDrawing).sort(orderDwgSort);
    const drawings = files.filter(f => f.isDrawing && !f.isOrderDrawing).sort((a, b) => b.name.localeCompare(a.name));
    const others = files.filter(f => !f.isDrawing && !f.isOrderDrawing).sort((a, b) => a.name.localeCompare(b.name));
    sortedFiles = [...od, ...drawings, ...others];
  }

  // Scan subfolders for "00" line orders
  const subfolders = [];
  if (line.padStart(2, '0') === '00') {
    const EXCLUDE_RE = /^\d{6,}$/;
    const EXCLUDE_NAMES = ['cust requested docs'];
    try {
      const entries = await readdirSafe(result.path);
      for (const name of entries) {
        const fp = path.join(result.path, name);
        if (!(await isDir(fp))) continue;
        const lower = name.toLowerCase();
        if (EXCLUDE_RE.test(name) || EXCLUDE_NAMES.includes(lower)) continue;
        const subFiles = [];
        try {
          for (const sf of await readdirSafe(fp)) {
            if (sf === '$' || sf.startsWith('~$')) continue;
            const sfExtLc = sf.slice(sf.lastIndexOf('.') + 1).toLowerCase();
            if (sfExtLc === 'dwl' || sfExtLc === 'dwl2') continue;
            // Filter Acknowledgment junk files
            if (lower === 'acknowledgments') {
              if (sf.toLowerCase().startsWith('email audit-')) continue;
              if (sf.toLowerCase() === 'engnotified.txt') continue;
            }
            const sfp = path.join(fp, sf);
            if (await isFile(sfp)) subFiles.push(await buildFileItem(sfp));
          }
        } catch {}
        const isKnown = ['acknowledgments', 'marketing'].includes(lower);
        // Sort: Acknowledgments by date (newest first), everything else by name
        if (lower === 'acknowledgments') {
          subFiles.sort((a, b) => (b.mtime || '').localeCompare(a.mtime || ''));
        } else {
          subFiles.sort((a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));
        }
        subfolders.push({ name, nameLower: lower, files: subFiles, isEmpty: subFiles.length === 0, isKnown });
      }
    } catch {}
    const SF_ORDER = { acknowledgments: 0, marketing: 1 };
    subfolders.sort((a, b) => {
      const oa = SF_ORDER[a.nameLower] ?? 2, ob = SF_ORDER[b.nameLower] ?? 2;
      if (oa !== ob) return oa - ob;
      return a.name.localeCompare(b.name);
    });
  }

  return {
    found: true, path: result.path, src: result.src,
    files: sortedFiles,
    cntFiles, subfolders,
    drawingCount: sortedFiles.filter(f => f.isDrawing).length,
    otherCount: sortedFiles.filter(f => !f.isDrawing).length,
  };
});

// Deep archive search — called after the fast search returns not-found
ipcMain.handle('deep-search-order', async (_, order, line) => {
  const result = await resolveOrderFolderDeep(order, line);
  if (!result) return { found: false };
  return { found: true, path: result.path, src: result.src };
});

ipcMain.handle('search-drawing', async (_, rawInput) => {
  const wildcard = /[*%_]/.test(rawInput);
  const cleaned = rawInput.replace(/[*%_]/g, '');
  const base = parseDrawingInput(cleaned);
  if (!base) return { found: false, error: 'Invalid drawing number' };
  // Pass the raw pattern (with wildcards preserved) for regex matching
  const wcPattern = wildcard ? rawInput.replace(/[\s-]/g, '').toUpperCase() : null;
  const { acad, ds, all } = await searchDrawings(base, wildcard, wcPattern);
  if (!all.length) {
    const offline = !(await networkReachable());
    return { found: false, base, wildcard, offline };
  }

  // Filter out BAK files unless they have a newer rev than DWG/PDF
  const newestNonBak = all.find(f => f._ext !== 'BAK');
  const newestBak = all.find(f => f._ext === 'BAK');
  let bakWarning = null;
  if (newestBak && newestNonBak && newestBak._revSort > newestNonBak._revSort) {
    bakWarning = `BAK file has newer rev (ʀ${newestBak._rev}) than any DWG/PDF (ʀ${newestNonBak._rev})`;
  }
  const filtered = all.filter(f => f._ext !== 'BAK');

  // Determine unique drawing base numbers
  const baseNums = new Set(filtered.map(f => {
    const m = DWG_RE.exec(f.name) || TL_RE.exec(f.name);
    return m ? m[1].toUpperCase() : f.name.toUpperCase();
  }));
  const isMultiDrawing = baseNums.size > 1;

  // Wildcard mode: always return flat list when wildcard was used
  if (wildcard) {
    return {
      found: true, base, wildcard: true,
      dir: path.dirname(filtered[0].filepath),
      files: filtered,
    };
  }

  // Single-drawing mode: card view
  const newestRev = filtered[0]._rev;
  const pref = config.preferred_format === 'pdf' ? 'PDF' : 'DWG';
  const newestFiles = filtered.filter(f => f._rev === newestRev);
  const primary = newestFiles.find(f => f._ext === pref) || newestFiles[0];

  // Cross-location warning
  let crossWarning = null;
  if (acad.length && ds.length) {
    const acadNewestRev = acad[0]?._revSort || '';
    const dsNewestRev = ds[0]?._revSort || '';
    if (dsNewestRev >= acadNewestRev) {
      crossWarning = `Drawing ʀ${ds[0]._rev} exists in both ACADDWGS and DSGNSTDS — this shouldn't happen!`;
    }
  }

  // Other revs
  const otherRevs = [];
  const seenRevs = new Set();
  for (const f of filtered) {
    if (f._rev === newestRev || seenRevs.has(f._rev)) continue;
    seenRevs.add(f._rev);
    const revFiles = filtered.filter(r => r._rev === f._rev);
    otherRevs.push(revFiles.find(r => r._ext === pref) || revFiles[0]);
  }

  const altFormat = newestFiles.find(f => f !== primary && f._ext !== 'BAK') || null;

  return {
    found: true, base, wildcard: false, newestRev,
    dir: path.dirname(primary.filepath),
    primary, altFormat, otherRevs,
    crossWarning, bakWarning,
    files: filtered,
  };
});

ipcMain.handle('read-link', async (_, filepath) => {
  return (await readLinkTarget(filepath)) || null;
});

ipcMain.handle('open-external', (_, url) => {
  shell.openExternal(url);
});

ipcMain.handle('winget-install', (_, packageId) => {
  return new Promise((resolve) => {
    const proc = spawn('winget', ['install', '--id', packageId, '--accept-source-agreements', '--accept-package-agreements'], {
      stdio: 'pipe', timeout: 300000
    });
    let stdout = '', stderr = '';
    proc.stdout?.on('data', d => stdout += d.toString());
    proc.stderr?.on('data', d => stderr += d.toString());
    proc.on('close', (code) => {
      if (code === 0) resolve({ ok: true });
      else resolve({ ok: false, msg: (stderr || stdout).trim().slice(-200) });
    });
    proc.on('error', (e) => {
      resolve({ ok: false, msg: 'winget not available: ' + e.message });
    });
  });
});

ipcMain.handle('restart-app', () => {
  app.relaunch();
  app.exit(0);
});

// ─── Marketing folder ───────────────────────────────────────────────────
ipcMain.handle('open-marketing', async (_, order) => {
  const marketingFolder = path.join(config.order_path, order.padStart(6, '0') + '00');
  if (!(await dirExists(marketingFolder))) return { found: false };
  try {
    const entries = await readdirSafe(marketingFolder);
    let hasFiles = false;
    for (const e of entries) { if (await isFile(path.join(marketingFolder, e))) { hasFiles = true; break; } }
    if (hasFiles) { shell.openPath(marketingFolder); return { found: true, path: marketingFolder }; }
    const mktSub = path.join(marketingFolder, 'Marketing');
    if (await dirExists(mktSub)) { shell.openPath(mktSub); return { found: true, path: mktSub }; }
    shell.openPath(marketingFolder);
    return { found: true, path: marketingFolder };
  } catch (e) { return { found: false }; }
});

// ─── Find checklist ─────────────────────────────────────────────────────
ipcMain.handle('find-checklist', async (_, order) => {
  const base00 = path.join(config.order_path, order.padStart(6, '0') + '00');
  const foldersToSearch = [base00, path.join(base00, 'Marketing')];
  let best = null; // { filepath, mtime }
  for (const folder of foldersToSearch) {
    if (!(await dirExists(folder))) continue;
    const entries = await readdirSafe(folder);
    for (const name of entries) {
      if (!name.toLowerCase().startsWith('1_checklist')) continue;
      const fp = path.join(folder, name);
      try {
        const st = await fsp.stat(fp);
        if (!st.isFile()) continue;
        if (!best || st.mtime > best.mtime) best = { filepath: fp, name, mtime: st.mtime };
      } catch {}
    }
  }
  if (best) return { found: true, filepath: best.filepath, name: best.name };
  return { found: false };
});

// ─── Sibling lines ──────────────────────────────────────────────────────
ipcMain.handle('get-sibling-lines', async (_, order) => {
  const orderPad = order.padStart(6, '0');
  const lineSet = new Set();
  // Search primary orders path
  try {
    for (const name of await readdirSafe(config.order_path)) {
      if (name.startsWith(orderPad) && name.length === 8) {
        const fp = path.join(config.order_path, name);
        if (await isDir(fp)) lineSet.add(name.slice(6, 8));
      }
    }
  } catch {}
  // Search archive path (folders are inside subdirectories)
  if (lineSet.size === 0 && config.archive_path) {
    try {
      const prefix = orderPad.slice(0, 2) + '0';
      const archSub = path.join(config.archive_path, prefix);
      for (const name of await readdirSafe(archSub)) {
        if (name.startsWith(orderPad) && name.length === 8) {
          const fp = path.join(archSub, name);
          if (await isDir(fp)) lineSet.add(name.slice(6, 8));
        }
      }
    } catch {}
    // Broadened archive search
    if (lineSet.size === 0) {
      try {
        for (const sub of await readdirSafe(config.archive_path)) {
          for (const name of await readdirSafe(path.join(config.archive_path, sub))) {
            if (name.startsWith(orderPad) && name.length === 8) {
              const fp = path.join(config.archive_path, sub, name);
              if (await isDir(fp)) lineSet.add(name.slice(6, 8));
            }
          }
        }
      } catch {}
    }
  }
  return [...lineSet].sort();
});

ipcMain.handle('open-file', async (_, filepath) => {
  const ext = getExt(path.basename(filepath)).replace(']', '');
  if (['DWG','DXF','PLT','BAK','MVW'].includes(ext)) {
    const exe = config.default_dwg_action === 'autocad' ? config.autocad_exe : config.dwg_viewer;
    if (exe && await isFile(exe)) {
      spawn(exe, [filepath], { detached: true, stdio: 'ignore' }).unref();
      return { ok: true };
    }
  }
  shell.openPath(filepath);
  return { ok: true };
});

ipcMain.handle('open-in-viewer', (_, filepath) => {
  if (config.dwg_viewer && fs.existsSync(config.dwg_viewer)) {
    spawn(config.dwg_viewer, [filepath], { detached: true, stdio: 'ignore' }).unref();
    return { ok: true };
  }
  shell.openPath(filepath);
  return { ok: true };
});

ipcMain.handle('open-in-autocad', (_, filepath) => {
  // Try configured path first
  if (config.autocad_exe && fs.existsSync(config.autocad_exe)) {
    spawn(config.autocad_exe, [filepath], { detached: true, stdio: 'ignore' }).unref();
    return { ok: true };
  }
  // Try auto-detection
  const detected = detectAutoCAD();
  if (detected) {
    config.autocad_exe = detected;
    saveConfig();
    spawn(detected, [filepath], { detached: true, stdio: 'ignore' }).unref();
    return { ok: true, autoDetected: detected };
  }
  return { ok: false, msg: 'AutoCAD not found. Set the path in Settings.' };
});

ipcMain.handle('detect-autocad', () => {
  return detectAutoCAD();
});

ipcMain.handle('check-accore', () => {
  return findPdfAccore() || null;
});

// Check for DXF conversion capability: ODA only
ipcMain.handle('check-dxf-converter', () => {
  const oda = findOdaConverter();
  if (oda) return { tool: 'oda', path: oda };
  return null;
});

function detectAutoCAD() {
  if (!IS_WIN) return null;
  // Scan C:\Program Files\Autodesk\ for AutoCAD installations
  const base = 'C:\\Program Files\\Autodesk';
  if (!fs.existsSync(base)) return null;
  try {
    const dirs = fs.readdirSync(base)
      .filter(d => d.toLowerCase().startsWith('autocad'))
      .sort()
      .reverse(); // newest first (alphabetical = year order)
    for (const dir of dirs) {
      const acad = path.join(base, dir, 'acad.exe');
      if (fs.existsSync(acad)) return acad;
    }
  } catch (e) {}
  return null;
}

ipcMain.handle('open-link', async (_, filepath) => {
  const ext = getExt(path.basename(filepath));
  if (ext === 'CNT]') {
    // Open contingency
    const name = path.basename(filepath);
    let ident = name;
    if (ident.startsWith('[')) ident = ident.slice(1);
    const dot = ident.toLowerCase().lastIndexOf('.cnt]');
    if (dot !== -1) ident = ident.slice(0, dot);
    if (config.contingency_exe && fs.existsSync(config.contingency_exe)) {
      spawn(config.contingency_exe, [ident], { detached: true, stdio: 'ignore' }).unref();
      return { ok: true };
    }
    return { ok: false, msg: 'Contingency program not found' };
  }

  const target = await readLinkTarget(filepath);
  if (!target) return { ok: false, msg: 'No target in link file' };
  if (!(await isFile(target))) return { ok: false, msg: `Target not found:\n${target}` };

  const targetExt = getExt(path.basename(target));
  if (['DWG','DXF','PLT','BAK','MVW'].includes(targetExt)) {
    const exe = config.default_dwg_action === 'autocad' ? config.autocad_exe : config.dwg_viewer;
    if (exe && fs.existsSync(exe)) {
      spawn(exe, [target], { detached: true, stdio: 'ignore' }).unref();
      return { ok: true };
    }
  }
  shell.openPath(target);
  return { ok: true };
});

// ─── Contingency reader ─────────────────────────────────────────────────
// Reads contingency data from BOTH Access databases (read-only, no lock files)
const CONTIN_DB_PATHS = [
  path.join(IS_WIN ? 'L:\\APPS\\HTDatabases\\ContingencyData' : path.join(os.homedir(), 'XylemTest'), 'ContingencyData.mdb'),
  path.join(IS_WIN ? 'L:\\APPS\\HTDatabases\\ContingencyData' : path.join(os.homedir(), 'XylemTest'), 'ContinData_Arc061207.mdb'),
];
let _continCache = null;  // { orderMap: Map<orderLine, data> }
let _continWatchers = [];

// Watch contingency DB files for changes — reload but keep old cache if reload fails
let _continReloadTimer = null;
function setupContinWatchers() {
  for (const w of _continWatchers) try { w.close(); } catch {}
  _continWatchers = [];
  for (const dbPath of CONTIN_DB_PATHS) {
    try {
      const watcher = fs.watch(dbPath, () => {
        console.log('Contingency DB changed:', dbPath);
        // Don't null the cache yet — keep serving stale data until reload succeeds
        if (_continReloadTimer) clearTimeout(_continReloadTimer);
        _continReloadTimer = setTimeout(async () => {
          _continReloadTimer = null;
          if (mainWindow) mainWindow.webContents.send('contingency-loading', true);
          const prev = _continCache;
          _continCache = null;  // Allow loadContingencyDBs to run
          try {
            await loadContingencyDBs();
            // If new cache is empty but old had data, restore old cache
            if (_continCache && _continCache.orderMap.size === 0 && prev && prev.orderMap.size > 0) {
              console.log('Contingency reload returned empty — keeping previous data');
              _continCache = prev;
            }
          } catch (e) {
            console.error('Contingency reload failed, keeping previous data:', e.message);
            if (!_continCache && prev) _continCache = prev;
          }
          if (mainWindow) mainWindow.webContents.send('contingency-loading', false);
        }, 500);
      });
      watcher.on('error', () => {});
      _continWatchers.push(watcher);
    } catch {}
  }
}

async function loadContingencyDBs() {
  if (_continCache) return _continCache;
  const MDBReader = (await import('mdb-reader')).default;
  const tables = ['OrderInfo', 'ConstructionData', 'InspectionData', 'TestingData',
    'CleaningData', 'PaintingData', 'NamePlateData', 'PackagingData', 'PurchasingData', 'QCdata', 'RevNotes', 'Attachments'];
  const orderMap = new Map();
  for (const dbPath of CONTIN_DB_PATHS) {
    try {
      const buf = await fsp.readFile(dbPath);
      const db = new MDBReader(Buffer.from(buf));
      for (const tn of tables) {
        try {
          const rows = db.getTable(tn).getData();
          for (const row of rows) {
            const ol = row.OrderLine;
            if (!ol) continue;
            if (!orderMap.has(ol)) orderMap.set(ol, {});
            const entry = orderMap.get(ol);
            if (tn === 'RevNotes') {
              if (!entry.RevNotes) entry.RevNotes = [];
              entry.RevNotes.push(row);
            } else {
              entry[tn] = row;
            }
          }
        } catch {}
      }
    } catch (e) { console.error('Contingency DB load error:', dbPath, e.message); }
  }
  // If load produced no data (likely network issue), throw so callers can fall back to previous cache
  if (orderMap.size === 0) throw new Error('Contingency DB load returned no data — network may be unavailable');
  _continCache = { orderMap, loadedAt: Date.now() };
  setupContinWatchers();  // Start watching for changes
  return _continCache;
}

ipcMain.handle('preload-contingency', async (_, force) => {
  if (force) _continCache = null; // Force reload
  const needsLoad = !_continCache;
  try {
    if (needsLoad && mainWindow) mainWindow.webContents.send('contingency-loading', true);
    await loadContingencyDBs();
    if (needsLoad && mainWindow) mainWindow.webContents.send('contingency-loading', false);
    return true;
  } catch {
    if (needsLoad && mainWindow) mainWindow.webContents.send('contingency-loading', false);
    return false;
  }
});

ipcMain.handle('get-contingency', async (_, orderLine) => {
  try {
    const cache = await loadContingencyDBs();
    const ol = parseInt(orderLine, 10);
    const entry = cache.orderMap.get(ol);
    if (!entry || !entry.OrderInfo) return { found: false, loadedAt: cache.loadedAt };
    return { found: true, data: entry, loadedAt: cache.loadedAt };
  } catch (e) {
    return { found: false, error: e.message, loadedAt: 0 };
  }
});

const CNT_PAGE_CSS = `
  * { margin:0; padding:0; box-sizing:border-box; }
  body { font:11px/1.6 'Segoe UI', Arial, sans-serif; color:#000; padding:48px 56px; }
  h2 { font:700 15px/1.2 'Segoe UI',sans-serif; margin:0 0 2px; border-bottom:2px solid #000; padding-bottom:4px; }
  .meta { font:10px/1.5 'Consolas','Courier New',monospace; color:#000; margin-bottom:14px; padding-top:4px; }
  .meta b { font-weight:700; }
  .section { margin-bottom:10px; page-break-inside:avoid; }
  .section-title { font:700 10px/1 'Segoe UI',sans-serif; color:#000; text-transform:uppercase; border-bottom:1px solid #888; padding-bottom:2px; margin-bottom:3px; letter-spacing:0.5px; }
  .section-body { padding-left:14px; font:11px/1.6 'Segoe UI',sans-serif; }
  .label { font-weight:700; font-size:9px; text-transform:uppercase; letter-spacing:0.4px; }
  .empty { color:#888; font-style:italic; }
  .footer { margin-top:14px; padding-top:8px; border-top:1px solid #888; font:9px/1.5 'Segoe UI',sans-serif; color:#444; text-align:center; }`;

function buildCntPage(html) {
  return `<!DOCTYPE html><html><head><meta charset="utf-8"><style>${CNT_PAGE_CSS}</style></head><body>${html}</body></html>`;
}

// Save contingency as PDF to Downloads
ipcMain.handle('save-contingency-pdf', async (_, orderLine, html) => {
  let win;
  try {
    const order = orderLine.slice(0, 6), line = orderLine.slice(6);
    win = new BrowserWindow({ show: false, width: 816, height: 1056, webPreferences: { offscreen: true } });
    win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(buildCntPage(html)));
    await new Promise(r => win.webContents.on('did-finish-load', r));
    const pdf = await win.webContents.printToPDF({
      pageSize: 'Letter',
      printBackground: false,
      landscape: false,
      margins: { top: 0, bottom: 0, left: 0, right: 0 },
    });
    const outPath = path.join(app.getPath('downloads'), `${order}-${line} contingency.pdf`);
    await fsp.writeFile(outPath, pdf);
    const { shell } = require('electron');
    shell.openPath(outPath);
    return { ok: true, path: outPath };
  } catch (e) {
    return { ok: false, msg: e.message };
  } finally {
    if (win) win.destroy();
  }
});

// Print contingency directly to a printer (no PDF, no Adobe)
ipcMain.handle('print-contingency', async (_, html) => {
  let win;
  try {
    win = new BrowserWindow({ show: false, width: 816, height: 1056, webPreferences: { offscreen: true } });
    win.loadURL('data:text/html;charset=utf-8,' + encodeURIComponent(buildCntPage(html)));
    await new Promise(r => win.webContents.on('did-finish-load', r));
    const printed = await new Promise((resolve) => {
      win.webContents.print({ silent: false, printBackground: false }, (success, reason) => {
        resolve({ success, reason });
      });
    });
    if (!printed.success && printed.reason !== 'cancelled') {
      return { ok: false, msg: printed.reason || 'Print failed' };
    }
    return { ok: true };
  } catch (e) {
    return { ok: false, msg: e.message };
  } finally {
    if (win) setTimeout(() => { if (win) win.destroy(); }, 2000); // Brief delay so print spooler can finish
  }
});

// ─── Checklist review ───────────────────────────────────────────────────
// Parses order review checklist PDFs to extract line items and drawing numbers

async function parseChecklistPdf(filepath) {
  const pdfParse = require('pdf-parse');
  const buf = await fsp.readFile(filepath);
  const d = await pdfParse(buf);
  const text = d.text;
  const stat = await fsp.stat(filepath);
  const fileDate = stat.mtime.toISOString();
  const result = { filepath, lineItems: [], poNumber: '', ae: '', date: '', fileDate, comments: '' };

  // Extract PO number
  const poMatch = text.match(/PO Number\s+(\S+)/i);
  if (poMatch) result.poNumber = poMatch[1];

  // Extract AE name
  const aeMatch = text.match(/AE\n(\w+)/i);
  if (aeMatch) result.ae = aeMatch[1];

  // Extract date
  const dateMatch = text.match(/Date Submitted:\s*\n(.+)/i);
  if (dateMatch) result.date = dateMatch[1].trim();

  // Extract line items from the section between "Line item" and "U1FORM"/"Standard Paint"/"PO approved"
  const sectionMatch = text.match(/Line item[\s\S]*?(?=U1FORM|Standard Paint|PO approved)/i);
  if (sectionMatch) {
    const section = sectionMatch[0];
    // Find all template/drawing entries — handles multiple separator styles:
    //   %BY5362 / 5360-04F-12-003$1,183.22112 weeks    (price=$1,183.22, qty=1, lead=12 weeks)
    //   %BY5362 / 5360-06F-12-003$1,484.961ship         (price=$1,484.96, qty=1)
    //   %BY5362 / ...$1,484.96  each212 weeks            (price=$1,484.96, qty=2)
    //   %SN5142 PN 5-142-03-024-064$2,415.0017 wk       (price=$2,415.00, qty=1)
    //   %BYRDBU ref 4-244-04-036-037$3,516.60 Net14 std (price=$3,516.60, qty=1)
    // Price ends at .XX, then optional whitespace/text, then qty is always a single digit.
    const itemRe = /(%\w+)\s*(?:\/|PN|ref)\s*([\d][\d\-A-Za-z]+)\$([\d,]+\.\d{2})\s*(?:Net|each)*\s*(\d)/g;
    let m, lineNum = 1;
    while ((m = itemRe.exec(section)) !== null) {
      const template = m[1];
      const dwgRaw = m[2].trim();
      const dwgNorm = dwgRaw.replace(/-/g, '');
      const price = m[3];
      const qty = parseInt(m[4]) || 1;
      result.lineItems.push({ line: String(lineNum).padStart(2, '0'), template, dwgRaw, dwgNorm, price, qty, isRepeat: true });
      lineNum++;
    }
  }

  // Line numbers will be resolved by cross-referencing the acknowledgment (see scan-checklists handler)

  // Extract comments
  const commentsMatch = text.match(/Comments and Additional Requirements:\s*\n([\s\S]*?)(?:Checklist for|Page \d|$)/i);
  if (commentsMatch) {
    result.comments = commentsMatch[1].trim().replace(/\n{2,}/g, '\n');
  }

  return result;
}

// Scan an order's 00 folder for checklist files
async function findChecklists(orderPad) {
  const folder00 = path.join(config.order_path, orderPad + '00');
  if (!(await dirExists(folder00))) return [];
  const files = await readdirSafe(folder00);
  const checklists = files.filter(f => /^1_checklist/i.test(f) && f.toLowerCase().endsWith('.pdf'));
  // Also check Marketing subfolder
  const mktgDir = path.join(folder00, 'Marketing');
  if (await dirExists(mktgDir)) {
    const mktgFiles = await readdirSafe(mktgDir);
    for (const f of mktgFiles) {
      if (/^1_checklist/i.test(f) && f.toLowerCase().endsWith('.pdf'))
        checklists.push(path.join('Marketing', f));
    }
  }
  return checklists.map(f => ({ name: f, path: path.join(folder00, f) }));
}

// Parse the latest order acknowledgment to get line number → drawing number mapping
async function getAckLineMapping(orderPad) {
  const pdfParse = require('pdf-parse');
  const ackDir = path.join(config.order_path, orderPad + '00', 'Acknowledgments');
  if (!(await dirExists(ackDir))) return null;
  const files = (await readdirSafe(ackDir))
    .filter(f => f.toLowerCase().endsWith('.pdf'))
    .sort().reverse();  // Latest first
  if (!files.length) return null;
  try {
    const buf = await fsp.readFile(path.join(ackDir, files[0]));
    const d = await pdfParse(buf);
    const mapping = [];
    // Parse SHIPPING LINE -> FPC ITEM (drawing number) pairs
    const tagRe = /SHIPPING\s+LINE\s+(\d+)[\s\S]*?FPC ITEM #\s+([\dA-Za-z]+)/g;
    let m;
    while ((m = tagRe.exec(d.text)) !== null) {
      mapping.push({ line: String(parseInt(m[1])).padStart(2, '0'), dwg: m[2] });
    }
    return mapping.length ? mapping : null;
  } catch { return null; }
}

ipcMain.handle('scan-checklists', async (_, orderPad) => {
  try {
    const checklists = await findChecklists(orderPad);
    if (!checklists.length) return { found: false };
    const results = [];
    for (const cl of checklists) {
      try {
        const data = await parseChecklistPdf(cl.path);
        results.push(data);
      } catch (e) {
        results.push({ filepath: cl.path, error: e.message, lineItems: [] });
      }
    }
    // Cross-reference with acknowledgment to get correct line numbers
    const ackMap = await getAckLineMapping(orderPad);
    for (const cl of results) {
      if (!cl.lineItems || !cl.lineItems.length) continue;
      if (ackMap && ackMap.length) {
        // Match checklist drawings to ack drawings by order of appearance
        // Both lists should be in the same sequence
        if (cl.lineItems.length === ackMap.length) {
          for (let i = 0; i < cl.lineItems.length; i++) {
            cl.lineItems[i].line = ackMap[i].line;
          }
        } else {
          // Different counts — match by drawing number where possible
          for (const li of cl.lineItems) {
            const match = ackMap.find(a => a.dwg === li.dwgNorm);
            if (match) li.line = match.line;
          }
          // For unmatched items, assign from remaining ack lines in order
          const usedLines = new Set(cl.lineItems.filter(li => li.line !== '00').map(li => li.line));
          const unusedAck = ackMap.filter(a => !usedLines.has(a.line));
          let ui = 0;
          for (const li of cl.lineItems) {
            if (li.line === String(cl.lineItems.indexOf(li) + 1).padStart(2, '0') && ui < unusedAck.length) {
              li.line = unusedAck[ui++].line;
            }
          }
        }
        cl.lineNumberSource = 'acknowledgment';
      } else {
        // No acknowledgment — use sequential odd numbering as best guess
        for (let i = 0; i < cl.lineItems.length; i++) {
          cl.lineItems[i].line = String(1 + i * 2).padStart(2, '0');
        }
        cl.lineNumberSource = 'estimated';
      }
    }
    return { found: true, checklists: results };
  } catch (e) {
    return { found: false, error: e.message };
  }
});

// Scan for new orders (orders with 00 folder + checklist but no line item folders yet)
ipcMain.handle('scan-new-orders', async () => {
  try {
    const entries = await readdirSafe(config.order_path);
    const newOrders = [];
    for (const name of entries) {
      if (name.length !== 8 || !/^\d{8}$/.test(name)) continue;
      if (!name.endsWith('00')) continue;  // Only look at 00 folders
      const orderPad = name.slice(0, 6);
      // Check if there's a checklist in this 00 folder
      const checklists = await findChecklists(orderPad);
      if (!checklists.length) continue;
      // Check if line 01 folder exists (if not, it's likely a new/unprocessed order)
      const line01 = path.join(config.order_path, orderPad + '01');
      if (await dirExists(line01)) continue;  // Already has line folders, skip
      newOrders.push({ order: orderPad, checklists: checklists.length });
    }
    return { orders: newOrders };
  } catch (e) {
    return { orders: [], error: e.message };
  }
});

ipcMain.handle('open-contingency-standalone', () => {
  if (config.contingency_exe && fs.existsSync(config.contingency_exe)) {
    spawn(config.contingency_exe, [], { detached: true, stdio: 'ignore' }).unref();
    return { ok: true };
  }
  return { ok: false, msg: `Not found: ${config.contingency_exe}` };
});

ipcMain.handle('open-contingency-order', (_, ident) => {
  if (config.contingency_exe && fs.existsSync(config.contingency_exe)) {
    spawn(config.contingency_exe, [ident], { detached: true, stdio: 'ignore' }).unref();
    return { ok: true };
  }
  return { ok: false, msg: `Not found: ${config.contingency_exe}` };
});

ipcMain.handle('open-folder', (_, folderPath) => {
  shell.openPath(folderPath);
  return { ok: true };
});

ipcMain.handle('create-link', async (_, sourceFilepath, orderFolder, createFolder) => {
  return await createDrawingLink(sourceFilepath, orderFolder, createFolder);
});

ipcMain.handle('copy-file-to-clipboard', async (_, filepath) => {
  if (IS_WIN) {
    try {
      await execAsync(`powershell -NoProfile -Command "Set-Clipboard -Path '${filepath.replace(/'/g, "''")}'"`);
      return { ok: true };
    } catch (e) { return { ok: false, msg: e.message }; }
  }
  // macOS: no direct file clipboard API easily, just copy path
  clipboard.writeText(filepath);
  return { ok: true, msg: 'Path copied (file copy not supported on macOS)' };
});

ipcMain.handle('paste-files', async (_, targetFolder) => {
  if (!IS_WIN) return { ok: false, msg: 'File paste only supported on Windows' };
  if (!(await dirExists(targetFolder))) return { ok: false, msg: 'Target folder not found' };

  let paths;
  try {
    const { stdout } = await execAsync(
      'powershell -NoProfile -Command "Get-Clipboard -Format FileDropList | ForEach-Object { $_.FullName }"',
      { encoding: 'utf-8', timeout: 5000 }
    );
    paths = stdout.trim().split('\n').map(s => s.trim()).filter(Boolean);
  } catch (e) { return { ok: false, msg: 'Cannot read clipboard' }; }

  if (!paths.length) return { ok: false, msg: 'No files on clipboard' };

  let copied = 0, skipped = 0;
  for (const src of paths) {
    if (!(await isFile(src))) continue;
    const dest = path.join(targetFolder, path.basename(src));
    if (await isFile(dest)) { skipped++; continue; }
    try { await fsp.copyFile(src, dest, fs.constants.COPYFILE_EXCL); copied++; }
    catch (e) { skipped++; }
  }

  return { ok: true, copied, skipped, files: paths.map(p => path.basename(p)) };
});

ipcMain.handle('copy-text', (_, text) => {
  clipboard.writeText(text);
  // Verify the clipboard actually contains what we wrote
  const readBack = clipboard.readText();
  return { ok: readBack === text };
});

ipcMain.handle('send-email', (_, email, subject) => {
  const params = subject ? `?subject=${encodeURIComponent(subject)}&body=%20` : '';
  shell.openExternal(`mailto:${email}${params}`);
});

ipcMain.handle('get-version', () => ({ version: CURRENT_VERSION, isDev: !app.isPackaged }));

// Window controls (frameless window)
ipcMain.handle('win-minimize', () => mainWindow?.minimize());
ipcMain.handle('win-maximize', () => {
  if (mainWindow?.isMaximized()) mainWindow.unmaximize();
  else mainWindow?.maximize();
});
ipcMain.handle('win-close', () => mainWindow?.close());
ipcMain.handle('win-is-maximized', () => mainWindow?.isMaximized() || false);

ipcMain.handle('set-glass-mode', (_, enabled) => {
  if (!mainWindow) return;
  if (IS_WIN) {
    mainWindow.setBackgroundMaterial(enabled ? 'acrylic' : 'none');
  }
});

// ─── Recent orders ──────────────────────────────────────────────────────
ipcMain.handle('get-recent-orders', () => {
  return config.recent_orders || [];
});

ipcMain.handle('add-recent-order', async (_, order, line) => {
  // Only add if the order folder actually exists
  const result = await resolveOrderFolder(order, line);
  if (!result) return config.recent_orders || [];
  const key = `${order.padStart(6,'0')}-${line.padStart(2,'0')}`;
  let list = config.recent_orders || [];
  list = list.filter(r => r.key !== key);
  list.unshift({ key, order: order.padStart(6,'0'), line: line.padStart(2,'0'), ts: Date.now() });
  // Keep last 15
  config.recent_orders = list.slice(0, 15);
  saveConfig();
  return config.recent_orders;
});

ipcMain.handle('get-recent-drawings', () => {
  return config.recent_drawings || [];
});

ipcMain.handle('add-recent-drawing', (_, query) => {
  let list = config.recent_drawings || [];
  list = list.filter(r => r.query !== query);
  list.unshift({ query, ts: Date.now() });
  config.recent_drawings = list.slice(0, 6);
  saveConfig();
  return config.recent_drawings;
});

ipcMain.handle('add-recent-opened-drawing', (_, displayName, searchQuery) => {
  let list = config.recent_opened_drawings || [];
  list = list.filter(r => r.name !== displayName);
  list.unshift({ name: displayName, query: searchQuery, ts: Date.now() });
  config.recent_opened_drawings = list.slice(0, 8);
  saveConfig();
  return config.recent_opened_drawings;
});

ipcMain.handle('get-recent-opened-drawings', () => {
  return config.recent_opened_drawings || [];
});

// ─── Rev-checking (async) ───────────────────────────────────────────────
// For each link file in an order folder, check if a newer revision exists
ipcMain.handle('check-revisions', async (_, files) => {
  const results = [];
  for (const f of files) {
    if (!f.isLink || !f.isDrawing) continue;

    let linkName = f.name;
    if (linkName.startsWith('[')) linkName = linkName.slice(1);
    const dot = linkName.lastIndexOf('.');
    if (dot !== -1) linkName = linkName.slice(0, dot);

    const m = DWG_RE.exec(linkName + '.dwg') || TL_RE.exec(linkName + '.dwg');
    if (!m) continue;

    const baseNum = m[1];
    const currentRev = m[2];

    const allRevs = [];
    const isTL = baseNum.toUpperCase().startsWith('TL');

    if (isTL) {
      await scanFolder(path.join(config.drawing_path, 'TL'), baseNum, allRevs);
    } else {
      const digit = baseNum[0] || '';
      const pc = baseNum.slice(0, 4);
      const sub = DIGIT_TO_FOLDER[digit];
      if (sub) await scanFolder(path.join(config.drawing_path, sub, pc), baseNum, allRevs);
      if (!allRevs.length) {
        const ds = DIGIT_TO_DS[digit];
        if (ds) await scanFolder(path.join(config.ds_path, ds, pc), baseNum, allRevs);
      }
    }

    // Find the highest revision
    let newestRev = currentRev;
    for (const rev of allRevs) {
      const rm = DWG_RE.exec(rev.name) || TL_RE.exec(rev.name);
      if (rm) {
        const r = rm[2];
        if (r.padStart(6, '0').toUpperCase() > newestRev.padStart(6, '0').toUpperCase()) {
          newestRev = r;
        }
      }
    }

    const isOutdated = newestRev.padStart(6, '0').toUpperCase() > currentRev.padStart(6, '0').toUpperCase();
    if (isOutdated) {
      const cleanCurrent = (currentRev.replace(/^0+/, '') || currentRev.slice(-1)).toUpperCase();
      const cleanNewest = (newestRev.replace(/^0+/, '') || newestRev.slice(-1)).toUpperCase();
      results.push({
        filepath: f.filepath,
        name: f.name,
        currentRev: cleanCurrent,
        newestRev: cleanNewest,
        rawNewestRev: newestRev,  // Preserve original padding for link creation
        isOutdated: true,
      });
    }
  }
  return results;
});

// ─── PDF preview path lookup ────────────────────────────────────────────
// Given a drawing file, find a matching PDF for preview
ipcMain.handle('find-pdf-preview', async (_, filepath) => {
  const name = path.basename(filepath);
  const dir = path.dirname(filepath);

  if (name.toLowerCase().endsWith('.pdf')) return { found: true, pdfPath: filepath };

  const baseName = name.replace(/\.(dwg|dxf|plt)\]?$/i, '');
  try {
    for (const f of await readdirSafe(dir)) {
      if (f.toLowerCase().startsWith(baseName.toLowerCase()) && f.toLowerCase().endsWith('.pdf')) {
        return { found: true, pdfPath: path.join(dir, f) };
      }
    }
  } catch (e) {}

  // For link files, read target and check its directory
  if (name.startsWith('[')) {
    const target = await readLinkTarget(filepath);
    if (target) {
      const targetDir = path.dirname(target);
      const targetBase = path.basename(target).replace(/\.(dwg|dxf|plt)$/i, '');
      try {
        for (const f of await readdirSafe(targetDir)) {
          if (f.toLowerCase().startsWith(targetBase.toLowerCase()) && f.toLowerCase().endsWith('.pdf')) {
            return { found: true, pdfPath: path.join(targetDir, f) };
          }
        }
      } catch (e) {}
    }
  }

  return { found: false };
});

// ─── Disco song picker ──────────────────────────────────────────────────
ipcMain.handle('pick-disco-song', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    title: 'Choose Your Disco Song',
    filters: [{ name: 'Audio', extensions: ['mp3', 'm4a', 'wav', 'ogg', 'aac', 'flac'] }],
    properties: ['openFile'],
  });
  if (result.canceled || !result.filePaths.length) return null;
  const fp = result.filePaths[0];
  config.disco_song = fp;
  saveConfig();
  const ext = path.extname(fp).slice(1).toLowerCase();
  const mime = { mp3:'audio/mpeg', m4a:'audio/mp4', wav:'audio/wav', ogg:'audio/ogg', aac:'audio/aac', flac:'audio/flac' }[ext] || 'audio/mpeg';
  const data = fs.readFileSync(fp);
  return { dataUrl: `data:${mime};base64,${data.toString('base64')}`, name: path.basename(fp) };
});

ipcMain.handle('get-disco-song', async () => {
  // 1. User's custom pick
  try {
    if (config.disco_song) {
      await fsp.access(config.disco_song);
      const fp = config.disco_song;
      const ext = path.extname(fp).slice(1).toLowerCase();
      const mime = { mp3:'audio/mpeg', m4a:'audio/mp4', wav:'audio/wav', ogg:'audio/ogg' }[ext] || 'audio/mpeg';
      const data = await fsp.readFile(fp);
      return { dataUrl: `data:${mime};base64,${data.toString('base64')}`, name: path.basename(fp) };
    }
  } catch {}
  // 2. Network default
  try {
    const networkDefault = IS_WIN ? path.join(SHARED_DATA_DIR, 'disco.mp3') : '';
    if (networkDefault) {
      await fsp.access(networkDefault);
      const data = await fsp.readFile(networkDefault);
      return { dataUrl: `data:audio/mpeg;base64,${data.toString('base64')}`, name: 'disco.mp3' };
    }
  } catch {}
  return null;
});

// ─── User info ──────────────────────────────────────────────────────────
ipcMain.handle('get-user-info', async () => {
  const info = await getUserInfoCached();
  // Also try to get profile picture
  let avatarDataUrl = null;
  if (IS_WIN) {
    try {
      const picDir = path.join('C:\\Users', info.username, 'AppData', 'Roaming', 'Microsoft', 'Windows', 'AccountPictures');
      const pics = (await readdirSafe(picDir)).filter(f => /\.(png|jpg|jpeg|bmp)$/i.test(f));
      let bestPic = null; let bestSize = 0;
      for (const p of pics) {
        const fp = path.join(picDir, p);
        try { const st = await fsp.stat(fp); if (st.size > bestSize) { bestSize = st.size; bestPic = fp; } } catch {}
      }
      if (bestPic) {
        const data = await fsp.readFile(bestPic);
        const ext = path.extname(bestPic).slice(1).toLowerCase();
        const mime = ext === 'jpg' || ext === 'jpeg' ? 'jpeg' : ext;
        avatarDataUrl = `data:image/${mime};base64,${data.toString('base64')}`;
      }
    } catch (e) {}
  }
  return { ...info, avatarDataUrl };
});

// ─── Rev update — replace old link files with newer revisions ────────────
// Safety: old link is only deleted after the new link is confirmed written,
// and only if it passes all safety checks (bracket name, tiny text file).
function isSafeLinkToDelete(filepath, name) {
  // Must be a bracket-wrapped drawing link filename
  if (!name.startsWith('[') || !name.endsWith(']')) return false;
  if (!/\.(dwg|dxf|plt|pdf)\]$/i.test(name)) return false;
  return true;
}

ipcMain.handle('update-rev-links', async (_, updates) => {
  const results = [];
  for (const u of updates) {
    const dir = path.dirname(u.oldLinkPath);
    const oldName = path.basename(u.oldLinkPath);
    const newLinkName = `[${u.newDrawingName}]`;
    const newLinkPath = path.join(dir, newLinkName);
    if (await isFile(newLinkPath)) {
      // New link already exists — just remove the old one if safe
      if (isSafeLinkToDelete(u.oldLinkPath, oldName)) {
        try {
          const st = await fsp.stat(u.oldLinkPath);
          if (st.size < 1024) await fsp.unlink(u.oldLinkPath);
        } catch {}
      }
      results.push({ ok: true, name: newLinkName, replaced: true });
      continue;
    }
    try {
      const oldTarget = await readLinkTarget(u.oldLinkPath);
      if (oldTarget) {
        const targetDir = path.dirname(oldTarget);
        const newTarget = path.join(targetDir, u.newDrawingName);
        // Write new link first — atomic 'wx' flag ensures no overwrite
        await fsp.writeFile(newLinkPath, newTarget, { flag: 'wx' });
        // Verify the new link was written successfully before touching the old one
        if (await isFile(newLinkPath)) {
          // Safe to remove old link: verified bracket name, small text file
          if (isSafeLinkToDelete(u.oldLinkPath, oldName)) {
            try {
              const st = await fsp.stat(u.oldLinkPath);
              if (st.size < 1024) await fsp.unlink(u.oldLinkPath);
            } catch {}  // Non-fatal: old link stays if delete fails
          }
        }
        results.push({ ok: true, name: newLinkName, replaced: true });
      } else {
        results.push({ ok: false, name: newLinkName, reason: 'Could not read old link' });
      }
    } catch (e) {
      results.push({ ok: false, name: newLinkName, reason: e.message });
    }
  }
  return results;
});

// ─── Konami faces ──────────────────────────────────────────────────────
ipcMain.handle('get-konami-cursor', async () => {
  const candidates = [
    path.join(config.order_path, '..', 'Router Generation', "Cursor of Will's face.cur"),
    'F:\\Group\\Router Generation\\Cursor of Will\'s face.cur',
    '\\\\01ckfp02-1\\vol1\\Group\\Router Generation\\Cursor of Will\'s face.cur',
  ];
  for (const fp of candidates) {
    try {
      const data = await fsp.readFile(fp);
      return `data:image/x-icon;base64,${data.toString('base64')}`;
    } catch {}
  }
  return null;
});

ipcMain.handle('get-konami-faces', async () => {
  const facesDir = path.join(SHARED_DATA_DIR, 'faces');
  try {
    const files = await readdirSafe(facesDir);
    const pngs = files.filter(f => /\.(png|jpg|jpeg|gif|webp)$/i.test(f));
    const faces = [];
    for (const f of pngs) {
      try {
        const data = await fsp.readFile(path.join(facesDir, f));
        const ext = path.extname(f).slice(1).toLowerCase();
        const mime = { jpg:'jpeg', jpeg:'jpeg', gif:'gif', webp:'webp' }[ext] || 'png';
        faces.push(`data:image/${mime};base64,${data.toString('base64')}`);
      } catch {}
    }
    return faces;
  } catch { return []; }
});

// ─── BOM Extract from DWG ─────────────────────────────────────────────

// Parse DXF text for BOM data — groups attributes by INSERT block
function parseDxfBom(dxfText) {
  const lines = dxfText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  const bomLines = [];
  const header = { order: '', lineItem: '', contract: '', mxpn: '', dwgno: '', model: '' };

  const BOM_TAGS = new Set(['ITEM', 'QTY', 'DESC', 'MATL', 'SPEC', 'PATT', 'PN', 'DWG']);
  const INFO_TAGS = new Set(['LENG', 'MXPN', 'ORDER', 'LINE_ITEM', 'CONTRACT', 'MODEL']);
  const ALL_TAGS = new Set([...BOM_TAGS, ...INFO_TAGS]);

  let i = 0;
  while (i < lines.length - 1) {
    const code = lines[i].trim();
    const val = lines[i + 1].trim();

    if (code === '0' && val === 'INSERT') {
      // Check if this INSERT has attributes (gc 66 = 1)
      let hasAttribs = false;
      let j = i + 2;
      while (j < lines.length - 1) {
        const c = lines[j].trim();
        const v = lines[j + 1].trim();
        if (c === '0') break;
        if (c === '66' && v === '1') hasAttribs = true;
        j += 2;
      }

      if (hasAttribs) {
        // Collect all ATTRIB entities until SEQEND
        const attrs = {};
        while (j < lines.length - 1) {
          const c = lines[j].trim();
          const v = lines[j + 1].trim();

          if (c === '0' && v === 'SEQEND') { j += 2; break; }
          if (c === '0' && v !== 'ATTRIB') break; // safety
          if (c === '0' && v === 'ATTRIB') {
            let attrTag = null, attrValue = null;
            let k = j + 2;
            while (k < lines.length - 1) {
              const cc = lines[k].trim();
              if (cc === '0') break;
              const vv = lines[k + 1].trim();
              if (cc === '2') {
                const lookback = lines.slice(Math.max(k - 4, 0), k + 1).join('\n');
                if (lookback.includes('AcDbAttribute')) attrTag = vv;
              }
              if (cc === '1' && attrValue === null) attrValue = lines[k + 1].trimEnd();
              k += 2;
            }
            if (attrTag && ALL_TAGS.has(attrTag.toUpperCase())) {
              attrs[attrTag.toUpperCase()] = (attrValue || '').trim();
            }
            j = k;
            continue;
          }
          j += 2;
        }

        // Check if this block is a BOM line (has ITEM + at least PN or DESC)
        if (attrs.ITEM && (attrs.PN || attrs.DESC)) {
          bomLines.push({
            item: attrs.ITEM || '', qty: attrs.QTY || '', desc: attrs.DESC || '',
            matl: attrs.MATL || '', spec: attrs.SPEC || '', patt: attrs.PATT || '',
            pn: attrs.PN || '', dwg: attrs.DWG || '', leng: attrs.LENG || '',
          });
        }

        // Extract header info (first non-empty occurrence wins)
        if (attrs.ORDER && !header.order) header.order = attrs.ORDER;
        if (attrs.LINE_ITEM && !header.lineItem) header.lineItem = attrs.LINE_ITEM;
        if (attrs.CONTRACT && !header.contract) header.contract = attrs.CONTRACT;
        if (attrs.MXPN && !header.mxpn) header.mxpn = attrs.MXPN;
        if (attrs.MODEL && !header.model) header.model = attrs.MODEL;

        i = j;
        continue;
      }
    }
    i += 2;
  }

  // Sort BOM lines by item number
  bomLines.sort((a, b) => (parseInt(a.item) || 0) - (parseInt(b.item) || 0));

  // Extract order/line from contract if not set directly
  if (!header.order && !header.lineItem && header.contract) {
    const digits = header.contract.replace(/\D/g, '');
    if (digits.length >= 8) {
      header.order = digits.substring(0, 6);
      header.lineItem = digits.substring(6, 8);
    }
  }
  // Pad line item to 2 digits
  if (header.lineItem.length === 1) header.lineItem = '0' + header.lineItem;

  // Extract drawing number from MXPN (digits + # only)
  if (header.mxpn) {
    header.dwgno = header.mxpn.replace(/[^0-9#]/g, '');
    if (header.dwgno.length < 3) header.dwgno = '';
  }

  // Derive parent PN from MXPN (the %PN field)
  header.parentPn = '';
  if (header.mxpn) {
    const cleaned = header.mxpn.replace(/[^0-9%]/g, '');
    if (cleaned.startsWith('%') && cleaned.length >= 10) header.parentPn = cleaned;
  }

  return { header, items: bomLines };
}

// Qty adjustment for 1- part numbers (matches VBA BOM_maker_A21 logic)
// For TIEROD/PIPE/TUBE/SPACER/PASS RIB items: BPCS_qty = qty × length × scrapFactor
// Length is parsed from the PATT (size/pattern) and LENG attributes
function adjustBomQuantities(items) {
  items.forEach(item => {
    const pn = (item.pn || '').trim();
    if (!pn.startsWith('1-') || pn.length <= 7) return;
    const qty = parseFloat(item.qty) || 0;
    if (qty <= 0.01) return;

    const desc = (item.desc || '').toUpperCase();
    let scrapFactor = 1.1;
    let calcType = 0; // 0=none, 1=length, 2=area (pass rib)

    if (desc.includes('TIEROD')) calcType = 1;
    if (desc.includes('SPACER')) { calcType = 1; scrapFactor = 1.02; }
    if (desc.includes('PIPE')) calcType = 1;
    if (desc.includes('TUBE')) calcType = 1;
    if (desc.includes('PASS RIB')) { calcType = 2; scrapFactor = 1.3; }
    if (calcType === 0) return;

    // Combine patt + leng (VBA appends LENG to PATT for SC_ABOM blocks)
    let spec = (item.patt || '').toUpperCase();
    const leng = (item.leng || '').toUpperCase();
    if (leng && leng.includes('LG') && !spec.includes('LG')) {
      spec = spec + ' X ' + leng;
    }

    let lgItem = 0;
    if (calcType === 1) {
      // Parse length: take last number before "LG" in pattern
      let txt = spec;
      let p = txt.indexOf('LG');
      if (p > 0) txt = txt.substring(0, p + 2);
      p = txt.indexOf('X');
      if (p > 0) txt = txt.substring(p + 1);
      p = txt.indexOf('X');
      if (p > 0) txt = txt.substring(p + 1);
      p = txt.indexOf('LG');
      if (p > 0) txt = txt.substring(0, p);
      lgItem = parseFloat(txt) || 0;
    } else if (calcType === 2) {
      // Area calc for pass ribs: width × length after "TK"
      let txt = spec;
      let p = txt.indexOf('TK');
      if (p >= 0) txt = txt.substring(p + 2);
      p = txt.indexOf('X');
      if (p > 0) txt = txt.substring(p + 1);
      p = txt.indexOf('X');
      if (p > 0) {
        const w = parseFloat(txt.substring(0, p)) || 0;
        const l = parseFloat(txt.substring(p + 1)) || 0;
        lgItem = w * l;
        if (w < 1.7 || l < 1.7) lgItem = 0.01;
      }
    }

    if (lgItem > 0.01) {
      const bpcsQty = qty * lgItem * scrapFactor;
      item.origQty = item.qty;
      item.qty = bpcsQty.toFixed(1);
    }
  });
}

// Format BOM as pipe-delimited .bom file (matches current BPCS upload format)
function formatBomFile(order, lineItem, parentPn, items, model) {
  const header = `REMARKS->>  ORDER: ${order}-${lineItem}, ....MODEL: ${model || ''}`;
  const padPn = parentPn.padEnd(15);
  const bomLines = items.map((item, idx) => {
    const seq = String(idx + 1).padStart(3, '0');
    const qty = String(parseFloat(item.qty) || 0);
    const paddedQty = qty.length < 5 ? qty.padStart(5) : qty;
    // Format PN with dashes (original format from drawing)
    const pn = (item.pn || '').trim();
    const desc = (item.desc || '').trim();
    return `${padPn}|${paddedQty}|${pn}|<${seq}> ${desc}`;
  });
  return [header, ...bomLines].join('\n') + '\n';
}

ipcMain.handle('extract-bom-from-dwg', async (_, dwgPath) => {
  const oda = findOdaConverter();
  if (!oda) return { ok: false, msg: 'ODA File Converter required for BOM extraction.' };
  if (!fs.existsSync(dwgPath)) return { ok: false, msg: `Drawing not found: ${path.basename(dwgPath)}` };

  const tmpDir = path.join(os.tmpdir(), 'xv_bom_' + Date.now());
  try {
    fs.mkdirSync(tmpDir, { recursive: true });
    const r = await odaConvertDwgToDxf(dwgPath, tmpDir);
    if (!r.ok) return { ok: false, msg: 'DWG→DXF conversion failed: ' + r.msg };

    const dxfText = await fsp.readFile(r.dxfPath, 'utf-8');
    try { fs.unlinkSync(r.dxfPath); fs.rmdirSync(tmpDir); } catch {}

    const result = parseDxfBom(dxfText);
    if (result.items.length === 0) return { ok: false, msg: 'No BOM lines found in drawing.' };

    adjustBomQuantities(result.items);
    return { ok: true, ...result };
  } catch (e) {
    try { fs.rmdirSync(tmpDir, { recursive: true }); } catch {}
    return { ok: false, msg: 'BOM extraction failed: ' + e.message };
  }
});

ipcMain.handle('save-bom', async (_, { order, lineItem, parentPn, items, upload, model }) => {
  if (!order || order.length !== 6) return { ok: false, msg: 'Order must be 6 digits.' };
  if (!lineItem || !/^\d{1,2}$/.test(lineItem)) return { ok: false, msg: 'Line item must be 1-2 digits.' };
  if (!parentPn) return { ok: false, msg: 'Parent part number is required.' };
  if (!items || items.length === 0) return { ok: false, msg: 'No BOM items to save.' };

  const li = lineItem.padStart(2, '0');
  // Filename based on parent item number (cleaned for filesystem safety)
  const safePn = parentPn.replace(/[<>:"/\\|?*]/g, '_');
  const baseName = `${safePn}_xvp`;
  const content = formatBomFile(order, li, parentPn, items, model);

  // Save to Downloads folder
  const dest = app.getPath('downloads');
  try {
    const target = path.join(dest, `${baseName}.bom`);
    let finalPath = target;
    if (fs.existsSync(target)) {
      for (let i = 2; i <= 99; i++) {
        const alt = path.join(dest, `${baseName}_(${i}).bom`);
        if (!fs.existsSync(alt)) { finalPath = alt; break; }
      }
    }
    await fsp.writeFile(finalPath, content, 'utf-8');
    const savedName = path.basename(finalPath);

    // Optionally upload to BPCSupload
    if (upload) {
      const bpcs = IS_WIN ? 'E:\\BPCSupload' : path.join(os.homedir(), 'XylemTest/BPCSupload');
      try {
        await fsp.mkdir(bpcs, { recursive: true });
        await fsp.copyFile(finalPath, path.join(bpcs, savedName));
        return { ok: true, msg: `Uploaded ${savedName} to BPCSupload`, fileName: savedName };
      } catch (e) { return { ok: false, msg: `Saved to Downloads but upload failed: ${e.message}` }; }
    }

    return { ok: true, msg: `Saved ${savedName} to Downloads`, fileName: savedName };
  } catch (e) { return { ok: false, msg: 'Save failed: ' + e.message }; }
});

// ─── Resubmit BOM/DRS ──────────────────────────────────────────────────
ipcMain.handle('resubmit-file', async (_, filepath) => {
  const name = path.basename(filepath);
  const ext = getExt(name).toUpperCase();

  if (ext === 'BOM') {
    const dest = IS_WIN ? 'E:\\BPCSupload' : path.join(os.homedir(), 'XylemTest/BPCSupload');
    try {
      await fsp.mkdir(dest, { recursive: true });
      const target = path.join(dest, name);
      await fsp.copyFile(filepath, target);
      return { ok: true, msg: `Uploaded ${name} to BPCSupload` };
    } catch (e) { return { ok: false, msg: e.message }; }
  }

  if (ext === 'DRS') {
    try {
      const content = await fsp.readFile(filepath, 'utf-8');
      const firstLine = content.split('\n')[0] || '';
      let dest;
      if (firstLine.startsWith('FAMILY|SDX')) {
        dest = IS_WIN ? 'E:\\Sondex\\DWGpoll_ORDERS' : path.join(os.homedir(), 'XylemTest/Sondex');
      } else if (firstLine.startsWith('FAMILY|BP')) {
        dest = IS_WIN ? 'E:\\brazepak\\DwgPoll' : path.join(os.homedir(), 'XylemTest/brazepak/DwgPoll');
      } else {
        return { ok: false, msg: `Unknown DRS type: "${firstLine.slice(0, 30)}…"` };
      }
      await fsp.mkdir(dest, { recursive: true });
      const target = path.join(dest, name);
      await fsp.copyFile(filepath, target);
      return { ok: true, msg: `Submitted ${name} to ${path.basename(dest)}` };
    } catch (e) { return { ok: false, msg: e.message }; }
  }

  return { ok: false, msg: 'Unsupported file type' };
});

// ─── Parse DRS for GPHE drawing number ─────────────────────────────────
ipcMain.handle('parse-drs-gphe', async (_, filepath) => {
  try {
    const content = await fsp.readFile(filepath, 'utf-8');
    const firstLine = content.split('\n')[0] || '';
    if (!firstLine.startsWith('FAMILY|SDX')) return { isGphe: false };
    const m = content.match(/~DWGNO\|([^~]+)~/);
    if (!m) return { isGphe: false };
    // Drawing number comes hyphenated like X-XXX-XX-XXX-XXX — strip hyphens for standard format
    const drawingNum = m[1].replace(/-/g, '');
    return { isGphe: true, drawingNum };
  } catch { return { isGphe: false }; }
});

// ─── Parse DRS for Brazed (FAMILY|BP) routing ─────────────────────────
ipcMain.handle('parse-drs-brazed', async (_, filepath) => {
  try {
    const content = await fsp.readFile(filepath, 'utf-8');
    const firstLine = content.split('\n')[0] || '';
    if (!firstLine.startsWith('FAMILY|BP')) return { ok: false };

    const tag = key => { const m = content.match(new RegExp('~' + key + '[|]([^~]*)')); return m ? m[1].trim() : ''; };

    // Model type: BPDW (double-wall), BPN (nickel), BP (single copper)
    const plateType = tag('PLATETYPE');
    const brazeMtl = tag('BRAZEMTL');
    let modType;
    if (plateType === 'DOUBLE') modType = 'BPDW';
    else if (brazeMtl === 'NICKEL') modType = 'BPN';
    else modType = 'BP';

    // Model number — first 3 digits after MODEL|
    const modelRaw = tag('MODEL');
    const modNo = parseInt(modelRaw.substring(0, 3)) || 0;

    // Plate count
    const plates = parseInt(tag('PLATEQUAN')) || 0;

    // Welded studs
    const studs = tag('MTGOPT') === 'STUDS';

    // ASME code
    const asme = tag('ASME') === 'YES';

    // Count front and back ports for swaged connections
    const portIds = new Set();
    for (const m of content.matchAll(/~([FB]\d+)\w+\|/g)) portIds.add(m[1]);
    const frontPorts = [...portIds].filter(p => p.startsWith('F')).length;
    const backPorts = [...portIds].filter(p => p.startsWith('B')).length;
    const swagedConns = (frontPorts > 0 && backPorts > 0) ? Math.min(frontPorts, backPorts) : 0;

    return { ok: true, modType, modNo, plates, studs, asme, swagedConns, frontPorts, backPorts, model: modelRaw };
  } catch { return { ok: false }; }
});

// ─── DWG → DXF Conversion ──────────────────────────────────────────────
ipcMain.handle('convert-dwg-dxf', async (_, dwgPath, outDir) => {
  // ODA only — faster and free, no accoreconsole fallback for DXF
  const result = await odaConvertDwgToDxf(dwgPath, outDir || path.dirname(dwgPath));
  if (result.ok) return result;
  return { ok: false, msg: 'ODA File Converter not found. Please install it (free) from opendesign.com.', needsInstall: true };
});

// ─── U-Tube Routing: Folder Scan ────────────────────────────────────────
ipcMain.handle('scan-folder-utube', async (_, folderPath) => {
  const entries = await readdirSafe(folderPath);
  const dwgs = [];
  for (const name of entries) {
    // Match real DWGs and link files: *.dwg or [*.dwg]
    const isLink = /^\[.+\.dwg\]$/i.test(name);
    const isReal = /\.dwg$/i.test(name) && !isLink;
    if (!isReal && !isLink) continue;
    const fp = path.join(folderPath, name);
    try { if (!(await fsp.stat(fp)).isFile()) continue; } catch { continue; }
    // For links, resolve the real path; strip brackets for classification
    let realPath = fp;
    let baseName = name;
    if (isLink) {
      const target = await readLinkTarget(fp);
      realPath = target || fp;  // Use link file path as fallback for classification
      baseName = name.slice(1, -1); // strip [ and ]
    }
    dwgs.push({ name, baseName, filepath: realPath, nameLower: baseName.toLowerCase(), isLink });
  }

  // Extract first 4 digits from the base filename (brackets already stripped)
  function first4(baseName) {
    const m = baseName.match(/^(\d{4})/);
    return m ? parseInt(m[1], 10) : null;
  }
  // Extract revision number from rNN suffix
  function revNum(baseName) {
    const m = baseName.match(/r(\d+)\.dwg$/i);
    return m ? parseInt(m[1], 10) : 0;
  }
  // Pick highest-rev file from a list
  function pickBest(list) {
    if (!list.length) return null;
    list.sort((a, b) => revNum(b.baseName) - revNum(a.baseName));
    return { filepath: list[0].filepath, name: list[0].name };
  }

  const hasGenericShell = dwgs.some(d => d.nameLower === 'shell.dwg');
  const hasGenericBundle = dwgs.some(d => d.nameLower === 'bundle.dwg');

  // Classify files
  const settings = [], shells = [], bundles = [];
  for (const d of dwgs) {
    const n4 = first4(d.baseName);
    if (d.nameLower === 'setting.dwg') {
      if (hasGenericShell || hasGenericBundle) settings.push(d);
    } else if (d.nameLower === 'shell.dwg') {
      shells.push(d);
    } else if (d.nameLower === 'bundle.dwg') {
      bundles.push(d);
    } else if (n4 !== null) {
      if (n4 >= 5254 && n4 <= 5299) settings.push(d);
      else if (n4 >= 4600 && n4 <= 4699) shells.push(d);
      else if (n4 >= 4220 && n4 <= 4249) bundles.push(d);
    }
  }

  const setting = pickBest(settings);
  if (!setting) return { ok: false, msg: 'No U-Tube setting drawing found in this order.' };

  const shell = pickBest(shells);
  const bundle = pickBest(bundles);
  const parts = [setting && 'setting', shell && 'shell', bundle && 'bundle'].filter(Boolean);
  return { ok: true, setting, shell, bundle, msg: `Found ${parts.join(' + ')}` };
});

// ─── U-Tube Routing: DWG → DXF (temp, return text) ─────────────────────
let _routeProc = null;

ipcMain.handle('convert-dwg-dxf-text', async (_, dwgPath) => {
  // Try ODA first (much faster for routing)
  const tmpDir = path.join(os.tmpdir(), 'xv_route_' + Date.now());
  const oda = findOdaConverter();
  if (oda) {
    try {
      fs.mkdirSync(tmpDir, { recursive: true });
      const odaResult = await odaConvertDwgToDxf(dwgPath, tmpDir);
      if (odaResult.ok) {
        const text = await fsp.readFile(odaResult.dxfPath, 'utf-8');
        try { fs.unlinkSync(odaResult.dxfPath); fs.rmdirSync(tmpDir); } catch {}
        return { ok: true, text, name: path.basename(dwgPath) };
      }
    } catch {}
    try { fs.rmdirSync(tmpDir, { recursive: true }); } catch {}
  }

  // ODA not available — no fallback to accoreconsole for DXF
  return { ok: false, msg: 'ODA File Converter not found. Please install it (free) from opendesign.com.', needsInstall: true };
});

ipcMain.handle('cancel-convert', () => {
  if (_routeProc) { try { _routeProc.kill(); } catch {} _routeProc = null; }
});

// ─── Chat Room ──────────────────────────────────────────────────────────
const CHAT_DIR_FALLBACK = path.join(os.homedir(), 'XylemTest');
function getChatDir() { return IS_WIN ? SHARED_DATA_DIR : CHAT_DIR_FALLBACK; }
function getChatFile() { return path.join(getChatDir(), 'chat.json'); }
function getCommentsDir() { return path.join(getChatDir(), 'comments'); }

async function readJsonSafe(fp) {
  try {
    await fsp.access(fp);
    return JSON.parse(await fsp.readFile(fp, 'utf-8'));
  } catch (e) {}
  return [];
}

async function appendJsonMsg(fp, msg) {
  const msgs = await readJsonSafe(fp);
  msgs.push(msg);
  const trimmed = msgs.slice(-200);
  await fsp.mkdir(path.dirname(fp), { recursive: true });
  await fsp.writeFile(fp, JSON.stringify(trimmed, null, 1));
  return trimmed;
}

ipcMain.handle('chat-read', async () => {
  return await readJsonSafe(getChatFile());
});

ipcMain.handle('chat-send', async (_, text) => {
  const info = await getUserInfoCached();
  const msg = { user: info.username, name: info.fullName, text, ts: Date.now() };
  return appendJsonMsg(getChatFile(), msg);
});

ipcMain.handle('comments-read', async (_, orderLine) => {
  return await readJsonSafe(path.join(getCommentsDir(), orderLine + '.json'));
});

ipcMain.handle('comments-send', async (_, orderLine, text) => {
  const info = await getUserInfoCached();
  const msg = { user: info.username, name: info.fullName, text, ts: Date.now() };
  return appendJsonMsg(path.join(getCommentsDir(), orderLine + '.json'), msg);
});

let _userInfoCache = null;
async function getUserInfoCached() {
  if (_userInfoCache) return _userInfoCache;
  const username = os.userInfo().username;
  let fullName = username;
  if (IS_WIN) {
    // Try domain-aware lookup first (works on corporate networks)
    try {
      const result = execSync(`net user "${username}" /domain`, { encoding: 'utf-8', timeout: 5000 });
      const m = result.match(/Full Name\s+(.+)/i);
      if (m && m[1].trim()) fullName = m[1].trim();
    } catch (e) {}
    // Fallback: local net user
    if (fullName === username) {
      try {
        const result = execSync(`net user "${username}"`, { encoding: 'utf-8', timeout: 3000 });
        const m = result.match(/Full Name\s+(.+)/i);
        if (m && m[1].trim()) fullName = m[1].trim();
      } catch (e) {}
    }
    // Fallback: PowerShell WMI (local accounts)
    if (fullName === username) {
      try {
        const result = execSync(
          `powershell -NoProfile -Command "(Get-WmiObject Win32_UserAccount -Filter \\"Name='${username}'\\").FullName"`,
          { encoding: 'utf-8', timeout: 5000 }
        ).trim();
        if (result && result !== username) fullName = result;
      } catch (e) {}
    }
  }
  _userInfoCache = { username, fullName };
  return _userInfoCache;
}

// Find AutoCAD's accoreconsole for PDF plotting (TrueView cannot do -PLOT)
function findPdfAccore() {
  if (!IS_WIN) return null;
  const base = 'C:\\Program Files\\Autodesk';
  try {
    if (!fs.existsSync(base)) return null;
    const dirs = fs.readdirSync(base)
      .filter(d => d.toLowerCase().startsWith('autocad'))
      .sort().reverse();
    for (const d of dirs) {
      const candidate = path.join(base, d, 'accoreconsole.exe');
      if (fs.existsSync(candidate)) return candidate;
    }
  } catch (e) {}
  return null;
}

// ─── DWG → PDF (B&W) Conversion ────────────────────────────────────────
const ANSI_PAPERS = {
  A: { name: 'ANSI_full_bleed_A_(8.50_x_11.00_Inches)',  long: 11,  short: 8.5 },
  B: { name: 'ANSI_full_bleed_B_(11.00_x_17.00_Inches)', long: 17,  short: 11  },
  C: { name: 'ANSI_full_bleed_C_(17.00_x_22.00_Inches)', long: 22,  short: 17  },
  D: { name: 'ANSI_full_bleed_D_(22.00_x_34.00_Inches)', long: 34,  short: 22  },
  E: { name: 'ANSI_full_bleed_E_(34.00_x_44.00_Inches)', long: 44,  short: 34  },
};

// ANSI paper sizes in mm for matching AcDbPlotSettings dimensions
const ANSI_MM = [
  { letter: 'A', w: 215.9, h: 279.4 },  // 8.5 x 11
  { letter: 'B', w: 279.4, h: 431.8 },  // 11 x 17
  { letter: 'C', w: 431.8, h: 558.8 },  // 17 x 22
  { letter: 'D', w: 558.8, h: 863.6 },  // 22 x 34
  { letter: 'E', w: 863.6, h: 1117.6 }, // 34 x 44
];

// Detect ANSI paper size + orientation from DXF AcDbPlotSettings (most reliable method)
// Returns { letter: 'A'-'E', orient: 'Landscape'|'Portrait' } or null
async function detectAnsiFromDxf(dwgPath) {
  const oda = findOdaConverter();
  if (!oda) return null;
  const tmpDir = path.join(os.tmpdir(), 'xv_pdf_' + Date.now());
  try {
    fs.mkdirSync(tmpDir, { recursive: true });
    const r = await odaConvertDwgToDxf(dwgPath, tmpDir);
    if (!r.ok) return null;
    const dxf = await fsp.readFile(r.dxfPath, 'utf-8');
    try { fs.unlinkSync(r.dxfPath); fs.rmdirSync(tmpDir); } catch {}

    const lines = dxf.split('\n');

    // Find AcDbPlotSettings with non-zero paper dimensions (skip empty/unused layouts)
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].trim() !== 'AcDbPlotSettings') continue;
      let paperW = 0, paperH = 0, rot = 0;
      for (let j = i + 1; j < Math.min(i + 80, lines.length); j++) {
        const gc = lines[j].trim();
        const val = (lines[j + 1] || '').trim();
        if (gc === '44') paperW = parseFloat(val);
        if (gc === '45') paperH = parseFloat(val);
        if (gc === '73') rot = parseInt(val);
        if (gc === '0') break;
      }
      if (paperW < 10 || paperH < 10) continue; // skip empty layouts

      // Match to closest ANSI size (within 5mm tolerance)
      const pShort = Math.min(paperW, paperH), pLong = Math.max(paperW, paperH);
      let match = null;
      for (const a of ANSI_MM) {
        if (Math.abs(a.w - pShort) < 5 && Math.abs(a.h - pLong) < 5) { match = a; break; }
      }
      if (!match) continue;

      // Determine orientation from rotation (gc 73)
      // Paper is naturally portrait (W < H). Rotation 0/2 = keep, 1/3 = flip.
      const naturalPortrait = paperW < paperH;
      const rotFlips = (rot === 1 || rot === 3);
      const isLandscape = naturalPortrait ? rotFlips : !rotFlips;
      const orient = isLandscape ? 'Landscape' : 'Portrait';

      console.log(`PDF DXF: AcDbPlotSettings ${paperW.toFixed(0)}x${paperH.toFixed(0)}mm rot=${rot} → ANSI ${match.letter} ${orient}`);
      return { letter: match.letter, orient };
    }

    // Fallback: look for xSIZE block names (ASIZE, BSIZE, etc.)
    for (let i = 0; i < lines.length; i++) {
      if (lines[i].trim() === 'AcDbBlockBegin') {
        for (let j = i + 1; j < Math.min(i + 6, lines.length); j++) {
          if (lines[j].trim() === '2') {
            const name = (lines[j + 1] || '').trim().toUpperCase();
            const m = name.match(/^([ABCDE])SIZE/);
            if (m) { console.log(`PDF DXF: block ${name} → ANSI ${m[1]}`); return { letter: m[1], orient: null }; }
            break;
          }
        }
      }
    }

    console.log('PDF DXF: no paper size found');
    return null;
  } catch (e) {
    console.log('PDF DXF error:', e.message);
    try { fs.rmdirSync(tmpDir, { recursive: true }); } catch {}
    return null;
  }
}

// Determine orientation from LIMMAX dimensions
function ansiOrient(limW, limH) { return limW >= limH ? 'Landscape' : 'Portrait'; }

// Fallback: try to match LIMMAX to an ANSI size using scale-factor consistency
function pickAnsiFromLimmax(limW, limH, dimScale) {
  const dLong = Math.max(limW, limH), dShort = Math.min(limW, limH);
  const candidates = [];
  for (const [letter, p] of Object.entries(ANSI_PAPERS)) {
    const sL = dLong / p.long, sS = dShort / p.short;
    if (sL > 0.5 && Math.abs(sL - sS) / sL < 0.02) {
      candidates.push({ letter, scale: (sL + sS) / 2 });
    }
  }
  if (candidates.length === 1) return candidates[0].letter;
  if (candidates.length > 1 && dimScale > 0) {
    candidates.sort((a, b) => Math.abs(a.scale - dimScale) - Math.abs(b.scale - dimScale));
    return candidates[0].letter;
  }
  return null;
}

// Helper: run accoreconsole with a script, return stdout
// Track spawned accoreconsole processes so we can kill them on app exit
const _accoreProcs = new Set();

function runAccore(accore, dwgPath, scrContent, timeoutMs) {
  const scrPath = path.join(os.tmpdir(), 'xv_' + Date.now() + '_' + Math.random().toString(36).slice(2, 6) + '.scr');
  fs.writeFileSync(scrPath, scrContent);
  const tm = timeoutMs || 120000;
  return new Promise((resolve) => {
    const proc = spawn(accore, ['/i', dwgPath, '/s', scrPath], {
      stdio: 'pipe', windowsHide: true,
      env: { ...process.env, ACAD_NOHARDWARE: '1', ACAD_NOGRAPHICS: '1' },
    });
    _accoreProcs.add(proc);
    let stdout = '', stderr = '', done = false;
    const finish = () => {
      if (done) return; done = true;
      _accoreProcs.delete(proc);
      try { fs.unlinkSync(scrPath); } catch {}
      resolve({ stdout, stderr });
    };
    const decode = (d) => {
      const buf = Buffer.from(d);
      if (buf.length >= 2 && buf[1] === 0 && buf[buf.length - 1] === 0) return buf.toString('utf16le');
      return buf.toString('utf-8');
    };
    proc.stdout?.on('data', d => stdout += decode(d));
    proc.stderr?.on('data', d => stderr += decode(d));
    proc.on('close', finish);
    proc.on('error', (e) => { stderr = e.message; finish(); });
    // Hard kill after timeout — Windows ignores SIGTERM, so use taskkill /T for the process tree
    setTimeout(() => {
      if (done) return;
      try {
        if (IS_WIN) execAsync(`taskkill /PID ${proc.pid} /T /F`, { timeout: 5000 }).catch(() => {});
        else proc.kill('SIGKILL');
      } catch {}
      setTimeout(finish, 1000); // Give it a moment, then resolve
    }, tm);
  });
}

// Kill any lingering accoreconsole processes on app exit
app.on('will-quit', () => {
  for (const proc of _accoreProcs) {
    try {
      if (IS_WIN) { require('child_process').execSync(`taskkill /PID ${proc.pid} /T /F`, { timeout: 3000 }); }
      else proc.kill('SIGKILL');
    } catch {}
  }
});

ipcMain.handle('convert-dwg-pdf', async (_, dwgPath, outDir) => {
  const accore = findPdfAccore();
  if (!accore) return { ok: false, msg: 'PDF conversion requires AutoCAD.' };
  if (!fs.existsSync(dwgPath)) return { ok: false, msg: `Drawing not found: ${path.basename(dwgPath)}` };

  const baseName = path.basename(dwgPath).replace(/\.dwg$/i, '.pdf');
  const baseDir = outDir || path.dirname(dwgPath);
  // Never overwrite — use numbered suffix if file exists (avoids locked-file issues)
  let pdfPath = path.join(baseDir, baseName);
  if (fs.existsSync(pdfPath)) {
    const ext = path.extname(pdfPath);
    const stem = pdfPath.slice(0, -ext.length);
    for (let i = 2; i <= 99; i++) {
      const alt = `${stem}_(${i})${ext}`;
      if (!fs.existsSync(alt)) { pdfPath = alt; break; }
    }
  }
  const pdfScr = pdfPath.replace(/\\/g, '/');

  // Step 1: Detect paper via LIMMAX + DXF AcDbPlotSettings (in parallel)
  // LIMMAX is most reliable WHEN it matches an ANSI size cleanly.
  // AcDbPlotSettings is the fallback (can be stale/wrong if drafter didn't set it up).
  const [dxfResult, extResult] = await Promise.all([
    detectAnsiFromDxf(dwgPath),
    runAccore(accore, dwgPath, ['SETVAR','LIMMAX','','SETVAR','DIMSCALE','','_QUIT','Y'].join('\r\n')+'\r\n', 120000),
  ]);

  // Parse LIMMAX + DIMSCALE (used for paper size detection)
  let limX = 0, limY = 0, dimScale = 1;
  try {
    const raw = (extResult.stdout + extResult.stderr).replace(/[\x00-\x1F]/g, '');
    const lm = raw.match(/LIMMAX\s*<([\d.]+)\s*,\s*([\d.]+)>/i);
    const ds = raw.match(/DIMSCALE\s*<([\d.]+)>/i);
    if (lm) { limX = parseFloat(lm[1]); limY = parseFloat(lm[2]); }
    if (ds) dimScale = parseFloat(ds[1]);
  } catch {}

  // Determine paper + orientation
  let paper, orient, detectedLetter = null;

  // Priority 1: LIMMAX — if it cleanly matches an ANSI size, trust it
  if (limX > 5 && limY > 5) {
    detectedLetter = pickAnsiFromLimmax(limX, limY, dimScale);
    if (detectedLetter) {
      console.log(`PDF: LIMMAX ${limX}x${limY} ds=${dimScale} → ANSI ${detectedLetter}`);
    }
  }

  // Priority 2: DXF AcDbPlotSettings — fallback when LIMMAX doesn't match
  if (!detectedLetter && dxfResult) {
    detectedLetter = dxfResult.letter;
    console.log(`PDF: DXF fallback → ANSI ${detectedLetter} (plot settings)`);
  }

  if (detectedLetter && ANSI_PAPERS[detectedLetter]) {
    const p = ANSI_PAPERS[detectedLetter];
    paper = p.name;
    // Orientation: LIMMAX aspect ratio (most reliable), DXF rotation as fallback
    if (limX > 0 && limY > 0) orient = ansiOrient(limX, limY);
    else if (dxfResult && dxfResult.orient) orient = dxfResult.orient;
    else orient = 'Landscape';
    console.log(`PDF: ${paper} ${orient}`);
  } else {
    paper = ANSI_PAPERS.B.name;
    orient = (limX > 0 && limY > 0) ? ansiOrient(limX, limY) : 'Landscape';
    detectedLetter = 'B';
    console.log(`PDF: detection failed, fallback ${paper} ${orient}`);
  }

  // Notify renderer of detected size (for toast)
  if (mainWindow) mainWindow.webContents.send('pdf-paper-detected', detectedLetter, orient);

  // Step 2: Plot to PDF
  const plotLines = [
    '_ZOOM _E',
    '-PLOT',
    'Y',                                    // Detailed config
    'Model',                                // Layout
    'DWG To PDF.pc3',                       // Device
    paper,                                  // Paper size (auto-detected)
    'Inches',                               // Units
    orient,                                 // Orientation (auto-detected)
    'N',                                    // Upside down
  ];
  plotLines.push(
    'E',                                    // Plot area: Extents
    'Fit',                                  // Scale
    'Center',                               // Centered on page
    'Y',                                    // Plot with styles
    'monochrome.ctb',                       // Style table
    'Y',                                    // Lineweights
    'A',                                    // Shade plot (As displayed)
    pdfScr,                                 // Output filename
    'N',                                    // Save changes to page setup
    'Y',                                    // Proceed with plot
    '_QUIT',
    'Y',
  );
  const plotScript = plotLines.join('\r\n') + '\r\n';

  const result = await runAccore(accore, dwgPath, plotScript, 120000);

  // Verify PDF
  try {
    const stat = fs.statSync(pdfPath);
    if (stat.size > 0) return { ok: true, pdfPath, msg: `Created ${path.basename(pdfPath)}` };
  } catch (e) {}
  const raw = (result.stderr || result.stdout).replace(/[\x00-\x09\x0B\x0C\x0E-\x1F]/g, '').trim();
  return { ok: false, msg: `PDF conversion failed: ${raw.slice(-300) || 'No output from AutoCAD'}` };
});

// ─── Auto-update check ─────────────────────────────────────────────────
const CURRENT_VERSION = '1.0.11';

async function findInstallerExe() {
  if (!IS_WIN) return null;
  // Check parent folder first, then subfolder
  for (const dir of [path.dirname(SHARED_DATA_DIR), SHARED_DATA_DIR]) {
    const files = (await readdirSafe(dir)).filter(f => /\.exe$/i.test(f) && /setup|install/i.test(f));
    if (files.length) return path.join(dir, files[0]);
  }
  return null;
}

ipcMain.handle('check-update', async () => {
  try {
    const exe = await findInstallerExe();
    if (!exe) return null;
    const { stdout } = await execAsync(
      `powershell -NoProfile -Command "(Get-Item '${exe.replace(/'/g, "''")}').VersionInfo.ProductVersion"`,
      { encoding: 'utf-8', timeout: 5000 }
    );
    const latest = stdout.trim();
    // Only show update if installer version is NEWER than current (not when dev is ahead)
    if (latest && latest !== CURRENT_VERSION) {
      const cur = CURRENT_VERSION.split('.').map(Number);
      const ins = latest.split('.').map(Number);
      const isNewer = ins[0] > cur[0] || (ins[0] === cur[0] && ins[1] > cur[1]) || (ins[0] === cur[0] && ins[1] === cur[1] && ins[2] > cur[2]);
      if (isNewer) return latest;
    }
  } catch (e) {}
  return null;
});

ipcMain.handle('open-installer', async () => {
  try {
    if (!IS_WIN) return { ok: false, msg: 'Not on Windows' };
    const exe = await findInstallerExe();
    if (exe) {
      shell.openPath(exe);
      // Quit after a short delay so the installer doesn't find us running
      setTimeout(() => app.quit(), 1000);
      return { ok: true, path: exe };
    }
    shell.openPath(dir);
    return { ok: true, path: dir };
  } catch (e) { return { ok: false, msg: e.message }; }
});
