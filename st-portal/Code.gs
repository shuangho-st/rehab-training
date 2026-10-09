/*  語言治療組 工作入口 — 伺服器端 v2
 *  ============================================================
 *  與 v1 的差別：網頁內容不再貼在 Apps Script 裡，
 *  改成從 Google 雲端硬碟讀取 Index.html。
 *  好處：上傳檔案不會像複製貼上那樣掉內容，之後要改版也只要換掉硬碟上的檔案。
 *  ------------------------------------------------------------
 *  安裝步驟
 *  1. 把 Index.html 拖到 Google 雲端硬碟（放哪個資料夾都可以）
 *  2. 打開「ST 雲端白板」那份試算表 →「擴充功能」→「Apps Script」
 *     ★ 一定要從白板進去，程式才讀得到白板的個案
 *  3. 把這份貼成 Code.gs（這份只有幾 KB，貼起來不易出錯）
 *  4. 在編輯器上方的函式下拉選單選「檢查設定」，按「執行」，第一次會要求授權
 *     執行紀錄會告訴你有沒有找到 Index.html、內容是否完整
 *  5. 顯示「一切正常」後，「部署」→「新增部署作業」→「網頁應用程式」
 *        執行身分：我
 *        誰可以存取：僅限臺北醫學大學
 *  ============================================================ */

let HTML_SOURCE = '';        // 記錄這次的網頁內容是從哪裡取得

const CONFIG = {
  PROJECT_HTML: 'Index',     // 入口頁；專案內的 HTML 檔名（不含 .html）
  // 其他頁面：網址加上 ?p=代號 即可開啟，值是專案內的 HTML 檔名
  // 分頁圖示（favicon）。需是可公開存取的圖片網址（png / ico / svg）。
  // 留空＝維持 Apps Script 預設圖示。三個頁面可各自不同，找不到就用 DEFAULT。
  FAVICON: {
    DEFAULT: '',
    index:   '',
    adult:   '',
    daily:   '',
    early:   ''
  },
  PAGES: {
    adult: 'Adult',          // 成人評估紀錄（分步驟表單）
    daily: 'Daily',          // 語言治療每日紀錄
    early: 'Early'           // 早療語言治療評估紀錄
  },
  HTML_NAME: 'Index.html',   // 備援：雲端硬碟上的檔名（需 Drive 授權）
  HTML_FILE_ID: '',          // 若有多個同名檔造成混淆，可直接指定檔案 ID
  SHEET_ID: '',              // 留空＝這支程式所附屬的試算表（白板）
  DEFAULT_TAB: ''            // 留空＝自動用第一個分頁
};

/* ---------------- 網頁進入點 ---------------- */
function doGet(e) {
  const p = (e && e.parameter && e.parameter.p) ? String(e.parameter.p) : '';
  const fileName = p ? (CONFIG.PAGES[p] || '') : CONFIG.PROJECT_HTML;
  const title = p ? (p === 'adult' ? '成人評估紀錄' : p === 'daily' ? '語言治療每日紀錄' : p === 'early' ? '早療評估紀錄' : '表單')
                  : '語言治療組 工作入口';
  let html;
  try {
    if (!fileName) throw new Error('未知的頁面代號「' + p + '」。目前可用：' +
      Object.keys(CONFIG.PAGES).join('、'));
    html = loadHtml_(fileName);
  } catch (e) {
    html = '<!DOCTYPE html><meta charset="utf-8">' +
           '<div style="font-family:sans-serif;padding:40px;line-height:1.8">' +
           '<h2>找不到網頁檔案</h2><p>' + escapeHtml_(String(e)) + '</p>' +
           '<p>請確認 Index.html 已上傳到你的 Google 雲端硬碟，' +
           '或在 Code.gs 的 CONFIG.HTML_FILE_ID 直接填入檔案 ID。</p></div>';
  }
  const out = HtmlService.createHtmlOutput(html)
    .setTitle(title)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  const icon = (CONFIG.FAVICON[p || 'index'] || CONFIG.FAVICON.DEFAULT || '').trim();
  if (icon) {
    try { out.setFaviconUrl(icon); } catch (e) { /* 網址無效就維持預設圖示 */ }
  }
  return out;
}

function loadHtml_(fileName) {
  return protectScript_(readHtmlFile_(fileName));
}

/**
 * Apps Script 送出 HTML 時會改寫 <script> 內容（實測 19142 字變 18781 字），
 * 導致主程式解析失敗。此處把主程式轉成 Base64，改由瀏覽器端解碼後以 DOM 方式插入，
 * 完全避開 HTML 解析，確保程式碼一字不差。
 */
function protectScript_(html) {
  const OPEN = '<' + 'script>';
  const CLOSE = '</' + 'script>';
  const a = html.lastIndexOf(OPEN);
  const b = html.lastIndexOf(CLOSE);
  if (a < 0 || b <= a) return html;
  const js = html.substring(a + OPEN.length, b);
  if (js.length < 2000) return html;   // 太短的不是主程式，不處理
  const b64 = Utilities.base64Encode(Utilities.newBlob(js, 'text/plain', 'x.js').getBytes());
  const loader = OPEN +
    '(function(){try{' +
      'var bin=atob("' + b64 + '");' +
      'var by=new Uint8Array(bin.length);' +
      'for(var i=0;i<bin.length;i++){by[i]=bin.charCodeAt(i);}' +
      'var src=new TextDecoder("utf-8").decode(by);' +
      'var el=document.createElement("script");' +
      'el.text=src;' +
      'document.body.appendChild(el);' +
    '}catch(e){' +
      'var b=document.getElementById("errbar");' +
      'if(b){b.style.display="block";b.textContent="主程式載入失敗："+e;}' +
    '}})();' + CLOSE;
  return html.substring(0, a) + loader + html.substring(b + CLOSE.length);
}

function readHtmlFile_(fileName) {
  const name = fileName || CONFIG.PROJECT_HTML;
  // 來源一：專案內的 HTML 檔。getRawContent() 取得未經處理的原始內容，且不需要額外授權。
  try {
    const raw = HtmlService.createTemplateFromFile(name).getRawContent();
    if (raw && raw.length > 1000) {
      HTML_SOURCE = '專案內檔案「' + name + '」';
      return raw;
    }
  } catch (e) { /* 專案內沒有就往下試雲端硬碟 */ }

  // 來源二：雲端硬碟（需要 Drive 授權；新增後必須重新授權並重新部署）
  HTML_SOURCE = '雲端硬碟「' + CONFIG.HTML_NAME + '」';
  let file;
  if (CONFIG.HTML_FILE_ID) {
    file = DriveApp.getFileById(CONFIG.HTML_FILE_ID);
  } else {
    const it = DriveApp.getFilesByName(CONFIG.HTML_NAME);
    if (!it.hasNext()) throw new Error('雲端硬碟中找不到檔案「' + CONFIG.HTML_NAME + '」');
    file = it.next();
    // 若有多個同名檔，取最後修改的那一個
    while (it.hasNext()) {
      const f = it.next();
      if (f.getLastUpdated() > file.getLastUpdated()) file = f;
    }
  }
  return file.getBlob().getDataAsString('UTF-8');
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

/* ---------------- 安裝前自我檢查 ---------------- */
function 檢查設定() {
  const log = [];
  // 1. 試算表
  try {
    const ss = ss_();
    log.push('✓ 試算表：' + ss.getName());
    log.push('  分頁：' + ss.getSheets().map(function (s) { return s.getName(); }).join('、'));
  } catch (e) {
    log.push('✗ 讀不到試算表：' + e);
    log.push('  → 請確認 Apps Script 是從「ST 雲端白板」按「擴充功能」進去建立的，');
    log.push('    或在 CONFIG.SHEET_ID 填入白板的試算表 ID。');
  }
  // 2. 網頁檔（入口頁與各表單）
  const pages = [['入口頁', CONFIG.PROJECT_HTML]];
  Object.keys(CONFIG.PAGES).forEach(function (k) {
    pages.push(['?p=' + k, CONFIG.PAGES[k]]);
  });
  pages.forEach(function (pair) {
    try {
      const raw = HtmlService.createTemplateFromFile(pair[1]).getRawContent();
      const prot = protectScript_(raw);
      const m = prot.match(/atob\("([A-Za-z0-9+\/=]+)"\)/);
      log.push('✓ ' + pair[0] + '（檔案 ' + pair[1] + '）' + raw.length + ' 字' +
               (m ? '，主程式已保護' : '，⚠ 未偵測到主程式'));
    } catch (e) {
      log.push('✗ ' + pair[0] + '：專案內找不到 HTML 檔「' + pair[1] + '」');
    }
  });
  try {
    const html = readHtmlFile_();
    const hasHead = html.indexOf('use strict') >= 0;
    const hasTail = html.indexOf('__PORTAL_BOOT="js"') >= 0;
    log.push('  含主程式開頭：' + hasHead);
    log.push('  含主程式結尾：' + hasTail);
    if (!hasHead || !hasTail) {
      log.push('✗ 檔案內容不完整，請重新上傳 Index.html 到雲端硬碟。');
    } else if (html.length < 30000) {
      log.push('✗ 長度偏短（正常約 36000 字），可能不是完整檔案。');
    } else {
      log.push('✓ 網頁檔完整');
      const prot = protectScript_(html);
      const m = prot.match(/atob\("([A-Za-z0-9+\/=]+)"\)/);
      log.push(m ? '✓ 主程式已轉為 Base64 保護，長度 ' + m[1].length + ' 字'
                 : '✗ Base64 保護未生效，請確認 Index.html 結尾有 window.__PORTAL_BOOT');
    }
  } catch (e) {
    log.push('✗ ' + e);
  }
  // 3. 今天的白板
  try {
    const r = getBoard(CONFIG.DEFAULT_TAB || null, 0);
    if (!r.ok) log.push('✗ 讀取白板失敗：' + r.error);
    else if (!r.found) log.push('△ 白板上找不到今天（' + r.date + '）的欄位');
    else log.push('✓ 今天（' + r.date + '）讀到 ' + r.total + ' 個時段');
  } catch (e) {
    log.push('✗ 讀取白板例外：' + e);
  }
  const out = log.join('\n');
  Logger.log(out);
  return out;
}

/* ---------------- 試算表存取 ---------------- */
function ss_() {
  return CONFIG.SHEET_ID
    ? SpreadsheetApp.openById(CONFIG.SHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function listTabs() {
  try {
    return { ok: true, tabs: ss_().getSheets().map(function (s) { return s.getName(); }) };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

const TIME_RE = /^\s*\d{1,2}:\d{2}\s*[-－–~至]\s*\d{1,2}:\d{2}/;

function sameDay_(a, b) {
  return a.getFullYear() === b.getFullYear() &&
         a.getMonth() === b.getMonth() &&
         a.getDate() === b.getDate();
}

function getBoard(tab, offset) {
  try {
    const ss = ss_();
    const name = tab || CONFIG.DEFAULT_TAB || ss.getSheets()[0].getName();
    const sh = ss.getSheetByName(name);
    if (!sh) return { ok: false, error: '找不到分頁「' + name + '」' };

    const rng = sh.getDataRange();
    const vals = rng.getValues();
    const disp = rng.getDisplayValues();
    const nR = vals.length, nC = vals[0] ? vals[0].length : 0;

    const target = new Date();
    target.setHours(0, 0, 0, 0);
    if (offset) target.setDate(target.getDate() + (offset | 0));

    // 找出今天在哪一欄（白板每週往右移，不能寫死）
    let col = -1;
    for (let r = 0; r < Math.min(nR, 6) && col < 0; r++) {
      for (let c = 0; c < nC; c++) {
        const v = vals[r][c];
        if (v instanceof Date && sameDay_(v, target)) { col = c; break; }
      }
    }
    if (col < 0) {
      return { ok: true, found: false, tab: name,
               date: Utilities.formatDate(target, Session.getScriptTimeZone(), 'yyyy/MM/dd'),
               msg: '白板上找不到這個日期的欄位' };
    }

    // 往左找時段欄
    let timeCol = -1, best = 0;
    for (let c = col; c >= 0 && c >= col - 12; c--) {
      let n = 0;
      for (let r = 0; r < nR; r++) if (TIME_RE.test(String(disp[r][c] || ''))) n++;
      if (n > best) { best = n; timeCol = c; }
    }
    if (timeCol < 0) return { ok: false, error: '找不到時段欄（形如 8:30-9:00）' };

    const slots = [];
    for (let r = 0; r < nR; r++) {
      const t = String(disp[r][timeCol] || '').trim();
      if (!TIME_RE.test(t)) continue;
      const type = String(disp[r][col] || '').trim();
      const who  = r + 1 < nR ? String(disp[r + 1][col] || '').trim() : '';
      let done = false;
      if (col + 1 < nC && r + 1 < nR) done = (vals[r + 1][col + 1] === true);
      if (!type && !who) continue;
      slots.push({ time: t.replace(/\s/g, ''), type: type, who: who, done: done });
    }

    const stats = [];
    ['本週總人次', '本月總人次', '距離目標人次尚餘',
     '本週總點數', '本月總點數', '距離目標點數尚餘'].forEach(function (label) {
      const v = findLabelValue_(vals, disp, label);
      if (v !== null) stats.push({ n: label, v: v });
    });

    return {
      ok: true, found: true, tab: name,
      date: Utilities.formatDate(target, Session.getScriptTimeZone(), 'yyyy/MM/dd'),
      weekday: '日一二三四五六'.charAt(target.getDay()),
      slots: slots, stats: stats,
      total: slots.length,
      doneCount: slots.filter(function (s) { return s.done; }).length
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function findLabelValue_(vals, disp, label) {
  for (let r = 0; r < vals.length; r++) {
    for (let c = 0; c < vals[r].length; c++) {
      if (String(disp[r][c] || '').trim() === label) {
        for (let k = 1; k <= 3 && r + k < vals.length; k++) {
          const v = vals[r + k][c];
          if (typeof v === 'number') return v;
          const s = String(disp[r + k][c] || '').trim();
          if (s && !isNaN(Number(s.replace(/,/g, '')))) return Number(s.replace(/,/g, ''));
        }
      }
    }
  }
  return null;
}

/* ---------------- 取得本網頁應用程式的正式網址 ----------------
   入口頁跑在 iframe 內，相對網址會接到內層的 googleusercontent 位址而失效，
   因此站內連結一律用這裡回傳的 /exec 網址組成。 */
function getAppUrl() {
  try {
    return { ok: true, url: ScriptApp.getService().getUrl() };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

/* ---------------- 設定：存在使用者帳號下 ---------------- */
const PKEY = 'st_portal_v1';

function getSettings() {
  try {
    const raw = PropertiesService.getUserProperties().getProperty(PKEY);
    return { ok: true, data: raw ? JSON.parse(raw) : null };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}

function saveSettings(json) {
  try {
    PropertiesService.getUserProperties().setProperty(PKEY, json);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}


/* =========================================================================
   成人評估紀錄：草稿（存在填寫者自己的 Google 帳號下）
   - 一份草稿 = 索引一筆 + 內容切成數段（UserProperties 每筆上限 9KB）
   - 內容先轉 Base64，避免中文字被切半或超過位元上限
   - 超過 DRAFT_TTL_DAYS 天或超過 DRAFT_MAX 份會自動清除
   ========================================================================= */
const DIDX = 'st_adult_idx_v1';     // 索引
const DPRE = 'st_adult_dft_v1_';    // 內容：DPRE + id + '_' + 段號
const DRAFT_MAX = 20;
const DRAFT_TTL_DAYS = 14;
const DCHUNK = 6000;                // Base64 為純 ASCII，6000 字元約 6KB

function dProps_() { return PropertiesService.getUserProperties(); }

function dIndex_() {
  try { return JSON.parse(dProps_().getProperty(DIDX) || '[]') || []; }
  catch (e) { return []; }
}
function dSetIndex_(arr) { dProps_().setProperty(DIDX, JSON.stringify(arr)); }

function dDropChunks_(p, m) {
  const n = m.chunks || 1;
  for (var i = 0; i < n; i++) { try { p.deleteProperty(DPRE + m.id + '_' + i); } catch (e) {} }
}

/** 清掉過期與超量的草稿，回傳清理後的索引 */
function dPurge_() {
  const p = dProps_();
  var idx = dIndex_();
  const now = Date.now(), ttl = DRAFT_TTL_DAYS * 86400000;
  var keep = [], drop = [];
  idx.forEach(function (m) { ((now - (m.ts || 0)) > ttl ? drop : keep).push(m); });
  keep.sort(function (a, b) { return (b.ts || 0) - (a.ts || 0); });
  while (keep.length > DRAFT_MAX) drop.push(keep.pop());
  if (drop.length) {
    drop.forEach(function (m) { dDropChunks_(p, m); });
    dSetIndex_(keep);
  }
  return keep;
}

function listDrafts() {
  try {
    const idx = dPurge_();
    return { ok: true, data: idx.map(function (m) {
      return { id: m.id, label: m.label, chart: m.chart, pct: m.pct, ts: m.ts };
    }) };
  } catch (e) { return { ok: false, msg: String(e) }; }
}

function saveDraft(payload) {
  try {
    payload = payload || {};
    const p = dProps_();
    var id = String(payload.id || '').replace(/[^0-9a-zA-Z]/g, '').substring(0, 32);
    if (!id) id = 'd' + Date.now();
    const json = String(payload.json || '{}');

    // Base64（UTF-8）→ 切段
    const b64 = Utilities.base64Encode(Utilities.newBlob(json, 'text/plain', 'd.json').getBytes());
    var chunks = [];
    for (var s = 0; s < b64.length; s += DCHUNK) chunks.push(b64.substring(s, s + DCHUNK));
    if (!chunks.length) chunks = [''];

    // 先清掉這個 id 的舊段落，避免殘留
    var idx = dIndex_(), old = null;
    idx = idx.filter(function (m) { if (m.id === id) { old = m; return false; } return true; });
    if (old) dDropChunks_(p, old);

    var obj = {};
    chunks.forEach(function (c, i) { obj[DPRE + id + '_' + i] = c; });
    p.setProperties(obj, false);

    idx.unshift({
      id: id,
      label: String(payload.label || '未命名個案').substring(0, 40),
      chart: String(payload.chart || '').substring(0, 20),
      pct: payload.pct,
      ts: Date.now(),
      chunks: chunks.length
    });
    dSetIndex_(idx);
    dPurge_();
    return { ok: true, id: id };
  } catch (e) { return { ok: false, msg: String(e) }; }
}

function getDraft(id) {
  try {
    const p = dProps_();
    var m = null;
    dIndex_().forEach(function (x) { if (x.id === id) m = x; });
    if (!m) return { ok: false, msg: '找不到這份草稿（可能已過期被清除）' };
    var b64 = '';
    const n = m.chunks || 1;
    for (var i = 0; i < n; i++) b64 += (p.getProperty(DPRE + id + '_' + i) || '');
    if (!b64) return { ok: false, msg: '這份草稿的內容已遺失' };
    const json = Utilities.newBlob(Utilities.base64Decode(b64)).getDataAsString('UTF-8');
    return { ok: true, json: json, label: m.label, ts: m.ts };
  } catch (e) { return { ok: false, msg: String(e) }; }
}

function deleteDraft(id) {
  try {
    const p = dProps_();
    var idx = dIndex_(), hit = null;
    idx = idx.filter(function (m) { if (m.id === id) { hit = m; return false; } return true; });
    if (hit) dDropChunks_(p, hit);
    dSetIndex_(idx);
    return { ok: true };
  } catch (e) { return { ok: false, msg: String(e) }; }
}

/** 手動執行：看目前帳號下有哪些草稿、佔多少空間 */
function 檢查草稿() {
  const idx = dIndex_();
  Logger.log('草稿數：' + idx.length);
  idx.forEach(function (m) {
    Logger.log('  ' + m.label + '　病歷號=' + (m.chart || '-') +
               '　完成=' + (m.pct == null ? '-' : m.pct + '%') +
               '　' + new Date(m.ts).toLocaleString('zh-TW') +
               '　段數=' + m.chunks);
  });
  Logger.log('保留天數：' + DRAFT_TTL_DAYS + '　上限份數：' + DRAFT_MAX);
}
