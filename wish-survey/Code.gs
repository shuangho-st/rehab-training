/* ══════════════════════════════════════════════════════════════════
   語言治療組｜年度排班許願問卷 — 獨立網頁應用程式（附屬於「排班許願問卷」試算表）
   - 每人一條專屬連結：.../exec?k=代碼，代碼決定身分
   - 第一次有人打開網頁時，自動在試算表建立「名單」「回覆」兩個分頁與每人連結
   - 名單改這裡的 ROSTER，下次打開網頁會自動補上新名字（舊代碼不變）
   - 連結外流：把那個人在「名單」的代碼（C 欄）清空，下次打開網頁會換新代碼與新連結
   - 由 GitHub（wish-survey 資料夾）自動部署，請不要直接在線上編輯器修改
   ══════════════════════════════════════════════════════════════════ */

/* 組員名單：role = lead（技術組長）／evening（明年兒童晚班）／staff */
const ROSTER = [
  { name: '慈芯', role: 'lead' },
  { name: '彥妤', role: 'staff' },
  { name: '園菁', role: 'staff' },
  { name: '品樺', role: 'staff' },
  { name: '郁方', role: 'evening' },
  { name: '宛宜', role: 'staff' },
  { name: '婉瑜', role: 'staff' },
  { name: '林寧', role: 'staff' },
  { name: '書寧', role: 'staff' }
];

const TITLE = '語言治療組｜年度排班許願問卷';
const TAB_LIST = '名單';
const TAB_RESP = '回覆';

/* 試算表開啟時加選單 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('排班問卷')
    .addItem('更新名單與每人連結', 'refreshLinks')
    .addToUi();
}

function refreshLinks() {
  sync_(SpreadsheetApp.getActive());
  SpreadsheetApp.getUi().alert('已更新。「名單」分頁 D 欄是每人的專屬連結，請個別寄送。');
}

/* ---------------- 網頁入口 ---------------- */
function doGet(e) {
  const k = String((e && e.parameter && e.parameter.k) || '').trim();
  let html;
  try {
    const ss = SpreadsheetApp.getActive();
    sync_(ss);
    const me = who_(k, ss);
    if (!me) {
      html = invalidPage_(ss);
    } else {
      const boot = JSON.stringify({ k: k, name: me.name, role: me.role, roster: ROSTER })
        .replace(/</g, '\\u003c');
      const raw = HtmlService.createTemplateFromFile('Index').getRawContent()
        .replace('<?!= boot ?>', function () { return boot; });
      html = protectScript_(raw);
    }
  } catch (err) {
    html = '<!DOCTYPE html><meta charset="utf-8"><div style="font-family:sans-serif;padding:40px;line-height:1.8">' +
           '<h2>問卷暫時打不開</h2><p>' + escapeHtml_(String(err)) + '</p><p>請聯絡組長。</p></div>';
  }
  return HtmlService.createHtmlOutput(html)
    .setTitle(TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function invalidPage_(ss) {
  // 試算表只有組長打得開，其他人看到網址也進不去
  return '<!DOCTYPE html><meta charset="utf-8"><base target="_top">' +
    '<div style="font-family:sans-serif;padding:40px;line-height:1.8;max-width:640px">' +
    '<h2>請使用專屬連結開啟</h2>' +
    '<p>這份問卷每人有一條自己的連結，請用組長寄給你的那一條。若連結已失效，請聯絡組長。</p>' +
    '<hr style="border:0;border-top:1px solid #ddd;margin:28px 0">' +
    '<p style="color:#667;font-size:14px">組長：每人連結在<a href="' + ss.getUrl() +
    '">「' + escapeHtml_(ss.getName()) + '」試算表</a>的「名單」分頁 D 欄（只有你打得開）。</p></div>';
}

/**
 * Apps Script 送出網頁時可能改寫較長的 <script>（工作入口實際遇過），導致整頁失效。
 * 這裡把主程式轉成 Base64，由瀏覽器解碼後再執行，確保程式一字不差。
 */
function protectScript_(html) {
  const OPEN = '<' + 'script>';
  const CLOSE = '</' + 'script>';
  const a = html.lastIndexOf(OPEN);
  const b = html.lastIndexOf(CLOSE);
  if (a < 0 || b <= a) return html;
  const js = html.substring(a + OPEN.length, b);
  const b64 = Utilities.base64Encode(Utilities.newBlob(js, 'text/plain', 'x.js').getBytes());
  const loader = OPEN +
    '(function(){try{' +
      'var bin=atob("' + b64 + '");' +
      'var by=new Uint8Array(bin.length);' +
      'for(var i=0;i<bin.length;i++){by[i]=bin.charCodeAt(i);}' +
      'var el=document.createElement("script");' +
      'el.text=new TextDecoder("utf-8").decode(by);' +
      'document.body.appendChild(el);' +
    '}catch(e){' +
      'var x=document.getElementById("errbar");' +
      'if(x){x.style.display="block";x.textContent="主程式載入失敗："+e;}' +
    '}})();' + CLOSE;
  return html.substring(0, a) + loader + html.substring(b + CLOSE.length);
}

function escapeHtml_(s) {
  return String(s).replace(/[&<>"]/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c];
  });
}

/* ---------------- 名單與代碼 ---------------- */

/* 「名單」：每人一列（姓名、身分、代碼、專屬連結）。已有的代碼不覆蓋，空白的補新代碼 */
function sync_(ss) {
  let list = ss.getSheetByName(TAB_LIST);
  const fresh = !list;
  if (!list) {
    list = ss.getSheets()[0];
    list.setName(TAB_LIST);
  }
  if (!ss.getSheetByName(TAB_RESP)) {
    const resp = ss.insertSheet(TAB_RESP);
    resp.getRange(1, 1, 1, 3).setValues([['姓名', '最後送出時間', '回覆內容（JSON，請勿手動修改）']]).setFontWeight('bold');
    resp.setFrozenRows(1);
  }

  const n = Math.max(list.getLastRow() - 1, 0);
  const old = n ? list.getRange(2, 1, n, 4).getDisplayValues() : [];
  const url = appUrl_();
  let changed = fresh || old.length !== ROSTER.length;
  const byName = {};
  old.forEach(function (r) { if (r[0]) byName[r[0]] = r; });
  const rows = ROSTER.map(function (m) {
    const o = byName[m.name];
    const token = (o && o[2]) ? o[2] : newToken_();
    const link = url ? url + '?k=' + token : '（部署後自動產生）';
    if (!o || o[1] !== m.role || o[2] !== token || o[3] !== link) changed = true;
    return [m.name, m.role, token, link];
  });
  if (!changed) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    list.clear();
    list.getRange(1, 1, 1, 4).setValues([['姓名', '身分', '代碼', '專屬連結（每人一條，個別寄送）']]).setFontWeight('bold');
    // 代碼一律存成純文字，避免被試算表轉成數字
    list.getRange(2, 3, rows.length, 1).setNumberFormat('@');
    list.getRange(2, 1, rows.length, 4).setValues(rows);
    list.setFrozenRows(1);
    list.autoResizeColumns(1, 4);
  } finally { lock.releaseLock(); }
}

function appUrl_() {
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  return /\/exec$/.test(url) ? url : '';
}

function newToken_() {
  // 開頭固定用英文字，代碼不會被當成數字
  return 'w' + Utilities.getUuid().replace(/-/g, '').slice(0, 15);
}

/* 依代碼找人 */
function who_(k, ss) {
  if (!k) return null;
  const list = (ss || SpreadsheetApp.getActive()).getSheetByName(TAB_LIST);
  if (!list || list.getLastRow() < 2) return null;
  const row = list.getRange(2, 1, list.getLastRow() - 1, 3).getDisplayValues()
    .find(function (r) { return r[2] && r[2] === k; });
  return row ? { name: row[0], role: row[1] } : null;
}

/* ---------------- 前端呼叫 ---------------- */

/* 送出（同一人重送會覆蓋，以最後一次為準） */
function wishSubmit(k, data) {
  const ss = SpreadsheetApp.getActive();
  const me = who_(k, ss);
  if (!me) throw new Error('連結無效，請聯絡組長');
  data = data || {};
  data.name = me.name;
  data.lead = me.role === 'lead';
  data.evening = me.role === 'evening';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(TAB_RESP);
    const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy-MM-dd HH:mm');
    data.submitted = now;
    const json = JSON.stringify(data);
    const n = sh.getLastRow() - 1;
    const names = n > 0 ? sh.getRange(2, 1, n, 1).getValues().map(function (r) { return r[0]; }) : [];
    const i = names.indexOf(me.name);
    if (i >= 0) sh.getRange(i + 2, 1, 1, 3).setValues([[me.name, now, json]]);
    else sh.appendRow([me.name, now, json]);
    return now;
  } finally {
    lock.releaseLock();
  }
}

/* 讀回自己上次送出的內容（回來修改用） */
function wishGetMine(k) {
  const ss = SpreadsheetApp.getActive();
  const me = who_(k, ss);
  if (!me) return null;
  return readAll_(ss).find(function (r) { return r.name === me.name; }) || null;
}

/* 組長讀取所有回覆 */
function wishGetAll(k) {
  const ss = SpreadsheetApp.getActive();
  const me = who_(k, ss);
  if (!me || me.role !== 'lead') throw new Error('只有組長可以查看彙整');
  return readAll_(ss);
}

/* 組長取得每人連結與試算表網址 */
function wishGetLinks(k) {
  const ss = SpreadsheetApp.getActive();
  const me = who_(k, ss);
  if (!me || me.role !== 'lead') throw new Error('只有組長可以查看連結');
  const list = ss.getSheetByName(TAB_LIST);
  const rows = list.getRange(2, 1, list.getLastRow() - 1, 4).getDisplayValues();
  return {
    sheet: ss.getUrl(),
    links: rows.map(function (r) { return { name: r[0], role: r[1], url: r[3] }; })
  };
}

function readAll_(ss) {
  const sh = ss.getSheetByName(TAB_RESP);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .map(function (r) { try { return JSON.parse(r[2]); } catch (e) { return null; } })
    .filter(Boolean);
}
