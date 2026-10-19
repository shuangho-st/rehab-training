/*  年度排班與行政業務許願問卷（工作入口 ?p=wish）
 *  ============================================================
 *  - 每人一條專屬連結：.../exec?p=wish&k=代碼，代碼決定身分，不必登入選名字
 *  - 回覆存在一份獨立的試算表「排班許願問卷（組長專用）」，由部署者（組長）擁有，
 *    第一次打開 ?p=wish 時自動建立，ID 記在指令碼屬性 WISH_SHEET_ID
 *  - 組長從試算表「名單」分頁 D 欄複製每人連結；組長自己的連結可看「彙整與試排」
 *  - 名單改這裡的 WISH_ROSTER，下次有人打開問卷時會自動補進新名字（舊代碼不變）
 *  - 連結外流：把那個人的代碼（C 欄）清空，下次打開問卷會自動換新代碼與新連結
 *  ============================================================ */

/* 組員名單：role = lead（技術組長）／evening（明年兒童晚班）／staff */
const WISH_ROSTER = [
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

const WISH = {
  HTML: 'Wish',
  TITLE: '語言治療組｜年度排班許願問卷',
  SHEET_NAME: '排班許願問卷（組長專用）',
  PROP: 'WISH_SHEET_ID',
  TAB_LIST: '名單',
  TAB_RESP: '回覆',
  // 讀不到 ScriptApp.getService().getUrl() 時用的正式網址
  FALLBACK_URL: 'https://script.google.com/a/macros/s.tmu.edu.tw/s/AKfycbwp_A-yvS0DYXf02a20ceiqjtJOQveEMi22PrvommX4VBVuemf6dpberFgMakZ42lT5ug/exec'
};

/* ---------------- 網頁入口（由 Code.gs 的 doGet 轉過來） ---------------- */
function wishDoGet_(e) {
  const k = String((e && e.parameter && e.parameter.k) || '').trim();
  let html;
  try {
    const ss = wishSs_();
    const me = wishWho_(k, ss);
    if (!me) {
      html = wishInvalidPage_(ss);
    } else {
      const boot = JSON.stringify({ k: k, name: me.name, role: me.role, roster: WISH_ROSTER })
        .replace(/</g, '\\u003c');
      html = protectScript_(readHtmlFile_(WISH.HTML).replace('<?!= boot ?>', function () { return boot; }));
    }
  } catch (err) {
    html = '<!DOCTYPE html><meta charset="utf-8"><div style="font-family:sans-serif;padding:40px;line-height:1.8">' +
           '<h2>問卷暫時打不開</h2><p>' + escapeHtml_(String(err)) + '</p><p>請聯絡組長。</p></div>';
  }
  return HtmlService.createHtmlOutput(html)
    .setTitle(WISH.TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function wishInvalidPage_(ss) {
  // 試算表只有建立者（組長）打得開，其他人看到網址也進不去
  return '<!DOCTYPE html><meta charset="utf-8"><base target="_top">' +
    '<div style="font-family:sans-serif;padding:40px;line-height:1.8;max-width:640px">' +
    '<h2>請使用專屬連結開啟</h2>' +
    '<p>這份問卷每人有一條自己的連結，請用組長寄給你的那一條。若連結已失效，請聯絡組長。</p>' +
    '<hr style="border:0;border-top:1px solid #ddd;margin:28px 0">' +
    '<p style="color:#667;font-size:14px">組長：每人連結在<a href="' + ss.getUrl() + '">「' +
    escapeHtml_(WISH.SHEET_NAME) + '」</a>的「名單」分頁 D 欄（只有你打得開）。</p></div>';
}

/* ---------------- 試算表：不存在就建立，並同步名單與代碼 ---------------- */
function wishSs_() {
  const props = PropertiesService.getScriptProperties();
  let ss = null;
  const id = props.getProperty(WISH.PROP);
  if (id) { try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; } }
  if (!ss) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      const id2 = props.getProperty(WISH.PROP);
      if (id2 && id2 !== id) { try { ss = SpreadsheetApp.openById(id2); } catch (e) { ss = null; } }
      if (!ss) {
        ss = SpreadsheetApp.create(WISH.SHEET_NAME);
        props.setProperty(WISH.PROP, ss.getId());
      }
    } finally { lock.releaseLock(); }
  }
  wishSync_(ss);
  return ss;
}

/* 名單分頁：每人一列（姓名、身分、代碼、專屬連結）。已有的代碼不覆蓋，空白的補新代碼 */
function wishSync_(ss) {
  let list = ss.getSheetByName(WISH.TAB_LIST);
  let resp = ss.getSheetByName(WISH.TAB_RESP);
  const fresh = !list;
  if (!list) {
    list = ss.getSheets()[0];
    list.setName(WISH.TAB_LIST);
  }
  if (!resp) {
    resp = ss.insertSheet(WISH.TAB_RESP);
    resp.getRange(1, 1, 1, 3).setValues([['姓名', '最後送出時間', '回覆內容（JSON，請勿手動修改）']]).setFontWeight('bold');
    resp.setFrozenRows(1);
  }

  const n = Math.max(list.getLastRow() - 1, 0);
  const old = n ? list.getRange(2, 1, n, 4).getDisplayValues() : [];
  const url = wishUrl_();
  let changed = fresh;
  const byName = {};
  old.forEach(function (r) { if (r[0]) byName[r[0]] = r; });
  const rows = WISH_ROSTER.map(function (m) {
    const o = byName[m.name];
    const token = (o && o[2]) ? o[2] : wishToken_();
    const link = url + '?p=wish&k=' + token;
    if (!o || o[1] !== m.role || o[2] !== token || o[3] !== link) changed = true;
    return [m.name, m.role, token, link];
  });
  if (old.length !== rows.length) changed = true;
  if (!changed) return;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    list.clear();
    list.getRange(1, 1, 1, 4).setValues([['姓名', '身分', '代碼', '專屬連結（每人一條，個別寄送）']]).setFontWeight('bold');
    // 代碼一律存成純文字，避免全數字或像 1e5 的代碼被試算表轉成數字
    list.getRange(2, 3, rows.length, 1).setNumberFormat('@');
    list.getRange(2, 1, rows.length, 4).setValues(rows);
    list.setFrozenRows(1);
    list.autoResizeColumns(1, 4);
  } finally { lock.releaseLock(); }
}

function wishUrl_() {
  let url = '';
  try { url = ScriptApp.getService().getUrl() || ''; } catch (e) {}
  return /\/exec$/.test(url) ? url : WISH.FALLBACK_URL;
}

function wishToken_() {
  // 開頭固定用英文字，代碼不會被當成數字
  return 'w' + Utilities.getUuid().replace(/-/g, '').slice(0, 15);
}

/* 依代碼找人 */
function wishWho_(k, ss) {
  if (!k) return null;
  const list = (ss || wishSs_()).getSheetByName(WISH.TAB_LIST);
  if (!list || list.getLastRow() < 2) return null;
  const row = list.getRange(2, 1, list.getLastRow() - 1, 3).getDisplayValues()
    .find(function (r) { return r[2] && r[2] === k; });
  return row ? { name: row[0], role: row[1] } : null;
}

/* ---------------- 前端呼叫 ---------------- */

/* 送出（同一人重送會覆蓋，以最後一次為準） */
function wishSubmit(k, data) {
  const ss = wishSs_();
  const me = wishWho_(k, ss);
  if (!me) throw new Error('連結無效，請聯絡組長');
  data = data || {};
  data.name = me.name;
  data.lead = me.role === 'lead';
  data.evening = me.role === 'evening';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = ss.getSheetByName(WISH.TAB_RESP);
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
  const ss = wishSs_();
  const me = wishWho_(k, ss);
  if (!me) return null;
  return wishReadAll_(ss).find(function (r) { return r.name === me.name; }) || null;
}

/* 組長讀取所有回覆 */
function wishGetAll(k) {
  const ss = wishSs_();
  const me = wishWho_(k, ss);
  if (!me || me.role !== 'lead') throw new Error('只有組長可以查看彙整');
  return wishReadAll_(ss);
}

/* 組長取得每人連結與試算表網址 */
function wishGetLinks(k) {
  const ss = wishSs_();
  const me = wishWho_(k, ss);
  if (!me || me.role !== 'lead') throw new Error('只有組長可以查看連結');
  const list = ss.getSheetByName(WISH.TAB_LIST);
  const rows = list.getRange(2, 1, list.getLastRow() - 1, 4).getDisplayValues();
  return {
    sheet: ss.getUrl(),
    links: rows.map(function (r) { return { name: r[0], role: r[1], url: r[3] }; })
  };
}

function wishReadAll_(ss) {
  const sh = ss.getSheetByName(WISH.TAB_RESP);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues()
    .map(function (r) { try { return JSON.parse(r[2]); } catch (e) { return null; } })
    .filter(Boolean);
}

/** 手動執行：建立／同步問卷試算表，執行紀錄會印出試算表網址與每人連結 */
function 排班問卷連結() {
  const ss = wishSs_();
  Logger.log('試算表：' + ss.getUrl());
  ss.getSheetByName(WISH.TAB_LIST).getRange(2, 1, WISH_ROSTER.length, 4).getDisplayValues()
    .forEach(function (r) { Logger.log(r[0] + '（' + r[1] + '）　' + r[3]); });
}
