/***********************************************************************
 * 兒科題庫前後測　網頁版後端（Google Apps Script）
 * ---------------------------------------------------------------------
 * 收網頁作答（peds-exam/index.html）送來的成績，寫進試算表；
 * 教師端（peds-exam/admin.html）用管理密碼讀回來做報表。
 *
 * 可以和「兒科題庫測驗系統_完整版.gs」放在同一個 Apps Script 專案：
 * 本檔只有 doGet / doPost 與 WEB_ 開頭的函式，不會和它衝突（它沒有 doGet/doPost）。
 *
 * 部署
 *   1. 試算表 →「擴充功能」→「Apps Script」→ 新增檔案 → 貼上本檔 → 存檔
 *   2. 改下面 WEB.ADMIN_KEY（教師端登入用的密碼）
 *   3. 函式下拉選單選 WEB_setup → 執行一次（建立工作表、走完授權）
 *   4. 右上「部署」→「新增部署作業」→ 類型「網頁應用程式」
 *        執行身分：我　／　誰可以存取：所有人
 *   5. 複製「網頁應用程式網址」（…/exec），貼到 peds-exam/config.js 的 API_URL
 *   之後改了程式要「管理部署作業」→ 編輯 → 版本選「新版本」，網址不變。
 ***********************************************************************/

var WEB = {
  ADMIN_KEY: 'change-me',          // ← 一定要改。教師端讀成績時要輸入
  SPREADSHEET_ID: '',              // 留空＝用本專案所屬的試算表；獨立專案才需要填
  SHEET: '網頁作答',               // 每份交卷一列
  SHEET_ITEMS: '網頁逐題',         // 每題一列（方便在試算表裡直接樞紐分析）
  MAX_BODY: 200000
};

var WEB_HEAD = ['收到時間', '提交ID', '梯次', '前後測', '卷別碼', '題庫', '測驗名稱',
                '姓名', 'Email/員編', '醫院', '職級', '分數', '題數', '答對率',
                '開始時間', '交卷時間', '作答秒數', '作答字串', '原始資料'];
var WEB_ITEM_HEAD = ['提交ID', '梯次', '前後測', '卷別碼', '姓名', 'Email/員編', '醫院',
                     '題號', '科別', '選擇', '正解', '對錯'];
var WEB_FMT = ['yyyy-mm-dd hh:mm:ss', '@', '@', '@', '@', '@', '@', '@', '@', '@', '@',
               '0', '0', '0.0', '@', '@', '0', '@', '@'];
var WEB_ITEM_FMT = ['@', '@', '@', '@', '@', '@', '@', '0', '@', '@', '@', '0'];
function WEB_iso_(v) { return v instanceof Date ? v.toISOString() : String(v); }

function WEB_ss_() {
  return WEB.SPREADSHEET_ID ? SpreadsheetApp.openById(WEB.SPREADSHEET_ID)
                            : SpreadsheetApp.getActiveSpreadsheet();
}
function WEB_sheet_(name, head) {
  var ss = WEB_ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head])
      .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function WEB_json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/** 第一次手動執行：建立兩張工作表並觸發授權 */
function WEB_setup() {
  WEB_sheet_(WEB.SHEET, WEB_HEAD);
  WEB_sheet_(WEB.SHEET_ITEMS, WEB_ITEM_HEAD);
  if (WEB.ADMIN_KEY === 'change-me') {
    try {
      SpreadsheetApp.getUi().alert('提醒', '請先把程式最上方的 WEB.ADMIN_KEY 改成你自己的密碼再部署。', SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) { Logger.log('請先修改 WEB.ADMIN_KEY'); }
  }
}

/* ---------- 交卷 ---------- */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var body = (e && e.postData && e.postData.contents) || '';
    if (!body || body.length > WEB.MAX_BODY) return WEB_json_({ ok: false, error: 'empty or too large' });
    var d = JSON.parse(body);
    if (d.type !== 'submit') return WEB_json_({ ok: false, error: 'unknown type' });

    var str = function (v, n) { return String(v == null ? '' : v).slice(0, n || 200); };
    var answers = (d.answers || []).slice(0, 300).map(function (a) {
      return { no: +a.no || 0, picked: str(a.picked, 2), ans: str(a.ans, 2), cat: str(a.cat, 40) };
    });
    var score = answers.filter(function (a) { return a.picked && a.picked === a.ans; }).length;   // 以伺服器重算為準
    var total = answers.length;

    lock.waitLock(20000);
    var sh = WEB_sheet_(WEB.SHEET, WEB_HEAD);
    var id = str(d.id, 40);
    // 去重：網路不穩時同一份可能被補傳兩次
    var last = sh.getLastRow();
    if (last > 1 && id) {
      var from = Math.max(2, last - 1999);
      var ids = sh.getRange(from, 2, last - from + 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) if (ids[i][0] === id) return WEB_json_({ ok: true, dup: true });
    }

    // 先把文字欄設成純文字格式，避免「0123」員編掉零、「2026-10」梯次被轉成日期
    var now = new Date();
    var row = sh.getLastRow() + 1;
    sh.getRange(row, 1, 1, WEB_HEAD.length).setNumberFormats([WEB_FMT]);
    sh.getRange(row, 1, 1, WEB_HEAD.length).setValues([[now, id, str(d.batch, 60), d.phase === 'post' ? '後測' : d.phase === 'pre' ? '前測' : str(d.phase, 10),
      str(d.paperCode, 12), str(d.bank, 60), str(d.title, 80),
      str(d.name, 60), str(d.pid, 120), str(d.hospital, 40), str(d.level, 20),
      score, total, total ? Math.round(score / total * 1000) / 10 : 0,
      str(d.startedAt, 40), str(d.submittedAt, 40), +d.durationSec || 0,
      answers.map(function (a) { return a.no + ':' + (a.picked || '-'); }).join(' '),
      JSON.stringify({ answers: answers, order: (d.order || []).slice(0, 300) })]]);

    var it = WEB_sheet_(WEB.SHEET_ITEMS, WEB_ITEM_HEAD);
    if (answers.length) {
      var block = it.getRange(it.getLastRow() + 1, 1, answers.length, WEB_ITEM_HEAD.length);
      block.setNumberFormats(answers.map(function () { return WEB_ITEM_FMT; }));
      block.setValues(answers.map(function (a) {
        return [id, str(d.batch, 60), d.phase === 'post' ? '後測' : d.phase === 'pre' ? '前測' : str(d.phase, 10),
                str(d.paperCode, 12), str(d.name, 60), str(d.pid, 120), str(d.hospital, 40),
                a.no, a.cat, a.picked || '', a.ans, a.picked && a.picked === a.ans ? 1 : 0];
      }));
    }
    return WEB_json_({ ok: true });
  } catch (err) {
    return WEB_json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ---------- 教師端讀取 ---------- */
function doGet(e) {
  var p = (e && e.parameter) || {};
  try {
    if (p.action === 'ping') return WEB_json_({ ok: true, time: new Date().toISOString() });
    if (String(p.key || '') !== WEB.ADMIN_KEY || WEB.ADMIN_KEY === 'change-me') {
      return WEB_json_({ ok: false, error: 'auth' });
    }
    var sh = WEB_sheet_(WEB.SHEET, WEB_HEAD);
    var last = sh.getLastRow();
    var vals = last > 1 ? sh.getRange(2, 1, last - 1, WEB_HEAD.length).getValues() : [];

    if (p.action === 'batches') {
      var map = {};
      vals.forEach(function (r) {
        var k = String(r[2]);
        var m = map[k] || (map[k] = { batch: k, pre: 0, post: 0, other: 0, last: '', titles: {}, papers: {} });
        if (r[3] === '前測') m.pre++; else if (r[3] === '後測') m.post++; else m.other++;
        var t = WEB_iso_(r[0]);
        if (t > m.last) m.last = t;
        if (r[6]) m.titles[r[6]] = 1;
        if (r[4]) m.papers[r[4]] = 1;
      });
      var list = Object.keys(map).map(function (k) {
        var m = map[k];
        return { batch: m.batch, pre: m.pre, post: m.post, other: m.other, last: m.last,
                 titles: Object.keys(m.titles), papers: Object.keys(m.papers) };
      }).sort(function (a, b) { return a.last < b.last ? 1 : -1; });
      return WEB_json_({ ok: true, batches: list });
    }

    if (p.action === 'results') {
      var want = String(p.batch == null ? '' : p.batch);
      var rows = vals.filter(function (r) { return String(r[2]) === want; }).map(function (r) {
        var raw = {}; try { raw = JSON.parse(r[18] || '{}'); } catch (x) {}
        return {
          receivedAt: WEB_iso_(r[0]),
          id: r[1], batch: String(r[2]), phase: r[3] === '前測' ? 'pre' : r[3] === '後測' ? 'post' : String(r[3]),
          paperCode: String(r[4]), bank: String(r[5]), title: String(r[6]),
          name: String(r[7]), pid: String(r[8]), hospital: String(r[9]), level: String(r[10]),
          score: +r[11], total: +r[12], startedAt: WEB_iso_(r[14]), submittedAt: WEB_iso_(r[15]),
          durationSec: +r[16], answers: raw.answers || [], order: raw.order || []
        };
      });
      return WEB_json_({ ok: true, rows: rows });
    }
    return WEB_json_({ ok: false, error: 'unknown action' });
  } catch (err) {
    return WEB_json_({ ok: false, error: String(err && err.message || err) });
  }
}
