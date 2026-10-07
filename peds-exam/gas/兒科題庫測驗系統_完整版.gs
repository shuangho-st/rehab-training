/***********************************************************************
 * 兒科題庫測驗系統  —  完整單一檔案版
 * =====================================================================
 * 安裝方式（最省事，也不會有 onOpen 衝突）：
 *
 *   1. 擴充功能 → Apps Script
 *   2. 把專案裡「現有的所有 .gs 檔案」刪掉
 *      （若不放心，先按檔案右邊的三個點 →「重新命名」加個 _backup，
 *        或整份複製到記事本留底）
 *   3. 新增一個檔案，把這一整份貼進去 → 存檔
 *   4. 回試算表，重新整理頁面
 *
 * 這份包含全部功能，不需要再裝其他檔案：
 *   · 重新抽題（可設定題數／年分／科別配比）
 *   · 產生表單（已修正正解比對與題幹截斷）
 *   · 題庫健檢
 *   · 完整測驗報告（個人成績／個人錯題／逐題分析／醫院科別）
 *   · 封存與還原考卷
 *
 * 設定都集中在下面 DRAW / CFG / FORMCFG 三個區塊。
 * 最常要改的是 DRAW.SHEET_BANK（題庫分頁名稱）和 FORMCFG.FOLDER_ID。
 ***********************************************************************/

var DRAW = {
  SHEET_BANK     : '題庫',       // ← 題庫主表名稱；改了名字就改這裡
  SHEET_SELECTED : '選出題目',
  DEFAULT_N      : 30
};

/* 選單（統一在這裡定義，報表模組那份 onOpen 請刪掉或改名） */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('題庫系統')
    .addItem('重新抽題（可設定）', 'drawQuestions')
    .addItem('一次產生前測＋後測（同卷）', 'generatePrePostForms')
    .addSeparator()
    .addItem('① 題庫健檢', 'checkQuestionBank')
    .addItem('② 產生報表', 'buildReport')
    .addItem('③ 封存成績報告', 'archiveReports')
    .addItem('④ 產生詳解講義', 'buildExplanationHandout')
    .addSeparator()
    .addItem('重新整理歷次表單（更新回應數）', 'refreshFormHistory')
    .addSubMenu(ui.createMenu('進階')
      .addItem('⓪ 環境診斷', 'checkSetup')
      .addItem('⓪ 雲端硬碟診斷', 'checkDriveAccess')
      .addSeparator()
      .addItem('產生單份表單', 'generateForm')
      .addItem('產生後測表單（同卷、打亂順序）', 'generatePostTestForm')
      .addItem('登記目前表單到歷次表單', 'archiveFormLink')
      .addSeparator()
      .addItem('匯出待補詳解的題目', 'exportForExplanationDraft')
      .addItem('補寫題目對照表', 'writeItemMap')
      .addSeparator()
      .addItem('封存本次考卷', 'archiveCurrentPaper')
      .addItem('還原封存的考卷', 'restoreArchivedPaper')
      .addItem('產生報告（舊版，需選出題目相符）', 'buildFullReport'))
    .addToUi();
}

var VERSION = 'v18 (2026-10-07) 支援五選項（A–E）';

/* ═══════════════════════════════════════════════════════════════
   標題列自動偵測
   ---------------------------------------------------------------
   把範圍轉成 Google 試算表「表格」之後，標題列不一定落在第 1 列
   （上方可能有標題文字或空白列）。這裡往下找前 10 列，
   找出 A 欄含「題目」的那一列當標題，其餘往後才是資料。
   ═══════════════════════════════════════════════════════════════ */
function findHeaderRow_(values) {
  var limit = Math.min(10, values.length);
  for (var i = 0; i < limit; i++) {
    var a = String(values[i][0] || '').trim();
    if (a === '題目' || a.indexOf('題目') === 0) return i;
  }
  return 0;
}

/* ═══════════════════════════════════════════════════════════════
   環境診斷 —— 出問題時先跑這個
   不會修改任何資料，只讀取並產生一張「環境診斷」報告。
   ═══════════════════════════════════════════════════════════════ */
function checkSetup() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = [];
  var push = function (a, b, c) { out.push([a, b === undefined ? '' : b, c === undefined ? '' : c]); };

  push('程式版本', VERSION, '看得到這一行就代表新程式已生效');
  push('診斷時間', Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm'), '');
  push('', '', '');

  push('── 工作表清單 ──', '', '');
  ss.getSheets().forEach(function (sh) {
    push(sh.getName(), sh.getLastRow() + ' 列 × ' + sh.getLastColumn() + ' 欄',
         sh.getName() === DRAW.SHEET_BANK ? '← 設定中的題庫來源' :
         sh.getName() === CFG.SHEET_SELECTED ? '← 抽出的考卷會寫在這裡' : '');
  });
  push('', '', '');

  push('── 題庫（' + DRAW.SHEET_BANK + '）──', '', '');
  var bank = ss.getSheetByName(DRAW.SHEET_BANK);
  if (!bank) {
    push('狀態', '✗ 找不到', '請確認分頁名稱，或改 DRAW.SHEET_BANK 設定');
  } else {
    var raw = bank.getDataRange().getValues();
    var h = findHeaderRow_(raw);
    push('標題列位置', '第 ' + (h + 1) + ' 列', h === 0 ? '正常' : '不在第 1 列，已自動偵測');
    push('標題內容', raw[h].slice(0, 5).join(' | '), '應為 題目 | 年分 | 正確答案 | 科別 | 詳解');

    var rows = [];
    for (var i = h + 1; i < raw.length; i++) if (raw[i][0]) rows.push(raw[i]);
    push('讀到的題數', rows.length, '');

    var ansMap = {}, ansBad = 0, hasExp = 0;
    rows.forEach(function (r) {
      var v = r[2];
      var shown = (v === '' || v === null) ? '(空白)' : String(v).replace(/ /g, '␣');
      ansMap[shown] = (ansMap[shown] || 0) + 1;
      if (!normalizeAnswer(v)) ansBad++;
      if (String(r[4] || '').trim()) hasExp++;
    });
    push('', '', '');
    push('正解欄的實際內容', '', '␣ 代表空白字元');
    Object.keys(ansMap).sort().forEach(function (k) {
      push('　「' + k + '」', ansMap[k] + ' 題',
           normalizeAnswer(k.replace(/␣/g, ' ')) ? '✓ 可讀' : '✗ 讀不出 A/B/C/D/E，這些題全班會算答錯');
    });
    push('正解讀不出來的題數', ansBad, ansBad ? '⚠ 一定要修' : '✓ 全部正常');
    push('已撰寫詳解的題數', hasExp + ' / ' + rows.length, 'E 欄；產生詳解講義時會自動帶入');

    var catMap = {};
    rows.forEach(function (r) { var c = String(r[3] || '(空白)').trim(); catMap[c] = (catMap[c] || 0) + 1; });
    var cats = Object.keys(catMap).sort(function (a, b) { return catMap[b] - catMap[a]; });
    push('', '', '');
    push('科別分布', cats.length + ' 個科別', '');
    cats.forEach(function (c) { push('　' + c, catMap[c] + ' 題', ''); });

    var yrMap = {};
    rows.forEach(function (r) { var y = String(r[1] || '(空白)').trim(); yrMap[y] = (yrMap[y] || 0) + 1; });
    push('', '', '');
    push('年分分布', '', '');
    Object.keys(yrMap).sort().forEach(function (y) { push('　' + y, yrMap[y] + ' 題', ''); });

    var parseBad = 0, parseLoose = 0;
    rows.forEach(function (r) {
      var pr = splitQuestion(r[0]);
      if (!pr.ok) parseBad++; else if (pr.loose) parseLoose++;
    });
    push('', '', '');
    push('選項可正確拆成 4 個', (rows.length - parseBad) + ' / ' + rows.length,
         parseBad ? '⚠ 有 ' + parseBad + ' 題拆不開' : '✓');
    if (parseLoose) push('需人工確認題幹', parseLoose + ' 題', '用寬鬆規則才切開');
  }

  push('', '', '');
  push('── 目前的考卷（' + CFG.SHEET_SELECTED + '）──', '', '');
  var sel = ss.getSheetByName(CFG.SHEET_SELECTED);
  if (!sel) push('狀態', '尚未建立', '按「重新抽題」就會產生');
  else {
    push('題數', Math.max(0, sel.getLastRow() - 1), '');
    var bg = sel.getRange(1, 1).getBackground();
    push('標題列底色', bg, bg.toLowerCase() === '#37474f'
      ? '✓ 是新版抽題寫的' : '← 還是舊版格式，代表新版「重新抽題」還沒成功跑過');
  }

  var sh2 = ss.getSheetByName('環境診斷') || ss.insertSheet('環境診斷');
  sh2.clear();
  sh2.getRange(1, 1, out.length, 3).setValues(out);
  sh2.getRange(1, 1, 1, 3).setFontWeight('bold');
  for (var k = 0; k < out.length; k++) {
    var t = String(out[k][0]);
    if (t.indexOf('──') === 0) sh2.getRange(k + 1, 1, 1, 3).setFontWeight('bold').setBackground('#eceff1');
    if (String(out[k][2]).indexOf('✗') === 0 || String(out[k][2]).indexOf('⚠') === 0)
      sh2.getRange(k + 1, 1, 1, 3).setBackground('#fce8e6');
  }
  sh2.setColumnWidth(1, 220); sh2.setColumnWidth(2, 180); sh2.setColumnWidth(3, 420);
  sh2.activate();

  ui.alert('診斷完成', '程式版本：' + VERSION + '\n\n結果寫在「環境診斷」工作表，請整份看過。',
    ui.ButtonSet.OK);
}

/* ═══════════════════════════════════════════════════════════════
   洗牌：Fisher-Yates（每個排列等機率）
   ═══════════════════════════════════════════════════════════════ */
function fyShuffle(arr) {
  var a = arr.slice();
  for (var i = a.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var t = a[i]; a[i] = a[j]; a[j] = t;
  }
  return a;
}

/* ═══════════════════════════════════════════════════════════════
   配額計算：最大餘數法
   mode = 'proportional' 依題庫比例 ／ 'equal' 每科平均 ／ 'min1' 依比例但每科至少 1 題
   保證：總數剛好等於 N，且任一科不超過它實際擁有的題數
   ═══════════════════════════════════════════════════════════════ */
function allocateQuota(catCounts, N, mode) {
  var cats = Object.keys(catCounts);
  var total = 0;
  cats.forEach(function (c) { total += catCounts[c]; });
  N = Math.min(N, total);

  var floors = {}, quota = {}, used = 0, rest;

  if (mode === 'min1') {
    if (N < cats.length) return null;                 // 題數不夠每科一題
    cats.forEach(function (c) { floors[c] = 1; });
    rest = N - cats.length;
    cats.forEach(function (c) { quota[c] = rest * catCounts[c] / total; });
    cats.forEach(function (c) {
      var add = Math.min(catCounts[c] - 1, Math.floor(quota[c]));
      floors[c] += add; used += add;
    });
    rest = rest - used;
  } else {
    cats.forEach(function (c) {
      quota[c] = (mode === 'proportional') ? N * catCounts[c] / total : N / cats.length;
      floors[c] = Math.min(catCounts[c], Math.floor(quota[c]));
      used += floors[c];
    });
    rest = N - used;
  }

  // 剩下的名額依「小數部分大小」依序發放，該科發滿了就跳過
  var order = cats.slice().sort(function (a, b) {
    return (quota[b] - Math.floor(quota[b])) - (quota[a] - Math.floor(quota[a]));
  });
  var guard = 0;
  while (rest > 0 && guard++ < 10000) {
    var moved = false;
    for (var i = 0; i < order.length && rest > 0; i++) {
      var c = order[i];
      if (floors[c] < catCounts[c]) { floors[c]++; rest--; moved = true; }
    }
    if (!moved) break;
  }
  return floors;
}

/* ═══════════════════════════════════════════════════════════════
   主流程
   ═══════════════════════════════════════════════════════════════ */
function drawQuestions() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    /* --- 1. 讀題庫（用名稱，不看你停在哪一頁） --- */
    var bank = ss.getSheetByName(DRAW.SHEET_BANK);
    if (!bank) {
      var names = ss.getSheets().map(function (s) { return s.getName(); }).join('、');
      throw new Error('找不到名為「' + DRAW.SHEET_BANK + '」的工作表。\n\n'
        + '目前有：' + names + '\n\n'
        + '請把題庫那張表改名成「' + DRAW.SHEET_BANK + '」，'
        + '或修改程式最上方 DRAW.SHEET_BANK 的設定。');
    }

    var raw = bank.getDataRange().getValues();
    var hRow = findHeaderRow_(raw);              // 表格化之後標題列不一定在第 1 列
    var pool = [];
    for (var i = hRow + 1; i < raw.length; i++) {
      if (!raw[i][0]) continue;
      pool.push({ q: raw[i][0], year: String(raw[i][1] || '').trim(),
                  ans: raw[i][2], cat: String(raw[i][3] || '（未分類）').trim() });
    }
    if (!pool.length) throw new Error('「' + DRAW.SHEET_BANK + '」裡讀不到任何題目。');

    /* --- 2. 先給使用者看清楚題庫長什麼樣 --- */
    var byYear = {}, byCat = {};
    pool.forEach(function (p) {
      byYear[p.year] = (byYear[p.year] || 0) + 1;
      byCat[p.cat]   = (byCat[p.cat] || 0) + 1;
    });
    var yearList = Object.keys(byYear).sort();
    var catList  = Object.keys(byCat).sort(function (a, b) { return byCat[b] - byCat[a]; });

    /* --- 3. 覆蓋警告 --- */
    var cur = ss.getSheetByName(DRAW.SHEET_SELECTED);
    if (cur && cur.getLastRow() > 1) {
      var ok = ui.alert('要覆蓋現在的考卷嗎？',
        '「' + DRAW.SHEET_SELECTED + '」目前有 ' + (cur.getLastRow() - 1) + ' 題。\n'
        + '重新抽題會直接覆蓋掉它。\n\n'
        + '如果那份考卷已經發出去、成績還沒統計完，'
        + '請先取消，去跑「② 產生完整測驗報告」和「③ 封存本次考卷」。\n\n'
        + '要繼續嗎？', ui.ButtonSet.YES_NO);
      if (ok !== ui.Button.YES) return;
    }

    /* --- 4. 題數 --- */
    var r1 = ui.prompt('要抽幾題？',
      '題庫共 ' + pool.length + ' 題，分成 ' + catList.length + ' 個科別。\n\n請輸入題數：',
      ui.ButtonSet.OK_CANCEL);
    if (r1.getSelectedButton() !== ui.Button.OK) return;
    var N = parseInt(r1.getResponseText(), 10);
    if (!N || N < 1) N = DRAW.DEFAULT_N;
    N = Math.min(N, pool.length);

    /* --- 5. 年分篩選 --- */
    var r2 = ui.prompt('要限定年分嗎？',
      '題庫裡的年分：' + yearList.map(function (y) { return y + '(' + byYear[y] + '題)'; }).join('、') + '\n\n'
      + '想全部都抽 → 直接按確定（留空）\n'
      + '只要某幾年 → 例如輸入  112,113',
      ui.ButtonSet.OK_CANCEL);
    if (r2.getSelectedButton() !== ui.Button.OK) return;
    var yearTxt = r2.getResponseText().trim();
    var wantYears = yearTxt ? yearTxt.split(/[,，\s]+/).filter(String) : null;

    var filtered = wantYears
      ? pool.filter(function (p) { return wantYears.indexOf(p.year) >= 0; })
      : pool;
    if (!filtered.length) throw new Error('年分「' + yearTxt + '」篩選後沒有任何題目。');
    if (filtered.length < N) {
      ui.alert('題目不夠', '篩選後只剩 ' + filtered.length + ' 題，將全部選入。', ui.ButtonSet.OK);
      N = filtered.length;
    }

    /* --- 6. 科別配比 --- */
    var fCat = {};
    filtered.forEach(function (p) { fCat[p.cat] = (fCat[p.cat] || 0) + 1; });
    var fCatList = Object.keys(fCat);

    var modeResp = ui.alert('科別要怎麼配？',
      '【是】依題庫比例\n'
      + '　　題多的科別就出比較多題，最接近真正的考試組成。\n'
      + '　　缺點：題目很少的冷門科別可能一題都不會出現。\n\n'
      + '【否】每科平均\n'
      + '　　每個科別出差不多的題數，方便比較各科強弱。\n'
      + '　　缺點：只有 3、4 題的冷門科別會被重複抽到，考幾次就背起來了。\n\n'
      + '【取消】依比例，但每科至少 1 題（折衷，通常選這個）',
      ui.ButtonSet.YES_NO_CANCEL);

    var mode = modeResp === ui.Button.YES ? 'proportional'
             : modeResp === ui.Button.NO  ? 'equal' : 'min1';

    if (mode === 'min1' && N < fCatList.length) {
      ui.alert('題數不夠',
        '篩選後有 ' + fCatList.length + ' 個科別，但你只要抽 ' + N + ' 題，'
        + '沒辦法每科至少 1 題。\n\n已改用「依題庫比例」。', ui.ButtonSet.OK);
      mode = 'proportional';
    }

    var quota = allocateQuota(fCat, N, mode);
    if (!quota) throw new Error('配額計算失敗。');

    /* --- 7. 抽題：每科先洗牌再取前 n 題 --- */
    var picked = [];
    fCatList.forEach(function (c) {
      var inCat = filtered.filter(function (p) { return p.cat === c; });
      picked = picked.concat(fyShuffle(inCat).slice(0, quota[c] || 0));
    });
    // 若因為某科題目用光而不足，從剩下的題目隨機補滿
    if (picked.length < N) {
      var chosen = {};
      picked.forEach(function (p) { chosen[p.q] = true; });
      var spare = fyShuffle(filtered.filter(function (p) { return !chosen[p.q]; }));
      picked = picked.concat(spare.slice(0, N - picked.length));
    }
    picked = fyShuffle(picked).slice(0, N);   // 打散科別順序，避免同科擠在一起

    /* --- 8. 寫入「選出題目」 --- */
    var tgt = ss.getSheetByName(DRAW.SHEET_SELECTED) || ss.insertSheet(DRAW.SHEET_SELECTED);
    tgt.clear();
    var headers = ['題目', '年分', '正確答案', '科別'];
    tgt.getRange(1, 1, 1, 4).setValues([headers])
       .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    tgt.getRange(2, 1, picked.length, 4).setValues(picked.map(function (p) {
      return [p.q, p.year, p.ans, p.cat];
    }));

    tgt.setColumnWidth(1, 500); tgt.setColumnWidth(2, 80);
    tgt.setColumnWidth(3, 100); tgt.setColumnWidth(4, 110);
    tgt.getRange(2, 1, picked.length, 1).setWrap(true).setVerticalAlignment('top');
    tgt.getRange(2, 2, picked.length, 3).setHorizontalAlignment('center').setVerticalAlignment('middle');
    tgt.getRange(1, 1, picked.length + 1, 4).setBorder(true, true, true, true, true, true);
    for (var r = 2; r <= picked.length + 1; r++) {
      if (r % 2 === 0) tgt.getRange(r, 1, 1, 4).setBackground('#f8f9fa');
    }
    tgt.setFrozenRows(1);

    /* --- 9. 抽題摘要 --- 
       寫在右側 F/G 欄，不寫在題目下方。
       寫在下方會讓 A 欄多出「抽題時間」「各科題數」「ID」這些非題目的內容，
       而讀取器是靠「A 欄非空」判斷有沒有題目的 —— 摘要會被當成題目讀進去，
       健檢報一堆假問題，產生表單也會被自己的摘要擋下來。 */
    var got = {};
    picked.forEach(function (p) { got[p.cat] = (got[p.cat] || 0) + 1; });
    var modeName = { proportional: '依題庫比例', equal: '每科平均', min1: '依比例＋每科至少1題' }[mode];

    var info = [
      ['本次抽題設定', ''],
      ['抽題時間', Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm')],
      ['題數',     picked.length],
      ['年分',     wantYears ? wantYears.join('、') : '全部'],
      ['科別配比', modeName],
      ['', ''],
      ['各科題數', '']
    ].concat(
      Object.keys(got).sort(function (a, b) { return got[b] - got[a]; })
        .map(function (c) { return [c, got[c] + ' 題　（題庫 ' + fCat[c] + ' 題）']; })
    );
    tgt.getRange(1, 6, info.length, 2).setValues(info);
    tgt.getRange(1, 6).setFontWeight('bold');
    tgt.getRange(7, 6).setFontWeight('bold');
    tgt.setColumnWidth(5, 24);          // E 欄留白，和題目區隔開
    tgt.setColumnWidth(6, 110);
    tgt.setColumnWidth(7, 190);
    tgt.getRange(1, 6, info.length, 2).setVerticalAlignment('middle');

    tgt.activate();

    ui.alert('抽題完成',
      '已抽出 ' + picked.length + ' 題（' + modeName + '）。\n\n'
      + Object.keys(got).sort(function (a, b) { return got[b] - got[a]; })
          .map(function (c) { return '· ' + c + '：' + got[c] + ' 題'; }).join('\n')
      + '\n\n下一步建議先跑「① 題庫健檢」，確認正解欄格式都正確，再產生表單。',
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}


var CFG = {
  SHEET_SELECTED : '選出題目',   // A題目 B年分 C正確答案 D科別
  SHEET_LINKS    : '表單連結',   // B1=表單連結  B2=編輯連結
  SHEET_MAP      : '題目對照',   // 表單 itemId ↔ 選出題目列號
  SHEET_HISTORY  : '歷次表單',   // 每一份發出去過的表單都登記在這裡
  OUT_PERSON     : '個人成績',
  OUT_WRONG      : '個人錯題',
  OUT_ITEM       : '逐題分析',
  OUT_HOSPITAL   : '醫院科別',
  WEAK_THRESHOLD : 60,           // 答對率低於這個數就標紅
  OUT_BATCH      : '梯次弱點',   // 依作答日期分批後的弱點分析
  OUT_COMPARE    : '前後測比較',
  BATCH_GAP_DAYS : 14            // 前後兩筆作答相隔超過這麼多天，就視為不同梯次
};

/* ═══════════════════════════════════════════════════════════════
   選單
   ═══════════════════════════════════════════════════════════════ */


/* ═══════════════════════════════════════════════════════════════
   共用工具
   ═══════════════════════════════════════════════════════════════ */

/**
 * 把「題幹＋四個選項」的整段文字拆開。
 *
 * 原版用 text.split(/(?=[A-D]\.)/) 然後取 parts[0] 當題幹、slice(-4) 當選項。
 * 問題：只要題幹裡出現 "vitamin D." 這種字串，就會多切一刀，
 *       而 parts[0] 只保留第一刀之前 → 題幹被默默截斷，你不會收到任何警告。
 *
 * 這裡的作法：
 *   1. 優先用「換行 + A./B./C./D.」來切（題庫幾乎都是這種格式，最安全）
 *   2. 切不出來才退回原本的寬鬆切法，但題幹改成「除了最後四段以外全部接回去」
 *   3. 一律回報 ok / reason，讓呼叫端知道這題可不可信
 */
function splitQuestion(raw) {
  var text = String(raw == null ? '' : raw).replace(/\r\n?/g, '\n').trim();
  if (!text) return { ok: false, reason: '題目是空的', stem: '', choices: [] };

  // v18：支援四選項（國考 A–D）與五選項（專科甄審 A–E）
  var parts = text.split(/\n(?=[A-EＡ-Ｅ][.、．]\s*)/);
  var loose = false;

  if (parts.length !== 5 && parts.length !== 6) {
    parts = text.split(/(?=[A-E][.、]\s*)/);
    loose = true;
  }
  if (parts.length < 5) {
    return { ok: false, reason: '找不到完整的四個（或五個）選項（只切出 ' + parts.length + ' 段）', stem: text, choices: [] };
  }

  // 選項數：最後一段是 E 才算五選項
  var k = (parts.length >= 6 && normalizeLetter(parts[parts.length - 1].charAt(0)) === 'E') ? 5 : 4;
  var opts = parts.slice(-k);
  var stem = parts.slice(0, parts.length - k).join('').trim();   // ← 不再丟掉中間段落

  var want = 'ABCDE'.slice(0, k);
  var letters = opts.map(function (o) { return normalizeLetter(o.charAt(0)); });
  if (letters.join('') !== want) {
    return { ok: false, reason: '選項標籤順序不是 ' + want.split('').join(' ') + '（讀到 ' + letters.join('') + '）', stem: stem, choices: [] };
  }

  return {
    ok: true,
    loose: loose,
    reason: loose ? '用寬鬆規則才切開，請人工確認題幹完整' : '',
    stem: stem,
    choices: opts.map(function (o, i) {
      return { letter: 'ABCDE'.charAt(i), full: o.trim(), body: o.trim().replace(/^[A-EＡ-Ｅ][.、．]\s*/, '') };
    })
  };
}

/** 全形轉半形、取出 A~E 這個字母；認不出來回 null */
function normalizeLetter(v) {
  var s = String(v == null ? '' : v).trim().toUpperCase()
          .replace(/[Ａ-Ｅ]/g, function (c) { return String.fromCharCode(c.charCodeAt(0) - 0xFEE0); });
  var m = s.match(/[A-E]/);
  return m ? m[0] : null;
}

/**
 * 正確答案欄的清洗。
 * 原版用 choice.startsWith(correctAnswer + '.')，只要 C 欄填的是
 * 「A 」「(A)」「Ａ」「a」「A.」任何一種，就沒有任何選項會被標成正解，
 * 該題在 Google 表單裡變成「沒有正解」→ 全班都算答錯，而且不會報錯。
 */
function normalizeAnswer(v) { return normalizeLetter(v); }

function sheetOf(name, create) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh && create) sh = ss.insertSheet(name);
  return sh;
}

/** 讀任一張題目表（欄位須為 題目/年分/正確答案/科別） */
function readQuestionSheet_(sheetName) {
  var sh = sheetOf(sheetName, false);
  if (!sh) throw new Error('找不到「' + sheetName + '」工作表。');
  var data = sh.getDataRange().getValues();
  var hRow = findHeaderRow_(data);
  var out = [];
  for (var i = hRow + 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    // 舊版把抽題摘要寫在題目下方；遇到摘要標記就停，不要把它當成題目
    var a0 = String(data[i][0]).trim();
    if (a0 === '本次抽題設定' || a0 === '各科題數') break;
    var pp = splitQuestion(data[i][0]);
    out.push({
      row       : i - hRow,          // 第幾題
      sheetRow  : i + 1,             // 實際列號，方便你直接跳過去改
      raw       : data[i][0],
      year      : data[i][1],
      answer    : normalizeAnswer(data[i][2]),
      rawAnswer : data[i][2],
      category  : String(data[i][3] || '（未分類）').trim(),
      explain   : String(data[i][4] || '').trim(),   // E 欄＝詳解，寫在題庫裡，同一題永久沿用
      stem      : pp.stem,
      choices   : pp.choices,
      parse     : pp
    });
  }
  return out;
}

/** 讀「選出題目」，回傳 [{row, stem, choices, answer, year, category, parse}] */
function readSelected() {
  var sh = sheetOf(CFG.SHEET_SELECTED, false);
  if (!sh) throw new Error('找不到「' + CFG.SHEET_SELECTED + '」工作表，請先執行「重新抽題」。');
  var data = sh.getDataRange().getValues();
  var hRow = findHeaderRow_(data);
  var out = [];
  for (var i = hRow + 1; i < data.length; i++) {
    if (!data[i][0]) continue;
    // 舊版把抽題摘要寫在題目下方；遇到摘要標記就停，不要把它當成題目
    var a0 = String(data[i][0]).trim();
    if (a0 === '本次抽題設定' || a0 === '各科題數') break;
    var p = splitQuestion(data[i][0]);
    out.push({
      row      : i,                      // 對應原本表單標題「第 i 題」
      raw      : data[i][0],
      year     : data[i][1],
      answer   : normalizeAnswer(data[i][2]),
      rawAnswer: data[i][2],
      category : data[i][3] || '（未分類）',
      stem     : p.stem,
      choices  : p.choices,
      parse    : p
    });
  }
  return out;
}

/* ═══════════════════════════════════════════════════════════════
   依作答日期自動分批
   ---------------------------------------------------------------
   同一份表單如果跨了好幾個梯次，成績混在一起算，平均分和答對率都會失真。
   這裡把回應依時間排序，前後相隔超過 BATCH_GAP_DAYS 天就切成新的一批。

   限制講清楚：靠日期間隔只能分出「不同時間的測驗」，
   分不出「同一梯次的前測 vs 後測」。要精確就每次測驗開一份新表單，
   一份表單＝一次測驗，不必用猜的。
   ═══════════════════════════════════════════════════════════════ */
function assignBatches_(people) {
  if (!people.length) return [];
  var sorted = people.slice().sort(function (a, b) { return a.ts - b.ts; });
  var gapMs = CFG.BATCH_GAP_DAYS * 24 * 3600 * 1000;
  var batches = [], cur = null;

  sorted.forEach(function (p) {
    if (!cur || (p.ts - cur.to) > gapMs) {
      cur = { idx: batches.length + 1, from: p.ts, to: p.ts, members: [] };
      batches.push(cur);
    }
    cur.to = p.ts;
    cur.members.push(p);
  });

  batches.forEach(function (b) {
    var f = Utilities.formatDate(b.from, 'GMT+8', 'yyyy/MM/dd');
    var t = Utilities.formatDate(b.to, 'GMT+8', 'MM/dd');
    b.label = '第' + b.idx + '批　' + f + (f.slice(5) === t ? '' : '–' + t);
    b.members.forEach(function (p) { p.batch = b.label; });
  });
  return batches;
}

/* ═══════════════════════════════════════════════════════════════
   梯次弱點分析
   每一批的平均分、最弱三個科別、最弱三題，以及科別×梯次答對率矩陣
   ═══════════════════════════════════════════════════════════════ */
function writeBatchSheet(batches, qs) {
  var sh = sheetOf(CFG.OUT_BATCH, true);
  sh.clear();
  if (!batches.length) { sh.getRange(1, 1).setValue('沒有資料'); return; }

  /* 每一批統計各科別與各題 */
  batches.forEach(function (b) {
    b.cat = {}; b.item = {}; b.sum = 0;
    b.members.forEach(function (p) {
      b.sum += p.total ? (p.ok / p.total * 100) : 0;
      var wrongK = {};
      p.wrongs.forEach(function (w) { wrongK[w.k] = true; });
      qs.forEach(function (q, k) {
        var c = b.cat[q.category] || (b.cat[q.category] = { t: 0, ok: 0 });
        c.t++; if (!wrongK[k]) c.ok++;
        var it = b.item[k] || (b.item[k] = { t: 0, ok: 0, q: q });
        it.t++; if (!wrongK[k]) it.ok++;
      });
    });
    b.avg = b.members.length ? b.sum / b.members.length : 0;
    b.catRank = Object.keys(b.cat).map(function (c) {
      return { c: c, rate: b.cat[c].t ? Math.round(b.cat[c].ok / b.cat[c].t * 100) : 0 };
    }).sort(function (x, y) { return x.rate - y.rate; });
    b.itemRank = Object.keys(b.item).map(function (k) {
      var it = b.item[k];
      return { k: +k, q: it.q, rate: it.t ? Math.round(it.ok / it.t * 100) : 0 };
    }).sort(function (x, y) { return x.rate - y.rate; });
  });

  var row = 1;
  sh.getRange(row, 1).setValue('梯次弱點分析').setFontWeight('bold').setFontSize(13); row += 2;

  /* --- 區塊 1：每批摘要 --- */
  var head1 = ['梯次', '人數', '平均分', '最弱科別 ①', '②', '③', '最弱的一題'];
  sh.getRange(row, 1, 1, head1.length).setValues([head1])
    .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  var r1 = row + 1;
  batches.forEach(function (b) {
    var top3 = b.catRank.slice(0, 3).map(function (x) { return x.c + ' ' + x.rate + '%'; });
    var worstQ = b.itemRank[0];
    sh.getRange(r1, 1, 1, head1.length).setValues([[
      b.label, b.members.length, Math.round(b.avg * 10) / 10,
      top3[0] || '', top3[1] || '', top3[2] || '',
      worstQ ? ('第' + (worstQ.k + 1) + '題 ' + worstQ.rate + '%　' + String(worstQ.q.stem).slice(0, 40)) : ''
    ]]);
    sh.getRange(r1, 3).setBackground(b.avg >= 80 ? '#e6f4ea' : b.avg >= CFG.WEAK_THRESHOLD ? '#fef7e0' : '#fce8e6')
                       .setFontWeight('bold');
    sh.getRange(r1, 4).setBackground('#fce8e6').setFontWeight('bold');
    r1++;
  });
  sh.getRange(row, 1, batches.length + 1, head1.length).setBorder(true, true, true, true, true, true);
  row = r1 + 2;

  /* --- 區塊 2：科別 × 梯次 答對率矩陣 --- */
  sh.getRange(row, 1).setValue('科別 × 梯次　答對率（%）').setFontWeight('bold'); row++;
  sh.getRange(row, 1).setValue('每一欄最低的那一格標紅。如果同一個科別每批都紅，那是課程問題，不是學生問題。')
    .setFontColor('#5f6368'); row++;

  var cats = {};
  batches.forEach(function (b) { Object.keys(b.cat).forEach(function (c) { cats[c] = 1; }); });
  var catList = Object.keys(cats).sort();

  var head2 = ['科別'].concat(batches.map(function (b) { return b.label; }));
  sh.getRange(row, 1, 1, head2.length).setValues([head2])
    .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');

  var matrix = catList.map(function (c) {
    return [c].concat(batches.map(function (b) {
      return b.cat[c] && b.cat[c].t ? Math.round(b.cat[c].ok / b.cat[c].t * 100) : '';
    }));
  });
  if (matrix.length) {
    sh.getRange(row + 1, 1, matrix.length, head2.length).setValues(matrix);
    // 每一欄（梯次）找最低值標紅
    batches.forEach(function (b, j) {
      var lo = 999, loRow = -1;
      matrix.forEach(function (m, i) {
        var v = m[j + 1];
        if (v !== '' && v < lo) { lo = v; loRow = i; }
      });
      if (loRow >= 0) sh.getRange(row + 1 + loRow, j + 2).setBackground('#fce8e6').setFontWeight('bold');
      // 其餘依門檻上色
      matrix.forEach(function (m, i) {
        var v = m[j + 1];
        if (v === '' || i === loRow) return;
        sh.getRange(row + 1 + i, j + 2)
          .setBackground(v >= 80 ? '#e6f4ea' : v >= CFG.WEAK_THRESHOLD ? '#fef7e0' : '#fbeae8');
      });
    });
    sh.getRange(row, 1, matrix.length + 1, head2.length).setBorder(true, true, true, true, true, true);
    sh.getRange(row + 1, 2, matrix.length, batches.length).setHorizontalAlignment('center');
  }

  sh.setColumnWidth(1, 130);
  for (var c2 = 2; c2 <= head2.length; c2++) sh.setColumnWidth(c2, 130);
  sh.setColumnWidth(7, 420);
  sh.setFrozenRows(1);
}

/* ═══════════════════════════════════════════════════════════════
   一次產生前測＋後測兩份表單
   ---------------------------------------------------------------
   兩份在同一個時間點、從同一份「選出題目」生出來，
   所以絕對是同一卷 —— 中間不可能有人按到「重新抽題」把卷換掉。

   前測：原順序。後測：題目順序打亂（可選連選項一起打亂）。
   四個連結一次寫進「表單連結」，兩份都登記到「歷次表單」。
   ═══════════════════════════════════════════════════════════════ */
function generatePrePostForms() {
  var ui = SpreadsheetApp.getUi();
  try {
    var qs = readSelected();
    if (!qs.length) throw new Error('「' + CFG.SHEET_SELECTED + '」裡沒有題目，請先執行「重新抽題」。');

    var bad = qs.filter(function (q) { return !q.parse.ok || !q.answer; });
    if (bad.length) {
      throw new Error('有 ' + bad.length + ' 題無法做成表單（第 '
        + bad.slice(0, 5).map(function (b) { return b.row; }).join('、')
        + (bad.length > 5 ? ' …' : '') + ' 題）。\n\n請先執行「① 題庫健檢」修正。');
    }

    var go = ui.alert('一次產生前測與後測',
      '會用「' + CFG.SHEET_SELECTED + '」現有的 ' + qs.length + ' 題建立兩份表單：\n\n'
      + '· 前測 —— 原順序\n'
      + '· 後測 —— 題目順序打亂（同一卷）\n\n'
      + '四個連結會一起寫進「' + CFG.SHEET_LINKS + '」，舊連結會先登記到「'
      + CFG.SHEET_HISTORY + '」。\n\n要繼續嗎？', ui.ButtonSet.YES_NO);
    if (go !== ui.Button.YES) return;

    var shufC = ui.alert('後測的選項順序也要打亂嗎？',
      '【是】連選項一起打亂（依新位置重編 A/B/C/D/E）\n'
      + '　　　更能避免背答案位置；報表裡的字母會變成顯示位置，\n'
      + '　　　但選項文字一律印在字母旁邊，仍看得出是哪一個\n\n'
      + '【否】只打亂題目順序，選項維持原樣', ui.ButtonSet.YES_NO_CANCEL);
    if (shufC === ui.Button.CANCEL) return;

    /* 舊連結先登記 */
    archiveFormLink_();

    /* 前測：原順序 */
    var pre = buildQuizForm_(qs, '前測', false, true);

    /* 後測：題目順序打亂（只動記憶體中的順序，不改試算表） */
    var shuffled = fyShuffle(qs);
    var post = buildQuizForm_(shuffled, '後測', shufC === ui.Button.YES, true);

    /* 四個連結一起寫 */
    var links = sheetOf(CFG.SHEET_LINKS, true);
    links.clear();
    links.getRange(1, 1, 4, 2).setValues([
      ['前測 表單連結', pre.formUrl],
      ['前測 編輯連結', pre.editUrl],
      ['後測 表單連結', post.formUrl],
      ['後測 編輯連結', post.editUrl]
    ]);
    links.getRange(1, 1, 4, 1).setFontWeight('bold');
    links.getRange(1, 1, 2, 2).setBackground('#e3eaf3');
    links.getRange(3, 1, 2, 2).setBackground('#e4f1e9');
    links.setColumnWidth(1, 130); links.setColumnWidth(2, 620);

    /* 兩份都登記 */
    archiveFormLink_();

    var warn = (pre.warn || post.warn) ? '\n\n⚠ ' + (pre.warn || post.warn) : '';
    ui.alert('前測與後測都建好了',
      '同一卷 ' + qs.length + ' 題，後測順序已打亂'
      + (shufC === ui.Button.YES ? '，選項也已打亂' : '') + '。\n\n'
      + '【前測】發這個：\n' + pre.formUrl + '\n\n'
      + '【後測】課程結束後再發：\n' + post.formUrl + '\n\n'
      + '四個連結都在「' + CFG.SHEET_LINKS + '」，也已登記到「' + CFG.SHEET_HISTORY + '」。'
      + warn, ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}

/* ═══════════════════════════════════════════════════════════════
   產生後測表單：同卷、打亂順序
   ---------------------------------------------------------------
   沿用「選出題目」現有的題目（也就是前測那一份），只重新排列順序。

   為什麼同卷最乾淨：同一題的難度是固定的，所以「這題前測 40% → 後測 85%」
   這種比較把題目難度的變異完全消掉，也大幅減少迴歸平均的問題。
   代價是記憶效應 —— 學生可能記得題目本身，所以進步幅度仍會被高估一些。

   做法上：打亂的是「選出題目」的列順序（這樣第 N 題仍然對應第 N 列，
   報表的對照關係不會壞），選項順序則只在建表單時打亂，不動試算表。
   ═══════════════════════════════════════════════════════════════ */
function generatePostTestForm() {
  var ui = SpreadsheetApp.getUi();
  try {
    var qs = readSelected();
    if (!qs.length) throw new Error('「' + CFG.SHEET_SELECTED + '」裡沒有題目。');

    var bad = qs.filter(function (q) { return !q.parse.ok || !q.answer; });
    if (bad.length) throw new Error('有 ' + bad.length + ' 題無法做成表單，請先執行「① 題庫健檢」。');

    var go = ui.alert('產生後測表單（同卷）',
      '會沿用「' + CFG.SHEET_SELECTED + '」現有的 ' + qs.length + ' 題，只重新排列順序。\n\n'
      + '請確認這' + qs.length + '題就是前測用的那一份 —— '
      + '中間如果按過「重新抽題」，就已經不是同一卷了。\n\n'
      + '要繼續嗎？', ui.ButtonSet.YES_NO);
    if (go !== ui.Button.YES) return;

    var shufC = ui.alert('選項順序也要打亂嗎？',
      '【是】連選項順序一起打亂（依新位置重編 A/B/C/D/E）\n'
      + '　　　更能避免背答案位置，但報表裡的字母不再對應題庫的原始字母\n'
      + '　　　（報表會把選項文字印在字母旁邊，仍然看得出是哪一個）\n\n'
      + '【否】只打亂題目順序，選項維持原樣', ui.ButtonSet.YES_NO_CANCEL);
    if (shufC === ui.Button.CANCEL) return;

    /* 打亂「選出題目」的列順序 —— 題目不變，只換位置 */
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(CFG.SHEET_SELECTED);
    var data = sh.getDataRange().getValues();
    var hRow = findHeaderRow_(data);
    var rows = [];
    for (var i = hRow + 1; i < data.length; i++) {
      var a0 = String(data[i][0]).trim();
      if (!data[i][0]) continue;
      if (a0 === '本次抽題設定' || a0 === '各科題數') break;
      rows.push(data[i].slice(0, 4));
    }
    var shuffled = fyShuffle(rows);
    sh.getRange(hRow + 2, 1, shuffled.length, 4).setValues(shuffled);

    /* 舊表單先登記，再建新的 */
    var logged = archiveFormLink_();

    var qs2 = readSelected();
    var info = buildQuizForm_(qs2, '後測', shufC === ui.Button.YES);

    ui.alert('後測表單已建立',
      '共 ' + info.count + ' 題（與前測同一卷，順序已打亂'
      + (shufC === ui.Button.YES ? '，選項也已打亂' : '') + '）。\n\n'
      + '連結：\n' + info.formUrl + '\n\n'
      + (logged ? '舊表單「' + logged.title + '」（' + logged.responses + ' 筆）已登記到「'
                  + CFG.SHEET_HISTORY + '」。\n\n' : '')
      + '收完之後跑「②″ 前後測比較」，會有逐題前後對照。'
      + (info.warn ? '\n\n⚠ ' + info.warn : ''),
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}

/* ═══════════════════════════════════════════════════════════════
   詳解講義
   ---------------------------------------------------------------
   詳解寫在「題庫」的 E 欄（標題請打「詳解」），跟著題目走，
   不跟著某一次測驗走 —— 同一題下次再考到就自動沿用。

   講義依全班答對率由低到高排列，並附上這次的實際作答分布，
   所以最該講的題目一定在最前面，而且看得到大家錯在哪個選項。
   ═══════════════════════════════════════════════════════════════ */
function buildExplanationHandout() {
  var ui = SpreadsheetApp.getUi();
  try {
    var form = openTheForm();
    var A = analyzeForm_(form);
    if (!A.people.length) throw new Error('這份表單還沒有人作答，先收到回應才做得出有意義的講義。');

    /* 詳解索引：題幹 → 題庫的詳解 */
    var expIdx = {};
    try {
      readQuestionSheet_(DRAW.SHEET_BANK).forEach(function (q) {
        var k = normKey_(q.stem);
        if (k) expIdx[k] = q.explain;
      });
    } catch (e) {}

    /* 每題的答對率與選項分布 */
    var stats = A.itemStat.map(function (st) {
      var counts = {};
      A.people.forEach(function (p) {
        var picked = null, wrong = false;
        p.wrongs.forEach(function (w) { if (w.k === st.k) { wrong = true; picked = w.picked; } });
        var L = wrong ? (picked || '未作答') : st.q.answer;
        counts[L] = (counts[L] || 0) + 1;
      });
      var trap = null, tn = 0;
      Object.keys(counts).forEach(function (L) {
        if (L !== st.q.answer && L !== '未作答' && counts[L] > tn) { trap = L; tn = counts[L]; }
      });
      return { k: st.k, q: st.q, n: A.people.length, ok: st.ok,
               rate: A.people.length ? Math.round(st.ok / A.people.length * 100) : 0,
               counts: counts, trap: trap, trapN: tn,
               explain: expIdx[normKey_(st.q.stem)] || '' };
    }).sort(function (a, b) { return a.rate - b.rate; });

    var missing = stats.filter(function (x) { return !x.explain; });

    /* --- 建立 Google 文件 --- */
    var tag = Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd');
    var doc = DocumentApp.create('兒科題庫測驗_詳解講義_' + tag);
    var body = doc.getBody();

    body.appendParagraph('兒科題庫測驗　詳解講義').setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph(form.getTitle());
    body.appendParagraph('作答人數 ' + A.people.length + ' 人　·　共 ' + stats.length + ' 題　·　'
      + Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd') + ' 製');
    body.appendParagraph('依全班答對率由低到高排列 —— 最前面的就是最該講的。')
        .setItalic(true);
    body.appendHorizontalRule();

    /* 摘要：最弱三題 */
    body.appendParagraph('本次最弱的題目').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    stats.slice(0, 3).forEach(function (x) {
      body.appendListItem('第 ' + (x.k + 1) + ' 題（' + x.q.category + '）答對率 ' + x.rate + '%'
        + (x.trap ? '　主要誤選 ' + x.trap + '（' + x.trapN + ' 人）' : ''));
    });
    body.appendHorizontalRule();

    /* 逐題 */
    stats.forEach(function (x) {
      body.appendParagraph('第 ' + (x.k + 1) + ' 題　' + x.q.category
        + '　答對率 ' + x.rate + '%（' + x.ok + '/' + x.n + '）')
        .setHeading(DocumentApp.ParagraphHeading.HEADING2);

      if (x.trap) {
        body.appendParagraph('主要誤選：' + x.trap + '　' + x.trapN + ' 人'
          + '　——　答錯的人集中選同一個選項，通常代表某個觀念被混淆，值得特別澄清。')
          .setItalic(true);
      }

      body.appendParagraph(x.q.stem);

      x.q.texts.forEach(function (t, j) {
        var L = x.q.letters[j];
        var mark = (L === x.q.answer) ? '　✓ 正解' : (L === x.trap ? '　← ' + x.trapN + ' 人選這個' : '');
        var cnt = x.counts[L] ? '　［' + x.counts[L] + ' 人］' : '';
        var li = body.appendListItem(t + cnt + mark);
        if (L === x.q.answer) li.setBold(true);
      });

      body.appendParagraph('詳解').setHeading(DocumentApp.ParagraphHeading.HEADING3);
      if (x.explain) {
        String(x.explain).split(/\n+/).forEach(function (line) {
          if (line.trim()) body.appendParagraph(line.trim());
        });
      } else {
        body.appendParagraph('（詳解待補 —— 請在「' + DRAW.SHEET_BANK + '」E 欄填寫，之後這一題再考到都會自動帶入）')
            .setItalic(true).setForegroundColor('#b4302a');
      }
      body.appendHorizontalRule();
    });

    if (missing.length) {
      body.appendParagraph('尚未撰寫詳解的題目（' + missing.length + ' 題）')
          .setHeading(DocumentApp.ParagraphHeading.HEADING2);
      missing.forEach(function (x) {
        body.appendListItem('第 ' + (x.k + 1) + ' 題　' + x.q.category + '　答對率 ' + x.rate + '%');
      });
    }
    doc.saveAndClose();

    /* 盡量搬進指定資料夾，失敗不影響 */
    if (FORMCFG.FOLDER_ID) try { DriveApp.getFileById(doc.getId()).moveTo(DriveApp.getFolderById(FORMCFG.FOLDER_ID)); } catch (e) {}

    ui.alert('詳解講義已產生',
      doc.getUrl() + '\n\n'
      + stats.length + ' 題，依答對率由低到高排列。\n'
      + (missing.length
          ? '其中 ' + missing.length + ' 題還沒有詳解 —— 請在「' + DRAW.SHEET_BANK + '」E 欄填寫（標題打「詳解」）。\n'
            + '寫在題庫裡，同一題以後再考到都會自動帶入。'
          : '全部題目都有詳解 ✓'),
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('錯誤', e.message + '\n\n' + (e.stack || ''), ui.ButtonSet.OK);
  }
}

/* ═══════════════════════════════════════════════════════════════
   匯出「還沒有詳解」的題目
   把題幹、四個選項、正解、答對率整理成純文字，方便整段複製出去請人撰寫。
   ═══════════════════════════════════════════════════════════════ */
function exportForExplanationDraft() {
  var ui = SpreadsheetApp.getUi();
  try {
    var form = openTheForm();
    var A = analyzeForm_(form);
    var bank = readQuestionSheet_(DRAW.SHEET_BANK);
    var byKey = {};
    bank.forEach(function (q) { var k = normKey_(q.stem); if (k) byKey[k] = q; });

    var out = [];
    A.itemStat.map(function (st) {
      var bk = byKey[normKey_(st.q.stem)];
      return { st: st, bk: bk,
               rate: A.people.length ? Math.round(st.ok / A.people.length * 100) : 0 };
    }).filter(function (x) { return !x.bk || !x.bk.explain; })
      .sort(function (a, b) { return a.rate - b.rate; })
      .forEach(function (x) {
        var lines = [];
        lines.push('【第 ' + (x.st.k + 1) + ' 題　' + x.st.q.category
                   + '　答對率 ' + x.rate + '%　題庫第 ' + (x.bk ? x.bk.sheetRow : '?') + ' 列】');
        lines.push(x.st.q.stem);
        x.st.q.texts.forEach(function (t, j) {
          lines.push(t + (x.st.q.letters[j] === x.st.q.answer ? '　✓正解' : ''));
        });
        out.push([lines.join('\n')]);
      });

    var sh = sheetOf('待補詳解', true);
    sh.clear();
    sh.getRange(1, 1).setValue('以下題目尚未撰寫詳解。整欄複製出去即可請人撰寫，'
      + '寫好後貼回「' + DRAW.SHEET_BANK + '」對應列的 E 欄。')
      .setFontWeight('bold').setWrap(true);
    if (out.length) {
      sh.getRange(3, 1, out.length, 1).setValues(out).setWrap(true).setVerticalAlignment('top');
    } else {
      sh.getRange(3, 1).setValue('全部題目都已經有詳解 ✓');
    }
    sh.setColumnWidth(1, 760);
    sh.activate();
    ui.alert('已匯出', out.length + ' 題待補詳解，見「待補詳解」工作表。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}

/* ═══════════════════════════════════════════════════════════════
   統計工具：配對 t 檢定的雙尾 p 值
   用不完全貝他函數計算，不是近似公式。
   已對照 t 表驗證：t=2.228/df=10→p=.050、t=3.169/df=10→p=.010、
   t=2.042/df=30→p=.050、t=1.96/df=∞→p=.050
   ═══════════════════════════════════════════════════════════════ */
function gammln_(x) {
  var cof = [76.18009172947146, -86.50532032941677, 24.01409824083091,
             -1.231739572450155, 0.1208650973866179e-2, -0.5395239384953e-5];
  var y = x, tmp = x + 5.5;
  tmp -= (x + 0.5) * Math.log(tmp);
  var ser = 1.000000000190015;
  for (var j = 0; j < 6; j++) ser += cof[j] / ++y;
  return -tmp + Math.log(2.5066282746310005 * ser / x);
}
function betacf_(a, b, x) {
  var MAXIT = 200, EPS = 3e-12, FPMIN = 1e-300;
  var qab = a + b, qap = a + 1, qam = a - 1;
  var c = 1, d = 1 - qab * x / qap;
  if (Math.abs(d) < FPMIN) d = FPMIN;
  d = 1 / d; var h = d;
  for (var m = 1; m <= MAXIT; m++) {
    var m2 = 2 * m;
    var aa = m * (b - m) * x / ((qam + m2) * (a + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d; h *= d * c;
    aa = -(a + m) * (qab + m) * x / ((a + m2) * (qap + m2));
    d = 1 + aa * d; if (Math.abs(d) < FPMIN) d = FPMIN;
    c = 1 + aa / c; if (Math.abs(c) < FPMIN) c = FPMIN;
    d = 1 / d; var del = d * c; h *= del;
    if (Math.abs(del - 1) < EPS) break;
  }
  return h;
}
function betai_(a, b, x) {
  if (x <= 0) return 0;
  if (x >= 1) return 1;
  var bt = Math.exp(gammln_(a + b) - gammln_(a) - gammln_(b) + a * Math.log(x) + b * Math.log(1 - x));
  return x < (a + 1) / (a + b + 2) ? bt * betacf_(a, b, x) / a : 1 - bt * betacf_(b, a, 1 - x) / b;
}
function pTwoSided_(t, df) {
  if (!isFinite(t) || df <= 0) return NaN;
  return betai_(df / 2, 0.5, df / (df + t * t));
}
function fmtP_(p) {
  if (isNaN(p)) return '—';
  return p < 0.001 ? '< 0.001' : p.toFixed(3);
}

/* ═══════════════════════════════════════════════════════════════
   前後測比較
   ---------------------------------------------------------------
   從「歷次表單」挑兩份表單（一份前測、一份後測），
   用姓名／Email 配對同一個人，算出個人進步、各科進步，
   並做配對 t 檢定。

   只納入兩次都有作答的人 —— 沒配對到的人單獨列出，不混進統計裡，
   否則「進步」會被出席率的變化污染。
   ═══════════════════════════════════════════════════════════════ */
function runComparison_(urlA, urlB) {
  var A = analyzeForm_(FormApp.openByUrl(urlA));
  var B = analyzeForm_(FormApp.openByUrl(urlB));
  writeComparison_(A, B, formTitle_(FormApp.openByUrl(urlA)), formTitle_(FormApp.openByUrl(urlB)));
  return { a: A, b: B };
}

/** 舊入口保留，內部改走統一的挑選流程 */
function comparePrePost() { buildReport(); }

function idKey_(who) { return String(who || '').trim().toLowerCase().replace(/\s/g, ''); }

function writeComparison_(A, B, titleA, titleB) {
  var sh = sheetOf(CFG.OUT_COMPARE, true);
  sh.clear();

  var mapA = {}, mapB = {};
  A.people.forEach(function (p) { mapA[idKey_(p.who)] = p; });
  B.people.forEach(function (p) { mapB[idKey_(p.who)] = p; });

  var paired = [];
  Object.keys(mapA).forEach(function (k) { if (mapB[k]) paired.push({ a: mapA[k], b: mapB[k] }); });
  var onlyA = A.people.filter(function (p) { return !mapB[idKey_(p.who)]; });
  var onlyB = B.people.filter(function (p) { return !mapA[idKey_(p.who)]; });

  var row = 1;
  sh.getRange(row, 1).setValue('前後測比較').setFontWeight('bold').setFontSize(13); row++;
  sh.getRange(row, 1, 2, 2).setValues([['前測', titleA], ['後測', titleB]]);
  sh.getRange(row, 1, 2, 1).setFontWeight('bold'); row += 3;

  if (!paired.length) {
    sh.getRange(row, 1).setValue('沒有任何人兩次都有作答，無法配對比較。')
      .setFontColor('#b4302a').setFontWeight('bold');
    row += 2;
  } else {
    /* --- 統計摘要 --- */
    var d = paired.map(function (x) {
      return (x.b.total ? x.b.ok / x.b.total * 100 : 0) - (x.a.total ? x.a.ok / x.a.total * 100 : 0);
    });
    var n = d.length;
    var mean = d.reduce(function (s, v) { return s + v; }, 0) / n;
    var sd = n > 1 ? Math.sqrt(d.reduce(function (s, v) { return s + (v - mean) * (v - mean); }, 0) / (n - 1)) : 0;
    var se = n > 1 ? sd / Math.sqrt(n) : 0;
    var t = se > 0 ? mean / se : NaN;
    var pv = pTwoSided_(t, n - 1);
    var avgA = paired.reduce(function (s, x) { return s + (x.a.total ? x.a.ok / x.a.total * 100 : 0); }, 0) / n;
    var avgB = paired.reduce(function (s, x) { return s + (x.b.total ? x.b.ok / x.b.total * 100 : 0); }, 0) / n;

    sh.getRange(row, 1).setValue('統計摘要（只納入兩次都有作答的人）').setFontWeight('bold'); row++;
    var summary = [
      ['配對人數 n', n],
      ['前測平均', Math.round(avgA * 10) / 10 + ' 分'],
      ['後測平均', Math.round(avgB * 10) / 10 + ' 分'],
      ['平均進步', (mean >= 0 ? '+' : '') + (Math.round(mean * 10) / 10) + ' 分'],
      ['進步的標準差', Math.round(sd * 10) / 10],
      ['配對 t', isNaN(t) ? '—' : Math.round(t * 1000) / 1000],
      ['自由度 df', n - 1],
      ['雙尾 p 值', fmtP_(pv)],
      ["Cohen's dz", sd > 0 ? Math.round(mean / sd * 100) / 100 : '—']
    ];
    sh.getRange(row, 1, summary.length, 2).setValues(summary);
    sh.getRange(row, 1, summary.length, 1).setFontWeight('bold');
    sh.getRange(row + 3, 2).setBackground(mean > 0 ? '#e6f4ea' : '#fce8e6').setFontWeight('bold');
    sh.getRange(row + 7, 2).setBackground(!isNaN(pv) && pv < 0.05 ? '#e6f4ea' : '#f1f3f4').setFontWeight('bold');
    row += summary.length + 2;

    /* --- 各科別進步 --- */
    sh.getRange(row, 1).setValue('各科別答對率變化（百分點）').setFontWeight('bold'); row++;
    var catA = catRates_(A, paired.map(function (x) { return x.a; }));
    var catB = catRates_(B, paired.map(function (x) { return x.b; }));
    var cats = {};
    Object.keys(catA).forEach(function (c) { cats[c] = 1; });
    Object.keys(catB).forEach(function (c) { cats[c] = 1; });
    var catRows = Object.keys(cats).map(function (c) {
      var a = catA[c] === undefined ? null : catA[c];
      var b = catB[c] === undefined ? null : catB[c];
      return { c: c, a: a, b: b, d: (a !== null && b !== null) ? b - a : null };
    }).sort(function (x, y) {
      if (x.d === null) return 1; if (y.d === null) return -1; return x.d - y.d;
    });
    var valid = catRows.filter(function (r) { return r.d !== null; });
    var overall = valid.length ? valid.reduce(function (s2, r) { return s2 + r.d; }, 0) / valid.length : 0;

    sh.getRange(row, 1).setValue('「扣掉整體進步」那一欄是差異中的差異：把記憶效應與迴歸平均造成的'
      + '共同上升（本次 ' + (overall >= 0 ? '+' : '') + (Math.round(overall * 10) / 10)
      + ' pp）扣掉之後，這一科相對其他科的表現。').setFontColor('#5f6368');
    row++;
    sh.getRange(row, 1, 1, 5).setValues([['科別', '前測答對率', '後測答對率', '變化', '扣掉整體進步']])
      .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    row++;
    catRows.forEach(function (r) {
      var rel = r.d === null ? null : Math.round((r.d - overall) * 10) / 10;
      sh.getRange(row, 1, 1, 5).setValues([[r.c,
        r.a === null ? '—' : r.a + '%', r.b === null ? '—' : r.b + '%',
        r.d === null ? '（兩份考卷科別不同）' : (r.d >= 0 ? '+' : '') + r.d + ' pp',
        rel === null ? '—' : (rel >= 0 ? '+' : '') + rel + ' pp']]);
      if (r.d !== null) {
        sh.getRange(row, 4).setBackground(r.d > 0 ? '#e6f4ea' : r.d < 0 ? '#fce8e6' : '#f1f3f4').setFontWeight('bold');
        sh.getRange(row, 5).setBackground(rel > 0 ? '#e6f4ea' : rel < 0 ? '#fce8e6' : '#f1f3f4');
      }
      row++;
    });
    row += 1;

    /* --- 逐題前後對照（同卷才有意義）--- */
    var pairedA = paired.map(function (x) { return x.a; });
    var pairedB = paired.map(function (x) { return x.b; });
    var itemPairs = matchItems_(A, B, pairedA, pairedB);
    if (itemPairs.length) {
      sh.getRange(row, 1).setValue('逐題前後對照（' + itemPairs.length + ' 題兩份考卷都有）').setFontWeight('bold'); row++;
      sh.getRange(row, 1).setValue('同一題的難度是固定的，所以這是最乾淨的比較 —— 題目難度的變異被完全消掉了。')
        .setFontColor('#5f6368'); row++;
      sh.getRange(row, 1, 1, 5).setValues([['科別', '前測答對率', '後測答對率', '變化', '題幹']])
        .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
      row++;
      itemPairs.sort(function (m, n2) { return (m.b - m.a) - (n2.b - n2.a); });
      var ipRows = itemPairs.map(function (it) {
        return [it.cat, it.a + '%', it.b + '%', (it.b - it.a >= 0 ? '+' : '') + (it.b - it.a) + ' pp',
                String(it.stem).slice(0, 60)];
      });
      sh.getRange(row, 1, ipRows.length, 5).setValues(ipRows);
      for (var ip = 0; ip < itemPairs.length; ip++) {
        var dg = itemPairs[ip].b - itemPairs[ip].a;
        sh.getRange(row + ip, 4).setBackground(dg > 0 ? '#e6f4ea' : dg < 0 ? '#fce8e6' : '#f1f3f4')
                                 .setFontWeight('bold');
      }
      sh.getRange(row, 2, ipRows.length, 3).setHorizontalAlignment('center');
      row += ipRows.length + 2;
    }

    /* --- 每人明細 --- */
    sh.getRange(row, 1).setValue('每人進步幅度').setFontWeight('bold'); row++;
    sh.getRange(row, 1, 1, 5).setValues([['作答者', '醫院', '前測', '後測', '進步']])
      .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    row++;
    var detail = paired.map(function (x) {
      var pa = x.a.total ? Math.round(x.a.ok / x.a.total * 100) : 0;
      var pb = x.b.total ? Math.round(x.b.ok / x.b.total * 100) : 0;
      return [x.a.who, x.b.hospital || x.a.hospital, pa, pb, pb - pa];
    }).sort(function (m, n2) { return m[4] - n2[4]; });
    sh.getRange(row, 1, detail.length, 5).setValues(detail);
    for (var i = 0; i < detail.length; i++) {
      sh.getRange(row + i, 5).setBackground(detail[i][4] > 0 ? '#e6f4ea' : detail[i][4] < 0 ? '#fce8e6' : '#f1f3f4')
                              .setFontWeight('bold');
    }
    sh.getRange(row, 3, detail.length, 3).setHorizontalAlignment('center');
    row += detail.length + 2;
  }

  /* --- 未配對 --- */
  sh.getRange(row, 1).setValue('沒有配對到的人（不列入上面的統計）').setFontWeight('bold'); row++;
  sh.getRange(row, 1, 1, 2).setValues([['只有前測', '只有後測']])
    .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  row++;
  var maxU = Math.max(onlyA.length, onlyB.length, 1);
  var unpaired = [];
  for (var u = 0; u < maxU; u++) {
    unpaired.push([onlyA[u] ? onlyA[u].who : '', onlyB[u] ? onlyB[u].who : '']);
  }
  sh.getRange(row, 1, unpaired.length, 2).setValues(unpaired);

  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 150);
  for (var cw = 3; cw <= 5; cw++) sh.setColumnWidth(cw, 110);
}

/** 兩份表單裡題幹相同的題目配對，回傳各自在配對者身上的答對率 */
function matchItems_(A, B, peopleA, peopleB) {
  var idxB = {};
  B.qs.forEach(function (q, k) { var key = normKey_(q.stem); if (key) idxB[key] = k; });

  function rateOf(people, k) {
    if (!people.length) return null;
    var ok = 0;
    people.forEach(function (p) {
      var wrong = false;
      p.wrongs.forEach(function (w) { if (w.k === k) wrong = true; });
      if (!wrong) ok++;
    });
    return Math.round(ok / people.length * 100);
  }

  var out = [];
  A.qs.forEach(function (qa, ka) {
    var key = normKey_(qa.stem);
    if (!key || idxB[key] === undefined) return;
    var kb = idxB[key];
    var ra = rateOf(peopleA, ka), rb = rateOf(peopleB, kb);
    if (ra === null || rb === null) return;
    out.push({ cat: qa.category, stem: qa.stem, a: ra, b: rb });
  });
  return out;
}

/** 針對指定的一群人，算各科別答對率 */
function catRates_(A, subset) {
  var acc = {};
  subset.forEach(function (p) {
    var wrongK = {};
    p.wrongs.forEach(function (w) { wrongK[w.k] = true; });
    A.qs.forEach(function (q, k) {
      var c = acc[q.category] || (acc[q.category] = { t: 0, ok: 0 });
      c.t++; if (!wrongK[k]) c.ok++;
    });
  });
  var out = {};
  Object.keys(acc).forEach(function (c) { out[c] = acc[c].t ? Math.round(acc[c].ok / acc[c].t * 100) : 0; });
  return out;
}

/* ═══════════════════════════════════════════════════════════════
   歷次表單登記簿
   ---------------------------------------------------------------
   表單本身永遠留在 Drive，被覆蓋的只是試算表裡的連結。
   只要把連結記下來，舊表單和它的回應就一輩子找得回來。
   產生新表單時會自動呼叫這個，不需要記得手動做。
   ═══════════════════════════════════════════════════════════════ */
var HIST_HEAD = ['登記時間', '表單名稱', '類型', '選擇題數', '回應數', '表單連結', '編輯連結'];

/**
 * 確保「歷次表單」的標題列是現行的 7 欄，並把舊的 6 欄資料往右挪一格對齊。
 *
 * 為什麼需要：舊版的標題只有 6 欄（沒有「類型」）。標題列只在建立工作表時寫一次，
 * 所以加了「類型」之後，既有的表標題沒更新，新資料卻是 7 個值 ——
 * 整排往右錯開一格，「回應數」欄位顯示的其實是題數。
 */
function ensureHistoryHeader_(h) {
  var head = h.getRange(1, 1, 1, 7).getValues()[0].map(function (v) { return String(v || '').trim(); });
  var isOld = head[2] === '選擇題數';       // 第 3 欄是「選擇題數」代表是舊的 6 欄版

  if (isOld && h.getLastRow() > 1) {
    // 舊資料：[時間, 名稱, 題數, 回應數, 表單連結, 編輯連結] → 插入空白的「類型」欄
    var n = h.getLastRow() - 1;
    var body = h.getRange(2, 1, n, 6).getValues();
    var fixed = body.map(function (r) {
      return [r[0], r[1], '', r[2], r[3], r[4], r[5]];
    });
    h.getRange(2, 1, n, 7).setValues(fixed);
  }
  if (head.join('|') !== HIST_HEAD.join('|')) {
    h.getRange(1, 1, 1, 7).setValues([HIST_HEAD])
     .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    h.setFrozenRows(1);
  }
}

/** 表單標題；getTitle() 空白時退回用 Drive 的檔名 */
function formTitle_(f) {
  var t = '';
  try { t = String(f.getTitle() || '').trim(); } catch (e) {}
  if (!t) { try { t = DriveApp.getFileById(f.getId()).getName(); } catch (e2) {} }
  return t || '（無標題）';
}

function archiveFormLink_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var L = getFormLinks_();
  var kinds = Object.keys(L);
  if (!kinds.length) return null;

  var h = ss.getSheetByName(CFG.SHEET_HISTORY);
  if (!h) {
    h = ss.insertSheet(CFG.SHEET_HISTORY);
    h.getRange(1, 1, 1, 7).setValues([HIST_HEAD])
     .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    h.setFrozenRows(1);
  }
  ensureHistoryHeader_(h);

  var exist = h.getLastRow() > 1 ? h.getRange(2, 7, h.getLastRow() - 1, 1).getValues()
                                     .map(function (r) { return String(r[0]); }) : [];

  var last = null;
  kinds.forEach(function (k) {
    var edit = L[k].edit || '', pub = L[k].pub || '';
    if (!edit && !pub) return;
    if (edit && exist.indexOf(edit) >= 0) return;

    var title = '（開不起來）', nQ = '', nR = '';
    try {
      var f = FormApp.openByUrl(edit || pub);
      title = formTitle_(f);
      nQ = f.getItems(FormApp.ItemType.MULTIPLE_CHOICE).length;
      nR = f.getResponses().length;
    } catch (e) {}

    var kind = k !== 'single' ? k : (/前測/.test(title) ? '前測' : /後測/.test(title) ? '後測' : '');
    h.appendRow([Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm'),
                 title, kind, nQ, nR, pub, edit]);
    last = { title: title, responses: nR };
  });
  h.setColumnWidth(2, 260); h.setColumnWidth(6, 300); h.setColumnWidth(7, 300);
  return last;
}

/* ═══════════════════════════════════════════════════════════════
   重新整理「歷次表單」
   ---------------------------------------------------------------
   登記時寫入的回應數是「當下的快照」—— 表單剛建好時登記，一定是 0，
   之後有人作答也不會自己更新。這個函式重新開啟每一份表單，
   把表單名稱、類型、題數、回應數全部更新成現況。
   順便修好舊版 6 欄標題造成的欄位錯位。
   ═══════════════════════════════════════════════════════════════ */
function refreshFormHistory() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var h = ss.getSheetByName(CFG.SHEET_HISTORY);
    if (!h || h.getLastRow() < 2) throw new Error('「' + CFG.SHEET_HISTORY + '」還沒有任何登記。');

    ensureHistoryHeader_(h);

    var n = h.getLastRow() - 1;
    var rows = h.getRange(2, 1, n, 7).getValues();
    var okCount = 0, failCount = 0, totalR = 0;

    var updated = rows.map(function (r) {
      var edit = String(r[6] || ''), pub = String(r[5] || '');
      if (!edit && !pub) return r;
      try {
        var f = FormApp.openByUrl(edit || pub);
        var title = formTitle_(f);
        var nQ = f.getItems(FormApp.ItemType.MULTIPLE_CHOICE).length;
        var nR = f.getResponses().length;
        var kind = String(r[2] || '').trim() || (/前測/.test(title) ? '前測' : /後測/.test(title) ? '後測' : '');
        okCount++; totalR += nR;
        return [r[0], title, kind, nQ, nR, pub || f.getPublishedUrl(), edit || f.getEditUrl()];
      } catch (e) {
        failCount++;
        return [r[0], String(r[1] || '') || '（開不起來）', r[2], r[3], r[4], pub, edit];
      }
    });

    h.getRange(2, 1, n, 7).setValues(updated);
    h.getRange(2, 4, n, 2).setHorizontalAlignment('center');
    // 有回應的列標綠，方便一眼找到有資料的表單
    for (var i = 0; i < updated.length; i++) {
      var nr = updated[i][4];
      h.getRange(i + 2, 5).setBackground(nr > 0 ? '#e6f4ea' : '#ffffff')
                          .setFontWeight(nr > 0 ? 'bold' : 'normal');
    }
    h.setColumnWidth(2, 260); h.setColumnWidth(6, 300); h.setColumnWidth(7, 300);
    h.activate();

    ui.alert('已重新整理',
      n + ' 份表單：' + okCount + ' 份讀取成功、' + failCount + ' 份開不起來。\n\n'
      + '總回應數 ' + totalR + ' 筆。\n\n'
      + '「回應數」欄現在是即時的；有回應的列已標綠。', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}

/** 手動登記目前的表單（平常不用，自動的就夠了） */
function archiveFormLink() {
  var ui = SpreadsheetApp.getUi();
  var r = archiveFormLink_();
  ui.alert(r ? '已登記' : '沒有東西可登記',
    r ? '「' + r.title + '」（' + r.responses + ' 筆回應）已寫入「' + CFG.SHEET_HISTORY + '」。'
      : '「' + CFG.SHEET_LINKS + '」裡沒有連結。', ui.ButtonSet.OK);
}

/* ═══════════════════════════════════════════════════════════════
   封存成績報告
   ---------------------------------------------------------------
   四張報告表每次產生都會被覆蓋。這裡把它們整份複製到一個
   獨立的新試算表存檔，本檔案的分頁不會愈積愈多。
   ═══════════════════════════════════════════════════════════════ */
function archiveReports() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var names = [CFG.OUT_PERSON, CFG.OUT_WRONG, CFG.OUT_ITEM, CFG.OUT_HOSPITAL, CFG.OUT_BATCH, CFG.OUT_COMPARE];
    var have = names.filter(function (n) { var sh = ss.getSheetByName(n); return sh && sh.getLastRow() > 1; });
    if (!have.length) throw new Error('目前沒有報告可以封存。請先跑「② 產生完整測驗報告（從表單）」。');

    var tag = Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd_HHmm');
    var target = SpreadsheetApp.create('兒科題庫測驗成績_' + tag);
    var defaultSheet = target.getSheets()[0];

    have.forEach(function (n) {
      var copied = ss.getSheetByName(n).copyTo(target);
      copied.setName(n);
    });
    if (target.getSheets().length > have.length) target.deleteSheet(defaultSheet);

    // 順便把來源表單一起記進去
    var info = null;
    try { info = archiveFormLink_(); } catch (e) {}

    var meta = target.insertSheet('說明', 0);
    meta.getRange(1, 1, 5, 2).setValues([
      ['封存時間', Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm')],
      ['來源表單', info ? info.title : '（不明）'],
      ['作答人數', info ? info.responses : ''],
      ['原始試算表', ss.getUrl()],
      ['', '']
    ]);
    meta.getRange(1, 1, 4, 1).setFontWeight('bold');
    meta.autoResizeColumns(1, 2);

    ui.alert('成績已封存',
      '已建立獨立試算表：\n兒科題庫測驗成績_' + tag + '\n\n'
      + target.getUrl() + '\n\n'
      + '包含 ' + have.join('、') + '。\n'
      + '現在可以安全地產生新表單了。', ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}

/* ═══════════════════════════════════════════════════════════════
   從表單直接產生報告
   ---------------------------------------------------------------
   Google 表單本身就存著題幹、四個選項、以及標記好的正解，
   所以「選出題目」被覆蓋掉之後，成績還是算得回來。
   科別和年分靠題幹回頭比對「題庫」補上，比對不到就標示出來。

   這個方式比 buildFullReport 穩：它不依賴「選出題目」和表單是同一批題目。
   ═══════════════════════════════════════════════════════════════ */

/** 題幹正規化：去掉開頭原始題號與所有空白，取前 12 字當比對鍵 */
function normKey_(t) {
  return String(t == null ? '' : t)
    .replace(/^\s*第?\s*\d+\s*題?\s*[:：.、．]?\s*/, '')
    .replace(/^\s*\d+\s*[.、．]\s*/, '')
    .replace(/\s/g, '')
    .slice(0, 12);
}

/**
 * 把一份表單解析成可分析的結構。
 * 回傳 { qs, people, itemStat, hospItem, nameItem, noCat, noAns }
 * buildReportFromForm 和 comparePrePost 共用這一段，兩邊的計分方式保證一致。
 */
function analyzeForm_(form) {
  /* 題庫索引，用來補回科別與年分 */
  var bankIdx = {};
  try {
    readQuestionSheet_(DRAW.SHEET_BANK).forEach(function (q) {
      var k = normKey_(q.stem);
      if (k && !bankIdx[k]) bankIdx[k] = q;
    });
  } catch (e) { /* 沒有題庫也還是能算分，只是沒有科別 */ }

  var all = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
  var hospItem = null, qItems = [];
  all.forEach(function (it) {
    if (/醫院/.test(it.getTitle())) hospItem = it; else qItems.push(it);
  });
  var nameItem = findNameItem(form);

  var noCat = 0, noAns = 0;
  var qs = qItems.map(function (it, k) {
    var mc = it.asMultipleChoiceItem();
    var choices = mc.getChoices();
    var letters = choices.map(function (c) { return normalizeLetter(String(c.getValue()).charAt(0)) || '?'; });

    var ansLetter = null;
    choices.forEach(function (c, j) { if (c.isCorrectAnswer()) ansLetter = letters[j]; });

    var title = it.getTitle().replace(/^第\s*\d+\s*題\s*[:：]?\s*/, '');
    var bk = bankIdx[normKey_(title)];
    if (!ansLetter && bk) ansLetter = bk.answer;
    if (!ansLetter) noAns++;
    if (!bk) noCat++;

    return {
      k: k, stem: title, answer: ansLetter,
      category: bk ? bk.category : '（對不到題庫）',
      year: bk ? bk.year : '',
      item: it, choices: choices, letters: letters,
      texts: choices.map(function (c) { return String(c.getValue()); })
    };
  });

  var people = [];
  var itemStat = qs.map(function (q, k) {
    return { k: k, q: q, counts: { A: 0, B: 0, C: 0, D: 0, E: 0 }, blank: 0, ok: 0, n: 0 };
  });

  form.getResponses().forEach(function (resp) {
    var hospital = '（未填）';
    if (hospItem) { var hr = resp.getResponseForItem(hospItem); if (hr) hospital = String(hr.getResponse()); }

    var who = '';
    if (nameItem) { var nr = resp.getResponseForItem(nameItem); if (nr) who = String(nr.getResponse()).trim(); }
    if (!who) { try { who = resp.getRespondentEmail() || ''; } catch (e) {} }
    if (!who) who = '（匿名 ' + Utilities.formatDate(resp.getTimestamp(), 'GMT+8', 'MM/dd HH:mm') + '）';

    var rec = { who: who, hospital: hospital, ts: resp.getTimestamp(), ok: 0, total: 0, wrongs: [] };

    qs.forEach(function (q, k) {
      rec.total++; itemStat[k].n++;
      var ir = resp.getResponseForItem(q.item);
      var picked = null;
      if (ir) {
        var v = String(ir.getResponse());
        var idx = q.texts.indexOf(v);
        picked = idx >= 0 ? q.letters[idx] : normalizeLetter(v.charAt(0));
      }
      if (!picked) { itemStat[k].blank++; rec.wrongs.push({ k: k, q: q, picked: null }); return; }
      if (itemStat[k].counts[picked] !== undefined) itemStat[k].counts[picked]++;
      if (picked === q.answer) { rec.ok++; itemStat[k].ok++; }
      else rec.wrongs.push({ k: k, q: q, picked: picked });
    });

    people.push(rec);
  });

  return { qs: qs, people: people, itemStat: itemStat,
           hospItem: hospItem, nameItem: nameItem, noCat: noCat, noAns: noAns };
}

/* ═══════════════════════════════════════════════════════════════
   統一報表入口
   ---------------------------------------------------------------
   一個選單項目做兩件事：
     輸入一個列號        → 產生那一份表單的完整報表（五張）
     輸入兩個列號（前,後）→ 產生前後測比較

   清單直接取自「歷次表單」，並會先把「表單連結」裡目前這份補登進去，
   所以不管是現在的還是以前的表單，都在同一份清單裡挑。
   ═══════════════════════════════════════════════════════════════ */
function buildReport() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    /* 先把目前「表單連結」裡的表單補登，清單才完整 */
    try { archiveFormLink_(); } catch (e) {}

    var h = ss.getSheetByName(CFG.SHEET_HISTORY);
    if (!h || h.getLastRow() < 2) {
      throw new Error('還沒有任何表單可以統計。\n\n'
        + '請先「一次產生前測＋後測（同卷）」，或用「進階 → 登記目前表單」把既有的表單補登進來。');
    }
    ensureHistoryHeader_(h);

    var rows = h.getRange(2, 1, h.getLastRow() - 1, 7).getValues();
    var usable = [];
    rows.forEach(function (r, i) {
      var url = String(r[6] || r[5] || '');
      if (url) usable.push({ row: i + 2, kind: String(r[2] || ''), title: String(r[1] || '（無標題）'),
                             n: r[4], url: url });
    });
    if (!usable.length) throw new Error('「' + CFG.SHEET_HISTORY + '」裡沒有可用的連結。');

    /* 只有一份就直接做，不必問 */
    if (usable.length === 1) { singleReport_(usable[0].url); return; }

    var listing = usable.map(function (u) {
      return '　' + u.row + '　' + (u.kind || '—') + '　'
             + u.title.slice(0, 26) + '　回應 ' + (u.n === '' ? '?' : u.n) + ' 筆';
    }).join('\n');

    var resp = ui.prompt('產生報表',
      listing + '\n\n'
      + '─────────────────────\n'
      + '輸入 一個 列號　→　那一份的完整報表\n'
      + '輸入 兩個 列號　→　前後測比較（前測,後測）\n\n'
      + '例如：  6      或      6,7\n\n'
      + '（回應數顯示 ? 或不是最新的，先跑「重新整理歷次表單」）',
      ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    var nums = String(resp.getResponseText()).split(/[^0-9]+/).filter(String).map(Number);
    var find = function (rowNo) {
      var hit = usable.filter(function (u) { return u.row === rowNo; })[0];
      if (!hit) throw new Error('清單裡沒有第 ' + rowNo + ' 列。');
      return hit;
    };

    if (nums.length === 1) {
      singleReport_(find(nums[0]).url);
    } else if (nums.length >= 2) {
      var a = find(nums[0]), b = find(nums[1]);
      if (a.url === b.url) throw new Error('前測和後測不能是同一份表單。');
      runComparison_(a.url, b.url);
      ss.getSheetByName(CFG.OUT_COMPARE).activate();
      ui.alert('前後測比較完成',
        '前測：' + a.title + '\n後測：' + b.title + '\n\n'
        + '結果在「' + CFG.OUT_COMPARE + '」，含統計摘要、各科變化、逐題前後對照、每人進步幅度。',
        ui.ButtonSet.OK);
    } else {
      throw new Error('沒讀到列號。請輸入數字，例如 6 或 6,7。');
    }

  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}

function singleReport_(url) {
  var ui = SpreadsheetApp.getUi();
  var form = FormApp.openByUrl(url);
  var A = analyzeForm_(form);
  if (!A.people.length) {
    ui.alert('這份表單還沒有人作答', formTitle_(form) + '\n\n換一份有回應的再試。', ui.ButtonSet.OK);
    return;
  }
  writeAllReports_(form, A);
}

/** 產生四＋一張報表。buildReportFromForm 與 buildReportFromHistory 共用。 */
function writeAllReports_(form, A) {
  var batches = assignBatches_(A.people);
  writePersonSheet(A.people);
  writeWrongSheet(A.people);
  writeItemSheet(A.itemStat);
  writeHospitalSheet(A.people, A.qs.map(function (q) { return { q: q }; }));
  writeBatchSheet(batches, A.qs);

  var warn = '';
  if (batches.length > 1) {
    warn += '\n\n★ 偵測到 ' + batches.length + ' 個作答批次（相隔超過 '
          + CFG.BATCH_GAP_DAYS + ' 天就切一批）。「' + CFG.OUT_BATCH + '」有各批的平均分與最弱科別。'
          + '「' + CFG.OUT_PERSON + '」的整體平均把不同批混在一起算，看各批請用「' + CFG.OUT_BATCH + '」。';
  }
  if (A.noCat) warn += '\n⚠ 有 ' + A.noCat + ' 題在「' + DRAW.SHEET_BANK + '」比對不到，這些題目沒有科別（成績仍然正確）。';
  if (A.noAns) warn += '\n⚠ 有 ' + A.noAns + ' 題連表單裡也沒有標記正解，這些題所有人都會被算答錯。';
  if (!A.hospItem) warn += '\n⚠ 表單裡沒有醫院題目，醫院統計會是空的。';
  if (!A.nameItem) warn += '\n（表單沒有姓名欄，報表用 Email 當作答者識別。）';

  sheetOf(CFG.OUT_PERSON, false).activate();
  SpreadsheetApp.getUi().alert('報告完成（資料直接取自表單）',
    formTitle_(form) + '\n\n'
    + A.people.length + ' 人作答、' + A.qs.length + ' 題。\n\n'
    + '已產生：' + CFG.OUT_PERSON + '、' + CFG.OUT_WRONG + '、' + CFG.OUT_ITEM + '、'
    + CFG.OUT_HOSPITAL + '、' + CFG.OUT_BATCH
    + warn, SpreadsheetApp.getUi().ButtonSet.OK);
}

function buildReportFromForm() {
  var ui = SpreadsheetApp.getUi();
  try {
    var form = openTheForm();
    var A = analyzeForm_(form);
    if (!A.people.length) {
      ui.alert('這份表單還沒有人作答',
        formTitle_(form) + '\n\n'
        + '如果你要統計的是「以前那一份」表單，請改用\n'
        + '「② 從歷次表單挑一份產生報告」。', ui.ButtonSet.OK);
      return;
    }
    writeAllReports_(form, A);
  } catch (e) {
    ui.alert('錯誤', e.message + '\n\n' + (e.stack || ''), ui.ButtonSet.OK);
  }
}

/* ═══════════════════════════════════════════════════════════════
   從「歷次表單」挑一份產生報告
   ---------------------------------------------------------------
   「表單連結」只放目前這一份。要統計以前發過的任何一份表單，
   從這裡挑就好 —— 舊表單和它的回應永遠留在 Drive，不會消失。
   ═══════════════════════════════════════════════════════════════ */
function buildReportFromHistory() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var h = ss.getSheetByName(CFG.SHEET_HISTORY);
    if (!h || h.getLastRow() < 2) {
      throw new Error('「' + CFG.SHEET_HISTORY + '」還沒有任何登記。\n\n'
        + '可以先執行「登記目前表單到歷次表單」把現在這份補登進去。');
    }
    ensureHistoryHeader_(h);

    var rows = h.getRange(2, 1, h.getLastRow() - 1, 7).getValues();
    var listing = rows.map(function (r, i) {
      var nR = r[4] === '' ? '?' : r[4];
      return '第 ' + (i + 2) + ' 列　' + (r[2] || '—') + '　'
             + String(r[1] || '（無標題）').slice(0, 28) + '　回應 ' + nR + ' 筆';
    }).join('\n');

    var resp = ui.prompt('要統計哪一份表單？',
      listing + '\n\n（回應數如果顯示 ? 或不是最新的，先跑「重新整理歷次表單」）\n\n請輸入列號：',
      ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;

    var idx = parseInt(resp.getResponseText(), 10) - 2;
    if (!rows[idx]) throw new Error('列號不在範圍內。');

    var url = String(rows[idx][6] || rows[idx][5] || '');
    if (!url) throw new Error('那一列沒有可用的連結。');

    var form = FormApp.openByUrl(url);
    var A = analyzeForm_(form);
    if (!A.people.length) { ui.alert('這份表單還沒有人作答。', ui.ButtonSet.OK); return; }
    writeAllReports_(form, A);

  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}


/* ═══════════════════════════════════════════════════════════════
   雲端硬碟與表單存取診斷
   ---------------------------------------------------------------
   「No item with the given ID could be found」這個錯誤可能來自三個地方，
   這裡逐一測試，直接告訴你是哪一個。
   ═══════════════════════════════════════════════════════════════ */
function checkDriveAccess() {
  var ui = SpreadsheetApp.getUi();
  var lines = [];

  lines.push('【1】表單存放資料夾  FORMCFG.FOLDER_ID');
  lines.push('　　' + FORMCFG.FOLDER_ID);
  try {
    var f = DriveApp.getFolderById(FORMCFG.FOLDER_ID);
    lines.push('　　✓ 可存取：' + f.getName());
    lines.push('　　　' + f.getUrl());
  } catch (e) {
    lines.push('　　✗ 存取失敗：' + e.message);
    lines.push('　　→ 這就是錯誤來源。資料夾被刪除、改權限，或搬到你沒權限的位置。');
    lines.push('　　→ v7 起搬檔失敗不會中斷建表單，但建議還是換一個有效的資料夾 ID。');
  }

  lines.push('');
  lines.push('【2】表單連結工作表');
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_LINKS);
  if (!sh) {
    lines.push('　　✗ 找不到「' + CFG.SHEET_LINKS + '」工作表');
  } else {
    var pub = String(sh.getRange('B1').getValue() || '');
    var edit = String(sh.getRange('B2').getValue() || '');
    lines.push('　　B1 表單連結：' + (pub ? pub.slice(0, 60) + '…' : '（空白）'));
    lines.push('　　B2 編輯連結：' + (edit ? edit.slice(0, 60) + '…' : '（空白）'));
    lines.push('');
    lines.push('【3】用編輯連結開啟表單');
    try {
      var form = openTheForm();
      var mc = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
      lines.push('　　✓ 開得起來：' + form.getTitle());
      lines.push('　　　選擇題 ' + mc.length + ' 題、回應 ' + form.getResponses().length + ' 筆');
    } catch (e2) {
      lines.push('　　✗ 開啟失敗：' + e2.message);
      lines.push('　　→ FormApp.openByUrl 需要「編輯連結」（結尾是 /edit），不是 /viewform。');
    }
  }

  lines.push('');
  lines.push('【4】建立測試檔案的權限');
  try {
    var probe = DriveApp.createFile('__權限測試__.txt', 'ok');
    lines.push('　　✓ 可以在雲端硬碟建立檔案');
    probe.setTrashed(true);
    lines.push('　　（測試檔已丟到垃圾桶）');
  } catch (e3) {
    lines.push('　　✗ 無法建立檔案：' + e3.message);
  }

  ui.alert('雲端硬碟診斷', lines.join('\n'), ui.ButtonSet.OK);
}

/* ═══════════════════════════════════════════════════════════════
   ① 題庫健檢 — 在做成表單之前先抓問題
   ═══════════════════════════════════════════════════════════════ */
function checkQuestionBank() {
  var ui = SpreadsheetApp.getUi();
  try {
    var targets = [DRAW.SHEET_BANK, CFG.SHEET_SELECTED];
    var problems = [], counts = {};

    targets.forEach(function (name) {
      var qs;
      try { qs = readQuestionSheet_(name); } catch (e) { return; }   // 表不存在就跳過
      counts[name] = qs.length;

      qs.forEach(function (q) {
        if (!q.parse.ok) {
          problems.push([name, q.sheetRow, q.row, '選項拆不開', q.parse.reason,
                         String(q.raw).replace(/\n/g, ' ⏎ ').slice(0, 90)]);
        } else if (q.parse.loose) {
          problems.push([name, q.sheetRow, q.row, '題幹可能被切壞', q.parse.reason,
                         q.stem.slice(0, 90)]);
        }
        if (!q.answer) {
          problems.push([name, q.sheetRow, q.row, '正解讀不出來',
            'C 欄是「' + q.rawAnswer + '」，必須是 A/B/C/D/E。這題在表單裡會沒有正解，全班算答錯。',
            q.stem.slice(0, 90)]);
        } else if (q.parse.ok && !q.parse.choices.some(function (c) { return c.letter === q.answer; })) {
          problems.push([name, q.sheetRow, q.row, '正解對不到選項',
            '正解是 ' + q.answer + '，但這題只有 ' + q.parse.choices.length + ' 個選項。這題在表單裡會沒有正解。',
            q.stem.slice(0, 90)]);
        }
      });

      // 重複題偵測
      // 比對「題幹＋四個選項」的完整內容（去掉開頭原始題號、去掉所有空白）。
      // 只有整題一字不差才算重複 —— 題幹相同但選項不同是不同的題目，
      // 例如多份考卷都有「關於兒童癌症，下列敘述何者最不適當？」，那是正常的，不該被報。
      var seen = {};
      qs.forEach(function (q) {
        var key = String(q.raw).replace(/^\s*\d+\s*[.、．]\s*/, '').replace(/\s/g, '');
        if (!key) return;
        if (seen[key]) problems.push([name, q.sheetRow, q.row, '完全重複',
          '和第 ' + seen[key] + ' 題（第 ' + seen[key + '_r'] + ' 列）題幹與四個選項完全相同',
          q.stem.slice(0, 90)]);
        else { seen[key] = q.row; seen[key + '_r'] = q.sheetRow; }
      });
    });

    var sh = sheetOf('題庫健檢', true);
    sh.clear();
    var head = ['工作表', '實際列號', '第幾題', '問題類型', '說明', '內容'];
    sh.getRange(1, 1, 1, head.length).setValues([head])
      .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');

    if (problems.length) {
      sh.getRange(2, 1, problems.length, head.length).setValues(problems);
      sh.getRange(2, 4, problems.length, 1).setBackground('#fce8e6').setFontWeight('bold');
      sh.getRange(2, 5, problems.length, 2).setWrap(true).setVerticalAlignment('top');
    } else {
      sh.getRange(2, 1).setValue('全部通過檢查 ✓');
    }
    sh.setColumnWidth(1, 110); sh.setColumnWidth(2, 80); sh.setColumnWidth(3, 70);
    sh.setColumnWidth(4, 130); sh.setColumnWidth(5, 340); sh.setColumnWidth(6, 420);
    sh.setFrozenRows(1);
    sh.activate();

    var scanned = Object.keys(counts).map(function (k) { return k + ' ' + counts[k] + ' 題'; }).join('、');
    ui.alert('題庫健檢完成',
      '掃描範圍：' + scanned + '\n\n'
      + '發現 ' + problems.length + ' 個問題。\n\n'
      + (problems.length
          ? '請看「題庫健檢」工作表，第 2 欄是實際列號，可以直接跳過去修。\n\n'
            + '★ 修要修在「' + DRAW.SHEET_BANK + '」，不要只修「' + CFG.SHEET_SELECTED + '」——'
            + '後者每次重新抽題都會被覆蓋掉。'
          : '沒有問題，可以放心產生表單。'),
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}

/* ═══════════════════════════════════════════════════════════════
   ③ 題目對照表
   報告要把「表單的第 N 題」對回「選出題目的第 N 列」。
   原本靠位置對應，只要產生表單時有任何一題被 catch 掉（continue），
   後面所有題目的科別就會整排錯位，而且不會有人發現。
   這裡改成用表單標題裡的「第 N 題」回推，並額外存一份 itemId 對照。
   ═══════════════════════════════════════════════════════════════ */
function writeItemMap() {
  var ui = SpreadsheetApp.getUi();
  try {
    var form = openTheForm();
    var qs = readSelected();
    var pairs = mapItemsToQuestions(form, qs);

    var sh = sheetOf(CFG.SHEET_MAP, true);
    sh.clear();
    sh.getRange(1, 1, 1, 4).setValues([['表單題序', 'itemId', '對應選出題目列', '科別']])
      .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
    if (pairs.length) {
      sh.getRange(2, 1, pairs.length, 4).setValues(pairs.map(function (p, i) {
        return [i + 1, p.item.getId(), p.q ? p.q.row : '對不到', p.q ? p.q.category : ''];
      }));
    }
    sh.setFrozenRows(1); sh.autoResizeColumns(1, 4);
    ui.alert('完成', '已寫入「' + CFG.SHEET_MAP + '」，共 ' + pairs.length + ' 題。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}

function getFormLinks_() {
  var sh = sheetOf(CFG.SHEET_LINKS, false);
  var out = {};
  if (!sh || sh.getLastRow() < 1) return out;
  var vals = sh.getRange(1, 1, sh.getLastRow(), 2).getValues();
  vals.forEach(function (r) {
    var label = String(r[0] || '').trim();
    var url = String(r[1] || '').trim();
    if (!url) return;
    var kind = /前測/.test(label) ? '前測' : /後測/.test(label) ? '後測' : 'single';
    var isEdit = /編輯/.test(label) || /\/edit\b/.test(url);
    var e = out[kind] || (out[kind] = {});
    if (isEdit) e.edit = url; else e.pub = url;
  });
  return out;
}

/**
 * 開啟表單。
 * 「表單連結」只有一份時直接開；前測後測都有時會問你要用哪一份。
 */
function openTheForm() {
  var L = getFormLinks_();
  var cands = [];
  ['前測', '後測', 'single'].forEach(function (k) {
    if (L[k] && (L[k].edit || L[k].pub)) cands.push({ kind: k, url: L[k].edit || L[k].pub });
  });
  if (!cands.length) {
    throw new Error('「' + CFG.SHEET_LINKS + '」裡沒有可用的連結，請先產生表單。');
  }

  var pick = cands[0];
  if (cands.length > 1) {
    var ui = SpreadsheetApp.getUi();
    var r = ui.alert('要用哪一份表單？',
      cands.map(function (c) { return '· ' + (c.kind === 'single' ? '（未分類）' : c.kind); }).join('\n')
      + '\n\n【是】前測　　【否】後測', ui.ButtonSet.YES_NO_CANCEL);
    if (r === ui.Button.CANCEL) throw new Error('已取消。');
    var want = r === ui.Button.YES ? '前測' : '後測';
    pick = cands.filter(function (c) { return c.kind === want; })[0] || cands[0];
  }

  try { return FormApp.openByUrl(pick.url); }
  catch (e) {
    throw new Error('無法開啟表單（' + pick.kind + '）：' + e.message
      + '\n\nFormApp 需要「編輯連結」（結尾是 /edit），不是 /viewform。');
  }
}

/** 表單題目 → 選出題目 的對應。優先讀標題裡的「第 N 題」，失敗才用位置。 */
function mapItemsToQuestions(form, qs) {
  var byRow = {};
  qs.forEach(function (q) { byRow[q.row] = q; });

  var mc = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE).filter(function (it) {
    return !/醫院/.test(it.getTitle());          // 排除「請選擇您的醫院」
  });

  var usedTitleMap = true;
  var pairs = mc.map(function (it, idx) {
    var m = it.getTitle().match(/^第\s*(\d+)\s*題/);
    if (m && byRow[Number(m[1])]) return { item: it, q: byRow[Number(m[1])], via: '標題' };
    usedTitleMap = false;
    return { item: it, q: qs[idx] || null, via: '位置' };
  });

  // ★ 防呆：確認表單裡的題目和「選出題目」現在的內容是同一批。
  //   如果中間重抽過題目，第 N 列已經換人了，報表會安靜地算出一份全錯的成績。
  //   這裡比對題幹前 10 個字（去空白）：足以分辨是不是同一題，
  //   又能容忍事後在題幹中段補字。
  var mismatch = [];
  pairs.forEach(function (p, i) {
    if (!p.q) { mismatch.push(i + 1); return; }
    var inForm  = p.item.getTitle().replace(/^第\s*\d+\s*題\s*[:：]?\s*/, '').replace(/\s/g, '').slice(0, 10);
    var inSheet = String(p.q.stem || '').replace(/\s/g, '').slice(0, 10);
    if (inForm && inSheet && inForm !== inSheet) mismatch.push(i + 1);
  });
  pairs.usedTitleMap = usedTitleMap;
  pairs.mismatch = mismatch;
  return pairs;
}

function findHospitalItem(form) {
  var all = form.getItems(FormApp.ItemType.MULTIPLE_CHOICE);
  for (var i = 0; i < all.length; i++) if (/醫院/.test(all[i].getTitle())) return all[i];
  return null;
}
function findNameItem(form) {
  var t = form.getItems(FormApp.ItemType.TEXT);
  for (var i = 0; i < t.length; i++) if (/姓名|名字/.test(t[i].getTitle())) return t[i];
  return null;
}

/* ═══════════════════════════════════════════════════════════════
   ② 完整測驗報告
   ═══════════════════════════════════════════════════════════════ */
function buildFullReport() {
  var ui = SpreadsheetApp.getUi();
  try {
    var form = openTheForm();
    var qs   = readSelected();
    var pairs = mapItemsToQuestions(form, qs);
    var hospItem = findHospitalItem(form);
    var nameItem = findNameItem(form);
    var responses = form.getResponses();

    if (!responses.length) { ui.alert('還沒有人作答', '這份表單目前 0 筆回應。', ui.ButtonSet.OK); return; }

    /* ---------- 把每一份回應攤平成一筆一筆的作答紀錄 ---------- */
    var people = [];      // {who, hospital, ts, ok, total, wrongs:[{k,q,picked}]}
    var itemStat = pairs.map(function (p, k) {
      return { k: k, q: p.q, counts: { A: 0, B: 0, C: 0, D: 0, E: 0 }, blank: 0, ok: 0, n: 0 };
    });

    responses.forEach(function (resp) {
      var hospital = '（未填）';
      if (hospItem) {
        var hr = resp.getResponseForItem(hospItem);
        // ★ 原版寫的是 resp.getResponseForItem(...) 直接當字串用。
        //   那回傳的是 ItemResponse 物件，拿去當 key 會全部擠進同一格，
        //   醫院分組因此完全失效。一定要再 .getResponse()。
        if (hr) hospital = String(hr.getResponse());
      }

      var who = '';
      if (nameItem) { var nr = resp.getResponseForItem(nameItem); if (nr) who = String(nr.getResponse()).trim(); }
      if (!who) { try { who = resp.getRespondentEmail() || ''; } catch (e) {} }
      if (!who) who = '（匿名 ' + Utilities.formatDate(resp.getTimestamp(), 'GMT+8', 'MM/dd HH:mm') + '）';

      var rec = { who: who, hospital: hospital, ts: resp.getTimestamp(), ok: 0, total: 0, wrongs: [] };

      pairs.forEach(function (p, k) {
        if (!p.q) return;
        rec.total++;
        itemStat[k].n++;
        var ir = resp.getResponseForItem(p.item);
        var picked = ir ? normalizeLetter(String(ir.getResponse()).charAt(0)) : null;

        if (!picked) { itemStat[k].blank++; rec.wrongs.push({ k: k, q: p.q, picked: null }); return; }
        itemStat[k].counts[picked]++;
        // ★ 正確與否一律以「選出題目」C 欄為準，不依賴表單自己的評分，
        //   因為表單的正解可能因為 C 欄格式問題而根本沒被設定。
        if (picked === p.q.answer) { rec.ok++; itemStat[k].ok++; }
        else rec.wrongs.push({ k: k, q: p.q, picked: picked });
      });

      people.push(rec);
    });

    writePersonSheet(people);
    writeWrongSheet(people);
    writeItemSheet(itemStat);
    writeHospitalSheet(people, pairs);

    if (pairs.mismatch && pairs.mismatch.length > pairs.length * 0.2) {
      ui.alert('這份報告不能做 —— 題目對不上',
        '表單裡的題目和目前「' + CFG.SHEET_SELECTED + '」的內容不一致（' + pairs.mismatch.length +
        ' / ' + pairs.length + ' 題對不上）。\n\n' +
        '最常見的原因：這份表單發出去之後，又按過「抽取30題」重新抽題，' +
        '把當時的考卷覆蓋掉了。\n\n' +
        '沒有原始考卷就算不出正確成績。請到 Drive 找當時封存的「考卷_日期」工作表，' +
        '把內容貼回「' + CFG.SHEET_SELECTED + '」後再跑一次。',
        ui.ButtonSet.OK);
      return;
    }

    var warn = '';
    if (pairs.mismatch && pairs.mismatch.length) warn += '\n⚠ 有 ' + pairs.mismatch.length + ' 題的題幹和表單對不上（第 ' + pairs.mismatch.slice(0, 8).join(', ') + ' 題），這幾題的統計不可信。';
    if (!pairs.usedTitleMap) warn += '\n⚠ 有題目的標題不是「第 N 題:」開頭，只能用位置對應，科別可能錯位。建議重新產生表單。';
    if (!hospItem) warn += '\n⚠ 表單裡找不到醫院題目，醫院統計會是空的。';
    var noAns = qs.filter(function (q) { return !q.answer; }).length;
    if (noAns) warn += '\n⚠ 有 ' + noAns + ' 題的正解欄讀不出 A/B/C/D/E，這些題目所有人都會被算答錯。請先跑「題庫健檢」。';

    sheetOf(CFG.OUT_PERSON, false).activate();
    ui.alert('報告完成',
      '共 ' + people.length + ' 人作答、' + pairs.length + ' 題。\n\n已產生四張工作表：\n' +
      '· ' + CFG.OUT_PERSON + '　每人分數與排名\n' +
      '· ' + CFG.OUT_WRONG  + '　每人錯了哪幾題、選了什麼\n' +
      '· ' + CFG.OUT_ITEM   + '　每題全班答對率與錯誤選項分布\n' +
      '· ' + CFG.OUT_HOSPITAL + '　醫院 × 科別' + warn,
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('錯誤', e.message + '\n\n' + (e.stack || ''), ui.ButtonSet.OK);
  }
}

/* ---------- 個人成績 ---------- */
function writePersonSheet(people) {
  var sorted = people.slice().sort(function (a, b) { return (b.ok / b.total) - (a.ok / a.total); });
  var sh = sheetOf(CFG.OUT_PERSON, true); sh.clear();

  var head = ['排名', '作答者', '醫院', '分數', '答對', '總題數', '答錯', '作答時間', '批次'];
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');

  var rows = sorted.map(function (p, i) {
    return [i + 1, p.who, p.hospital,
            p.total ? Math.round(p.ok / p.total * 100) : 0,
            p.ok, p.total, p.total - p.ok,
            Utilities.formatDate(p.ts, 'GMT+8', 'yyyy/MM/dd HH:mm'),
            p.batch || ''];
  });
  if (rows.length) {
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    for (var i = 0; i < rows.length; i++) {
      var pct = rows[i][3];
      sh.getRange(i + 2, 4).setBackground(pct >= 80 ? '#e6f4ea' : pct >= CFG.WEAK_THRESHOLD ? '#fef7e0' : '#fce8e6')
                            .setFontWeight('bold');
    }
    sh.getRange(2, 1, rows.length, 1).setHorizontalAlignment('center');
    sh.getRange(2, 4, rows.length, 4).setHorizontalAlignment('center');
  }
  // 全班摘要
  var pcts = rows.map(function (r) { return r[3]; }).sort(function (a, b) { return a - b; });
  var r0 = rows.length + 3;
  sh.getRange(r0, 1).setValue('全班摘要').setFontWeight('bold');
  sh.getRange(r0 + 1, 1, 4, 2).setValues([
    ['作答人數', rows.length],
    ['平均分',   pcts.length ? (pcts.reduce(function (a, b) { return a + b; }, 0) / pcts.length).toFixed(1) : ''],
    ['中位數',   pcts.length ? (pcts.length % 2 ? pcts[(pcts.length - 1) / 2] : ((pcts[pcts.length / 2 - 1] + pcts[pcts.length / 2]) / 2)).toFixed(1) : ''],
    ['最低 / 最高', pcts.length ? pcts[0] + ' / ' + pcts[pcts.length - 1] : '']
  ]);
  sh.setFrozenRows(1); sh.autoResizeColumns(1, head.length); sh.setColumnWidth(2, 200);
  sh.getRange(1, 1, rows.length + 1, head.length).setBorder(true, true, true, true, true, true);
}

/* ---------- 個人錯題 ---------- */
function writeWrongSheet(people) {
  var sh = sheetOf(CFG.OUT_WRONG, true); sh.clear();
  var head = ['作答者', '醫院', '題號', '科別', '年分', '你選的', '正解', '題幹'];
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');

  var rows = [];
  people.slice().sort(function (a, b) { return b.wrongs.length - a.wrongs.length; })
    .forEach(function (p) {
      p.wrongs.forEach(function (w) {
        rows.push([p.who, p.hospital, w.k + 1, w.q.category, w.q.year,
                   w.picked || '（未作答）', w.q.answer || '?', w.q.stem]);
      });
    });

  if (rows.length) {
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    sh.getRange(2, 8, rows.length, 1).setWrap(true).setVerticalAlignment('top');
    sh.getRange(2, 3, rows.length, 5).setHorizontalAlignment('center');
    sh.getRange(2, 6, rows.length, 1).setBackground('#fce8e6');
    sh.getRange(2, 7, rows.length, 1).setBackground('#e6f4ea');
  } else {
    sh.getRange(2, 1).setValue('沒有任何人答錯 —— 題目可能太簡單了。');
  }
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 90); sh.setColumnWidth(3, 55);
  sh.setColumnWidth(4, 100); sh.setColumnWidth(5, 60);
  sh.setColumnWidth(6, 70);  sh.setColumnWidth(7, 60); sh.setColumnWidth(8, 520);
  sh.getRange(1, 1, rows.length + 1, head.length).setBorder(true, true, true, true, true, true);
}

/* ---------- 逐題分析（含錯誤選項分布） ---------- */
function writeItemSheet(itemStat) {
  var sh = sheetOf(CFG.OUT_ITEM, true); sh.clear();
  var head = ['題號', '科別', '年分', '答對率', '作答人數', 'A', 'B', 'C', 'D', 'E', '未作答', '正解', '主要誤選', '誤選人數', '題幹'];
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');

  var rows = itemStat.filter(function (s) { return s.q; })
    .map(function (s) {
      var rate = s.n ? Math.round(s.ok / s.n * 100) : 0;
      var ans = s.q.answer;
      var trap = null, trapN = 0;
      ['A', 'B', 'C', 'D', 'E'].forEach(function (L) {
        if (L !== ans && s.counts[L] > trapN) { trap = L; trapN = s.counts[L]; }
      });
      return { rate: rate, row: [s.k + 1, s.q.category, s.q.year, rate, s.n,
              s.counts.A, s.counts.B, s.counts.C, s.counts.D, s.counts.E, s.blank,
              ans || '?', trapN > 0 ? trap : '—', trapN, s.q.stem] };
    })
    .sort(function (a, b) { return a.rate - b.rate; });   // 最不會的排最上面

  if (rows.length) {
    sh.getRange(2, 1, rows.length, head.length).setValues(rows.map(function (r) { return r.row; }));
    for (var i = 0; i < rows.length; i++) {
      var pct = rows[i].rate;
      sh.getRange(i + 2, 4).setBackground(pct >= 80 ? '#e6f4ea' : pct >= CFG.WEAK_THRESHOLD ? '#fef7e0' : '#fce8e6')
                            .setFontWeight('bold');
      // 正解那一欄的人數上綠底，主要誤選上紅底 —— 一眼看出概念被什麼帶走
      var ansCol = { A: 6, B: 7, C: 8, D: 9, E: 10 }[rows[i].row[11]];
      if (ansCol) sh.getRange(i + 2, ansCol).setBackground('#e6f4ea');
      var trapCol = { A: 6, B: 7, C: 8, D: 9, E: 10 }[rows[i].row[12]];
      if (trapCol && rows[i].row[13] > 0) sh.getRange(i + 2, trapCol).setBackground('#fce8e6');
    }
    sh.getRange(2, 15, rows.length, 1).setWrap(true).setVerticalAlignment('top');
    sh.getRange(2, 1, rows.length, 14).setHorizontalAlignment('center');
  }
  sh.setFrozenRows(1);
  sh.setColumnWidth(1, 55); sh.setColumnWidth(2, 100); sh.setColumnWidth(3, 60);
  for (var c = 4; c <= 14; c++) sh.setColumnWidth(c, 68);
  sh.setColumnWidth(15, 520);
  sh.getRange(1, 1, rows.length + 1, head.length).setBorder(true, true, true, true, true, true);

  var note = rows.length + 3;
  sh.getRange(note, 1).setValue('依答對率由低到高排序。紅底的「主要誤選」代表：答錯的人集中選同一個選項，'
    + '通常是某個觀念被混淆，而不只是「不會」—— 這種題最值得拿到 teaching 討論。').setFontColor('#5f6368');
}

/* ---------- 醫院 × 科別 ---------- */
function writeHospitalSheet(people, pairs) {
  var sh = sheetOf(CFG.OUT_HOSPITAL, true); sh.clear();
  var head = ['醫院', '科別', '答對率', '答對', '作答數', '人數'];
  sh.getRange(1, 1, 1, head.length).setValues([head])
    .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');

  var agg = {};   // hospital -> category -> {ok,total}
  var headcount = {};
  people.forEach(function (p) {
    headcount[p.hospital] = (headcount[p.hospital] || 0) + 1;
    if (!agg[p.hospital]) agg[p.hospital] = {};
    var wrongSet = {};
    p.wrongs.forEach(function (w) { wrongSet[w.k] = true; });
    pairs.forEach(function (pr, k) {
      if (!pr.q) return;
      var c = pr.q.category;
      if (!agg[p.hospital][c]) agg[p.hospital][c] = { ok: 0, total: 0 };
      agg[p.hospital][c].total++;
      if (!wrongSet[k]) agg[p.hospital][c].ok++;
    });
  });

  var rows = [];
  Object.keys(agg).sort().forEach(function (h) {
    Object.keys(agg[h]).sort().forEach(function (c) {
      var s = agg[h][c];
      rows.push([h, c, s.total ? Math.round(s.ok / s.total * 100) : 0, s.ok, s.total, headcount[h]]);
    });
  });

  if (rows.length) {
    sh.getRange(2, 1, rows.length, head.length).setValues(rows);
    for (var i = 0; i < rows.length; i++) {
      var pct = rows[i][2];
      sh.getRange(i + 2, 3).setBackground(pct >= 80 ? '#e6f4ea' : pct >= CFG.WEAK_THRESHOLD ? '#fef7e0' : '#fce8e6')
                            .setFontWeight('bold');
    }
    sh.getRange(2, 3, rows.length, 4).setHorizontalAlignment('center');
  }
  sh.setFrozenRows(1); sh.autoResizeColumns(1, head.length);
  sh.getRange(1, 1, rows.length + 1, head.length).setBorder(true, true, true, true, true, true);
}


/* ═══════════════════════════════════════════════════════════════
   ③ 封存本次考卷 —— 重新抽題之前務必先跑這個
   ---------------------------------------------------------------
   「抽取30題」會直接清空並覆寫「選出題目」，
   「產生表單」會清空並覆寫「表單連結」。
   一旦覆蓋，舊表單雖然還躺在 Drive 裡、回應也都在，
   但因為對不回原始題目，成績就再也算不出來了。
   這個函式把當下的考卷與連結存成一張帶日期的快照，並登記到索引表。
   ═══════════════════════════════════════════════════════════════ */
function archiveCurrentPaper() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var src = sheetOf(CFG.SHEET_SELECTED, false);
    if (!src || src.getLastRow() < 2) throw new Error('「' + CFG.SHEET_SELECTED + '」是空的，沒有東西可以封存。');

    var tag  = Utilities.formatDate(new Date(), 'GMT+8', 'yyyyMMdd_HHmm');
    var name = '考卷_' + tag;
    if (ss.getSheetByName(name)) name = name + '_' + Math.floor(Math.random() * 900 + 100);

    var copy = src.copyTo(ss).setName(name);
    ss.setActiveSheet(copy);
    ss.moveActiveSheet(ss.getNumSheets());

    // 把表單連結一併寫進快照，日後才找得回那份表單
    var links = sheetOf(CFG.SHEET_LINKS, false);
    var pub = '', edit = '';
    if (links) { pub = links.getRange('B1').getValue(); edit = links.getRange('B2').getValue(); }
    var last = copy.getLastRow() + 2;
    copy.getRange(last, 1, 3, 2).setValues([
      ['封存時間', Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm')],
      ['表單連結', pub],
      ['編輯連結', edit]
    ]);
    copy.getRange(last, 1, 3, 1).setFontWeight('bold');

    // 索引表
    var idx = sheetOf('歷次考卷', true);
    if (idx.getLastRow() === 0) {
      idx.getRange(1, 1, 1, 5).setValues([['封存時間', '快照工作表', '題數', '表單連結', '編輯連結']])
         .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
      idx.setFrozenRows(1);
    }
    idx.appendRow([Utilities.formatDate(new Date(), 'GMT+8', 'yyyy/MM/dd HH:mm'),
                   name, src.getLastRow() - 1, pub, edit]);
    idx.autoResizeColumns(1, 3);

    ui.alert('已封存',
      '這份考卷已存成「' + name + '」，並登記在「歷次考卷」。\n\n' +
      '現在可以安全地重新抽題了。\n\n' +
      '重要：按「抽取30題」之前，記得先點回你的題庫主分頁 —— ' +
      '程式是從「目前所在的工作表」讀題庫的，停在「' + CFG.SHEET_SELECTED + '」上按會把資料洗掉。',
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}

/* ═══════════════════════════════════════════════════════════════
   還原某一次封存的考卷（為了重算舊表單的成績）
   在「歷次考卷」找到那一列的快照名稱，執行本函式並輸入即可。
   ═══════════════════════════════════════════════════════════════ */
function restoreArchivedPaper() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var resp = ui.prompt('還原考卷', '請輸入快照工作表名稱（例如 考卷_20260822_1430）：', ui.ButtonSet.OK_CANCEL);
    if (resp.getSelectedButton() !== ui.Button.OK) return;
    var name = resp.getResponseText().trim();
    var snap = ss.getSheetByName(name);
    if (!snap) throw new Error('找不到工作表「' + name + '」。');

    if (ui.alert('確認',
        '這會用「' + name + '」覆蓋目前的「' + CFG.SHEET_SELECTED + '」。\n' +
        '若目前那份考卷還沒封存，請先取消並跑「封存本次考卷」。\n\n要繼續嗎？',
        ui.ButtonSet.YES_NO) !== ui.Button.YES) return;

    // 只取題目區（封存時在尾端多寫了 3 行連結資訊，要排除）
    var vals = snap.getDataRange().getValues();
    var rows = [];
    for (var i = 0; i < vals.length; i++) {
      if (String(vals[i][0]).indexOf('封存時間') === 0) break;
      rows.push(vals[i].slice(0, 4));
    }

    var tgt = sheetOf(CFG.SHEET_SELECTED, true);
    tgt.clear();
    tgt.getRange(1, 1, rows.length, 4).setValues(rows);
    tgt.getRange(1, 1, 1, 4).setBackground('#f3f3f3').setFontWeight('bold');
    tgt.setColumnWidth(1, 500);
    tgt.getRange(2, 1, Math.max(1, rows.length - 1), 1).setWrap(true).setVerticalAlignment('top');

    // 連結也一併還原，報表才開得到正確的那份表單
    var pub = '', edit = '';
    for (var j = 0; j < vals.length; j++) {
      if (String(vals[j][0]) === '表單連結') pub = vals[j][1];
      if (String(vals[j][0]) === '編輯連結') edit = vals[j][1];
    }
    if (edit) {
      var links = sheetOf(CFG.SHEET_LINKS, true);
      links.clear();
      links.getRange('A1:B2').setValues([['表單連結', pub], ['編輯連結', edit]]);
    }
    tgt.activate();
    ui.alert('已還原', '「' + CFG.SHEET_SELECTED + '」已還原成 ' + name + '（' + (rows.length - 1) + ' 題）。\n現在可以跑「產生完整測驗報告」了。', ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}


/* ═══════════════════════════════════════════════════════════════
   產生測驗表單（修正版）
   ---------------------------------------------------------------
   相對於原版修掉三件事：
   1. 正解比對改用 normalizeAnswer()。原版 choice.startsWith(正解 + '.')
      只要 C 欄是「A 」「Ａ」「(A)」「a」「A.」任何一種就標不到正解，
      該題在表單裡變成沒有答案，全班都算答錯，而且不會報錯。
   2. 題幹改用 splitQuestion()，不會再因為題幹裡出現 "vitamin D." 這類字串
      而被默默截斷。
   3. 產生表單時「同時」寫下題目對照表，報表才對得回正確的題目與科別。
      原版遇到有問題的題目是 continue 跳過，表單題數就和「選出題目」對不上，
      後面所有題目的科別會整排錯位。這裡改成事前擋下並要求你先修。
   ═══════════════════════════════════════════════════════════════ */

var FORMCFG = {
  FOLDER_ID              : '',   // ← 貼上表單存放資料夾 ID（網址 /folders/ 後那串）；留空＝存在「我的雲端硬碟」根目錄
  ASK_NAME               : true,   // 在表單最前面加一題「姓名」，報表就能顯示姓名而不是 email
  STRIP_ORIGINAL_NUMBER  : true,   // 去掉題幹開頭的原始題號（例如「9.」），避免和表單題號疊在一起
  HOSPITALS              : ['北醫附醫', '雙和醫院', '萬芳醫院']
};

function generateForm() {
  var ui = SpreadsheetApp.getUi();
  try {
    var qs = readSelected();

    /* 前測 / 後測 —— 一份表單＝一次測驗，靠表單本身區別，不必猜日期 */
    var kindResp = ui.alert('這份是前測還是後測？',
      '【是】前測　　課程開始前施測，用來抓弱點\n'
      + '【否】後測　　課程結束後施測，和前測配對比較\n'
      + '【取消】不分　單純的一次測驗\n\n'
      + '選了之後會寫進表單名稱和「' + CFG.SHEET_HISTORY + '」，之後才做得出前後測比較。',
      ui.ButtonSet.YES_NO_CANCEL);
    var kind = kindResp === ui.Button.YES ? '前測'
             : kindResp === ui.Button.NO  ? '後測' : '';
    if (!qs.length) throw new Error('「' + CFG.SHEET_SELECTED + '」裡沒有題目，請先執行「重新抽題」。');

    /* --- 事前擋下有問題的題目，不再默默跳過 --- */
    var bad = qs.filter(function (q) { return !q.parse.ok || !q.answer; });
    if (bad.length) {
      throw new Error('有 ' + bad.length + ' 題無法做成表單（第 '
        + bad.slice(0, 5).map(function (b) { return b.row; }).join('、')
        + (bad.length > 5 ? ' …' : '') + ' 題）。\n\n'
        + '原因多半是正解欄格式不對，或選項拆不出四個。\n'
        + '請先執行「① 題庫健檢」看詳細清單並修正，再回來產生表單。\n\n'
        + '（如果讓這些題目進到表單，它們會變成「沒有正解」，全班都會被算答錯。）');
    }

    /* --- 舊表單連結會被覆蓋：先自動登記到「歷次表單」，再確認 --- */
    var links = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CFG.SHEET_LINKS);
    if (links && links.getRange('B2').getValue()) {
      var logged = archiveFormLink_();          // 舊表單存進「歷次表單」，之後永遠找得回來
      var go = ui.alert('會蓋掉舊表單的連結',
        (logged
          ? '舊表單「' + logged.title + '」（' + logged.responses + ' 筆回應）'
            + '已經自動登記到「' + CFG.SHEET_HISTORY + '」，之後隨時可以從那裡開回去。\n\n'
          : '（舊連結無法開啟，沒能登記，但表單本身仍在 Drive。）\n\n')
        + '接下來會產生新表單並覆蓋「' + CFG.SHEET_LINKS + '」的連結。\n\n'
        + '如果舊考卷的成績還沒統計完，請先取消，跑「② 產生完整測驗報告（從表單）」'
        + '和「封存成績報告」。\n\n要繼續嗎？', ui.ButtonSet.YES_NO);
      if (go !== ui.Button.YES) return;
    }

    var info = buildQuizForm_(qs, kind);
    ui.alert('表單已建立',
      '共 ' + info.count + ' 題。\n\n'
      + '發給住院醫師的連結：\n' + info.formUrl + '\n\n'
      + '（連結已寫入「' + CFG.SHEET_LINKS + '」，題目對照已寫入「' + CFG.SHEET_MAP + '」）'
      + (info.warn ? '\n\n⚠ ' + info.warn : ''),
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}

function buildQuizForm_(qs, kind, shuffleChoices, skipLinkSheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dateString = Utilities.formatDate(new Date(), 'GMT+8', 'yyyy-MM-dd');

  var form = FormApp.create('兒科題庫測驗' + (kind ? '_' + kind : '') + '_' + dateString);
  var responseSheet = SpreadsheetApp.create('兒科題庫測驗回應' + (kind ? '_' + kind : '') + '_' + dateString);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, responseSheet.getId());

  /* 把表單和回應試算表搬進指定資料夾。
     這一步純粹是檔案整理，失敗不該讓整個建表單流程中斷 ——
     原版用 folder.addFile() + getRootFolder().removeFile()，資料夾 ID 失效
     或檔案在共用雲端硬碟時會直接丟出
     「No item with the given ID could be found」而讓一切停擺。
     改用 moveTo()（共用雲端硬碟也支援），並包在 try 裡。 */
  var moveWarn = '';
  if (FORMCFG.FOLDER_ID) try {   // 沒設定資料夾就留在根目錄，不必警告
    var folder = DriveApp.getFolderById(FORMCFG.FOLDER_ID);
    DriveApp.getFileById(form.getId()).moveTo(folder);
    DriveApp.getFileById(responseSheet.getId()).moveTo(folder);
  } catch (moveErr) {
    moveWarn = '表單已建立，但沒能搬進指定資料夾（' + moveErr.message + '）。'
             + '檔案留在你的「我的雲端硬碟」根目錄，功能完全不受影響。'
             + '要修的話，把 FORMCFG.FOLDER_ID 換成一個你有權限的資料夾 ID。';
  }

  form.setDescription('這是一份兒科題庫測驗，請仔細作答。')
      .setIsQuiz(true)
      .setCollectEmail(true)
      .setLimitOneResponsePerUser(true)
      .setShowLinkToRespondAgain(false);

  if (FORMCFG.ASK_NAME) {
    form.addTextItem().setTitle('姓名').setRequired(true);
  }

  var hosp = form.addMultipleChoiceItem();
  hosp.setTitle('請選擇您的醫院:')
      .setChoices(FORMCFG.HOSPITALS.map(function (h) { return hosp.createChoice(h); }))
      .setRequired(true);

  var mapRows = [];
  qs.forEach(function (q) {
    var stem = q.stem;
    if (FORMCFG.STRIP_ORIGINAL_NUMBER) stem = stem.replace(/^\s*\d+\s*[.、．]\s*/, '');

    var item = form.addMultipleChoiceItem();
    // 標題保留「第 N 題」，N 是「選出題目」的列號 —— 報表靠這個對回題目與科別
    item.setTitle('第' + q.row + '題: ' + stem).setPoints(1).setRequired(true);

    var opts = q.choices.slice();
    if (shuffleChoices) {
      /* 打亂選項顯示順序，並依新位置重新編上 A/B/C/D。
         必須重新編號 —— 否則畫面會出現「B. … D. … A. …」這種看起來像壞掉的順序。
         代價：報表裡的字母是「顯示位置」，不再對應題庫的原始字母。
         但報表一律把選項文字印在字母旁邊，所以看得出來是哪一個選項。 */
      opts = fyShuffle(opts);
    }
    item.setChoices(opts.map(function (c, i) {
      var text = shuffleChoices
        ? 'ABCDE'.charAt(i) + '.' + c.body      // 去掉原字母、依新位置重編
        : c.full;                                // 保留原本的 "A.xxx"
      return item.createChoice(text, c.letter === q.answer);
    }));
    mapRows.push([q.row, item.getId(), q.category, q.answer]);
  });

  if (!skipLinkSheet) {
    var links = sheetOf(CFG.SHEET_LINKS, true);
    links.clear();
    links.getRange('A1:B2').setValues([
      ['表單連結', form.getPublishedUrl()],
      ['編輯連結', form.getEditUrl()]
    ]);
    links.getRange('A1:A2').setFontWeight('bold');
    links.autoResizeColumns(1, 2);
  }

  var map = sheetOf(CFG.SHEET_MAP + (kind ? '_' + kind : ''), true);
  map.clear();
  map.getRange(1, 1, 1, 4).setValues([['選出題目列', 'itemId', '科別', '正解']])
     .setBackground('#37474f').setFontColor('#ffffff').setFontWeight('bold');
  if (mapRows.length) map.getRange(2, 1, mapRows.length, 4).setValues(mapRows);
  map.setFrozenRows(1); map.autoResizeColumns(1, 4);

  return { count: qs.length, formUrl: form.getPublishedUrl(), editUrl: form.getEditUrl(), warn: moveWarn };
}
