#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
由題目 JSON 產生一次性匯入用的 Apps Script。

用法:
    python3 make_import_gs.py rows.json 115 [函式後綴] [> 匯入115年題目.gs]
    python3 make_import_gs.py --from-bank ../data/bank-114-specialty-A.json [排除題號,…] [函式後綴] > 匯入.gs

選項可以是四個（A–D，國考）或五個（A–E，專科甄審）。
年分欄可以是非數字（例如「114專科」），抽題時的年分篩選是用字串比對。
函式名稱是 import<後綴>，後綴預設等於年分；年分有中文時請另給英數後綴。

rows.json 的格式是陣列，每筆四個元素:
    ["1.題幹\nA.選項\nB.選項\nC.選項\nD.選項[\nE.選項]", 115, "B", "General"]
     題目全文                                    年分  正解  科別

產生的 .gs 只往後新增、自動跳過重複、執行前會跳確認視窗。
"""
import io, json, re, sys


TEMPLATE = '''/***********************************************************************
 * 一次性匯入：__YEAR__ 兒科題目（年分欄填「__YEAR__」）
 * ---------------------------------------------------------------------
 * 使用方式
 *   1. Apps Script → 新增一個檔案 → 貼上這一整份 → 存檔
 *   2. 上方函式下拉選單選 import__FN__ → 按「執行」
 *   3. 跑完後回試算表看「題庫」，新題目會加在最後面
 *   4. 確認沒問題後，這個檔案就可以刪掉了
 *
 * 特性
 *   · 只往後新增，不會動到既有的任何一列
 *   · 自動跳過題庫裡已經有的題目（比對題幹＋全部選項的完整內容）
 *   · 五選項（A–E）題目需要「兒科題庫測驗系統_完整版.gs」v18 以上
 *   · 自動偵測標題列位置，相容 Google 表格格式
 ***********************************************************************/

var IMPORT___FN__ = __DATA__;

function import__FN__() {
  var ui = SpreadsheetApp.getUi();
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var name = (typeof DRAW !== 'undefined' && DRAW.SHEET_BANK) ? DRAW.SHEET_BANK : '題庫';
    var sh = ss.getSheetByName(name);
    if (!sh) throw new Error('找不到「' + name + '」工作表。');

    var data = sh.getDataRange().getValues();
    var hRow = 0;
    for (var i = 0; i < Math.min(10, data.length); i++) {
      if (String(data[i][0] || '').indexOf('題目') === 0) { hRow = i; break; }
    }
    var key = function (t) {
      return String(t).replace(/^\\s*\\d+\\s*[.\\u3001\\uFF0E]\\s*/, '').replace(/\\s/g, '');
    };
    var seen = {}, existing = 0;
    for (var r = hRow + 1; r < data.length; r++) {
      if (!data[r][0]) continue;
      seen[key(data[r][0])] = true;
      existing++;
    }

    var fresh = IMPORT___FN__.filter(function (row) { return !seen[key(row[0])]; });
    var dup = IMPORT___FN__.length - fresh.length;

    if (!fresh.length) {
      ui.alert('沒有新題目', '這 ' + IMPORT___FN__.length + ' 題在「' + name + '」裡都已經有了。',
               ui.ButtonSet.OK);
      return;
    }

    var go = ui.alert('確認匯入',
      '要把 ' + fresh.length + ' 題加進「' + name + '」嗎？\\n\\n'
      + '· 目前題庫有 ' + existing + ' 題，匯入後會變成 ' + (existing + fresh.length) + ' 題\\n'
      + (dup ? '· 有 ' + dup + ' 題已經存在，會自動略過\\n' : '')
      + '\\n只會往後新增，不會動到既有的任何一列。', ui.ButtonSet.YES_NO);
    if (go !== ui.Button.YES) return;

    var start = sh.getLastRow() + 1;
    sh.getRange(start, 1, fresh.length, 4).setValues(fresh);
    sh.getRange(start, 1, fresh.length, 1).setWrap(true).setVerticalAlignment('top');
    sh.getRange(start, 2, fresh.length, 3).setHorizontalAlignment('center').setVerticalAlignment('middle');

    ui.alert('匯入完成',
      '已新增 ' + fresh.length + ' 題（第 ' + start + ' 列起）。\\n\\n'
      + '建議接著跑「① 題庫健檢」確認格式，再「重新抽題」就會抽到這些新題目了。',
      ui.ButtonSet.OK);

  } catch (e) {
    ui.alert('錯誤', e.message, ui.ButtonSet.OK);
  }
}
'''


def from_bank(path, exclude):
    """把網頁版題庫 JSON 轉成匯入用的列；有附圖的題目一律排除（Google 表單看不到圖）"""
    bank = json.loads(io.open(path, encoding="utf-8").read())
    letters = bank.get("optionLetters", "ABCDE")
    rows, skipped = [], []
    for q in bank["questions"]:
        if q.get("image") or q["no"] in exclude:
            skipped.append(q["no"]); continue
        text = "%d.%s\n" % (q["no"], q["stem"].replace("\n", " ")) + "\n".join(
            "%s.%s" % (letters[i], o) for i, o in enumerate(q["options"]))
        rows.append([text, q["year"], q["answer"], q["category"]])
    return rows, skipped


def main():
    if len(sys.argv) >= 3 and sys.argv[1] == "--from-bank":
        exclude = {int(x) for x in sys.argv[3].split(",")} if len(sys.argv) > 3 and sys.argv[3] else set()
        rows, skipped = from_bank(sys.argv[2], exclude)
        year = str(rows[0][1])
        bank_id = json.loads(io.open(sys.argv[2], encoding="utf-8").read())["id"]
        fn = sys.argv[4] if len(sys.argv) > 4 else re.sub(r"[^A-Za-z0-9_]", "_", bank_id)
        if skipped:
            sys.stderr.write("排除（有附圖或指定排除）：第 %s 題\n" % "、".join(map(str, skipped)))
    elif len(sys.argv) >= 3:
        rows = json.loads(io.open(sys.argv[1], encoding="utf-8").read())
        year = str(sys.argv[2])
        fn = sys.argv[3] if len(sys.argv) > 3 else year
    else:
        sys.exit(__doc__)
    if not re.match(r"^[A-Za-z0-9_]+$", fn):
        sys.exit("函式後綴只能用英數字與底線：%s" % fn)

    bad = []
    for i, r in enumerate(rows):
        if len(r) != 4:
            bad.append("第 %d 筆欄位數不是 4" % (i + 1)); continue
        parts = r[0].split("\n")
        letters = "".join(p[0] for p in parts[1:] if p[:1] in "ABCDE")
        if letters not in ("ABCD", "ABCDE") or len(parts) != len(letters) + 1:
            bad.append("第 %d 筆選項標籤是「%s」，應為 ABCD 或 ABCDE" % (i + 1, letters))
        if len(str(r[2])) != 1 or r[2] not in letters:
            bad.append("第 %d 筆正解是「%s」，應為單一字母且對得到選項" % (i + 1, r[2]))
        if not any(p.startswith(str(r[2]) + ".") for p in parts[1:]):
            bad.append("第 %d 筆的正解對不到任何選項" % (i + 1))
    if bad:
        sys.exit("資料有問題，未產生檔案：\n  " + "\n  ".join(bad))

    data = json.dumps(rows, ensure_ascii=False).replace("</", "<\\/")
    sys.stdout.write(TEMPLATE.replace("__FN__", fn).replace("__YEAR__", year).replace("__DATA__", data))
    sys.stderr.write("已產生 %d 題的匯入檔（%s 年）\n" % (len(rows), year))


if __name__ == "__main__":
    main()
