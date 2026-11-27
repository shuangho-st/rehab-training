#!/usr/bin/env python3
"""由成人吞嚥評估課程頁產生其他課程頁（共用版型：單元、測驗、錯題重考、模擬情境、週次、紀錄回傳）。
用法：python3 tools/gen_course.py <課程資料夾>…；課程設定寫在 COURSES。"""
import re, sys, os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MULTI_API = ''  # 共用 Apps Script 的 /exec 網址（部署後填入）

COURSES = {
 'foundations': dict(title='臨床基礎與專業規範', sub='實習第 1–2 週', ver='foundations-1', prefix='fnd', logbook=True,
   lead='病人安全、感控、病歷閱讀與書寫、專業倫理與溝通、學習方法。每個單元約 8 到 12 分鐘，最後有小測驗；讀完後用模擬情境練習，並每天寫觀察日誌。',
   order='第 1 週讀「病人安全與感控」與「病歷閱讀與書寫」，第 2 週讀「專業素養」與「學習方法」。第一次進病房前，至少完成「病人安全與感控」。',
   sim='模擬情境', simintro='把規範用在真實情境：進隔離病房、治療中狀況改變、隱私與錄影。可以重複練習，紀錄取最高分。',
   foot='本課程為科內自學教材；感控、病人安全與隱私規範以院內最新規定與督導指導為準。'),
 'instrumental': dict(title='儀器吞嚥檢查判讀', sub='VFSS 與 FEES', ver='instr-1', prefix='ins',
   lead='螢光吞嚥攝影與內視鏡吞嚥檢查的選擇、流程、影像判讀與報告。讀完後用模擬個案練習判讀與建議。',
   order='依序讀「檢查的選擇與準備」「影像判讀」「報告與建議」。第一次見習儀器檢查前，先讀完前兩個領域。',
   sim='模擬個案', simintro='從檢查前準備、判讀到給建議，每一步都要做決定。',
   foot='本課程為科內自學教材；儀器檢查的執行依院內規範與資格，判讀請與督導確認。'),
 'head-neck': dict(title='頭頸癌的吞嚥與溝通', sub='治療前、中、後', ver='hnc-1', prefix='hnc',
   lead='頭頸癌手術與放化療對吞嚥與說話的影響、預防與復健、張口困難，以及喉切除後的發聲。讀完後用模擬個案練習。',
   order='依序讀「治療與影響」「吞嚥介入」「溝通」。',
   sim='模擬個案', simintro='從治療前諮詢到治療後追蹤，每一步都要做決定。',
   foot='本課程為科內自學教材；術後介入須依外科醫師指示與院內規範。'),
 'cognitive-communication': dict(title='成人認知溝通障礙', sub='右腦損傷、腦外傷、失智症', ver='cogcomm-1', prefix='cgc',
   lead='認知與溝通的關係、右腦損傷、腦外傷與失智症的評估與介入。讀完後用模擬個案練習。',
   order='依序讀「基礎」「右腦損傷與腦外傷」「失智症」。',
   sim='模擬個案', simintro='依病人的認知溝通特性選擇評估與介入。',
   foot='本課程為科內自學教材；篩檢與評估工具請依手冊施測，並與團隊共同判斷。'),
 'voice': dict(title='嗓音障礙', sub='評估與治療', ver='voice-1', prefix='voi',
   lead='發聲生理、嗓音評估、常見嗓音障礙與治療方法。讀完後用模擬個案練習。',
   order='依序讀「發聲與評估」「常見嗓音障礙」「治療」。',
   sim='模擬個案', simintro='依喉鏡結果與嗓音表現選擇治療方向。',
   foot='本課程為科內自學教材；嗓音治療前須有耳鼻喉科喉部檢查。'),
 'tracheostomy': dict(title='氣切與說話瓣膜', sub='溝通與吞嚥', ver='trach-1', prefix='trc',
   lead='氣切管的種類、氣囊管理、說話瓣膜的評估與試戴、氣切病人的吞嚥與拔管。讀完後用模擬個案練習。',
   order='依序讀「氣切基礎」「說話瓣膜」「吞嚥與拔管」。第一次參與說話瓣膜試戴前，先讀完前兩個領域。',
   sim='模擬個案', simintro='從確認條件、試戴到監測，每一步都要做決定。',
   foot='本課程為科內自學教材；氣囊與說話瓣膜的操作依醫囑、院內規範並與呼吸治療師、護理師合作。'),
 'elderly-care': dict(title='高齡與護理之家吞嚥照護', sub='衰弱、失智與照護目標', ver='elder-1', prefix='eld',
   lead='老化與吞嚥、衰弱與肌少症、多重用藥、機構照護、舒適進食與共享決策。讀完後用模擬個案練習。',
   order='依序讀「高齡吞嚥」「機構照護」「照護目標與倫理」。到護理之家見習前讀完。',
   sim='模擬個案', simintro='在機構與家庭的情境中，兼顧安全、營養與生活品質。',
   foot='本課程為科內自學教材；照護決策由團隊與病人、家屬共同討論。'),
 'aac': dict(title='輔助溝通系統（AAC）', sub='評估與建置', ver='aac-1', prefix='aac',
   lead='輔助溝通的類型、評估與特徵配對、建置與溝通夥伴訓練。讀完後用模擬個案練習。',
   order='依序讀「認識輔助溝通」「評估」「建置與訓練」。',
   sim='模擬個案', simintro='依使用者的能力與需求選擇並建置輔助溝通。',
   foot='本課程為科內自學教材；輔具補助與申請流程依各縣市規定。'),
 'interprofessional': dict(title='跨團隊合作與共享決策', sub='團隊、家庭會議與 SDM', ver='ipe-1', prefix='ipe',
   lead='醫療團隊的角色分工、跨團隊溝通、共享決策與困難對話。讀完後用模擬情境練習。',
   order='依序讀「團隊合作」「共享決策」「困難對話與倫理」。',
   sim='模擬情境', simintro='在團隊會議與家庭會議中練習溝通與共享決策。',
   foot='本課程為科內自學教材；醫療決策由團隊與病人、家屬共同決定。'),
 'evidence': dict(title='實證臨床與個案報告', sub='從問題到發表', ver='ebp-1', prefix='ebp',
   lead='臨床問題、文獻搜尋與評讀、單一個案設計與成效測量、個案報告的撰寫與發表。讀完後用模擬情境練習。',
   order='依序讀「找證據」「評讀與應用」「個案報告」。開始準備個案報告前讀完。',
   sim='模擬情境', simintro='從臨床問題、找文獻到設計成效測量與報告。',
   foot='本課程為科內自學教材；個案報告須去識別化並依院內規定取得同意。'),
 'capstone': dict(title='綜合複習與總結評量', sub='第 19–20 週', ver='capstone-1', prefix='cap',
   lead='跨主題的綜合個案：成人與兒童、評估與治療、安全與溝通。用來檢視是否具備獨立執業的能力。',
   order='先讀「總結評量說明」，再完成所有綜合個案。每個個案都要達 80 分以上且沒有不適當的選擇。',
   sim='綜合個案', simintro='每個個案都可能跨越多個主題，請整合 20 週所學。',
   foot='本課程為總結評量的一部分；最終是否可獨立執業，由總結里程碑會議決定。'),
}

LOGBOOK_CODE = r'''/* ---------- 觀察日誌 ---------- */
const SITES=['成人病房 BS','成人門診','兒童門診','早療','NICU','護理之家','其他'];
let LOGS=ld(K.out.replace('outbox','logs'),[]);
function logForm(){
 cur='log';header('觀察日誌');window.scrollTo(0,0);
 const fs='width:100%;font:inherit;font-size:16px;padding:9px 11px;border-radius:10px;border:1px solid var(--line);background:var(--surface);color:var(--ink);margin:4px 0 10px';
 app.innerHTML=`<p class="small muted" style="margin-top:22px">觀察日誌</p><h1>今天的觀察</h1>
 <div class="box goal">每天寫一篇。病人資料要去識別化（不寫姓名、床號、病歷號、日期組合）。寫出一個「為什麼」，比記流水帳更有學習價值。</div>
 <label class="small muted" for="lg-site">場域</label><select id="lg-site" style="${fs}">${SITES.map(s=>`<option>${s}</option>`).join('')}</select>
 <label class="small muted" for="lg-what">發生什麼（客觀描述）</label><textarea id="lg-what" style="${fs};min-height:90px" placeholder="例如：72 歲男性，左側中風第 3 天，治療師先做口腔清潔，再用 IDDSI 4 試餵…"></textarea>
 <label class="small muted" for="lg-think">我的想法、感受或疑問</label><textarea id="lg-think" style="${fs};min-height:70px" placeholder="例如：我原本以為口腔清潔是多餘的…"></textarea>
 <label class="small muted" for="lg-learn">學到什麼</label><textarea id="lg-learn" style="${fs};min-height:70px"></textarea>
 <label class="small muted" for="lg-next">下次我會怎麼做</label><textarea id="lg-next" style="${fs};min-height:60px"></textarea>
 <div id="lg-err"></div>
 <div class="nav"><button class="btn ghost" id="bk">回課程首頁</button><button class="btn" id="lg-go">送出給督導</button></div>
 ${LOGS.length?`<h2>我寫過的日誌</h2>${LOGS.slice().reverse().map(l=>`<div class="chartpaper" style="margin:8px 0"><b>${esc(l.d)}｜${esc(l.site)}</b><p style="white-space:pre-line;margin:4px 0 0">${esc(l.what)}</p></div>`).join('')}`:''}`;
 document.getElementById('bk').onclick=home;
 document.getElementById('lg-go').onclick=()=>{const v=id=>document.getElementById(id).value.trim();
  const e={d:new Date().toLocaleDateString('zh-TW'),site:v('lg-site'),what:v('lg-what'),think:v('lg-think'),learn:v('lg-learn'),next:v('lg-next')};
  if(e.what.length<10||e.learn.length<4){document.getElementById('lg-err').innerHTML='<div class="fb no">請至少寫下「發生什麼」與「學到什麼」。</div>';return}
  LOGS.push(e);sv(K.out.replace('outbox','logs'),LOGS);
  queue({type:'觀察日誌',domain:'觀察日誌',lesson:'log',title:e.site,n:LOGS.length,note:`發生什麼：${e.what}｜想法：${e.think}｜學到：${e.learn}｜下次：${e.next}`});
  logForm();document.getElementById('lg-err').innerHTML='<div class="fb ok">已送出，督導會在紀錄中看到。</div>';};
}

/* ---------- 畫面 ---------- */'''

def build(d, m):
    src = open(os.path.join(ROOT, 'dysphagia', 'index.html'), encoding='utf-8').read()
    s = src
    def rep(a, b, n=1):
        nonlocal s
        assert s.count(a) == n, (d, a, s.count(a)); s = s.replace(a, b)
    rep('<title>成人吞嚥評估｜病房床邊實務</title>', f"<title>{m['title']}｜實習生自學課程</title>")
    rep('id="home">成人吞嚥評估<small id="crumb">病房床邊實務</small>', f"id=\"home\">{m['title']}<small id=\"crumb\">{m['sub']}</small>")
    s = re.sub(r"^const API_URL = '.*';$", f"const API_URL = '{MULTI_API}';  // 共用 Apps Script 的 /exec 網址", s, flags=re.M)
    rep("const APP_VER = 'dysphagia-1';\nconst COURSE_NAME = '吞嚥';", f"const APP_VER = '{m['ver']}';\nconst COURSE_NAME = '{m['title']}';")
    p = m['prefix']
    rep("const K={done:'dys-done',quiz:'dys-quiz',secs:'dys-secs',me:'st-me',out:'dys-outbox',sim:'dys-sim'};",
        f"const K={{done:'{p}-done',quiz:'{p}-quiz',secs:'{p}-secs',me:'st-me',out:'{p}-outbox',sim:'{p}-sim'}};")
    rep("app.innerHTML=`<h1>成人吞嚥評估｜病房床邊實務</h1>", f"app.innerHTML=`<h1>{m['title']}</h1>")
    rep("L.push(`成人吞嚥評估自學紀錄｜", f"L.push(`{m['title']}自學紀錄｜")
    rep("app.innerHTML=`<h1>成人吞嚥評估</h1>\n <p class=\"lead\">從吞嚥基礎、病房床邊實務到臨床評估與決策。每個單元約 8 到 12 分鐘，最後有小測驗；讀完後用模擬個案練習完整的床邊評估。</p>",
        f"app.innerHTML=`<h1>{m['title']}</h1>\n <p class=\"lead\">{m['lead']}</p>")
    i = s.index('<b>建議順序：</b>'); j = s.index('</div>', i)
    s = s[:i] + '<b>建議順序：</b>' + m['order'] + s[j:]
    rep("<h2>模擬個案 ${wkTag(SCH.cases)}</h2>", f"<h2>{m['sim']} ${{wkTag(SCH.cases)}}</h2>")
    rep("從看病歷、進病房、檢查、試餵到給建議，每一步都要做決定。可以重複練習，紀錄取最高分。", m['simintro'] + ('' if '紀錄取最高分' in m['simintro'] else '可以重複練習，紀錄取最高分。'))
    i = s.index(' <details class="box goal"><summary><b>成人核心階段進度'); j = s.index('</details>\n', i) + len('</details>\n')
    s = s[:i] + ' <p class="small muted">週次依 <a href="../roadmap/">20 週學習地圖</a>，依實際排程調整。</p>\n' + s[j:]
    i = s.index('<div class="foot">'); j = s.index('</div>', i)
    s = s[:i] + '<div class="foot">' + m['foot'] + s[j:]
    for a, b in [("<b>會診資料：</b>", "<b>情境：</b>"), ("<h2>看病歷</h2>", "<h2>開始前</h2>"), ("<h3>病歷查到的資訊</h3>", "<h3>確認到的資訊</h3>"),
                 ("cf.textContent='進病房'", "cf.textContent='開始'"), ("'<b>不安全的選擇。</b>'", "'<b>不適當的選擇。</b>'"),
                 ("另選了 '+extra+' 項與評估無關的資料", "另選了 '+extra+' 項關係較小的資料"),
                 ("病歷關鍵 ${SM.chartHit}/${keys}｜最佳決策 ${SM.ok}/${n}｜次佳 ${subN}｜不安全 ${SM.bad}", "開始前確認 ${SM.chartHit}/${keys}｜最佳決策 ${SM.ok}/${n}｜次佳 ${subN}｜不適當 ${SM.bad}"),
                 ("<b>病歷關鍵項目</b>", "<b>開始前確認</b>"), ("<b>不安全的選擇</b>", "<b>不適當的選擇</b>"),
                 ("計分：病歷 20 分、決策 80 分（次佳選擇給一半），每個不安全選擇扣 10 分。通過標準：80 分以上且沒有不安全的選擇。", "計分：開始前確認 20 分、決策 80 分（次佳選擇給一半），每個不適當選擇扣 10 分。通過標準：80 分以上且沒有不適當的選擇。")]:
        rep(a, b)
    s = re.sub(r"有不安全的選擇，臨床上可能[^'`]*。", "有不適當的選擇，臨床上可能影響病人安全或效果，請回頭複習相關單元。", s)
    s = re.sub(r'content\.js\?v=[0-9a-z]*', 'content.js?v=20261127', s)
    if '.chartpaper{' not in s:
        rep('.copyarea{', '.chartpaper{background:var(--surface);border:1px solid var(--line);border-radius:12px;padding:10px 12px;margin:8px 0}\n.copyarea{')
    if m.get('logbook'):
        rep("/* ---------- 畫面 ---------- */", LOGBOOK_CODE)
        rep(" ${(()=>{const n=wrongList().length;", """ <h2>觀察日誌 ${wkTag(1)}</h2>
 <p class="small muted">第 1–2 週跟診時，每天寫一篇觀察日誌，送給督導。</p>
 <button class="card" id="logbtn"><b>寫今天的觀察日誌</b><br><span class="small muted">發生什麼、我的想法、學到什麼、下次怎麼做</span><span class="meta"><span class="small muted">已寫 ${LOGS.length} 篇</span></span></button>
 ${(()=>{const n=wrongList().length;""")
        rep("app.querySelectorAll('.card[data-c]').forEach(b=>b.onclick=()=>simStart(b.dataset.c));",
            "app.querySelectorAll('.card[data-c]').forEach(b=>b.onclick=()=>simStart(b.dataset.c));\n const lb=document.getElementById('logbtn');if(lb)lb.onclick=logForm;")
    open(os.path.join(ROOT, d, 'index.html'), 'w', encoding='utf-8').write(s)

if __name__ == '__main__':
    for d in (sys.argv[1:] or [k for k in COURSES if os.path.exists(os.path.join(ROOT, k, 'content.js'))]):
        build(d, COURSES[d]); print('built', d)
