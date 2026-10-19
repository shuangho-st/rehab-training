# 年度排班許願問卷（獨立 Apps Script）

獨立於「工作入口」的網頁應用程式，附屬於「排班許願問卷（組長專用）」試算表。
合併到 `main` 時，`.github/workflows/deploy-wish-survey.yml` 會推送程式並更新同一個部署，網址不變。
第一次執行時會自動建立試算表與程式，並把 `.clasp.json`（scriptId）和 `DEPLOYMENT_ID` 存回這個資料夾。

- `Code.gs`：伺服器端（名單 `ROSTER` 改這裡）
- `Index.html`：問卷頁面與組長彙整
- 每人連結：`.../exec?k=代碼`，在試算表「名單」分頁 D 欄，組長的「彙整與試排」也看得到
