# 語言治療組 工作入口（Apps Script）

這個資料夾是「語言治療組 工作入口」Apps Script 專案的原始碼（附屬於「ST 雲端白板」試算表）。
合併到 `main` 時，GitHub Actions（`.github/workflows/deploy-apps-script.yml`）會用 clasp
自動推送到 Apps Script，並更新原本的網頁應用程式部署，網址維持不變。

- `Code.gs`：伺服器端程式
- `Index.html`、`Adult.html`、`Daily.html`、`Early.html`：各頁面
- `appsscript.json`：專案設定（僅限臺北醫學大學存取）
- `.clasp.json`：指定要推送到哪個 Apps Script 專案（scriptId）

注意：推送會**整個覆蓋**線上專案，之後請以這裡的檔案為準，不要直接在線上編輯器改。

需要的 GitHub Secret：`CLASPRC_JSON`（在自己電腦執行 `clasp login` 後，`~/.clasprc.json` 的內容）。
