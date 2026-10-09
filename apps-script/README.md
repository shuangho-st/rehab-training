# Apps Script（資料上傳後端）

這個資料夾的程式碼會在合併到 `main` 時，由 GitHub Actions
（`.github/workflows/deploy-apps-script.yml`）自動用 clasp 推送到 Google Apps Script，
並更新原本的網頁應用程式部署，網址維持不變。

- `Code.gs`：後端程式碼
- `appsscript.json`：專案設定（manifest）
- `.clasp.json`：指定要推送到哪個 Apps Script 專案（scriptId）

注意：推送會**覆蓋**線上編輯器裡的程式碼，之後請以這裡的檔案為準，不要直接在線上編輯器改。

需要的 GitHub Secret：`CLASPRC_JSON`（在自己電腦執行 `clasp login` 後，`~/.clasprc.json` 的內容）。
