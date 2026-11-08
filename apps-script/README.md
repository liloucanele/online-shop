# Apps Script 後端

這個資料夾是 Lilou 後端（Google Apps Script）的程式。合併到 main 後，GitHub Actions 會自動上傳並更新正式部署，網址不變。

- 密碼、試算表 ID、GitHub 權杖都放在 Apps Script 的「指令碼屬性」，不在這裡。
- 不要再直接在 Apps Script 編輯器改程式；如果改了，請先到 Actions 執行「Apps Script 發布」選 pull，把改動抓回來。
