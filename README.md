# 公文產生系統 v2 修正版

## 本版修正
- 修正「新增公文已寫入資料庫，但儲存後出現 500 Internal Server Error」：改由 INSERT cursor 正確取得 lastrowid。
- 正式公文預覽重新排版，依提供範本調整標題、右上聯絡資訊、受文者、主旨、說明、正副本與裝訂線位置。
- Word 匯出重新排版並指定中文字型「標楷體」。
- 發文字號日期改採民國年月日，例如 2026-09-18 產生 1150918-01，不再使用西元 20260918-01。
- 發文日期以公文日期自動換算民國年，不再固定使用專案年度。

## 部署
把本 ZIP 解壓縮後的全部內容覆蓋／上傳至 GitHub gongwen-system repository，保持 templates 與 static 資料夾結構。Commit 後 Render 可自動重新部署；若未自動部署，於 Render 選 Manual Deploy → Deploy latest commit。

## Render 免費方案提醒
目前系統仍以 SQLite 儲存資料。Render Free Web Service 的本機檔案系統不是永久資料庫，重新部署或執行個體重建時資料可能消失。正式長期使用建議下一版改接 PostgreSQL 或其他持久化資料庫。
