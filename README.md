# 公文產生系統｜可部署網頁版

依使用者提供的「徐英豪建築師事務所 函」範例製作。核心流程：專案 → 公文 → 固定發文字號前綴＋流水號 → 預覽 → Word。

## 直接部署

### Render
1. 將本資料夾上傳到 GitHub。
2. 在 Render 建立 Web Service，連接此 GitHub repository。
3. Build Command：`pip install -r requirements.txt`
4. Start Command：`gunicorn app:app`
5. Python runtime。
6. 若使用 `render.yaml`，可直接依設定建立服務與持久磁碟。

部署完成後 Render 會提供 `https://xxxx.onrender.com` 網址，使用者只要開網址即可，不需安裝 Python。

## 本機測試（可選）
`pip install -r requirements.txt` → `python app.py` → 開 `http://127.0.0.1:5000`

## 注意
目前版本以 SQLite + Render persistent disk 做單機型部署，適合小型內部使用。若需要多人同時大量使用，建議改 PostgreSQL。
