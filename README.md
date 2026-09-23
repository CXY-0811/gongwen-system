# 公文產生系統 V3.1

新增：專案內歷史公文直接套用、自訂公文範本、材料送審快速按鈕、Word 字級規則（一般欄位 12pt；主旨/說明 14pt且不粗體）、Word 左側裝訂線定位修正。

Render：Build `pip install -r requirements.txt`；Start `gunicorn app:app`。

注意：Render Free 的本機 SQLite 可能因重新部署/重啟而遺失，正式長期使用建議改用持久化資料庫。
