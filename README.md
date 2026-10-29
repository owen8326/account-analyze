# 事件地區分析系統 (Streamlit)

Excel 匯入 → GPS（Excel 座標優先，無效時用地址轉換）→ 時間／單位篩選 → 距離分群 →
事件地圖、熱力圖、Top N 熱點（含 24 小時密集度）、點擊統計 → 匯出 Excel。

## 本機執行
```bash
pip install -r requirements.txt
streamlit run app.py
```

## 部署到 Streamlit Community Cloud
1. 將本資料夾推到 GitHub（`app.py`、`requirements.txt` 放在 repo 根目錄）。
2. 到 https://share.streamlit.io 選擇該 repo，Main file 填 `app.py`。
3. （選用）App → Settings → Secrets 加入：
   ```toml
   GOOGLE_MAPS_API_KEY = "你的 Google Geocoding API Key"
   ```
   有設定則用 Google 轉地址；沒設定則用 OpenStreetMap Nominatim（免費、每秒 1 筆，較慢）。

本機測試時，金鑰放在 `.streamlit/secrets.toml`，**不要** commit 這個檔案。

## .gitignore 建議
```
.streamlit/secrets.toml
__pycache__/
```

## Excel 欄位
A 編號、B 日期、C 管轄單位、D 地點、I 受傷人數、K 經度、L 緯度；第一列為標題。
日期支援 Excel 日期、`2024-05-03 14:30`、民國 `113年5月3日 14時30分`。
