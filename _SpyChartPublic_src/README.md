# SPY 公開 K 線網頁 — 產生器原始碼

公開網頁：https://shuye1668.github.io/spy-intraday-chart/
發布 repo：https://github.com/shuye1668/spy-intraday-chart（工作目錄在 D:\fileserver_D\SpyChartPublic\site）

## 檔案
- `export_candles.py` — history_minute.xlsx（全匯）或 yfinance（--recent，Actions 用）→ data/*.json（只 candles，無 trades）
- `build_site.py` — 從 trade_review_app.py 前端建置靜態 index.html（隱藏所有交易 UI）
- `github_workflow/` — 要放到發布 repo 的 .github/workflows/（需 gh 有 workflow scope 才能推）

## 重建/更新網頁
```
cd D:\fileserver_D\SpyChartPublic
python export_candles.py --from-history D:\fileserver_D\TradeReview\history_minute.xlsx  # 全匯歷史
python build_site.py                                                                      # 重建 index.html（app 前端有改時）
cd site && git add -A && git commit -m "..." && git push
```

## 為什麼這個能完全雲端自動
SPY 是美股，yfinance 從 GitHub 海外 runner 抓得到（台股資料源才會封鎖海外 IP）。
所以 .github/workflows/update.yml 的排程完全不依賴 502/Boss —— 主機全關也照更新。

## 隱私
data/*.json 只含 candles，不含 trades。index.html 的 JS 仍保留原 app 的交易渲染
邏輯（死程式碼，因為 data 無 trades），但不含任何實際交易數字/日期/損益。
