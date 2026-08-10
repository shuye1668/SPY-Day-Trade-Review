# TradeReview 遷移指南（移到另一台電腦）

最後更新：2026-07-29。定位：把**目前真正在用的**流程搬到新機，不是復活舊的 Drive/Gmail 自動管線。

---

## 0. 【2026-07-29 已執行】資料夾位置變更

```
舊：C:\TradeReview新：D:ileserver_D\TradeReview\      ← 現行位置，所有文件與排程都已指到這裡
```

搬遷當下做過的事（避免下次再踩）：

| 項目 | 處理 |
|---|---|
| 文件內的路徑（操作手冊／人工SOP／搬遷清單／MIGRATION／SETUP_FREE） | 全部改成新路徑 |
| 資料夾外的 `~/.claude/scheduled-tasks/spy-daytrade-auto-csv/SKILL.md`（06:47 排程） | 已改（**整包複製帶不到，最容易漏**） |
| 周邊工具（daily_fetch／historical_backfill／build_history_xlsx／make_colors／make_dividends／trade_review_app_free／trade_review_local） | 由寫死路徑改成**相對腳本所在目錄**，之後再搬不必改碼 |
| 兩支舊啟動 bat（`SPY_Day_Trade_Review.bat`／`start_app_free.bat`） | `cd /d C:\TradeReview` → `cd /d "%~dp0"` |
| `決策樹.png` 與產圖腳本 | 圖上的路徑文字已重生 |
| 舊資料夾殘留的 app 行程 | 舊資料夾曾有一個 app 同時佔用 5500（兩個 listener），已關閉；`每日一鍵複盤.bat` 的 `killapp` 也改成**殺光所有 5500 listener**，不再只殺第一個 |

**沒有改**（刻意保留）：`archive\`、`app.log`／`_logs\*.log`、`*_backup_*.py`、`HANDOFF_export_label_fix.md`
——這些是歷史紀錄，改了等於竄改；以及 CLAUDE.md 明訂**已停用**的每日拋棄式腳本
（`fix_and_rerun.py`、`process_*.py`、`rebuild_cs_*.py`、`write_616.py`、`migrate_trades.py`），
它們仍寫死舊路徑，**本來就不該再跑**。

> ⚠️ 舊的 `C:\TradeReview\` 若還留著：**不要再從那裡雙擊任何東西**。
> 那邊的 `每日一鍵複盤.bat` 仍會寫入舊資料夾的帳本，兩邊帳一分岔就很難併回來。

---

## 0. 先搞清楚：現在到底哪條流程在跑

| 流程 | 現況 | 遷移策略 |
|---|---|---|
| **手動貼截圖 / Gmail draft → skill → 引擎+writer → app** | ✅ 現役（每交易日） | **這條才要搬**，見下 |
| `C:\TradeReview_Cloud` + `C:\TradeReview_AutoExport` 的 Drive/Gmail 自動管線 | ❌ **已死 70 天**（最後 2026-05-15），且從未真正接通（`process_csv.py` 路徑一直是斷的） | **不建議復活**；要搬的話 blocker 見 §5 |

2026-07-20 引擎化改寫後，每日處理的「唯一真相源」是券商 Account Statement CSV，
`spy_daytrade_engine.py`（分析）+ `spy_daytrade_writer.py`（寫入）已能全自動處理標準日、
異常自動攔截。這讓遷移大幅簡化——**不需要**搬那條三機相依（Drive 磁碟代號＋語系路徑＋OAuth）的舊管線。

---

## 1. 可攜核心：要複製到新機的東西

整個 `D:\fileserver_D\TradeReview\` 資料夾複製過去即可（引擎/writer 路徑已改為「相對腳本所在目錄」，換路徑不必改碼）。

**必要檔**：
- 程式：`trade_review_app.py`、`spy_daytrade_engine.py`、`spy_daytrade_writer.py`、`war_chart.py`
- SOP：`CLAUDE.md`、`SKILL.md`、`ROUTINE_診斷與改寫_20260720.md`、`MIGRATION.md`（本檔）
- 資料：`trades_all.xlsx`、`CS交易紀錄.xlsx`、`history_minute.xlsx`、`colors.xlsx`、`dividends.xlsx`、`events.xlsx`
- 狀態：`offset_state.json`（⚠️ 關鍵，見 §3）
- 目錄：`notes/`、`daily_cache/`

**不必搬**：`archive/`（歷史雜物）、`blank_out/`（手動標註工具產物）、`__pycache__`（自動重生）。

**routine 設定**（若要新機也自動跑）：`~/.claude/scheduled-tasks/spy-daytrade-auto-csv/`——
這是本機 Claude Code 排程，不隨資料夾走，需在新機的 Claude Code 重新建立（見全域 `rules/60-multi-surface.md`）。

---

## 2. 新機一次性設定

```bat
:: 1. 裝 Python（PATH 要有 python）+ 套件
pip install flask pandas numpy openpyxl yfinance

:: 2. 複製整個 D:\fileserver_D\TradeReview 過去（或放任一路徑，程式用相對路徑）

:: 3. 驗證引擎可跑（拿任一天 CSV 乾跑）
python spy_daytrade_engine.py <某日statement.csv> --date YYYY-MM-DD
```

無 Bloomberg 不影響日常（`build_history_xlsx.py` 才需要，K 線缺料時 app 自動用 yfinance 補）。

---

## 3. offset_state.json —— 遷移最容易漏、且會導致錯數字的一項

`Day Header B（CS 每日期初餘額）= 券商當日開盤餘額 + offset`。
`offset_state.json` 存的就是這個 offset 常數（目前 = 0，見 lessons 2026-07-21）。

- offset **只在**留倉建立／跨日平倉／核心持股調整日跳動——而那些日子引擎會攔下（`⚠️需人工確認`）並 abort，
  人工處理完後**手動更新** `offset_state.json` 的 `offset` 值。
- 遷移時**務必**把 `offset_state.json` 一起搬，且確認 `offset` 值與來源機一致。漏搬＝新機第一天期初餘額就錯。

---

## 4. 每日流程（遷移後在新機這樣跑）

1. 取得當日券商 Account Statement：
   - **有 CSV**（Gmail draft 的 `===== CSV CONTENT =====` 段、或本機檔）→ 存成 `.csv`。
   - **只有截圖**（無 draft）→ 見 §4.1 轉成 CSV。
2. 乾跑檢查全綠：
   ```
   python spy_daytrade_engine.py <statement.csv> --date YYYY-MM-DD
   ```
   看 `Σcol1 == Σpnl == 收-開`、無 `BALANCE 對不上`、無 `⚠️`。
3. 有任何 `⚠️` → 依 §0 人工處理，**不 commit**。
4. 全綠 → 寫入（自動備份、寫後 §4 全核驗、失敗自動還原）：
   ```
   python spy_daytrade_writer.py <statement.csv> --date YYYY-MM-DD --commit
   ```
5. 啟動 app 看 K 線：`taskkill /f /im python.exe` → `python trade_review_app.py` → 開 `http://localhost:5500/`。

> writer 已對 2026-07-22/07-23 做過 cell-identical 回歸驗證（與人工核對結果逐格相同），
> 且對「漏單／留倉／BALANCE 對不上」負向測試會正確 abort 不寫入。

### 4.1 無 Gmail draft 的備用路徑：截圖 → CSV

當只有券商 Cash Balance 截圖（無 CSV export、無 draft）時，把截圖**逐列轉成引擎吃的 CSV**：

- 檔頭一行必含：`DATE,TIME,TYPE,REF #,DESCRIPTION,Misc Fees,Commissions & Fees,AMOUNT,BALANCE`
- 每列照截圖抄：`m/d/yy,HH:MM:SS,TRD,="<ref>",BOT +100 SPY @<px>,<misc_fee負值或空>,,<amount>,<balance>`
- BAL 列：`m/d/yy,13:00:00,BAL,,Cash balance at the start of business day DD.MM CST,,,,<開盤餘額>`
- 拆單（同 REF# 多列，如 +20/+27/+53）**照抄成多列**，引擎會自動合併。
- ⚠️ **跨午夜列的日期用「真實日曆日」**：截圖上午夜後（CST 00:xx）那幾筆，`DATE` 欄要寫**次一日曆日**
  （券商 UI 的 Trade Date 欄是營業日、會是前一日，別照抄；引擎靠 DATE 欄算 CST→EDT，寫錯會差一天）。
  例：某營業日盤中最後兩筆顯示在午夜後 → 用它們的實際日曆日期。
- 存成 `.csv` 後接 §4 步驟 2。引擎的 BALANCE 交叉驗證會逐列擋下抄錯（這正是截圖 OCR 最容易錯的地方）。

範例 CSV：`D:\fileserver_D\TradeReview\_inbox\2026-07-22_23_example.csv`（07-22/07-23 實例，含拆單與跨午夜列）。

---

## 5. 舊 Drive/Gmail 自動管線：不建議復活（若堅持要搬的 blocker）

`C:\TradeReview_Cloud`（watch_inbox.py / pull_from_gmail.py / process_csv.py）+ `C:\TradeReview_AutoExport`。
判定**已死且從未真正接通**（詳見 `scratchpad\pipeline_audit.md`）。若仍要在新機救回：

1. **硬阻斷**：`PROCESS_SCRIPT` 路徑指向不存在的 `C:\TradeReview_AutoExport\scripts\process_csv.py`
   （實檔在 `C:\TradeReview_Cloud\`）——這在**舊機上就已經斷**，搬過去一樣斷。
2. **硬阻斷**：Gmail OAuth 需在新機重走同意流程（`token.json` 已 71 天未用、預期失效；且檔內沒記是哪個信箱，
   要用能收到 Boss 端 `[SPY-DayTrade-Autosend]` 郵件的帳號）；需可互動瀏覽器，不能無頭。
3. **硬阻斷**：Google Drive Desktop 路徑三重機器相依——磁碟代號不保證 G:、資料夾名依語系
   （我的雲端硬碟／My Drive／…）、需先登入 sync。`watch_inbox.py:57` 要按新機實際路徑改。
4. 三支 .py + 兩支 .bat 共 11+ 處寫死絕對路徑，無設定層，需逐處改。
5. `credentials.json` / `token.json`（Gmail OAuth，敏感）——搬機要一起帶但預期重新授權；
   決定不救的話，這兩個資料夾可整包歸檔或刪除（含憑證，處置前確認）。
6. `process_csv.py` docstring 提及的觸發器 `tos_export.ahk`（AutoHotkey）在本機找不到，疑在 Boss PC，來源不明。

**建議**：新機直接用 §1–4 的引擎流程，舊管線兩個資料夾（`TradeReview_Cloud`/`TradeReview_AutoExport`）
留在舊機或歸檔即可，不隨遷移復活。
