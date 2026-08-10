# Boss PC（交易機）建置與每日流程

> 對象：在 Boss PC 上做建置的人。
> 這份只講 Boss PC。502 那台的日常操作看《操作手冊》。
> 建立於 2026-08-10。

---

## 0. 這台機器在整套系統裡的角色

依 `CLAUDE.md` 的兩段式架構，**Boss PC 只做第一段**：

```
Boss PC（交易機）                              502 PC
──────────────────                            ──────────────────
TOS → 採集 → _inbox\YYYY-MM-DD.csv
  ├─ engine 乾跑（§1b BALANCE 閘門）
  ├─ writer --trades-only → trades_all.xlsx    ← 第一段到此為止
  ├─ Gmail 草稿（保留，當退路）
  ├─ app（pythonw 常駐，5500）
  └─ sync_push -Owner boss ──→ GitHub ──→ sync_pull → 第二段
                                              cs_from_trades.py
                                              → CS交易紀錄.xlsx（含人工修正）
```

### 🔴 這台機器的鐵則

1. **writer 一定要帶 `--trades-only`**。不帶會去寫 `CS交易紀錄.xlsx` 和
   `offset_state.json` —— 那兩個是 502 的所有物，你會製造出永遠推不出去、
   還會卡住自己 pull 的本地改動。
2. **不要手動編輯任何 .xlsx**。所有寫入只能經由 engine / writer。
3. **`git add -A` 是禁止的**。一律用 `sync_push.ps1 -Owner boss`，
   它只 stage 這台該推的三樣：`trades_all.xlsx`、`_inbox\`、`sync_state.json`。

---

## 1. 一次性建置

### 1.1 需求

| 項目 | 值 |
|---|---|
| 安裝路徑 | `C:\TradeReview`（與採集系統 `C:\TradeReview_Boss` 分開） |
| Python | `C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe` |
| pythonw | `C:\Users\TS USER\AppData\Local\Programs\Python\Python313\pythonw.exe` |
| git | https://git-scm.com/download/win（若尚未安裝） |
| Port | 5500 |
| Repo | `https://github.com/shuye1668/SPY-Day-Trade-Review.git`（**private**） |

### 1.2 執行

用**系統管理員**開 PowerShell：

```powershell
# 先把 bootstrap_boss.ps1 放到桌面（或先手動 clone 再從 repo 內執行）
powershell -ExecutionPolicy Bypass -File .\bootstrap_boss.ps1
```

腳本會依序：前置檢查 → clone → 裝套件 → 註冊開機排程 → 檢查 K 線資料。
第一次 clone 私有 repo 會要求 GitHub 認證（瀏覽器授權或 PAT）。

### 1.3 手動補一件事：K 線資料

`history_minute.xlsx`（4.2 MB）**刻意不走 git** —— 它每天被 app 用 yfinance
追加，進 git 一年會把 repo 撐到 1 GB。

請從 502 的 `D:\fileserver_D\TradeReview\history_minute.xlsx` 用隨身碟或共享
複製到 `C:\TradeReview\`。

不複製也能跑，但 app 首次要靠 yfinance 補 318 天，會很慢。

### 1.4 驗收

```powershell
Start-ScheduledTask -TaskName TradeReview_Dashboard
Start-Sleep 8
(Invoke-WebRequest http://localhost:5500 -UseBasicParsing).StatusCode   # 應為 200
```

重開機後不必手動啟動，排程會自己起（`pythonw` 無視窗）。
Log 在 `C:\TradeReview\_logs\trade_review_app.log`。

---

## 2. 每日採集流程要新增的步驟

現行採集 prompt（`C:\TradeReview_Boss\CLAUDE.md`）跑完步驟 6（取得對帳單資料）之後，
**在建 Gmail 草稿之前**插入本機寫入，然後在最後推送。

### 2.1 新步驟 6.5：寫入本機 trades_all

```powershell
cd C:\TradeReview
$py = "C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe"
$csv = "C:\TradeReview\_inbox\$today.csv"

# 先乾跑：看 Σcol1 與 Σpnl 是否一致、有無 🔴
& $py spy_daytrade_engine.py $csv

# 全綠才寫入（--trades-only！）
& $py spy_daytrade_writer.py $csv --commit --trades-only
```

`--commit` 的結束碼：`0` = 成功或冪等跳過；`2` = 被 gate 擋下（有 🔴）；其他 = 例外。

### 2.2 新步驟 9.5：推送

```powershell
cd C:\TradeReview
.\sync_push.ps1 -Owner boss
```

### 2.3 被 gate 擋下時怎麼辦（無人看管的正確行為）

**什麼都不要改，讓它擋著。** 這是 `CLAUDE.md §0` 的核心：寧可今天不寫，
也不要為了跑完流程而寫入猜值。

會被擋下的情況：`BALANCE 對不上`、`收盤留倉`、`非交易現金異動（DOI/JRN）`、
`session 開盤/收盤餘額不明`。

此時：
- `trades_all.xlsx` 停在前一日（正確，不是壞掉）
- Gmail 草稿照建，正文寫明被擋原因 → 502 那邊看得到
- 502 走既有的人工 SOP 處理該日
- log 記 `FAIL - <step> - <reason>`

---

## 3. 採集 prompt 必須修改的兩條禁止事項

現行 prompt 的「禁止事項」寫著：

> * 不操作 `C:\TradeReview\` 下任何檔案（那是使用者電腦端的）
> * 不操作 trade_review_app

**這兩條與新流程正面衝突**，必須改寫為：

> * 只操作 `C:\TradeReview\_inbox\` 與 `trades_all.xlsx`，且**只透過
>   spy_daytrade_engine.py / spy_daytrade_writer.py --trades-only**；
>   不得手動編輯任何 .xlsx
> * 不得寫 `CS交易紀錄.xlsx` / `offset_state.json`（502 的所有物）
> * 不得修改 `trade_review_app.py`；不得為了讓流程跑完而修改任何數字（§0）

---

## 4. ⚠️ 尚未解決：對帳單 CSV 的欄位格式不相容

TOS 直接匯出的 CSV 表頭是 **10 欄**：

```
Trade Date,Exec Date,Exec Time,Type,Ref #,Description,Misc Fees,Commissions & Fees,Amount,Balance
```

但引擎要的是 **9 欄**（`spy_daytrade_engine.py:150` 找不到就直接拒收）：

```
DATE,TIME,TYPE,REF #,DESCRIPTION,Misc Fees,Commissions & Fees,AMOUNT,BALANCE
```

`_inbox\2026-08-05.csv` 就是 10 欄格式，引擎目前對它報
`找不到 Cash Balance 表頭 — Statement 格式不符`。

**這很可能就是 2026-07-17 停用 CSV 匯出、改走 OCR-only 的原因** ——
OCR 路徑是由 AI 在轉錄時順手把 10 欄併成 9 欄（Trade Date/Exec Date → DATE、
Exec Time → TIME），所以感覺「OCR 能用、CSV 不能用」。

→ 恢復 CSV 匯出路徑時**必須加一支格式轉換**（10 欄 → 9 欄）。
這件事還沒做，屬於階段 3/4。

### 另一個已知缺口：OCR 漏抓 BAL / DOI / JRN 列

現行 prompt 步驟 6 只抓 `TRD` 列。但引擎靠 `BAL` 列切 session、取期初餘額、
跑 §1b 逐列餘額驗證（`spy_daytrade_engine.py:191-203`）。沒有 BAL 列 →
`session 開頭缺 BAL 列` → writer 一定 abort。

`offset_state.json` 自己記著這件事已經發生過：

> ⚠️注意:8/05 open-bal 115851.99 是由成交鏈回推(今日OCR草稿漏BAL/DOI/JRN列)

→ OCR 備援路徑必須改成「Cash Balance 區塊**每一列都抓**」（含 BAL/DOI/JRN/ADJ），
步驟 7 的 `qty×price≈|amount|` 驗算只對 `TRD` 列做，其餘列原樣帶入。

---

## 5. 檔案所有權對照表

| 檔案 | 擁有者 | Boss PC 可以…… |
|---|---|---|
| `_inbox\*.csv` | **Boss** | 寫、推 |
| `trades_all.xlsx` | **Boss** | 經 writer `--trades-only` 寫、推 |
| `sync_state.json` | **Boss** | 寫、推 |
| `CS交易紀錄.xlsx` | 502 | **唯讀，絕不可寫** |
| `offset_state.json` | 502 | **唯讀，絕不可寫** |
| `notes\` | 502 | 唯讀 |
| `*.py` `*.md` `*.bat` `*.ps1` | 502 | 唯讀（改程式碼在 502 做，Boss pull 即可） |
| `history_minute.xlsx` | 各機獨立 | 本機自行追加，不進 git |

`sync_pull.ps1 -Owner boss` 會自動丟棄本機對「502 所有物」的任何改動，
所以就算誤跑了完整 writer，下次 pull 也會清乾淨、不會卡住。

---

## 6. 疑難排解

| 狀況 | 處理 |
|---|---|
| `.ps1` 一執行就一堆語法錯誤 | PowerShell 5.1 讀沒有 BOM 的 UTF-8 會當 ANSI。跑 `python _fix_ps1_bom.py` |
| `git pull` 說 unstaged changes | 跑 `.\sync_pull.ps1 -Owner boss`，它會先丟棄對方所有物再 pull |
| port 5500 被佔用 | `Get-NetTCPConnection -LocalPort 5500 -State Listen` 找出 PID 後處理 |
| app 起不來但沒有錯誤訊息 | pythonw 無主控台，看 `_logs\trade_review_app.log` |
| writer 說「鎖檔存在（Excel 開著？）」 | 有人開著 xlsx，關掉再跑 |
| writer 印 🔴 | **正常保護**。照第 2.3 節處理，不要改數字 |
