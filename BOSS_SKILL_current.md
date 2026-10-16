---
name: tos-daily-collector
description: 美股交易日的台灣早上 06:00 TOS 採集；有新成交才寄給 juliahsu13@gmail.com
---

# 每日 TOS 採集 routine（Boss PC，權威版）

> 本檔是自足的 —— 執行時**不需要**再去讀別的規則檔。
> 最後更新：2026-10-16。
>
> ⚠️ **規則只能有一份。** 2026-08-11 查出的事故：這個排程原本內嵌 2026-07-17 的
> OCR-only 流程，同時 `C:\TradeReview_Boss\CLAUDE.md` 另有一份較新的規則，兩者
> 互相矛盾（前者明文禁止 CSV 匯出與寫入 `C:\TradeReview\`）。直接指令會贏，
> 導致本機寫入與 git 同步從未執行、每天都要人工補帳。
> **要改流程請改本檔，不要在別處另寫一份。** 改完要做兩件事：
> ① 把正文整份複製回 `C:\TradeReview_Boss\CLAUDE.md`（開頭說明保留）；
> ② 複製本檔到 `C:\TradeReview\BOSS_SKILL_current.md`（步驟 9.5 每次推送前也會自動做），502 端母版靠它刷新。

---

## 這台機器在整套系統裡的角色

依兩段式架構，**Boss PC 只做第一段**：對帳單 → `trades_all.xlsx`。

```
Boss PC（這台，C:\TradeReview）                 502（D:\fileserver_D\TradeReview）
──────────────────────────────                 ────────────────────────────────
TOS → 採集（CSV 為主／OCR 備援）
  ├─ engine 乾跑（BALANCE 閘門）
  ├─ writer --trades-only → trades_all.xlsx     ← 第一段結束
  ├─ write_sync_state.py（新鮮度戳記）
  ├─ Gmail 草稿（保留為退路）
  └─ sync_push -Owner boss ─→ GitHub ─→ sync_pull ─→ 第二段
                                                       cs_from_trades.py
                                                       → CS交易紀錄.xlsx（人工修正在這）
```

**Email（2026-10-16 起）**：只有出現「尚未寄過的新成交」時，才建草稿並**由本 routine 直接寄出**到
**juliahsu13@gmail.com**（見步驟 8～9）。沒有新成交就不建草稿、不寄信。
（舊做法「草稿交給 Apps Script 寄到 sunbeamichelle」已停用：Apps Script 自 2026-05-14 起
就沒有寄出過任何一封這類草稿。）

### 🔴 三條鐵則

1. **writer 一定要帶 `--trades-only`。** 不帶會去寫 `CS交易紀錄.xlsx` 與
   `offset_state.json` —— 那是 502 的所有物，會製造出永遠推不出去、還會卡住
   自己 pull 的本機改動。
2. **不得手動編輯任何 `.xlsx`。** 所有寫入只能經由 engine / writer。
3. **不得用 `git add -A`。** 一律用 `sync_push.ps1 -Owner boss`。

---

## 執行時機（這台週末關機）

排程：**台灣週二～週六 06:00**（cron `0 6 * * 2-6`，對應美股週一～週五收盤後；排程器會加幾分鐘延遲）。
Windows 工作排程器的 `Claude-TOS-Collector-Launch` 在同一天 **05:45** 喚醒電腦並開啟 Claude Desktop，
讓這個排程能觸發（Claude 沒開著的話，排程不會跑）。美股休市日由步驟 0 判斷後直接結束。
機器關機（不是睡眠）時喚醒不了，那天就不會跑，開機後手動補跑即可。所以：

- **一次跑會補完所有欠帳**：步驟 4 要求「至少 5 days back」（平常是 7）不是保險，是必要的。
  writer 會處理 CSV 裡所有 session，已寫過的自動冪等跳過。
- **週一落後一個交易日是正常的**，不是故障。502 的 App 標頭那段時間會顯示
  琥珀色「⚠待補前一交易日」，補完轉回灰色。
- 顯示**紅色「✕落後 N 個交易日」**（N≥2）才代表真的有事沒跑成。

---

## 執行環境（會踩到的四個坑）

- **Claude Desktop 會搶走前景視窗**（2026-10-16 實測）：`Activate-TOSWindow` 回報成功後，
  Claude 視窗常在一兩秒內又跳回前面，點擊就落在 Claude 上而不是 TOS（症狀：截圖出現 Claude 畫面、
  選單沒開、表格某列被選取）。**每一次點擊 TOS 之前，都要在同一個 call 裡先跑
  `Activate-TOSWindowVerified`**（反覆啟用直到前景真的是 TOS；回傳 False 就不要點）。
  需要連續操作的（開選單→選項目）一律放在同一個 call。

- **此 session 不會有 desktop computer-use MCP。**
  不要 tool_search 找 computer use、不要用 `Claude_in_Chrome__computer`
  （那個要 browser tabId，控制不了桌面程式）。一律走 PowerShell + Win32，
  所有需要的功能 `tos_automation.ps1` 都有。

- **ExecutionPolicy 會擋 `.ps1`**（本機全 scope 為 Undefined＝實際 Restricted）。
  每個需要 dot-source 的 call 開頭都要先下 `Set-ExecutionPolicy -Scope Process`；
  執行獨立腳本則用 `powershell -ExecutionPolicy Bypass -File`。

- **每次 PowerShell tool call 都是新 process**，變數不跨 call 保留。
  每個 call 都要重新 dot-source，重要狀態要落檔。

工作流程：每個關鍵步驟後 `Capture-Screen` 存 PNG 到 `C:\TradeReview_Boss\logs\`，
再用 view tool 讀圖確認畫面，確認後才進下一步。

### 常用變數

```powershell
$py    = "C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe"
$root  = "C:\TradeReview"          # 複盤系統（git repo）
$boss  = "C:\TradeReview_Boss"     # 採集系統（本檔相關資源）
$today = (Get-Date).ToString("yyyy-MM-dd")
```

---

## 【步驟 0】前一晚美股有沒有開盤

```powershell
$py = "C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe"
& $py C:\TradeReview_Boss\us_trading_day.py
```

- `US_SESSION=YES` → 繼續步驟 1。
- `US_SESSION=NO`（週末或 NYSE 休市）→ **什麼都不碰**（不動 TOS、不建草稿、不寄信、
  不跑 writer／sync_state），只在 log 寫一行
  `[時間] SKIP - no US session (<日期> <原因>)` 後結束。
- 輸出含 `WARNING: no NYSE holiday list` → 照常繼續，但在最終回報提醒要更新 `us_trading_day.py` 的休市日清單。
- 手動補跑（例如週一開機補週五）時不受此限：直接從步驟 1 開始。

**沒有新成交時「回復原樣」**：步驟 8 判定 `NEW_TRADES=0` → 不建草稿、不寄信，
照樣做步驟 10 把 TOS 切回 Charts（採集前的看盤狀態），寫 log 後結束。

---

## 【步驟 1】載入 helper + 取得今天日期

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview_Boss\tos_automation.ps1
$today = (Get-Date).ToString("yyyy-MM-dd")
Write-Host "Today: $today"
```

預期看到：

```
tos_automation.ps1 loaded. Available functions:
  Find-TOSWindow / Activate-TOSWindow / Get-TOSWindowRect / Get-TOSWindowTitle
  Click-At / TripleClick-At / Send-Keys / Send-Hotkey
  Capture-Window / Capture-Screen / Capture-Region
Today: 2026-XX-XX
```

---

## 【步驟 2】啟用 TOS 視窗

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview_Boss\tos_automation.ps1
$ok = Activate-TOSWindowVerified
Write-Host "Activated: $ok"
if ($ok) { Capture-Screen "C:\TradeReview_Boss\logs\dbg_tos_activated_$today.png" }
```

- `$ok = False` → 存 `logs\debug_no_tos_$today.png`，建 FAILED 草稿
  「TOS window not found」後結束。
- `$ok = True` → view 該 PNG，確認 TOS 在前景且最大化、顯示交易介面（不是登入頁）。

---

## 【步驟 3】切到 Account Statement

從步驟 2 截圖判斷當下狀態：

- **情境 1**：已在 Monitor → Account Statement → 跳到步驟 4
- **情境 2**：在 Monitor 但不是 Account Statement 子頁籤
  → 從截圖找到該子頁籤座標 → `Click-At <x> <y>` → 等 1 秒截圖確認
- **情境 3**：在其他主頁籤（上次收尾會停在 Charts，這是正常起點）
  → 同一個 call：`Activate-TOSWindowVerified` → `Click-At 262 33`（Monitor）→ 等 1 秒
  → `[Win32Api]::SetCursorPos(1000, 900)`（游標移開，主頁籤 hover 選單才會收起）→ 截圖
  → 通常會直接停在 Account Statement 子頁籤；若不是，再點子頁籤 `Click-At 455 55`（同樣先 Verified）

座標（2026-10-16 實測，TOS build 1993，1920x1080）：主頁籤列 **y≈33**，
Monitor=262、Charts=557；子頁籤列 y≈55，Activity and Positions≈318、Account Statement≈455。
新版本 TOS 可能再變，**以當次截圖為準**。

---

## 【步驟 4】確認日期範圍「至少 5 days back」

⚠️ 要 ≥5 天（涵蓋週末＋假日＋補跑）。**N≥5 就不要動**：TOS 這台平常顯示
`7 days back from today`，7 天已涵蓋需求。2026-10-02 曾為了把 7 改成 5 在輸入框打字，
**TOS 跳出「Unexpected error」**（按 Continue 可恢復，**絕不按 Restart**，並取消勾選回報）。

**4.1 觀察 status bar**（Account Statement 上方）：

| 看到 | 情境 |
|---|---|
| `N days back from today`，N≥5（常見是 7） | A → **不動**，跳步驟 5 |
| `1`～`4 days back from today` | B → 要改 |
| `from X to Y` 自訂範圍且涵蓋不到最近 5 天 | B → 改回 days back |
| 完全找不到 | C |

**4.3 情境 B：改成 ≥5 days back**（只有真的不夠才做；優先用輸入框旁的上下箭頭調整，
每點一次截圖確認，**避免直接打字**；下方 a～h 的打字做法是最後手段）

a) 從截圖找出「N days back from today」文字的精確座標。
   1920x1080 下 y 軸約 78-80、x 軸約 600-700。
   **y 軸差 3-5 px 就會 miss（實戰教訓）**，不確定先 `Capture-Region` zoom。

b) 點該文字：
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview_Boss\tos_automation.ps1
Click-At <x> <y>
Start-Sleep -Milliseconds 500
Capture-Screen "C:\TradeReview_Boss\logs\dbg_dialog_$today.png"
```

c) view 確認對話框已開，預期長相：
```
Statement for:                    [Reset]
◉  [1] days back from [Today ▼]
○  from [MM/DD/YYYY] to [MM/DD/YYYY]
Maximum period length is 370 days.
```
沒開就 zoom-in 重找座標再點。

d) 確認上方 radio 被選中（**不要動下方 from/to**）。

e) 三連點輸入框 + 輸入 5 + Enter（單一 call）：
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview_Boss\tos_automation.ps1
TripleClick-At <input_x> <input_y>
Start-Sleep -Milliseconds 200
Send-Keys "5"
Start-Sleep -Milliseconds 100
Send-Keys "{ENTER}"
```

f) `Start-Sleep 8` 等表格重載
g) `Send-Keys "{ESC}"` 關掉殘留 tooltip
h) `Capture-Screen` + view 驗證 status bar 已變成 `N days back from today`（N≥5）

**4.4 情境 C**：存 `logs\debug_no_status_bar_$today.png`，建 FAILED 草稿
「Cannot find date range status bar」後結束。

---

## 【步驟 5】★ 主路徑：CSV 匯出

**2026-08-10 起改回 CSV 為主。** 引擎已能直接吃 TOS 原生 10 欄格式
（`Trade Date,Exec Date,Exec Time,Type,Ref #,Description,Misc Fees,Commissions & Fees,Amount,Balance`），
**不需要任何轉換** —— 這正是當初停用 CSV 路徑的原因，已在引擎端修好。

log 實證：2026-05-13～07-16 走 CSV 連續 **44 次成功**（max diff 多在 $0.004 以下）；
07-17 之後改 OCR，覆蓋率反而降到約 6 成。CSV 才是穩的那條。

**5.1 開匯出對話框**（單一 call；2026-10-16 實測可行的唯一做法）：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview_Boss\tos_automation.ps1
$ok = Open-ExportDialog      # Verified 啟用 → Click-At 1899 60 開選單 → 鍵盤 7×DOWN + ENTER
Write-Host "dialog open: $ok"
Capture-Screen "C:\TradeReview_Boss\logs\dbg_savedialog_$today.png"
```

- 選單共 8 項，「Export to file...」是第 7 項。**用滑鼠點選單項目不會觸發**（實測多次失敗），
  只能用鍵盤；開選單則要用 `Click-At`（太快的原始點擊開不了）。
- view 截圖確認對話框標題是 `Account Statement for ... since M/D/YY through M/D/YY`
  （這代表點到的是正確的三條線），資料夾 `C:\TradeReview\_inbox`、類型 `CSV (*.csv)`、
  檔名預填 `<today>-AccountStatement`。

**5.2 存檔 + 改名**：⚠️ **不要在對話框裡打字**——這台輸入法是注音，數字會被吃成中文字。
接受預設檔名，按存檔後再用 PowerShell 改名：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview_Boss\tos_automation.ps1
Confirm-ExportSave           # 把對話框叫到前景（確認後）點 存檔(S) (1155,775)
Move-Item "C:\TradeReview\_inbox\$today-AccountStatement.csv" "C:\TradeReview\_inbox\$today.csv" -Force
```

**5.3** `Capture-Screen` + view 確認：Save dialog 已關、沒有「檔案已存在是否覆蓋」殘留。

**5.4** 確認檔案產生且非空：
```powershell
$f = "C:\TradeReview\_inbox\$today.csv"
if (Test-Path $f) { (Get-Item $f).Length } else { "MISSING" }
```

- 存在且 **> 500 bytes** → **跳到步驟 6.5**（不需要 OCR）
- 否則 → 走步驟 5'

### ⚠️ 嚴格限制重試次數

2026-10-01、10-16 已實機驗證可行；仍限制重試，避免在 TOS 裡亂點：

- 找不到三條線／選單裡沒有 Export to file → **最多找 2 次**
  （可 `Capture-Region` zoom 右上角再確認一次），找不到就**直接走步驟 5'**
- Save dialog 沒出現或存檔失敗 → **最多重試 1 次**，仍失敗走步驟 5'
- **任何情況都不要去點 TOS 其他功能表探索**

降級到 OCR 是完全正常的結果，不是失敗。log 與草稿的 `Source:` 欄要註明走了哪條，
累積幾天就知道 CSV 路徑能不能用。

---

## 【步驟 5'】備援路徑：截圖

⚠️ 只有步驟 5 失敗才走這裡。

**5'.1** 關掉 tooltip／對話框：
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview_Boss\tos_automation.ps1
Send-Keys "{ESC}"
Start-Sleep 1
```

**5'.2** `Capture-Screen "C:\TradeReview_Boss\logs\screenshot_$today.png"`

**5'.3** view 確認：
- 看得到「N days back from today」（N≥5）
- 看得到 Cash Balance 表格多列資料
- 沒有對話框或 tooltip 遮住
- ⚠️ **最關鍵**：確認「最後一列（最新交易）」有入鏡。Cash Balance 通常最舊在上、
  最新在下；底部被截斷＝有 overflow → 見 5'.4

**5'.4 列數 overflow 時**：在表格內側點一下取得焦點（點資料列空白處，不要點任何
按鈕/連結），按 End 捲到最底：
```powershell
Click-At <x_table_body> <y_table_body>
Send-Keys "{END}"
Start-Sleep 1
Capture-Screen "C:\TradeReview_Boss\logs\screenshot_${today}_b.png"
```
一張截不完就「捲一段截一張」（`_b.png`、`_c.png`…），步驟 6 OCR 後靠 Ref#／時間
去重合併，確保每列都涵蓋且沒重複計入。合併後仍無法確認完整 → 建 FAILED 草稿
「OCR rows overflow, cannot capture all rows」後結束。

---

## 【步驟 6】OCR（僅備援路徑需要）

### 🔴 最重要：Cash Balance 區塊的「每一列」都要抓

**不是只抓 TRD。** 必抓類型：

| TYPE | 意義 | 為什麼非抓不可 |
|---|---|---|
| `TRD` | 成交 | 主資料 |
| **`BAL`** | 期初餘額 | **引擎靠它切交易日、取期初餘額、跑逐列餘額驗證。沒有它引擎必定 abort，當天一筆都寫不進去** |
| **`DOI`** | 配息 | 影響餘額鏈，漏抓會讓 BALANCE 對不上 |
| **`JRN`** | 轉帳 | 同上 |
| `ADJ` 等 | 調整 | 同上 |

**這件事已經出事過。** `offset_state.json` 自己記著：

> ⚠️注意:8/05 open-bal 115851.99 是由成交鏈回推(今日OCR草稿漏BAL/DOI/JRN列)，
> 非直接讀BAL行 ... 唯 offset 數值為暫定值，待補齊含BAL列的完整對帳單後覆核

`BAL` 列通常在表格最上面一列，很容易被當標題略過：
```
7/31/26,13:00:00,BAL,,Cash balance at the start of business day 31.07 CST,,,,36372.84
```
`DOI`（配息）長這樣：
```
7/31/26,16:39:50,DOI,126779764549,STATE STREET SPDR S&P 500 ETF TRUST 5900.9 US$,,,5900.90,42273.74
```

### 每列要抓的欄位

```
Trade Date / Exec Date / Exec Time / Type / Ref# / Description
/ Misc Fees / Comm & Fees / Amount / Balance
```

⚠️ **兩個日期欄都要抓。** Trade Date 是券商營業日，午夜後成交的列會掛在前一個
營業日；**Exec Date 才是真正的日曆日**。實測 2026-07-29 的對帳單有 5 列
`Trade Date=7/27` 但 `Exec Date=7/28 01:18` —— 取錯就整整差一天。
引擎會自動取 Exec Date，照實抓兩欄即可。

### OCR 注意事項

- Description 的 qty（`+100` / `-100` / 分批 `+80` `+20`）與 price（`@xxx.xxxx`）要看清楚
- Amount / Balance 有千分位逗號，去掉再存
- 分批成交（同 Ref#、同秒、qty 拆兩列）各存一列
- 非 SPY 標的照樣抓，Description 保留原標的字樣
- 任一 **TRD** 列的 qty/price/amount 辨識不清 → 見步驟 7 失敗處理

### 寫成 CSV

寫入 `C:\TradeReview\_inbox\<today>.csv`，第一行欄位表頭、之後每列一行、逗號分隔。
9 欄或 10 欄格式都可以，引擎依欄名對應。

---

## 【步驟 6.5】★ 寫入本機 trades_all

```powershell
$py  = "C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe"
$csv = "C:\TradeReview\_inbox\$today.csv"
cd C:\TradeReview

# 先乾跑，不寫任何檔案
& $py spy_daytrade_engine.py $csv

# 全綠才寫入
& $py spy_daytrade_writer.py $csv --commit --trades-only
$wr = $LASTEXITCODE

# 記錄狀態供 502 的 App 顯示新鮮度
# ⚠️ 不要帶 --date $today。這台週末關機，「今天」不等於「資料涵蓋到哪一天」；
#    省略 --date 會自動讀 trades_all 裡真正的最新交易日。
if ($wr -eq 0) {
    & $py write_sync_state.py
} else {
    & $py write_sync_state.py --status blocked --note "writer gate 擋下，需人工確認"
}
```

**「全綠」判準**（引擎輸出）：最後一行 `Σcol1` 與 `Σpnl` **兩數字相同**、
`留倉=False`，且下面**沒有**「警告／需人工確認」那一段。

**writer 結束碼**：`0`＝成功或冪等跳過；`2`＝被 gate 擋下（有 🔴）；其他＝例外。

**catch-up 是自動的**：writer 會處理 CSV 裡所有 session（步驟 4 的「至少 5 days back」
正是為此），已寫過的自動跳過。週一開機一次跑完會把週五（甚至更早漏掉的）一起補上。

### 被 gate 擋下時：什麼都不要改

`trades_all` 停在前一日**是正確的，不是壞掉**。照常繼續步驟 8 建草稿，把被擋原因
寫進正文，交由 502 人工處理。

會被擋下的情況：`BALANCE 對不上`、`收盤留倉`、`非交易現金異動（DOI/JRN）`、
`session 開盤/收盤餘額不明`。

> 這是最高原則：**絕不為了跑完流程而寫入估算／推算／猜值。**

---

## 【步驟 7】驗算

對每一列 **TRD** 驗算 `qty × price ≈ abs(amount)`，容差 **$0.02**。

⚠️ **`BAL` / `DOI` / `JRN` / `ADJ` 列不做這個驗算** —— 它們沒有 qty/price，
驗算必然失敗。這些列原樣帶入 CSV 即可，引擎會自行處理：`BAL` 取期初餘額；
`DOI`/`JRN` 會被 gate 標為「非交易現金異動」而擋下該日交人工判斷
（**這是正確行為，不是錯誤**）。

- 逐列算 `diff = |qty × price − abs(amount)|`，累積 maxDiff
- `diff > $0.02` 的 TRD 列 → 草稿主旨改 FAILED，正文記該列詳情後結束
- 任一 TRD 列 qty/price/amount 讀不出來 → 建 FAILED 草稿
  「OCR unreadable row: `<定位資訊>`」後結束

> 走 CSV 主路徑時這是額外保險（資料直接來自 TOS，不會有 OCR 誤判）；
> 走 OCR 備援時**這是抵擋誤判的主要防線**，務必逐列都算，不要抽樣。

---

## 【步驟 8】判斷有沒有新成交 + 防重複

**8.1 有沒有「沒寄過的新成交」**（對帳單一次涵蓋 7 天，所以不能只看有沒有 TRD 列，
否則同一筆交易會連續寄好幾天）：

```powershell
$py = "C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe"
$env:PYTHONIOENCODING = "utf-8"
& $py C:\TradeReview_Boss\check_new_trades.py "C:\TradeReview\_inbox\$today.csv"
```

- 輸出 `NEW_TRADES=0` → **不建草稿、不寄信**，跳到步驟 9.5（git）／10／11，
  log 註記 `no new trades, no email`。
- 輸出 `NEW_TRADES=N`（N≥1）→ 繼續 8.2 與步驟 9。
- 已寄過的成交記在 `C:\TradeReview_Boss\logs\reported_trades.txt`（每筆一行），
  **只在信寄出成功後**用 `--mark` 寫入（見 9.2）。不要手動編輯這個檔。

**8.2 防重複**：用 Gmail MCP 搜尋草稿 `subject:"[SPY-DayTrade-Autosend] $today"`，
若有殘留的當日草稿（不論成功或 FAILED）**先刪除**。

---

## 【步驟 9】建草稿並寄出（只有 NEW_TRADES≥1 才做）

用 Gmail MCP 的 `create_draft` 建草稿，再用 `send_message` 帶 `draftId` 寄出：

- **收件人**：`juliahsu13@gmail.com`
- **主旨**：`[SPY-DayTrade-Autosend] $today`
- **正文**（嚴格純文字，不要 HTML）

**9.0 先產生對帳單文字表**（取代截圖：不附圖片，用純文字對齊表格呈現
TOS 畫面上的 Cash Balance／Order History／Trade History，人和另一個 Claude 都能直接讀）：

```powershell
$py = "C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe"
$env:PYTHONIOENCODING = "utf-8"
& $py C:\TradeReview_Boss\render_statement.py "C:\TradeReview\_inbox\$today.csv" |
    Out-File "C:\TradeReview_Boss\logs\statement_$today.txt" -Encoding utf8
```

把 `logs\statement_<today>.txt` 的輸出**原樣**貼進下方 STATEMENT TABLE 區塊（不要手改、不要重排）。
⚠️ 不要附截圖或任何圖片附件 —— base64 附件太大、容易抄壞（2026-10-16 決定）。
OCR 備援路徑時同樣對 OCR 產出的 CSV 跑這支（只有 Cash Balance 一段，其他段會顯示 section not found，正常）。

【步驟 9.0b】驗證截圖本機留存（2026-10-16 起，當日 CSV 含至少一列 TRD 才做）
目的：502 端把圖貼進 Julia 的 Excel 母檔，供讀者對照 TOS 原始畫面核驗。 不附進任何信件或草稿（base64 超過可靠處理長度）。純 BAL 無交易日略過。

擷取範圍：Cash Balance（含 BAL/TRD/DOI/JRN/TOTAL 每一列）＋ Order History（全部 FILLED 列）＋ Trade History（全部 fill 列），無遮擋。優先從流程中已確認過的全螢幕 截圖裁切，避免重拍時視窗搶前景；不動日期範圍欄位（改 days back 曾讓 TOS 當掉）， 現有範圍涵蓋即可。
存 C:\TradeReview_Boss\logs\TOS_verify_$today.png（放不下分段 _b/_c）， 再複製到 C:\TradeReview\verify_shots\（不存在就建；不放 _inbox\，那裡只放 CSV）。
失敗一律非致命：log 註記 Screenshot: FAILED - <原因>，流程照常。

```
TOS Account Statement Daily Export
Date: $today
Generated: <現在時間 ISO8601，例 2026-08-12T06:05:14+08:00>
Boss PC: $env:COMPUTERNAME
Source: <CSV export | screenshot OCR>
Local write: <OK | BLOCKED: 原因 | FAILED: 原因>
New trades: N（貼上 check_new_trades.py 列出的新成交，每筆一行）

===== STATEMENT TABLE START =====
<把 logs\statement_<today>.txt 的內容整段貼上>
===== STATEMENT TABLE END =====

===== CSV CONTENT START =====
<把 _inbox\<today>.csv 的內容整段貼上>
===== CSV CONTENT END =====

Validation: N TRD rows, all pass qty*price=abs(amount) within $0.02 tolerance (max diff $X).

--
Auto-generated by Claude Code on Boss PC
```

**9.2 寄出 + 標記**：

1. `send_message`（只帶 `draftId`）寄出剛建的草稿。收件人只能是 juliahsu13@gmail.com。
2. 寄出成功後才標記這些成交為已寄：
   ```powershell
   & $py C:\TradeReview_Boss\check_new_trades.py "C:\TradeReview\_inbox\$today.csv" --mark
   ```
3. 寄出失敗 → **不要** `--mark`（下次 routine 會重寄），草稿留著，log 記 `SEND FAILED`。

⚠️ `Source:` 與 `Local write:` 這兩行很重要 —— 收件人從信裡就能看出走了哪條路徑、
本機帳做成了沒。

---

## 【步驟 9.5】★ 推送到 GitHub

```powershell
# 先把現行 SKILL 全文複製進 repo（502 端母版靠這個檔刷新；每次推送前都做，確保永遠是最新版）
Copy-Item "C:\Users\TS USER\.claude\scheduled-tasks\tos-daily-collector\SKILL.md" "C:\TradeReview\BOSS_SKILL_current.md" -Force
powershell -ExecutionPolicy Bypass -File C:\TradeReview\sync_push.ps1 -Owner boss
```

只會推 `_inbox\` / `verify_shots\` / `BOSS_SKILL_current.md` 三樣（2026-10-16 所有權定案）。
`trades_all.xlsx` 與 `sync_state.json` **歸 502**：本機照樣由 writer／write_sync_state 寫入供乾跑驗證，
但**不推送**；新版 sync_pull 會在每次 pull 前丟棄這兩個檔的本機改動，以 502 版本為準。
⚠️ 2026-08-11 起 repo 卡在未完成的 interactive rebase（`.git\rebase-merge` 存在）：
存在時**不要跑 sync_push**（只會在 detached HEAD 上堆孤兒 commit），log 記 `PUSH SKIPPED (repo in rebase)`，
**不要自行修復**，等使用者手動處理。
push 失敗會自動重試 3 次；三次都失敗時本機 commit 仍在（**資料沒掉**），
但要把這件事寫進步驟 11 的 log 與 Gmail 草稿正文。

---

## 【步驟 10】收尾：切回 Charts

目的：採集完把 TOS 還原到看盤狀態，開盤時直接看到圖表。

⚠️ **非致命**：失敗也不建 FAILED 草稿、不改主旨，只在 log 尾註記 + 存 debug 截圖。
⚠️ 前提：步驟 8～9 已完成（信已寄出，或 NEW_TRADES=0 而略過）才做。若更早步驟就 FAILED 結束，保持畫面在失敗當下的
狀態（方便人工檢查），不切 Charts。

**10.1**（單一 call）：
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview_Boss\tos_automation.ps1
Switch-ToCharts          # Verified 啟用 → Click-At 557 33（Charts）→ 游標移到 (960,400)
Capture-Screen "C:\TradeReview_Boss\logs\dbg_charts_$today.png"
```

⚠️ 主頁籤 hover 會彈出下拉小選單（Charts / Flexible Grid / Product Depth / Bookmap），
游標停在頁籤上選單不會消失（ESC 也關不掉）→ `Switch-ToCharts` 會把游標移到 (960,400)。
這只是 hover 在圖表上（顯示十字線），無點擊、不改變任何狀態。

主頁籤列 y≈33（2026-10-16 實測，build 1993）：Monitor=262、Charts=557。
（2026-07-13 舊版 build 是 y≈52、Charts=336，已不適用。）

**10.2** view 驗證：Charts 呈選中、畫面主體是圖表、下拉選單已消失、無彈出對話框。

**10.3** 沒切成功：從全螢幕截圖重新定位 Charts 的 x 座標（不要另外 `Capture-Region`，
那樣拍到的常是搶到前景的 Claude 視窗）
→ 用新座標重點一次（點完一樣要 `SetCursorPos` 移開）→ 再截圖驗證。
第二次仍失敗 → 放棄，log 註記「Charts switch FAILED (non-fatal)」，繼續步驟 11。

**10.4 安全限制**：只點 Charts 頁籤本身；進 Charts 頁後不點頁面內任何元件；
不點下拉選單裡的項目；不改 symbol／時間週期；不點任何下單相關按鈕。

---

## 【步驟 11】寫 log

append 到 `C:\TradeReview_Boss\logs\daily_run.log`：

```
成功（CSV）：
[YYYY-MM-DD HH:MM:SS] OK (CSV) - N rows, local write OK, pushed <sha>, Charts restored

成功（OCR 備援）：
[YYYY-MM-DD HH:MM:SS] OK (OCR) - N TRD rows valid (max diff $X), local write OK, pushed <sha>, Charts restored

每行都要註明寄信結果：`emailed N new trades to juliahsu13 (<message-id>)` ／
`no new trades, no email` ／ `SEND FAILED: <原因>`

本機寫入被 gate 擋下（有新成交時信仍照寄）：
[YYYY-MM-DD HH:MM:SS] PARTIAL - emailed N new trades, local write BLOCKED: <原因>

push 失敗：
[YYYY-MM-DD HH:MM:SS] PARTIAL - local write OK, PUSH FAILED after 3 retries

Charts 收尾失敗但其餘成功：
[YYYY-MM-DD HH:MM:SS] OK (...) - ..., Charts switch FAILED (non-fatal)

失敗：
[YYYY-MM-DD HH:MM:SS] FAIL - <step> - <reason>
```

---

## 【步驟 12】最終回報

對話中輸出簡短摘要表格：

| 項目 | 值 |
|------|-----|
| Date | YYYY-MM-DD |
| Source | CSV export ／ screenshot OCR |
| CSV | `C:\TradeReview\_inbox\<date>.csv`（N bytes） |
| 驗算 | N TRD rows — all pass within $0.02 (max diff $X) |
| 本機寫入 | ✓ 已寫入 trades_all ／ ⚠ 被 gate 擋下：`<原因>` |
| 新成交 | N 筆（0 筆＝未寄信） |
| Email | ✓ 已寄 juliahsu13@gmail.com（`<message-id>`）／ — 無新成交未寄 ／ ✗ 寄送失敗 |
| Subject | `[SPY-DayTrade-Autosend] <date>` |
| git push | ✓ `<sha>` ／ ✗ 三次重試均失敗 |
| Charts 還原 | ✓（`logs/dbg_charts_<date>.png`）／ ✗ non-fatal |

---

## 失敗處理通則

任一步失敗時：

1. `Capture-Screen` 存 debug 截圖到 `logs\debug_<step>_<timestamp>.png`
2. 建 FAILED 主旨草稿：`[SPY-DayTrade-Autosend] FAILED <date> - <reason>`
   （**只建草稿留紀錄，不寄出** —— 只有成功取得新成交時才寄信）
3. log 記 FAIL 行
4. 執行 `& $py write_sync_state.py --status failed --note "<reason>"`
   （讓 502 的 App 標頭轉紅；同樣**不要帶 `--date`**）
5. 對話中回報具體失敗位置 + 建議下一步

**例外**：步驟 10（切回 Charts）是收尾動作，失敗為非致命 —— 不建 FAILED 草稿、
不改主旨，只在 log 的 OK 行尾註記並存 debug 截圖。

**不嘗試**：
- tool_search 找 computer use（不會有）
- AutoIt、SikuliX 等其他自動化工具
- 重啟 TOS、輸入密碼
- 寄出步驟 9.2 以外的任何信（FAILED 草稿、舊日期草稿都不寄）

---

## 禁止事項

### TOS 操作
- ❌ 不點任何 **Buy / Sell / Place Order** 按鈕
- ❌ 採集過程不切到 Account Statement 以外的頁面
  （唯一例外：步驟 10 收尾切回 Charts，且只點頁籤本身）
- ❌ 不重啟 TOS、不嘗試輸入密碼

### Gmail
- ✅ 唯一可寄的信：步驟 9.2，`NEW_TRADES≥1` 時寄給 **juliahsu13@gmail.com**
- ❌ 沒有新成交不寄；FAILED 草稿不寄；不寄給其他收件人；不寄草稿匣裡的舊草稿

### `C:\TradeReview\` 的存取規則

**✅ 可以寫**（但只能透過指定工具）：

| 檔案 | 只能用 |
|---|---|
| `_inbox\*.csv` | TOS 匯出 或 OCR 產出 |
| `trades_all.xlsx` | `spy_daytrade_writer.py --trades-only` |
| `sync_state.json` | `write_sync_state.py` |
| `verify_shots\TOS_verify_*.png` | 步驟 9.0b 複製的驗證截圖（2026-10-16 起） |
| `BOSS_SKILL_current.md` | 步驟 9.5 推送前從 SKILL.md 原樣複製（2026-10-16 起） |

**❌ 絕對不可寫**（502 的所有物）：
`CS交易紀錄.xlsx`、`CS交易紀錄_dump.txt`、`offset_state.json`、`notes\`

**❌ 其他禁止**：
- 不可用 Excel 或任何方式**手動編輯**任何 `.xlsx`
- 不可修改 `trade_review_app.py` 或任何 `.py`（程式碼在 502 改，Boss 靠 git pull）
- 不可執行**不帶 `--trades-only`** 的 `spy_daytrade_writer.py`
- 不可用 `git add -A` —— 一律 `sync_push.ps1 -Owner boss`
- 不可為了讓流程跑完而修改任何數字

---

## 實戰經驗

**✅ 有效**（2026-10-16 版）
- `Activate-TOSWindowVerified`：啟用後確認前景真的是 TOS 才點（單純 `Activate-TOSWindow` 會被 Claude 搶回）
- `Open-ExportDialog` + `Confirm-ExportSave` + `Move-Item` 改名：CSV 匯出整套
- `Switch-ToCharts`：收尾切回 Charts
- 操作後 `Capture-Screen` + view，能精準判斷下一步座標

**❌ 會踩的坑**
- 改 days back 輸入框（打字）會讓 TOS 跳 Unexpected error → N≥5 就不動
- 用滑鼠點三條線選單裡的項目不會觸發 → 用鍵盤
- Save dialog 打字會被注音輸入法吃掉 → 接受預設檔名再改名
- `Capture-Region` 拍局部常拍到 Claude 視窗 → 從確認過的全螢幕截圖裁切
- 點擊座標 y 軸差 3-5 px 就 miss → 不確定時 `Capture-Region` zoom-in
- 主頁籤 hover 會彈下拉選單，游標不移開就不消失、ESC 關不掉
  → 點完頁籤後 `SetCursorPos(960, 400)`
- 不要在 PowerShell 5.1 用 C# discard (`_`) 變數 → 用具名變數
- **PowerShell 5.1 讀沒有 BOM 的 UTF-8 `.ps1` 會當 ANSI**，中文註解會炸掉語法
  → 若改過 `C:\TradeReview\*.ps1`，跑 `python _fix_ps1_bom.py`

**⚠️ OCR 備援路徑特有風險**
- 只截得到 Cash Balance 區塊，拿不到完整 CSV 的其他 section
- 列數 overflow 要捲動多截幾張再合併（步驟 5'.4）
- OCR 誤判靠步驟 7 逐列驗算攔截，**每列都要驗**
- **最容易漏的是 BAL / DOI / JRN 列** —— 漏了引擎必定 abort

---

## 修訂紀錄

| 日期 | 變更 |
|---|---|
| 2026-10-16 | **推送改三樣**：9.5 只推 `_inbox\`／`verify_shots\`／`BOSS_SKILL_current.md`（trades_all／sync_state 歸 502、不推）；9.5 推送前自動複製 SKILL 到 `BOSS_SKILL_current.md`；檔頭寫明改 SKILL 後要同步 CLAUDE.md 與 BOSS_SKILL_current.md |
| 2026-10-16 | **舊內容校正**：座標改為 build 1993 實測（主頁籤 y≈33、Monitor 262、Charts 557）；新增「Claude 搶前景」坑與 `Activate-TOSWindowVerified`；步驟 4 改為 N≥5 不動；步驟 5 改用 `Open-ExportDialog`／`Confirm-ExportSave`（鍵盤選單、不打檔名再改名）；步驟 9.5 推送清單加 `verify_shots\`、repo 卡 rebase 時不跑 push；步驟 10 改用 `Switch-ToCharts` |
| 2026-10-16 | 新增 9.0b 驗證截圖本機留存：交易日擷取 Cash Balance＋Order History＋Trade History，存 `logs\` 與 `verify_shots\`；不附信件、不動日期範圍欄位、不放 `_inbox\` |
| 2026-10-16 | **美股交易日 06:00**：排程改台灣週二～週六 06:00；Windows 工作 `Claude-TOS-Collector-Launch` 05:45 喚醒並開 Claude；新增步驟 0（`us_trading_day.py`，休市日直接結束、不動 TOS） |
| 2026-10-16 | **改為有新成交才寄**：排程改每天 07:00（已再改為上一列）；`check_new_trades.py` 判斷有沒有沒寄過的成交，有才建草稿並直接寄給 juliahsu13@gmail.com，寄出後 `--mark`；沒有就不寄。停用 Apps Script → sunbeamichelle 路徑 |
| 2026-10-16 | **步驟 9 新增 STATEMENT TABLE 區塊**：`render_statement.py` 把 CSV 轉成純文字對齊表（Cash Balance／Order History／Trade History），取代截圖；不附圖片 |
| 2026-08-11 | **整併為單一權威檔**。原本排程 SKILL.md 內嵌 07-17 的 OCR-only 流程，與 `C:\TradeReview_Boss\CLAUDE.md` 的新規則衝突（前者禁止 CSV 匯出與寫入 `C:\TradeReview\`），直接指令勝出，導致步驟 6.5／9.5 從未執行 |
| 2026-08-11 | 補上 ExecutionPolicy 繞過寫法（本機全 scope Undefined＝實際 Restricted） |
| 2026-08-10 | 改回 CSV 為主：引擎改依欄名對應，直接吃 TOS 原生 10 欄，不需轉換 |
| 2026-08-10 | OCR 改抓 Cash Balance **每一列**（含 BAL/DOI/JRN/ADJ）；步驟 7 驗算只對 TRD |
| 2026-08-10 | 新增步驟 6.5（`--trades-only` 本機寫入 + `write_sync_state`）與 9.5（`sync_push`） |
| 2026-08-10 | 禁止事項改為精確的可寫／不可寫清單 |
| 2026-07-17 | 曾改為 OCR-only（現已解除） |
| 2026-07-13 | 實測步驟 10 切回 Charts |
| 2026-05-13 | 跑通全流程 |