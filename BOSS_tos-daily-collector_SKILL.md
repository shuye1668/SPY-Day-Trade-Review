# 每日 TOS 採集 routine（Boss PC，權威版）

> **這個檔要整檔取代：**
> `C:\Users\TS USER\.claude\scheduled-tasks\tos-daily-collector\SKILL.md`
>
> 本檔是自足的 —— 執行時**不需要**再去讀別的規則檔。
> 最後更新：2026-08-11。修訂紀錄見文末。
>
> ⚠️ **規則只能有一份。** 2026-08-11 查出的事故：這個排程原本內嵌 2026-07-17 的
> OCR-only 流程，同時 `C:\TradeReview_Boss\CLAUDE.md` 另有一份較新的規則，兩者
> 互相矛盾（前者明文禁止 CSV 匯出與寫入 `C:\TradeReview\`）。直接指令會贏，
> 導致本機寫入與 git 同步從未執行、每天都要人工補帳。
> **要改流程請改本檔，不要在別處另寫一份。**

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

Gmail 草稿由 boss 端 Gmail 的 Apps Script 每 5 分鐘自動送到
**sunbeamichelle@gmail.com**。草稿保留為退路：git push 掛掉時 502 仍可人工處理。

### 🔴 三條鐵則

1. **writer 一定要帶 `--trades-only`。** 不帶會去寫 `CS交易紀錄.xlsx` 與
   `offset_state.json` —— 那是 502 的所有物，會製造出永遠推不出去、還會卡住
   自己 pull 的本機改動。
2. **不得手動編輯任何 `.xlsx`。** 所有寫入只能經由 engine / writer。
3. **不得用 `git add -A`。** 一律用 `sync_push.ps1 -Owner boss`。

---

## 執行時機（這台週末關機）

週五的交易通常等到**週一較晚（10 點後）開機**才補完成。所以：

- **一次跑會補完所有欠帳**：步驟 4 設「5 days back」不是保險，是必要的。
  writer 會處理 CSV 裡所有 session，已寫過的自動冪等跳過。
- **週一落後一個交易日是正常的**，不是故障。502 的 App 標頭那段時間會顯示
  琥珀色「⚠待補前一交易日」，補完轉回灰色。
- 顯示**紅色「✕落後 N 個交易日」**（N≥2）才代表真的有事沒跑成。

---

## 執行環境（會踩到的三個坑）

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
$ok = Activate-TOSWindow
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
  → 找「Monitor」主頁籤（y≈52, x≈41）→ `Click-At` → 等 1 秒截圖
  → 再找「Account Statement」子頁籤 → `Click-At` → 等 1 秒截圖確認

---

## 【步驟 4】日期範圍設為「5 days back from today」

⚠️ TOS 預設「1 days back」只看今天。要 5 天（涵蓋週末＋假日＋補跑）。

**4.1 觀察 status bar**（Account Statement 上方）：

| 看到 | 情境 |
|---|---|
| `5 days back from today` | A → 跳步驟 5 |
| `1 days back from today` | B → 要改 |
| 其他文字（含 `from X to Y`） | B → 改回 5 days back |
| 完全找不到 | C |

**4.3 情境 B：改成 5 days back**

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
h) `Capture-Screen` + view 驗證 status bar 已變成 `5 days back from today`

**4.4 情境 C**：存 `logs\debug_no_status_bar_$today.png`，建 FAILED 草稿
「Cannot find date range status bar」後結束。

---

## 【步驟 5】★ 主路徑：CSV 匯出

**2026-08-10 起改回 CSV 為主。** 引擎已能直接吃 TOS 原生 10 欄格式
（`Trade Date,Exec Date,Exec Time,Type,Ref #,Description,Misc Fees,Commissions & Fees,Amount,Balance`），
**不需要任何轉換** —— 這正是當初停用 CSV 路徑的原因，已在引擎端修好。

log 實證：2026-05-13～07-16 走 CSV 連續 **44 次成功**（max diff 多在 $0.004 以下）；
07-17 之後改 OCR，覆蓋率反而降到約 6 成。CSV 才是穩的那條。

**5.1** Account Statement 右上三條線圖示 → **Export to file**

**5.2** Save dialog 出現後，路徑填 `C:\TradeReview\_inbox\<today>.csv`，
類型選 CSV，存檔。

**5.3** `Capture-Screen` + view 確認：Save dialog 已關、沒有「檔案已存在是否覆蓋」殘留。

**5.4** 確認檔案產生且非空：
```powershell
$f = "C:\TradeReview\_inbox\$today.csv"
if (Test-Path $f) { (Get-Item $f).Length } else { "MISSING" }
```

- 存在且 **> 500 bytes** → **跳到步驟 6.5**（不需要 OCR）
- 否則 → 走步驟 5'

### ⚠️ 嚴格限制重試次數

TOS 介面這一段自 07-17 後就沒實際跑過，未經驗證：

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
- 看得到「5 days back from today」
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

**catch-up 是自動的**：writer 會處理 CSV 裡所有 session（步驟 4 的「5 days back」
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

## 【步驟 8】防重複：檢查當日草稿

建草稿前用 Gmail MCP 搜尋：

```
subject:"[SPY-DayTrade-Autosend] $today"
```

若已存在當日草稿（不論成功或 FAILED），**先刪除舊的再建新的**。
確保每個交易日只有一份草稿。

---

## 【步驟 9】建 Gmail 草稿

用 Gmail MCP 的 `create_draft`：

- **收件人**：`sunbeamichelle@gmail.com`
- **主旨**：`[SPY-DayTrade-Autosend] $today`
- **正文**（嚴格純文字，不要 HTML）：

```
TOS Account Statement Daily Export
Date: $today
Generated: <現在時間 ISO8601，例 2026-08-12T06:05:14+08:00>
Boss PC: $env:COMPUTERNAME
Source: <CSV export | screenshot OCR>
Local write: <OK | BLOCKED: 原因 | FAILED: 原因>

===== CSV CONTENT START =====
<把 _inbox\<today>.csv 的內容整段貼上>
===== CSV CONTENT END =====

Validation: N TRD rows, all pass qty*price=abs(amount) within $0.02 tolerance (max diff $X).

--
Auto-generated by Claude Code on Boss PC
```

⚠️ **只 `drafts.create`，不要 `drafts.send`**（Apps Script 會在 5 分鐘內送出）。

⚠️ `Source:` 與 `Local write:` 這兩行很重要 —— 502 從 Gmail 就能看出走了哪條路徑、
本機帳做成了沒。

---

## 【步驟 9.5】★ 推送到 GitHub

```powershell
powershell -ExecutionPolicy Bypass -File C:\TradeReview\sync_push.ps1 -Owner boss
```

只會推 `trades_all.xlsx` / `_inbox\` / `sync_state.json` 三樣。
push 失敗會自動重試 3 次；三次都失敗時本機 commit 仍在（**資料沒掉**），
但要把這件事寫進步驟 11 的 log 與 Gmail 草稿正文。

---

## 【步驟 10】收尾：切回 Charts

目的：採集完把 TOS 還原到看盤狀態，開盤時直接看到圖表。

⚠️ **非致命**：失敗也不建 FAILED 草稿、不改主旨，只在 log 尾註記 + 存 debug 截圖。
⚠️ 前提：步驟 9 草稿已建好才做。若更早步驟就 FAILED 結束，保持畫面在失敗當下的
狀態（方便人工檢查），不切 Charts。

**10.1**（單一 call）：
```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview_Boss\tos_automation.ps1
Activate-TOSWindow | Out-Null
Start-Sleep -Milliseconds 500
Click-At 336 52          # Charts 主頁籤（2026-07-13 實測，1920x1080）
Start-Sleep -Milliseconds 500
# ⚠️ 實測：主頁籤 hover 會彈出下拉小選單（Charts / Flexible Grid / Product Depth /
#    Bookmap）。游標停在頁籤上時選單不會消失（ESC 也關不掉）→ 必須把游標移到
#    中性位置，選單才會自動收起。
[Win32Api]::SetCursorPos(960, 400) | Out-Null
Start-Sleep -Seconds 2
Capture-Screen "C:\TradeReview_Boss\logs\dbg_charts_$today.png"
```

主頁籤列 y≈52，各頁籤中心 x（2026-07-13 實測）：
```
Monitor=41  Trade=97  Analyze=151  Scan=202  MarketWatch=266
Charts=336  Tools=385  Education=444  Help=502
```
游標移到 (960, 400) 只是 hover 在圖表上（顯示十字線），無點擊、不改變任何狀態。

**10.2** view 驗證：Charts 呈選中、畫面主體是圖表、下拉選單已消失、無彈出對話框。

**10.3** 沒切成功：`Capture-Region 0 40 560 30` → view 重新定位 Charts 的 x 座標
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

本機寫入被 gate 擋下（草稿仍已建立）：
[YYYY-MM-DD HH:MM:SS] PARTIAL - draft OK, local write BLOCKED: <原因>

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
| Gmail draft | `<draft-id>` → sunbeamichelle@gmail.com |
| Subject | `[SPY-DayTrade-Autosend] <date>` |
| git push | ✓ `<sha>` ／ ✗ 三次重試均失敗 |
| Charts 還原 | ✓（`logs/dbg_charts_<date>.png`）／ ✗ non-fatal |

---

## 失敗處理通則

任一步失敗時：

1. `Capture-Screen` 存 debug 截圖到 `logs\debug_<step>_<timestamp>.png`
2. 建 FAILED 主旨草稿：`[SPY-DayTrade-Autosend] FAILED <date> - <reason>`
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
- 送出 Gmail 草稿

---

## 禁止事項

### TOS 操作
- ❌ 不點任何 **Buy / Sell / Place Order** 按鈕
- ❌ 採集過程不切到 Account Statement 以外的頁面
  （唯一例外：步驟 10 收尾切回 Charts，且只點頁籤本身）
- ❌ 不重啟 TOS、不嘗試輸入密碼

### Gmail
- ❌ 不送出草稿（只 `drafts.create`，Apps Script 會送）

### `C:\TradeReview\` 的存取規則

**✅ 可以寫**（但只能透過指定工具）：

| 檔案 | 只能用 |
|---|---|
| `_inbox\*.csv` | TOS 匯出 或 OCR 產出 |
| `trades_all.xlsx` | `spy_daytrade_writer.py --trades-only` |
| `sync_state.json` | `write_sync_state.py` |

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

**✅ 有效**
- `Activate-TOSWindow` 用 SetForegroundWindow + ShowWindow **100% 可靠**
- `TripleClick` + 輸入新值可改 days back 數字
- 操作後 `Capture-Screen` + view，能精準判斷下一步座標
- `Click-At 336 52` 可靠切到 Charts 主頁籤

**❌ 會踩的坑**
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
| 2026-08-11 | **整併為單一權威檔**。原本排程 SKILL.md 內嵌 07-17 的 OCR-only 流程，與 `C:\TradeReview_Boss\CLAUDE.md` 的新規則衝突（前者禁止 CSV 匯出與寫入 `C:\TradeReview\`），直接指令勝出，導致步驟 6.5／9.5 從未執行、每天人工補帳 |
| 2026-08-11 | 補上 ExecutionPolicy 繞過寫法（本機全 scope Undefined＝實際 Restricted） |
| 2026-08-10 | 改回 CSV 為主：引擎改依欄名對應，直接吃 TOS 原生 10 欄，不需轉換 |
| 2026-08-10 | OCR 改抓 Cash Balance **每一列**（含 BAL/DOI/JRN/ADJ）；步驟 7 驗算只對 TRD |
| 2026-08-10 | 新增步驟 6.5（`--trades-only` 本機寫入 + `write_sync_state`）與 9.5（`sync_push`） |
| 2026-08-10 | 禁止事項改為精確的可寫／不可寫清單 |
| 2026-07-17 | 曾改為 OCR-only（現已解除） |
| 2026-07-13 | 實測步驟 10 切回 Charts |
| 2026-05-13 | 跑通全流程 |
