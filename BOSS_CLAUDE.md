# Boss PC 每日 TOS 採集 routine（權威版）

> **這個檔要放到：`C:\TradeReview_Boss\CLAUDE.md`（整檔取代）**
> 本檔為 Boss PC（交易機）每日採集流程的唯一權威定義。
> 最後更新：2026-08-10。修訂重點見文末〈修訂紀錄〉。

---

## 模式：CSV 匯出為主，OCR 為備援

**2026-08-10 起改回 CSV 為主。** 引擎已能直接吃 TOS 原生匯出的 10 欄格式，
不需要任何轉換 —— 這正是 2026-07-17 當初停用 CSV 路徑的原因，現已在引擎端修好。

```
主路徑：Export to file → C:\TradeReview\_inbox\<date>.csv      （步驟 5）
備援  ：截圖 → OCR「Cash Balance 每一列」→ 同一個 CSV 路徑     （步驟 5' + 6）
```

## 架構

```
Boss PC（這台）                                502 PC
─────────────────                             ─────────────────
TOS → 採集 → _inbox\<date>.csv
  ├─ engine 乾跑（BALANCE 閘門）
  ├─ writer --trades-only → trades_all.xlsx    ← 第一段到此為止
  ├─ write_sync_state.py（新鮮度戳記）
  ├─ Gmail 草稿（保留，當退路）
  └─ sync_push -Owner boss ──→ GitHub ──→ sync_pull → 第二段
                                            cs_from_trades.py
                                            → CS交易紀錄.xlsx（含人工修正）
```

Gmail 草稿由 boss 端 Gmail 帳號的 Apps Script 每 5 分鐘自動送到
**sunbeamichelle@gmail.com**。草稿路徑保留為退路：git push 掛掉時 502 仍可人工作業。

## 自動化方式：PowerShell + Win32（已驗證跑通）

- ⚠️ 此 routine session **不會有** desktop computer-use MCP。
- ⚠️ 不要 tool_search 找 computer use、不要嘗試 `Claude_in_Chrome__computer`
  （那個要 browser tabId，不能控制 desktop app）。
- ⚠️ 直接走 PowerShell + Win32，所有功能都有 helper 提供。

工作流程：每個關鍵步驟完成後，用 `Capture-Screen` 存 PNG 到 `logs\`，
再用 view tool 讀 PNG 確認當下螢幕狀態。視覺確認後再進下一步。

## 常用變數

```powershell
$py   = "C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe"
$root = "C:\TradeReview"          # 複盤系統（git repo）
$boss = "C:\TradeReview_Boss"     # 採集系統（本檔所在）
```

---

## 【步驟 1】載入 helper + 取得今天日期

```powershell
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
. C:\TradeReview_Boss\tos_automation.ps1
$ok = Activate-TOSWindow
Write-Host "Activated: $ok"
if ($ok) { Capture-Screen "C:\TradeReview_Boss\logs\dbg_tos_activated_$today.png" }
```

- `$ok = False` → `Capture-Screen` 存 `logs\debug_no_tos_$today.png`，
  建 FAILED 草稿「TOS window not found」後結束。
- `$ok = True` → view 該 PNG 確認 TOS 在前景且最大化、顯示交易介面（不是登入頁）。

---

## 【步驟 3】操作 TOS 切到 Account Statement

從步驟 2 截圖判斷當下分頁狀態：

- **情境 1**：已在 Monitor → Account Statement → 跳到步驟 4
- **情境 2**：在 Monitor 但不是 Account Statement 子頁籤
  → 從截圖找到「Account Statement」子頁籤座標 → `Click-At <x> <y>`
  → 等 1 秒，`Capture-Screen` 確認頁面已切換
- **情境 3**：在其他主頁籤（Trade / Analyze / Charts 等；上次 routine 收尾會停在
  Charts，這是正常的預期起點）
  → 找到「Monitor」主頁籤座標（y≈52, x≈41）→ `Click-At` → 等 1 秒截圖確認
  → 再找「Account Statement」子頁籤 → `Click-At` → 等 1 秒截圖確認

---

## 【步驟 4】確認/設定日期範圍為「5 days back from today」

⚠️ TOS 預設「1 days back from today」只看今天。我們要 5 天（涵蓋週末＋假日）。

**4.1 觀察 status bar** — 看 Account Statement 上方狀態列：

| 看到 | 情境 |
|---|---|
| `5 days back from today` | A → 跳到步驟 5 |
| `1 days back from today` | B → 要改 |
| 其他文字（含 `from X to Y` 自訂範圍） | B → 要改回 5 days back |
| 完全找不到該文字 | C |

**4.3 情境 B：改成 5 days back from today**

a) 從截圖找出「N days back from today」文字的精確座標。
   1920x1080 下 y 軸通常約 78-80，x 軸約 600-700。
   **注意：y 軸差 3-5 px 就會 miss（實戰學到的）。** 不確定就先 `Capture-Region` zoom。

b) 點該文字：
```powershell
. C:\TradeReview_Boss\tos_automation.ps1
Click-At <x> <y>
Start-Sleep -Milliseconds 500
Capture-Screen "C:\TradeReview_Boss\logs\dbg_dialog_$today.png"
```

c) view 該 PNG 確認對話框已打開。預期長相：
```
Statement for:  [Reset]
  ◉ [1] days back from [Today ▼]
  ○ from [MM/DD/YYYY] to [MM/DD/YYYY]
  Maximum period length is 370 days.
```
沒打開 → 重新截圖 zoom-in 找精確座標再點。

d) 確認上方 radio「N days back from Today」被選中（**不要動下方 from/to**）。

e) 對「1」輸入框三連點 + 輸入 5 + Enter（單一 PowerShell call）：
```powershell
. C:\TradeReview_Boss\tos_automation.ps1
TripleClick-At <input_x> <input_y>
Start-Sleep -Milliseconds 200
Send-Keys "5"
Start-Sleep -Milliseconds 100
Send-Keys "{ENTER}"
```

f) 等表格重新載入：`Start-Sleep 8`
g) 按 ESC 關掉殘留 tooltip：`Send-Keys "{ESC}"`
h) `Capture-Screen` + view 驗證 status bar 已改為 `5 days back from today`

**4.4 情境 C**：找不到該文字 → `Capture-Screen` 存
`logs\debug_no_status_bar_$today.png` → 建 FAILED 草稿
「Cannot find date range status bar」後結束。

---

## 【步驟 5】★ 主路徑：CSV 匯出

**5.1** 在 Account Statement 右上找到選單（三條線圖示）→ **Export to file**

**5.2** Save dialog 出現後，路徑輸入 `C:\TradeReview\_inbox\<today>.csv`，
存檔類型選 CSV，按存檔。

**5.3** `Capture-Screen` 後 view 確認：
- Save dialog 已關閉
- 沒有「檔案已存在，是否覆蓋」之類殘留對話框

**5.4** 用 PowerShell 確認檔案確實產生且非空：
```powershell
$f = "C:\TradeReview\_inbox\$today.csv"
if (Test-Path $f) { (Get-Item $f).Length } else { "MISSING" }
```

- 檔案存在且 **> 500 bytes** → 成功，**跳到步驟 6.5**（不需要 OCR）
- 否則 → 走步驟 5'

> 💡 匯出的表頭是 10 欄
> （`Trade Date,Exec Date,Exec Time,Type,Ref #,Description,Misc Fees,Commissions & Fees,Amount,Balance`），
> 引擎會依欄名自動對應，**不需要任何轉換**。

---

## 【步驟 5'】備援路徑：截圖

⚠️ 只有步驟 5 失敗才走這裡。

**5'.1** 確保 tooltip / 對話框已關閉：
```powershell
. C:\TradeReview_Boss\tos_automation.ps1
Send-Keys "{ESC}"
Start-Sleep 1
```

**5'.2** 截整個畫面：
```powershell
Capture-Screen "C:\TradeReview_Boss\logs\screenshot_$today.png"
```

**5'.3** view 該 PNG 確認：
- 看得到「5 days back from today」字樣
- 看得到 Cash Balance 表格的多列資料
- 沒有對話框或 tooltip 遮住
- ⚠️ **最關鍵**：確認「最後一列（最新的 today 交易）」有出現在畫面內。
  Cash Balance 通常最舊在上、最新在下。底部被截斷＝有列 overflow → 見 5'.4。

**5'.4 列數超出可視範圍時**：在表格內側點一下取得焦點（點資料列空白處，
不要點任何按鈕/連結），按 End 或 PageDown 捲到最底：
```powershell
Click-At <x_table_body> <y_table_body>
Send-Keys "{END}"
Start-Sleep 1
Capture-Screen "C:\TradeReview_Boss\logs\screenshot_${today}b.png"
```
一張截不完就「捲一段截一張」抓多張重疊截圖（`_b.png`、`_c.png`…），
步驟 6 OCR 後靠 Ref# / 時間去重合併，確保每一列都被涵蓋且沒有重複計入。
合併後仍無法確認涵蓋完整 → 建 FAILED 草稿
「OCR rows overflow, cannot capture all rows」後結束。

---

## 【步驟 6】OCR 截圖（僅備援路徑需要）

### 🔴 最重要的一條：Cash Balance 區塊的「每一列」都要抓

**不是只抓 TRD。** 必抓的類型：

| TYPE | 意義 | 為什麼非抓不可 |
|---|---|---|
| `TRD` | 成交 | 主資料 |
| **`BAL`** | 期初餘額 | **引擎靠它切交易日、取期初餘額、跑逐列餘額驗證。沒有它引擎一定 abort，當天一筆都寫不進去** |
| **`DOI`** | 配息 | 影響餘額鏈；漏抓會讓 BALANCE 對不上 |
| **`JRN`** | 轉帳 | 同上 |
| `ADJ` 等 | 調整 | 同上 |

**這件事已經出事過。** `offset_state.json` 自己記著：

> ⚠️注意:8/05 open-bal 115851.99 是由成交鏈回推(今日OCR草稿漏BAL/DOI/JRN列)，
> 非直接讀BAL行 ... 唯 offset 數值為暫定值，待補齊含BAL列的完整對帳單後覆核

`BAL` 列通常在表格最上面一列，很容易被當成標題略過，長這樣：
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

⚠️ **日期欄：Trade Date 與 Exec Date 兩個都要抓。**
Trade Date 是券商營業日，午夜之後成交的列會掛在前一個營業日；
**Exec Date 才是真正的日曆日**。實測 2026-07-29 的對帳單有 5 列
`Trade Date=7/27` 但 `Exec Date=7/28 01:18` —— 取錯就整整差一天。
引擎會自動取 Exec Date，所以你只要照實抓兩欄即可。

### OCR 注意事項

- Description 裡的 qty（`+100` / `-100` / 分批 `+80` `+20`）與 price（`@xxx.xxxx`）要看清楚
- Amount / Balance 有千分位逗號，去掉逗號再存數值
- 分批成交（同 Ref#、同秒、qty 拆兩列）要各存一列
- 有 QQQ 等非 SPY 標的也照樣抓（Description 保留原標的字樣）
- 任一 **TRD** 列的 qty / price / amount 辨識不清 → 見步驟 7 失敗處理

### 寫成 CSV

把 OCR 結果寫入 `C:\TradeReview\_inbox\<today>.csv`，第一行為欄位表頭，
之後每列一行、逗號分隔。9 欄或 10 欄格式都可以，引擎依欄名對應。

---

## 【步驟 6.5】★ 寫入本機 trades_all

```powershell
$py   = "C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe"
$csv  = "C:\TradeReview\_inbox\$today.csv"
cd C:\TradeReview

# 先乾跑，不寫任何檔案
& $py spy_daytrade_engine.py $csv

# 全綠才寫入
& $py spy_daytrade_writer.py $csv --commit --trades-only
$wr = $LASTEXITCODE

# 記錄狀態，供 502 端 App 顯示資料新鮮度
if ($wr -eq 0) {
    & $py write_sync_state.py --date $today
} else {
    & $py write_sync_state.py --date $today --status blocked --note "writer gate 擋下，需人工確認"
}
```

**「全綠」的判準**（引擎輸出）：最後一行 `Σcol1` 與 `Σpnl` **兩個數字一樣**、
`留倉=False`，而且下面**沒有**「警告 / 需人工確認」那一段。

**writer 結束碼**：`0` = 成功或冪等跳過；`2` = 被 gate 擋下（有 🔴）；其他 = 例外。

### 🔴 writer 一定要帶 `--trades-only`

不帶會去寫 `CS交易紀錄.xlsx` 與 `offset_state.json` —— 那兩個是 **502 的所有物**，
會製造出永遠推不出去、還會卡住自己 pull 的本機改動。

### 被 gate 擋下時：什麼都不要改

`trades_all` 停在前一日**是正確的，不是壞掉**。照常繼續步驟 8 建草稿，
把被擋原因寫進正文，交由 502 人工處理。

會被擋下的情況：`BALANCE 對不上`、`收盤留倉`、`非交易現金異動（DOI/JRN）`、
`session 開盤/收盤餘額不明`。

> 這是 `CLAUDE.md §0` 的核心：**絕不為了跑完流程而寫入估算／推算／猜值。**

---

## 【步驟 7】驗算

對每一列 **TRD** 驗算 `qty × price ≈ abs(amount)`，容差 **$0.02**。

⚠️ **`BAL` / `DOI` / `JRN` / `ADJ` 列不做這個驗算** —— 它們沒有 qty/price，
驗算必然失敗。這些列原樣帶入 CSV 即可，引擎會自行處理：
`BAL` 用來取期初餘額；`DOI`/`JRN` 會被 gate 標為「非交易現金異動」而擋下該日
交由人工判斷（**這是正確行為，不是錯誤**）。

- 逐列計算 `diff = |qty × price − abs(amount)|`，累積 maxDiff
- `diff > $0.02` 的 TRD 列 → 草稿主旨改 FAILED，正文記該列詳情後結束
- 任一 TRD 列 qty/price/amount 無法讀出 → 建 FAILED 草稿
  「OCR unreadable row: `<定位資訊>`」後結束

> 走 CSV 主路徑時這一步是額外保險（資料直接來自 TOS，不會有 OCR 誤判）；
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
Generated: <現在時間 ISO8601，例 2026-08-10T06:05:14+08:00>
Boss PC: $env:COMPUTERNAME
Source: <CSV export | screenshot OCR>
Local write: <OK | BLOCKED: 原因 | FAILED: 原因>

===== CSV CONTENT START =====
<把 _inbox\<today>.csv 的內容整段貼上>
===== CSV CONTENT END =====

Validation: N TRD rows, all pass qty*price=abs(amount) within $0.02 tolerance (max diff $X).

Auto-generated by Claude Code on Boss PC
```

⚠️ **只 `drafts.create`，不要 `drafts.send`**（Apps Script 會在 5 分鐘內送出）。

⚠️ `Local write:` 這一行很重要 —— 502 從 Gmail 就能看到本機帳做成了沒。

---

## 【步驟 9.5】★ 推送到 GitHub

```powershell
cd C:\TradeReview
.\sync_push.ps1 -Owner boss
```

只會推 `trades_all.xlsx` / `_inbox\` / `sync_state.json` 三樣。
push 失敗會自動重試 3 次；三次都失敗時本機 commit 仍在（**資料沒掉**），
但要把這件事寫進步驟 11 的 log 與 Gmail 草稿正文。

---

## 【步驟 10】收尾：切回 Charts 頁面

目的：採集完成後把 TOS 畫面還原到看盤狀態，boss 開盤時直接看到圖表。

⚠️ 此步驟**非致命**：失敗也不建 FAILED 草稿、不改主旨、不影響本次採集結果 ——
只在 log 尾註記 + 存 debug 截圖。

⚠️ 前提：步驟 9 草稿已建好才做這步。若流程在更早步驟就 FAILED 結束，
保持畫面在失敗當下的狀態（方便人工檢查），不切 Charts。

**10.1** 點 Charts 主頁籤 + 移開游標（單一 PowerShell call）：
```powershell
. C:\TradeReview_Boss\tos_automation.ps1
Activate-TOSWindow | Out-Null
Start-Sleep -Milliseconds 500
Click-At 336 52          # Charts 主頁籤（2026-07-13 實測，1920x1080）
Start-Sleep -Milliseconds 500
# ⚠️ 實測：主頁籤 hover 會彈出下拉小選單（Charts / Flexible Grid /
#    Product Depth / Bookmap）。游標停在頁籤上時選單不會消失（ESC 也關不掉）
#    → 必須把游標移到中性位置，選單才會自動收起。
[Win32Api]::SetCursorPos(960, 400) | Out-Null
Start-Sleep -Seconds 2
Capture-Screen "C:\TradeReview_Boss\logs\dbg_charts_$today.png"
```

主頁籤列 y≈52，各頁籤中心 x 參考（2026-07-13 實測）：

```
Monitor=41  Trade=97  Analyze=151  Scan=202  MarketWatch=266
Charts=336  Tools=385  Education=444  Help=502
```

游標移到 (960, 400) 只是 hover 在圖表上（顯示十字線），無任何點擊、
不改變任何狀態，安全。

**10.2** view 該 PNG 驗證：
- 主頁籤列上「Charts」呈選中（highlight）狀態
- 畫面主體是圖表（K 線/線圖），不再是 Account Statement 表格
- 頁籤下方的下拉小選單已消失
- 沒有任何彈出對話框

**10.3** 若沒切成功：`Capture-Region 0 40 560 30` 存 PNG → view 重新定位
「Charts」文字的 x 座標 → 用新座標重點一次（點完一樣要 `SetCursorPos` 移開游標）
→ 再截圖驗證。第二次仍失敗 → 放棄重試，log 註記
「Charts switch FAILED (non-fatal)」，繼續步驟 11。

**10.4 安全限制**：
- 只點「Charts」頁籤本身，進到 Charts 頁後不點頁面內任何其他元件
- 不點下拉小選單裡的任何項目（Flexible Grid 等）—— 讓它自己收起就好
- 不改 symbol、不改時間週期、不點任何下單相關按鈕

---

## 【步驟 11】寫 log

PowerShell append 到 `C:\TradeReview_Boss\logs\daily_run.log`：

```
成功（CSV 路徑）：
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
4. 執行 `& $py write_sync_state.py --date $today --status failed --note "<reason>"`
   （讓 502 的 App 標頭轉紅）
5. 對話中回報具體失敗位置 + 建議下一步

**例外**：步驟 10（切回 Charts）是收尾動作，失敗為非致命 —— 不建 FAILED 草稿、
不改主旨，只在 log 的 OK 行尾註記「Charts switch FAILED (non-fatal)」並存 debug 截圖。

**不嘗試**：
- tool_search 找 computer use（不會有）
- AutoIt、SikuliX 等其他自動化工具
- 重啟 TOS、輸入密碼
- 送出 Gmail 草稿

---

## 禁止事項

### TOS 操作

- ❌ 不點 TOS 任何 **Buy / Sell / Place Order** 按鈕
- ❌ 採集過程（步驟 2~7）不切到 Account Statement 以外的 TOS 頁面；
  唯一例外是步驟 10 收尾切回 Charts（且只點頁籤本身）
- ❌ 不重啟 TOS、不嘗試輸入密碼

### Gmail

- ❌ 不送出草稿（只 `drafts.create`，Apps Script 會送）

### `C:\TradeReview\` 的存取規則（2026-08-10 起）

這台機器只做 **CLAUDE.md 的第一段**：對帳單 → `trades_all.xlsx`。

**✅ 可以寫**（但只能透過指定工具）：

| 檔案 | 只能用 |
|---|---|
| `_inbox\*.csv` | TOS 匯出 或 OCR 產出 |
| `trades_all.xlsx` | `spy_daytrade_writer.py --trades-only` |
| `sync_state.json` | `write_sync_state.py` |

**❌ 絕對不可寫**（502 的所有物）：

- `CS交易紀錄.xlsx`
- `CS交易紀錄_dump.txt`
- `offset_state.json`
- `notes\`

**❌ 其他禁止**：

- 不可用 Excel 或任何方式**手動編輯**任何 `.xlsx`
- 不可修改 `trade_review_app.py` 或任何 `.py`（程式碼在 502 改，Boss 靠 git pull 取得）
- 不可執行**不帶 `--trades-only`** 的 `spy_daytrade_writer.py`
- 不可用 `git add -A` —— 一律用 `.\sync_push.ps1 -Owner boss`
- 不可為了讓流程跑完而修改任何數字

---

## 實戰經驗

（2026-05-13 跑通全流程；2026-07-13 實測步驟 10；2026-07-17 曾改 OCR-only；
2026-08-10 改回 CSV 為主並接上本機寫入與 git 同步）

**✅ 有效**
- `Activate-TOSWindow` 用 SetForegroundWindow + ShowWindow **100% 可靠**
- `TripleClick` + 輸入新值，可改 days back 的數字
- 操作後 `Capture-Screen` + view，能精準判斷下一步座標
- `Click-At 336 52` 可靠切到 Charts 主頁籤
- 每次 PowerShell tool call 是新 process，變數不會跨 call 保留
  → 每個 call 開頭都要重新 `. tos_automation.ps1`；重要狀態存檔（Out-File 到 temp）

**❌ 會踩的坑**
- 點擊座標 y 軸差 3-5 px 就 miss → 不確定時 `Capture-Region` zoom-in
- 主頁籤 hover 會彈下拉小選單，游標不移開選單就不消失、ESC 關不掉
  → 點完頁籤後 `SetCursorPos(960, 400)` 移開游標
- 不要在 PowerShell 5.1 用 C# discard (`_`) 變數 → 用具名變數如 `out IntPtr lpdwProcessId`
- **PowerShell 5.1 讀沒有 BOM 的 UTF-8 `.ps1` 會當成 ANSI**，中文註解會炸掉語法
  → 若改過 `C:\TradeReview\*.ps1`，跑 `python _fix_ps1_bom.py`

**⚠️ OCR 備援路徑特有風險**
- 只截得到 Cash Balance 區塊，拿不到完整 CSV 的其他 section
- 列數 overflow 時要捲動多截幾張再合併（步驟 5'.4）
- OCR 誤判靠步驟 7 逐列驗算攔截；**每列都要驗**
- **最容易漏的是 BAL / DOI / JRN 列** —— 漏了引擎一定 abort（見步驟 6）

---

## 修訂紀錄

| 日期 | 變更 |
|---|---|
| 2026-08-10 | **改回 CSV 匯出為主**：引擎改為依欄名對應，直接吃 TOS 原生 10 欄格式，不需轉換。這是 07-17 停用 CSV 路徑的原因，已在引擎端修好 |
| 2026-08-10 | **OCR 改為抓 Cash Balance 每一列**（含 BAL/DOI/JRN/ADJ），不只 TRD。漏抓 BAL 會讓引擎 abort，8/05 的 offset 就是因此被迫回推 |
| 2026-08-10 | **步驟 7 驗算只對 TRD 做** —— BAL/DOI/JRN 沒有 qty/price |
| 2026-08-10 | **新增步驟 6.5**：本機 `--trades-only` 寫入 + `write_sync_state.py` |
| 2026-08-10 | **新增步驟 9.5**：`sync_push -Owner boss` 推到 GitHub |
| 2026-08-10 | **禁止事項改寫**：原本「不操作 `C:\TradeReview\`」「不操作 trade_review_app」與新流程衝突，改為精確的可寫／不可寫清單 |
| 2026-07-17 | 改為 OCR-only（現已解除） |
| 2026-07-13 | 實測步驟 10 切回 Charts |
| 2026-05-13 | 跑通全流程 |
