# Boss PC 採集 prompt — 需要替換的段落

> Boss PC 的採集 prompt 在 `C:\TradeReview_Boss\CLAUDE.md`，不在本 repo 內。
> 這份列出**必須替換的段落**，直接照抄過去即可。
> 建立於 2026-08-10。

替換的四處：**步驟 5（改為 CSV 為主）**、**步驟 6（OCR 要抓每一列）**、
**步驟 7（驗算只對 TRD）**、**禁止事項（解除對 C:\TradeReview 的封鎖）**，
另外新增 **步驟 6.5（本機寫入）** 與 **步驟 9.5（推送）**。

---

## 【替換】步驟 5 — 取得對帳單（CSV 為主，OCR 為備援）

```
──────────────────────────────────────────────────────────────
【步驟 5】取得對帳單 —— 主路徑：CSV 匯出
──────────────────────────────────────────────────────────────
⚠️ 2026-08-10 起改回 CSV 為主。引擎已能直接吃 TOS 原生的 10 欄格式
   （Trade Date,Exec Date,Exec Time,Type,Ref #,Description,Misc Fees,
     Commissions & Fees,Amount,Balance），不需要任何轉換。
   ——— 這是先前停用 CSV 路徑的原因，現已在引擎端修好。

5.1 在 Account Statement 右上找到選單（三條線圖示）→ Export to file
5.2 Save dialog 出現後，路徑輸入：
        C:\TradeReview\_inbox\<today>.csv
    存檔類型選 CSV。按存檔。
5.3 Capture-Screen 後 view 確認：
    - Save dialog 已關閉
    - 沒有「檔案已存在，是否覆蓋」之類殘留對話框
5.4 用 PowerShell 確認檔案確實產生且非空：
        $f = "C:\TradeReview\_inbox\$today.csv"
        if (Test-Path $f) { (Get-Item $f).Length } else { "MISSING" }
    - 檔案存在且 > 500 bytes → 成功，跳到步驟 6.5
    - 否則 → 走 5' OCR 備援

──────────────────────────────────────────────────────────────
【步驟 5'】備援路徑 —— 截圖 OCR
──────────────────────────────────────────────────────────────
（原步驟 5 的截圖流程照舊：ESC 關 tooltip → Capture-Screen →
  view 確認最新一列有入鏡 → 列數 overflow 就捲動多截幾張）
截圖完成後進步驟 6。
```

---

## 【替換】步驟 6 — OCR 要抓「每一列」，不只 TRD

```
──────────────────────────────────────────────────────────────
【步驟 6】OCR 截圖（僅備援路徑需要）
──────────────────────────────────────────────────────────────
🔴 最重要的修正（2026-08-10）：
   Cash Balance 區塊的【每一列都要抓】，不是只抓 TRD。
   必抓的類型：TRD（成交）、BAL（期初餘額）、DOI（配息）、
               JRN（轉帳）、ADJ（調整），以及任何其他 TYPE 值。

   為什麼：引擎靠 BAL 列切出交易日、取得期初餘額、並執行 §1b
   「逐列餘額重建」驗證。沒有 BAL 列 → 引擎報「session 開頭缺 BAL 列」
   → writer 一定 abort，當天一筆都寫不進去。

   這件事已經出事過。offset_state.json 自己記著：
     「8/05 open-bal 115851.99 是由成交鏈回推(今日OCR草稿漏BAL/DOI/JRN列)，
       非直接讀BAL行 ... 唯 offset 數值為暫定值，待補齊含BAL列的完整
       對帳單後覆核」

   BAL 列長這樣，通常在表格最上面一列，很容易被當成標題略過：
     7/31/26,13:00:00,BAL,,Cash balance at the start of business day 31.07 CST,,,,36372.84
   DOI（配息）長這樣：
     7/31/26,16:39:50,DOI,126779764549,STATE STREET SPDR S&P 500 ETF TRUST 5900.9 US$,,,5900.90,42273.74

每列要抓的欄位：
  Trade Date / Exec Date / Exec Time / Type / Ref# / Description
  / Misc Fees / Comm & Fees / Amount / Balance

⚠️ 日期欄若截圖同時有 Trade Date 與 Exec Date，兩個都要抓。
   Trade Date 是券商營業日，午夜之後的成交會掛在前一個營業日；
   Exec Date 才是真正的日曆日。實測 2026-07-29 的對帳單有 5 列
   Trade Date=7/27 但 Exec Date=7/28 01:18 —— 取錯就整整差一天。

OCR 注意事項：
  - Description 裡的 qty（+100 / -100 / 分批 +80 +20）與 price（@xxx.xxxx）要看清楚
  - Amount / Balance 有千分位逗號，去掉逗號再存
  - 分批成交（同 Ref#、同秒、qty 拆兩列）要各存一列
  - 非 SPY 標的照樣抓，Description 保留原標的字樣
  - 任一 TRD 列的 qty/price/amount 辨識不清 → 見步驟 7 失敗處理
```

---

## 【替換】步驟 7 — 驗算只對 TRD 做

```
──────────────────────────────────────────────────────────────
【步驟 7】驗算
──────────────────────────────────────────────────────────────
對每一列【TRD】驗算 qty × price ≈ abs(amount)，容差 $0.02。

⚠️ BAL / DOI / JRN / ADJ 列【不做這個驗算】—— 它們沒有 qty/price，
   驗算必然失敗。這些列原樣帶入 CSV 即可，引擎會自行處理：
   BAL 用來取期初餘額，DOI/JRN 會被 gate 標為「非交易現金異動」
   而擋下該日交由人工判斷（這是正確行為，不是錯誤）。

  - 逐列計算 diff = |qty × price − abs(amount)|，累積 maxDiff
  - diff > $0.02 的 TRD 列：草稿主旨改 FAILED，正文記該列詳情後結束
  - 任一 TRD 列 qty/price/amount 無法讀出：視為驗算失敗，
    建 FAILED 草稿「OCR unreadable row: <定位資訊>」後結束
```

---

## 【新增】步驟 6.5 — 寫入本機 trades_all

```
──────────────────────────────────────────────────────────────
【步驟 6.5】本機寫入（2026-08-10 新增）
──────────────────────────────────────────────────────────────
$py  = "C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe"
$csv = "C:\TradeReview\_inbox\$today.csv"
cd C:\TradeReview

# 先乾跑，不寫任何檔案
& $py spy_daytrade_engine.py $csv
# 全綠（Σcol1 == Σpnl、留倉=False、無 ⚠️/🔴）才寫入
& $py spy_daytrade_writer.py $csv --commit --trades-only
$wr = $LASTEXITCODE

# 記錄狀態供 502 端的 App 顯示新鮮度
if ($wr -eq 0) { & $py write_sync_state.py --date $today }
else           { & $py write_sync_state.py --date $today --status blocked --note "writer gate 擋下，需人工確認" }

🔴 writer 一定要帶 --trades-only。
   不帶會去寫 CS交易紀錄.xlsx 與 offset_state.json —— 那兩個是 502 的
   所有物，會製造出永遠推不出去、還會卡住自己 pull 的本地改動。

exit code：0=成功或冪等跳過  2=被 gate 擋下（有 🔴）  其他=例外

被擋下時【什麼都不要改】。trades_all 停在前一日是正確的，不是壞掉。
照常繼續步驟 8 建草稿，把被擋原因寫進正文，交由 502 人工處理。
（CLAUDE.md §0：絕不為了跑完流程而寫入估算／推算／猜值。）
```

---

## 【新增】步驟 9.5 — 推送到 GitHub

```
──────────────────────────────────────────────────────────────
【步驟 9.5】推送（2026-08-10 新增）
──────────────────────────────────────────────────────────────
cd C:\TradeReview
.\sync_push.ps1 -Owner boss

只會推 trades_all.xlsx / _inbox\ / sync_state.json 三樣。
push 失敗會自動重試 3 次；三次都失敗時本機 commit 仍在（資料沒掉），
請把這件事寫進步驟 11 的 log 與 Gmail 草稿正文。
```

---

## 【替換】禁止事項

原本這兩條與新流程直接衝突，必須改寫：

```
  ❌ 舊：不操作 C:\TradeReview\ 下任何檔案（那是使用者電腦端的）
  ❌ 舊：不操作 trade_review_app
```

改為：

```
──────────────────────── 禁止事項 ────────────────────────
* 不點 TOS 任何 Buy/Sell/Place Order 按鈕
* 採集過程不切到 Account Statement 以外的 TOS 頁面
  （唯一例外是收尾切回 Charts，且只點頁籤本身）
* 不重啟 TOS、不嘗試輸入密碼
* 不送出 Gmail 草稿（只 drafts.create）

【C:\TradeReview\ 的存取規則（2026-08-10 起）】
* 可以寫：_inbox\*.csv、trades_all.xlsx、sync_state.json
  —— 但 trades_all 只能透過 spy_daytrade_writer.py --trades-only 寫，
     絕不可用 Excel 或任何方式手動編輯
* 絕對不可寫：CS交易紀錄.xlsx、offset_state.json、notes\
  —— 這三樣是 502 的所有物
* 不可修改 trade_review_app.py 或任何 .py（程式碼在 502 改，Boss 靠 git pull）
* 不可執行不帶 --trades-only 的 spy_daytrade_writer.py
* 不可為了讓流程跑完而修改任何數字（CLAUDE.md §0）
* 不可用 git add -A —— 一律用 sync_push.ps1 -Owner boss
```
