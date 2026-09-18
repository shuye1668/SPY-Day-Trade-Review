# Spy daytrade auto csv — 每日寫入 SOP（權威版）

本檔為「Spy daytrade auto csv」routine 的唯一權威定義。每日流程必須完全遵循本檔，
不得憑記憶重新發明寫法。輸入來源：使用者提供當日券商 Account Statement
（CSV 匯出或截圖 OCR），Cash Balance 區段含欄位：
`DATE, TIME, TYPE, REF #, DESCRIPTION, Misc Fees, Commissions & Fees, AMOUNT, BALANCE`。

## 鐵則（違反任一條即停止並回報，不得自行變通）

**【兩段式架構，2026-07-24】流程切成兩段，把「需要 AI」的部分壓到最小：
第一段 對帳單→`trades_all.xlsx`（有 CSV 用引擎全自動；只有截圖才需 AI，見 `PROMPT_轉trades_all.md`）；
第二段 `trades_all.xlsx`→`CS交易紀錄.xlsx` 由 **`cs_from_trades.py`** 純程式完成、**永不需要 AI**。
淨額精確反推：買進 net = −(價×股數)（無手續費），賣出 net = 損益 − 買進net，可逐筆還原真實手續費
（1.54/1.55/1.56 各異），**不違反 §1 禁止估算**。已對 07-22/23 共 36 筆做逐格回歸驗證。
AI 之後的角色＝檢查／修正／維護；AI 不可用時人工照《操作手冊》跑同一支腳本。**

**【引擎化，2026-07-20；寫入層 2026-07-24 接上並驗證】每日產出一律用引擎，禁止再手抄成
`process_*.py`/`write_*.py` literal，也禁止在對話裡手工推算 LIFO/損益/offset：
- 分析（dry-run）：`spy_daytrade_engine.py`（對帳單 CSV 為唯一真相源）。
- 寫入：`spy_daytrade_writer.py <csv> --date YYYY-MM-DD --commit`——已對 07-22/07-23 做過
  cell-identical 回歸（與人工核對逐格相同）、對漏單/留倉/BALANCE 對不上做過負向測試會正確 abort。
  自動備份、寫後 §四 全核驗、失敗自動還原；任何 `⚠️需人工確認` 一律不寫入。
先 dry-run 全綠才 `--commit`。詳見 `SKILL.md`、`MIGRATION.md`。舊每日腳本一律停用。**

0. **【最高優先，2026-07-15】本 routine 無人看管自動跑（每交易日 06:47）。凡遇任何「停下問使用者」情境（offset 對不上、Cash Balance 不完整無法校正、留倉/平倉致 offset 跳動、Day Header B 無法確定），一律把該儲存格填 `⚠️需人工確認`、對話首行加 `🔴 今日 N 處需人工確認`、受牽連下游（餘額鏈／合計／核驗／SUMIF／累積損益check）留空；絕不為了跑完流程而靜默寫入估算／推算／猜值。這是 2026-07 多筆 silent 錯誤的共同根因。詳見 SKILL.md §0。**

1. **手續費一律取對帳單 `Misc Fees` 欄的實際值。嚴禁任何估算、比例分攤、平均推算。**
   若某筆成交在紀錄中找不到費用，停下來問使用者，不要猜。
1b. **【BALANCE 交叉驗證閘門，2026-07-20】逐筆用 `開盤餘額 + Σ(AMOUNT+Misc Fees)` 重建餘額，
   必須與對帳單 `BALANCE` 欄逐列相符（誤差 < 0.01）。任一列對不上＝漏單／多單／費用錯，
   立即停止並標 `⚠️需人工確認`，不得寫入。這是攔截 OCR 漏單（如 7/16）的關鍵閘門。**
2. **每日區塊之間固定隔 2 個空行。**
   寫入位置 = 「用儲存格值掃描找到的最後一個非空白列」+ 3。
   **嚴禁使用 `ws.max_row + 1`**；`ws.max_row` 會被殘留格式撐大或誤判，必須以值掃描為準：
   ```python
   def last_nonempty_row(ws):
       for r in range(ws.max_row, 0, -1):
           if any(ws.cell(row=r, column=c).value not in (None, "") for c in range(1, 5)):
               return r
   header_row = last_nonempty_row(ws) + 3
   ```
3. **寫入前先備份** `CS交易紀錄.xlsx` 與 `trades_all.xlsx` 為 `*_backup_YYYYMMDD.xlsx`。
4. **寫入後必須通過本檔末尾的全部核驗**才算完成；任一失敗即還原備份並回報。

## 一、trades_all.xlsx（append）

欄位順序：`Date, Exec Time(EDT), Symbol, Price, Type, 損益(AI辨識), Shares, Action, Status, Pair_Date, Pair_Time`

- `Date`：**字串** `'YYYY-MM-DD'`。**嚴禁寫入 datetime 物件**（2026-07-10、07-13 曾因此漂移）。
- `損益(AI辨識)`：字串（如 `'97.08'`、`'-29.55'`），只寫在平倉列。
- 拆單合併：同一 REF # 的部分成交（如 40+60 股）合併為一列 100 股，
  價格按股數加權平均，`Misc Fees` 加總。
- `Type`：多 / 空 / 平（跨日平倉）；`Status`：當沖 `'0'`，跨日平倉 `'2'`。
- 從最後非空列 +1 直接續寫（此表不隔空行）。

## 二、CS交易紀錄.xlsx →「個別識別法」工作表

### 明細列數值定義（直接來自對帳單，不做任何運算以外的調整）
- 每列 A 欄 = 該筆成交的**淨現金流** = `AMOUNT + Misc Fees`（兩者皆取自對帳單；
  Misc Fees 本身是負數，BOT 通常無費用）。
  例：`SOLD -100 SPY @753.07`，AMOUNT 75,307.00、Misc Fees -1.57 → A = 75305.43。
  BOT +100 @751.935 → A = -75193.50。
- C 欄 = 該筆的平倉損益，只寫在平倉列（多單平倉在 SELL 列、空單平倉在 BOT 列），數值型別為 float。

### 區塊結構（逐列，嚴格照抄）
```
header_row   : A='YYYY.MM.DD'  B=<日首B值>  C='損益'
detail 第1列 : A=<淨額>  B='=+B{header_row}+A{r}'  C=<損益或空>
detail 其餘  : A=<淨額>  B='=+B{r-1}+A{r}'         C=<損益或空>
totals_row   : B='合計 '  C='=SUM(C{detail_start}:C{detail_end})'
verify_row   : B='核驗'   C='=+B{detail_end}-B{header_row}'
(1 空行)
spy_buy_row  : A='YYYY.MM.DD當沖'  B='買'  C='=SUMIF(A{ds}:A{de},"<0")'
spy_sell_row : A='="SPY " & COUNT(C{ds}:C{de}) * 100 & "股"'  B='賣'  C='=SUMIF(A{ds}:A{de},">0")'
spy_pnl_row  : B='當日損益'    C='=+C{spy_buy_row}+C{spy_sell_row}'
spy_pct_row  : B='當日損益 %'  C='=C{spy_pnl_row}/ABS(C{spy_buy_row})'
```
- 數字格式：買/賣/當日損益 C 欄 `'0.00_);[Red](0.00)'`，當日損益 % `'0.00%'`。
- `日首B值` = 前一日區塊尾端 B 值延續：`券商當日開盤餘額 + offset_prev`
  （`offset_prev` = 上一日 header_B − 上一日券商開盤餘額，需 assert 頭尾一致；
  2026-07-14 校驗時 offset = −122.18）。
- **券商會追溯重編日開盤餘額**（同一天在不同日期匯出的對帳單數字可能不同），
  核驗一律以**最新一份匯出**的 BAL 列為準。
- 若當日有留倉/平留倉等特例（如 2026.07.09），依當次情況在當沖摘要後加註區塊，
  並相應調整 核驗 與 SUMIF 排除留倉列（參照 1866–1903 列的既有寫法）。

## 三、CS交易紀錄.xlsx →「累積損益check」工作表

最後非空列 +1 續寫：A=當日日期（datetime，維持既有型別）、B=當日損益(float)、
C=`'=C{上一列}+B{本列}'`。

## 四、寫入後核驗（全部必須通過）

1. `sum(A明細) == sum(C損益)`（當日全平時），誤差 < 0.01。
2. 頭尾雙 offset：`header_B − 券商開盤餘額 == 尾端B − 券商收盤餘額 == offset_prev`，誤差 < 0.01。
3. 獨立自洽：`sum(A明細) == 券商收盤餘額 − 開盤餘額`，誤差 < 0.01。
4. 重新掃描工作表：新 header 前面的空白列數 **必須正好等於 2**。
5. 明細 B 欄公式逐列比對 `'=+B{r-1}+A{r}'` 樣式無誤。
6. trades_all 新寫入列的 Date 欄型別為 str。
7. 當日損益與「累積損益check」B 欄一致。

## 五、Trade Review App 對接注意

- `trade_review_app.py` 以 `pd.read_excel(dtype=str)` + `_norm_date` 讀 trades_all，
  並會在分析後把整表 write-back（字串化）。routine 只負責 append，不要動舊列。
- App 不讀寫 CS交易紀錄.xlsx，CS 格式問題與 App 無關。
- 寫入時确保兩個 xlsx 沒有在 Excel 中開啟（檢查 `~$` 鎖定檔）。
