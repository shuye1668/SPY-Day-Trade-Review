# Spy daytrade auto csv — 每日寫入 SOP（權威版）

本檔為「Spy daytrade auto csv」routine 的唯一權威定義。每日流程必須完全遵循本檔，
不得憑記憶重新發明寫法。輸入來源：當日券商 Account Statement，
Cash Balance 區段含欄位（**9 欄與 TOS 原生 10 欄皆可，引擎依欄名對應**）：
`DATE, TIME, TYPE, REF #, DESCRIPTION, Misc Fees, Commissions & Fees, AMOUNT, BALANCE`
或 `Trade Date, Exec Date, Exec Time, Type, Ref #, Description, ..., Amount, Balance`。

**2026-08-10 起改為雙機運作**：對帳單不再由使用者手動提供，而是由 Boss PC
（交易機）採集後經 GitHub private repo 同步過來。見下方〈§〇 雙機架構〉。

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

**【雙機分工，2026-08-10】兩段式架構現在跑在兩台機器上，各自只寫自己的檔，
寫錯邊會製造出永遠推不出去、還會卡住對方 git pull 的本機改動：
- **Boss PC（交易機）只做第一段**：對帳單 → `trades_all.xlsx`，
  一律用 `spy_daytrade_writer.py <csv> --commit --trades-only`。
  **絕不可寫 `CS交易紀錄.xlsx` / `offset_state.json` / `notes/`。**
- **502 只做第二段**：`trades_all.xlsx` → `CS交易紀錄.xlsx`，用 `cs_from_trades.py`。
  **不主動改 `trades_all.xlsx` / `_inbox/`**（那是 Boss 的產物；App 的配對回寫除外，
  同步時會被丟棄重取）。
`--trades-only` 保留 `gate()` 全部安全閘門（§1b BALANCE、留倉、非交易現金、
needs_review），一條都沒放寬。詳見 §〇。**

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

## 〇、雙機架構與同步（2026-08-10 起）

```
Boss PC（交易機，C:\TradeReview）              502（D:\fileserver_D\TradeReview）
────────────────────────────────              ──────────────────────────────────
TOS → 採集（CSV 匯出為主／OCR 備援）
  ├─ engine 乾跑（§1b BALANCE 閘門）
  ├─ writer --trades-only → trades_all.xlsx     ← 第一段結束
  ├─ write_sync_state.py（新鮮度戳記）
  ├─ Gmail 草稿（保留為退路）
  └─ sync_push -Owner boss ─→ GitHub(private) ─→ sync_pull -Owner p502
                                                   ├─ cs_from_trades.py --from-statement
                                                   │    （期初/收盤餘額自動帶入）
                                                   ├─ CS交易紀錄.xlsx ← 人工修正在這
                                                   └─ sync_push -Owner p502
```

### 檔案所有權（這是同步不衝突的關鍵）

xlsx 是 binary，git 無法 merge；兩台若改同一個 .xlsx 就是死結。分區之後
兩邊改的檔案不重疊，`pull --rebase` 永遠不需要人工解衝突。

| 檔案 | 擁有者 | 另一台 |
|---|---|---|
| `_inbox/*.csv`（原始對帳單） | Boss | 只讀 |
| `trades_all.xlsx` | Boss | 只讀 |
| `sync_state.json` | Boss | 只讀 |
| `CS交易紀錄.xlsx`、`CS交易紀錄_dump.txt` | **502** | Boss 只讀 |
| `offset_state.json` | **502** | Boss 只讀 |
| `notes/` | **502** | Boss 只讀 |
| `*.py` `*.md` `*.bat` `*.ps1`（程式碼） | **502** | Boss 靠 pull 取得 |
| `history_minute.xlsx` | 各機獨立，**不進 git** | 4.3MB，每天被 app 追加 |

`sync_pull.ps1 -Owner <boss|p502>` 會先丟棄本機對「對方所有物」的改動再 pull，
所以誤寫也能自動清乾淨、不會卡住。

### 同步指令

```powershell
powershell -ExecutionPolicy Bypass -File sync_pull.ps1 -Owner p502   # 取得 Boss 產出
powershell -ExecutionPolicy Bypass -File sync_push.ps1 -Owner p502   # 推出 CS/notes/程式碼
```

⚠️ **禁止 `git add -A`** —— 一律用 `sync_push.ps1`，它只 stage 該台擁有的檔。

### 502 的自動排程（2026-08-19 建立，取代每日手按 .bat）

| 排程 | 時間 | 做什麼 |
|---|---|---|
| `TradeReview_Dashboard` | 開機 + 每 10 分看門狗 | pythonw 常駐 app（含下方「主動補 K 線」） |
| `TradeReview_SyncPull_p502` | 09:30、11:30 | `sync_pull.ps1 -Owner p502` 取回 Boss 的 `trades_all` / `_inbox` |
| `TradeReview_SyncPush_p502` | 10:00、12:00 | `sync_push.ps1 -Owner p502` 推出 CS 帳／notes／程式碼 |

順序刻意是**先拉再推**（09:30 拉 → 10:00 推 → 11:30 拉 → 12:00 推）。
`StartWhenAvailable`：當時沒開機會在開機後補跑。

`每日一鍵複盤.bat` **保留為手動退路**（互動式、有 `choice`/`pause`，不能排程）。
正常情況不需要按它；要按的時機只剩「Boss 沒跑成、要用手邊 CSV 自己補」。

⚠️ **`sync_pull` 的 dirty 閘門只看已追蹤檔**（`--untracked-files=no`）。
未追蹤檔不會擋 rebase pull，但原本的 `--porcelain` 會把它們也算成 dirty，
於是資料夾裡只要躺著一個沒進 git 的檔（例如 `_inbox\2026-08-11.csv`），
就會每天 exit 2 靜默跳過 —— 無人看管下等於同步永久停擺。**不要改回去。**

### 三個容易踩的環境陷阱（都已在腳本內處理，改動時勿破壞）

1. **Boss PC 的 ExecutionPolicy 全 scope 為 Undefined（實際 = Restricted）**，
   直接跑 `.ps1` 或 dot-source 會被擋。腳本一律用
   `powershell -ExecutionPolicy Bypass -File`；需保留函式時在該 process 內先下
   `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force`。
2. **PowerShell 5.1 讀沒有 BOM 的 UTF-8 `.ps1` 會當 ANSI**，中文註解會炸掉語法。
   改過任何 `.ps1` 後跑 `python _fix_ps1_bom.py`。
3. **`.gitattributes` 把 `*.csv` 設為 `-text`**（不做行尾轉換）。
   曾因 `text eol=lf` 與 `core.autocrlf=true` 互相拉扯，造成 index 與工作區永遠
   對不上、每次 `git pull` 都以 "You have unstaged changes" 失敗。**不要改回去。**

### 資料新鮮度（防止同步靜默失敗）

Boss 每次跑完寫 `sync_state.json`（最後成功的交易日 + 狀態），隨 git 同步；
App 標頭顯示「資料截至 X」。判準是**落後幾個美股交易日**，不是幾小時 ——
Boss PC 週末關機，週五的帳常等週一較晚才補，用時數判斷會每週誤報一次。

| 狀態 | 意義 |
|---|---|
| 灰「資料截至 X」 | 正常 |
| 琥珀「⚠待補前一交易日」 | 落後 1 日。週一早上出現是正常待處理，不是故障 |
| 琥珀「⚠需人工確認」 | Boss 的 gate 擋下了某日，需依 §0 人工處理 |
| 紅「✕落後 N 個交易日」(N≥2) | 真的有事沒跑成 |
| 紅「✕採集失敗」 | Boss 採集階段就失敗了，看 Gmail 草稿的失敗原因 |

### 相關文件

| 檔案 | 內容 |
|---|---|
| `BOSS_CLAUDE.md` | Boss PC 採集 routine 完整版（要複製到 `C:\TradeReview_Boss\CLAUDE.md`） |
| `SETUP_BOSS_PC.md` | Boss PC 建置與每日流程 |
| `_BossPC_安裝包/` | 帶去 Boss PC 的檔案（含無法走 git 的 `history_minute.xlsx`） |

---

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

### 主動補 K 線（2026-08-19 加，勿移除）

App 內建 `candle-refresher` daemon 執行緒：每 60 分鐘把最近 7 個日曆日中
**已收盤**的交易日補進 `history_minute.xlsx`。

修的是一個死鎖：`_append_to_history` 只由 `fetch_intraday` 呼叫，而 `fetch_intraday`
只在 `/api/data`（有人點開某天）時才跑。沒有交易的日子不在 `trades_all` 裡、也還不在
`history` 裡 → 不會出現在 `/api/dates` 清單 → 點不到 → 永遠不會被補。
結果就是「沒交易的日子看不到股價線」。先主動補進 history，那天就會自己出現。

三個刻意的設計，改動時勿破壞：

1. **只收已收盤的盤（16:05 ET 之後）**。盤中抓只會拿到半天 bar，而
   `_append_to_history` 見日期已存在就不再更新 —— 那筆殘缺資料會被**永久凍結**。
2. **放在 app 內部，不另寫排程腳本**。`history_minute.xlsx` 也被 app 回寫，
   兩個 process 同時重寫這個 4.4MB 活頁簿正是 2026-08-10 把它寫成截斷 zip 的成因；
   同一 process 內走既有 `_file_lock` + `_atomic_to_excel`，沒有跨程序競態。
3. **每輪清掉該日的 `_negative_cache`**，但用 `REFRESH_MAX_ATTEMPTS=3` 收斂。
   常駐 app 下，一次網路抖動會讓那天在整個 process 生命週期內不再重試；
   而假日若無上限則會每小時空敲 yfinance。

### 美股行事曆與「安靜要能被證明」（2026-09-08 加）

`us_market_calendar.py`：全天休市日（含耶穌受難日、觀察日順延）與
`latest_completed_session()`。已用 `history_minute.xlsx` 裡 339 個真實交易日
反向檢驗，**休市判定 0 誤判**。

起因：2026-09-07 勞動節休市，使用者看到「K 線最新只到 9/4」卻無從分辨那是
正確還是壞掉 —— 原本 refresher 只排除週末，遇到平日休市會照樣問 yfinance、
抓空、靜默重試三次後放棄，**完全沒有留下任何紀錄**。

三項對策，改動時勿破壞：

1. `_settled_sessions` 用行事曆排除休市日，不再空敲 yfinance。
2. `_refresh_recent_candles` **每輪都寫日誌**（`_rlog`，強制 flush ——
   pythonw 下 stdout 是檔案，區塊緩衝會讓運維訊息卡在記憶體裡好幾小時，
   正好在你要查問題時看不到）。無需補件時也會說明「為什麼今天沒有新資料」。
3. `/api/version` 增加 `candles` 區塊、標頭增加 **K 線徽章**（`#cdlbadge`），
   直接回答「應該要有哪天、有沒有到手、今天為何沒有盤」。
   它與交易帳徽章（`#syncbadge`）**刻意分開** —— 股價線與交易帳是兩條獨立的鏈，
   任一條停掉都必須各自看得出來。

- `trade_review_app.py` 以 `pd.read_excel(dtype=str)` + `_norm_date` 讀 trades_all，
  並會在分析後把整表 write-back（字串化）。routine 只負責 append，不要動舊列。
- App 不讀寫 CS交易紀錄.xlsx，CS 格式問題與 App 無關。
- 寫入時确保兩個 xlsx 沒有在 Excel 中開啟（檢查 `~$` 鎖定檔）。

### 寫入互斥（2026-08-10 補強，勿移除）

App 是 pythonw 常駐、每 4 秒偵測 mtime，而它的 write-back 會重寫**整張**
trades_all。`to_excel()` 是就地寫入，中途被另一個程序介入留下的是**不可讀的
截斷 zip**，不是舊資料 —— 2026-08-10 就是這樣把 `history_minute.xlsx`
（4.4MB）寫壞的：當時同時有兩個 app 實例在跑。三道防護：

1. **`.writer.lock`** — `spy_daytrade_writer.py` 寫入期間持鎖，App 見鎖即跳過
   write-back；鎖逾時 600 秒視為殘留（避免 writer 被中斷後 App 永遠不再回寫）。
2. **`_atomic_to_excel()`** — 先寫完整暫存檔再 `os.replace()`（同磁碟區為原子操作），
   讀取端永遠只會看到完整檔案。`history_minute.xlsx` 與 `trades_all.xlsx` 皆適用。
3. **`_file_lock()`** — O_EXCL 鎖檔，擋第二個 app 實例同時重寫。
   「工作排程常駐 + 手動再開一個」在這裡是常態，不是異常。

⚠️ `history_minute.xlsx` **沒有備份來源**：yfinance 的 1 分鐘資料只回得了最近
約 30 天，Bloomberg 的 IntradayBarRequest 也回不了一年多前 —— 且它刻意不進 git。
請另外保留備份。
