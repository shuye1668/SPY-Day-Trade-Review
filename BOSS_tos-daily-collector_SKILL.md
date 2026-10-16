# 每日 TOS 採集 routine（Boss PC）—— 502 端鏡像

> **本檔性質**：Boss PC 現行 SKILL 的**鏡像/存檔**，權威版在
> `C:\Users\TS USER\.claude\scheduled-tasks\tos-daily-collector\SKILL.md`（Boss 本機）。
> **不要拿本檔整檔取代 Boss 端。** 下方全文是 2026-10-16（第二次）由使用者貼回的現行版
> （「最後更新：2026-10-16」，含座標 build 1993 校正、Activate-TOSWindowVerified、
> 步驟 4 N≥5 不動、Open-ExportDialog/Confirm-ExportSave、9.0b、9.5 rebase 防呆、Switch-ToCharts）。
> 貼上過程 markdown 已扁平化；位元組級正本將在 git 修復後以 `BOSS_SKILL_current.md` 隨 sync_push 送達。
>
> **剩餘已知差異（git 修復後要對齊）**：
> 1. 步驟 9.5 內文說推「trades_all / _inbox / sync_state / verify_shots 四樣」——Boss pull 到新版
>    sync_push.ps1 後實際只推 `_inbox\`、`verify_shots\`、`BOSS_SKILL_current.md` 三樣
>    （trades_all/sync_state 歸 502，由 Boss 端 sync_pull 丟棄其本機改動）。Boss 的 SKILL 文字屆時要再改。
> 2. Boss 端 `C:\TradeReview_Boss\CLAUDE.md` 是否已同步 9.0b 未確認。
> 3. 修復 runbook 要求的「每次改 SKILL 後 copy 到 C:\TradeReview\BOSS_SKILL_current.md」尚未寫進 Boss SKILL 本文。

---

以下為 Boss 端 2026-10-16 現行全文（使用者貼回，格式扁平化）：

每日 TOS 採集 routine（Boss PC，權威版）
本檔是自足的 —— 執行時不需要再去讀別的規則檔。 最後更新：2026-10-16。

⚠️ 規則只能有一份。 2026-08-11 查出的事故：這個排程原本內嵌 2026-07-17 的 OCR-only 流程，同時 C:\TradeReview_Boss\CLAUDE.md 另有一份較新的規則，兩者 互相矛盾（前者明文禁止 CSV 匯出與寫入 C:\TradeReview\）。直接指令會贏， 導致本機寫入與 git 同步從未執行、每天都要人工補帳。 要改流程請改本檔，不要在別處另寫一份。

這台機器在整套系統裡的角色
依兩段式架構，Boss PC 只做第一段：對帳單 → trades_all.xlsx。

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

Email（2026-10-16 起）：只有出現「尚未寄過的新成交」時，才建草稿並由本 routine 直接寄出到 juliahsu13@gmail.com（見步驟 8～9）。沒有新成交就不建草稿、不寄信。 （舊做法「草稿交給 Apps Script 寄到 sunbeamichelle」已停用：Apps Script 自 2026-05-14 起 就沒有寄出過任何一封這類草稿。）

🔴 三條鐵則
1. writer 一定要帶 --trades-only。 不帶會去寫 CS交易紀錄.xlsx 與 offset_state.json —— 那是 502 的所有物，會製造出永遠推不出去、還會卡住 自己 pull 的本機改動。
2. 不得手動編輯任何 .xlsx。 所有寫入只能經由 engine / writer。
3. 不得用 git add -A。 一律用 sync_push.ps1 -Owner boss。

執行時機（這台週末關機）
排程：台灣週二～週六 06:00（cron 0 6 * * 2-6，對應美股週一～週五收盤後；排程器會加幾分鐘延遲）。 Windows 工作排程器的 Claude-TOS-Collector-Launch 在同一天 05:45 喚醒電腦並開啟 Claude Desktop， 讓這個排程能觸發（Claude 沒開著的話，排程不會跑）。美股休市日由步驟 0 判斷後直接結束。 機器關機（不是睡眠）時喚醒不了，那天就不會跑，開機後手動補跑即可。所以：
- 一次跑會補完所有欠帳：步驟 4 要求「至少 5 days back」（平常是 7）不是保險，是必要的。 writer 會處理 CSV 裡所有 session，已寫過的自動冪等跳過。
- 週一落後一個交易日是正常的，不是故障。502 的 App 標頭那段時間會顯示 琥珀色「⚠待補前一交易日」，補完轉回灰色。
- 顯示紅色「✕落後 N 個交易日」（N≥2）才代表真的有事沒跑成。

執行環境（會踩到的四個坑）
- Claude Desktop 會搶走前景視窗（2026-10-16 實測）：Activate-TOSWindow 回報成功後， Claude 視窗常在一兩秒內又跳回前面，點擊就落在 Claude 上而不是 TOS（症狀：截圖出現 Claude 畫面、 選單沒開、表格某列被選取）。每一次點擊 TOS 之前，都要在同一個 call 裡先跑 Activate-TOSWindowVerified（反覆啟用直到前景真的是 TOS；回傳 False 就不要點）。 需要連續操作的（開選單→選項目）一律放在同一個 call。
- 此 session 不會有 desktop computer-use MCP。 不要 tool_search 找 computer use、不要用 Claude_in_Chrome__computer （那個要 browser tabId，控制不了桌面程式）。一律走 PowerShell + Win32， 所有需要的功能 tos_automation.ps1 都有。
- ExecutionPolicy 會擋 .ps1（本機全 scope 為 Undefined＝實際 Restricted）。 每個需要 dot-source 的 call 開頭都要先下 Set-ExecutionPolicy -Scope Process； 執行獨立腳本則用 powershell -ExecutionPolicy Bypass -File。
- 每次 PowerShell tool call 都是新 process，變數不跨 call 保留。 每個 call 都要重新 dot-source，重要狀態要落檔。

工作流程：每個關鍵步驟後 Capture-Screen 存 PNG 到 C:\TradeReview_Boss\logs\， 再用 view tool 讀圖確認畫面，確認後才進下一步。

常用變數
$py    = "C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe"
$root  = "C:\TradeReview"          # 複盤系統（git repo）
$boss  = "C:\TradeReview_Boss"     # 採集系統（本檔相關資源）
$today = (Get-Date).ToString("yyyy-MM-dd")

【步驟 0】前一晚美股有沒有開盤
& $py C:\TradeReview_Boss\us_trading_day.py
- US_SESSION=YES → 繼續步驟 1。
- US_SESSION=NO（週末或 NYSE 休市）→ 什麼都不碰（不動 TOS、不建草稿、不寄信、不跑 writer／sync_state），只在 log 寫一行 [時間] SKIP - no US session (<日期> <原因>) 後結束。
- 輸出含 WARNING: no NYSE holiday list → 照常繼續，但在最終回報提醒要更新 us_trading_day.py 的休市日清單。
- 手動補跑（例如週一開機補週五）時不受此限：直接從步驟 1 開始。
- 沒有新成交時「回復原樣」：步驟 8 判定 NEW_TRADES=0 → 不建草稿、不寄信，照樣做步驟 10 把 TOS 切回 Charts，寫 log 後結束。

【步驟 1】載入 helper + 取得今天日期
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview_Boss\tos_automation.ps1
$today = (Get-Date).ToString("yyyy-MM-dd")
預期：tos_automation.ps1 loaded.（Find/Activate/Get-TOSWindow*、Click-At/TripleClick-At/Send-Keys/Send-Hotkey、Capture-Window/Screen/Region）

【步驟 2】啟用 TOS 視窗
$ok = Activate-TOSWindowVerified
if ($ok) { Capture-Screen "C:\TradeReview_Boss\logs\dbg_tos_activated_$today.png" }
- $ok = False → 存 logs\debug_no_tos_$today.png，建 FAILED 草稿「TOS window not found」後結束。
- $ok = True → view 該 PNG，確認 TOS 在前景且最大化、顯示交易介面（不是登入頁）。

【步驟 3】切到 Account Statement
從步驟 2 截圖判斷當下狀態：
- 情境 1：已在 Monitor → Account Statement → 跳到步驟 4
- 情境 2：在 Monitor 但不是 Account Statement 子頁籤 → 從截圖找到子頁籤座標 → Click-At → 等 1 秒截圖確認
- 情境 3：在其他主頁籤（上次收尾停在 Charts，正常起點）→ 同一個 call：Activate-TOSWindowVerified → Click-At 262 33（Monitor）→ 等 1 秒 → [Win32Api]::SetCursorPos(1000, 900)（游標移開，hover 選單才收）→ 截圖 → 通常直接停在 Account Statement；不是再點 Click-At 455 55（同樣先 Verified）
座標（2026-10-16 實測，TOS build 1993，1920x1080）：主頁籤列 y≈33，Monitor=262、Charts=557；子頁籤列 y≈55，Activity and Positions≈318、Account Statement≈455。新版本 TOS 可能再變，以當次截圖為準。

【步驟 4】確認日期範圍「至少 5 days back」
⚠️ 要 ≥5 天。N≥5 就不要動：這台平常顯示 7 days back from today，已涵蓋需求。2026-10-02 曾為了把 7 改成 5 在輸入框打字，TOS 跳「Unexpected error」（按 Continue 可恢復，絕不按 Restart，取消勾選回報）。
4.1 status bar 判斷：
| 看到 | 情境 |
|---|---|
| N days back from today，N≥5（常見 7）| A → 不動，跳步驟 5 |
| 1～4 days back | B → 要改 |
| from X to Y 且涵蓋不到最近 5 天 | B → 改回 days back |
| 完全找不到 | C |
4.3 情境 B（只有真的不夠才做；優先用輸入框旁上下箭頭、每點一次截圖確認，打字是最後手段）：
a) 截圖找「N days back from today」精確座標（1920x1080 y≈78-80、x≈600-700；y 差 3-5px 就 miss）。
b) Click-At 後截圖 dbg_dialog。c) 確認對話框（Statement for: / ◉ [N] days back from Today / ○ from…to…）。
d) 確認上方 radio 選中（不動 from/to）。e) TripleClick-At 輸入框 + "5" + {ENTER}（單一 call）。
f) Sleep 8 g) {ESC} h) 截圖驗證 N≥5。
4.4 情境 C：存 debug_no_status_bar，建 FAILED 草稿後結束。

【步驟 5】★ 主路徑：CSV 匯出
引擎直接吃 TOS 原生 10 欄格式。log 實證 CSV 連續 44 次成功；OCR 覆蓋率僅約 6 成。
5.1 開匯出對話框（單一 call；2026-10-16 實測唯一可行做法）：
$ok = Open-ExportDialog      # Verified 啟用 → Click-At 1899 60 開選單 → 鍵盤 7×DOWN + ENTER
Capture-Screen dbg_savedialog
- 選單共 8 項，「Export to file...」是第 7 項。滑鼠點選單項目不會觸發（實測多次失敗），只能用鍵盤；開選單則要用 Click-At。
- view 確認對話框標題 Account Statement for ... since M/D/YY through M/D/YY、資料夾 C:\TradeReview\_inbox、類型 CSV、檔名預填 <today>-AccountStatement。
5.2 存檔＋改名：⚠️ 不要在對話框打字——這台輸入法是注音，數字會被吃成中文字。接受預設檔名：
Confirm-ExportSave           # 對話框前景（確認後）點 存檔(S) (1155,775)
Move-Item "_inbox\$today-AccountStatement.csv" "_inbox\$today.csv" -Force
5.3 截圖確認 dialog 已關、無覆蓋提示。
5.4 檔案存在且 > 500 bytes → 跳步驟 6.5；否則走步驟 5'。
重試限制（10-01、10-16 已實機驗證可行仍限制）：三條線/Export 最多找 2 次；Save dialog 失敗最多重試 1 次；不探索其他功能表。降級 OCR 是正常結果。

【步驟 5'】備援路徑：截圖（僅步驟 5 失敗）
5'.1 {ESC}。5'.2 Capture-Screen screenshot_$today.png。
5'.3 view 確認：N days back（N≥5）可見、Cash Balance 多列、無遮擋、最後一列入鏡（截斷＝overflow→5'.4）。
5'.4 overflow：點表格空白處 {END} 捲底，分段 _b/_c，OCR 後靠 Ref#/時間去重合併；仍不完整 → FAILED 草稿結束。

【步驟 6】OCR（僅備援路徑）
🔴 Cash Balance 每一列都要抓（TRD/BAL/DOI/JRN/ADJ）——BAL 漏抓引擎必 abort（8/05 出過事，offset_state.json 有記）。
每列欄位：Trade Date / Exec Date / Exec Time / Type / Ref# / Description / Misc Fees / Comm & Fees / Amount / Balance。
兩個日期欄都要抓（Exec Date 才是真日曆日；7/29 實測差一天）。千分位逗號去掉；分批成交各存一列；非 SPY 照抓。
寫入 _inbox\<today>.csv（9 或 10 欄皆可，引擎依欄名對應）。

【步驟 6.5】★ 寫入本機 trades_all
engine 乾跑 → 全綠才 writer --commit --trades-only → write_sync_state.py
（成功不帶 --date；gate 擋下 --status blocked --note "writer gate 擋下，需人工確認"；
⚠️ 不要帶 --date $today——週末關機，省略會自動讀 trades_all 最新交易日）。
全綠判準：Σcol1==Σpnl、留倉=False、無警告段。writer 結束碼 0=成功/冪等、2=gate 擋下。
被 gate 擋下：什麼都不要改，照常步驟 8，把原因寫進信件正文交 502 人工。絕不寫估算值。
catch-up 自動：writer 處理 CSV 裡所有 session，已寫過的跳過。

【步驟 7】驗算
每列 TRD 驗 qty×price≈abs(amount)，容差 $0.02；BAL/DOI/JRN/ADJ 不驗。
diff>$0.02 或讀不出 → FAILED 草稿結束。CSV 主路徑＝保險；OCR＝主防線，逐列不抽樣。

【步驟 8】判斷有沒有新成交 + 防重複
8.1 & $py C:\TradeReview_Boss\check_new_trades.py "_inbox\$today.csv"（PYTHONIOENCODING=utf-8）
- NEW_TRADES=0 → 不建草稿不寄信，跳步驟 9.5/10/11，log 註記 no new trades, no email。
- NEW_TRADES=N≥1 → 繼續 8.2 與步驟 9。
- 已寄過的記在 logs\reported_trades.txt，只在寄出成功後 --mark；不要手編。
8.2 防重複：搜尋草稿 subject:"[SPY-DayTrade-Autosend] $today"，殘留當日草稿（成功或 FAILED）先刪。

【步驟 9】建草稿並寄出（只有 NEW_TRADES≥1 才做）
create_draft → send_message（draftId）。收件人 juliahsu13@gmail.com；主旨 [SPY-DayTrade-Autosend] $today；嚴格純文字。
9.0 render_statement.py 產純文字對帳表 logs\statement_$today.txt，原樣貼進 STATEMENT TABLE 區塊
（取代截圖；⚠️ 不附截圖或任何圖片附件——base64 太大易抄壞，2026-10-16 決定。OCR 路徑同樣跑，僅 Cash Balance 段）。
【步驟 9.0b】驗證截圖本機留存（2026-10-16 起，CSV 含 TRD 才做）：
擷取 Cash Balance（BAL/TRD/DOI/JRN/TOTAL 每列）＋Order History（全部 FILLED）＋Trade History（全部 fill），無遮擋。
優先從已確認的全螢幕截圖裁切；不動日期範圍欄位。存 logs\TOS_verify_$today.png（分段 _b/_c），
再複製到 C:\TradeReview\verify_shots\（不存在就建；不放 _inbox\）。失敗非致命：log 註記 Screenshot: FAILED。
正文骨架：Date / Generated / Boss PC / Source / Local write / New trades: N（每筆一行）
===== STATEMENT TABLE START/END ===== 與 ===== CSV CONTENT START/END ===== / Validation 行 / -- Auto-generated。
9.2 send_message 寄出（只能 juliahsu13）。成功才 check_new_trades.py --mark；失敗不 mark、草稿留著、log 記 SEND FAILED。

【步驟 9.5】★ 推送到 GitHub
powershell -ExecutionPolicy Bypass -File C:\TradeReview\sync_push.ps1 -Owner boss
只會推 trades_all.xlsx / _inbox\ / sync_state.json / verify_shots\ 四樣（verify_shots\ 由 502 端在 sync_push.ps1 的 boss 清單加入，git pull 後生效）。
（⚠️ 502 註：pull 到新版後實際為 _inbox / verify_shots / BOSS_SKILL_current.md 三樣，trades_all/sync_state 不再推，見檔頭差異 1。）
⚠️ 2026-08-11 起 repo 卡在未完成的 interactive rebase（.git\rebase-merge 存在）：存在時不要跑 sync_push（只會堆孤兒 commit），log 記 PUSH SKIPPED (repo in rebase)，不要自行修復，等使用者手動處理。
push 失敗自動重試 3 次；三次失敗本機 commit 仍在，寫進 log 與信件正文。

【步驟 10】收尾：切回 Charts
非致命。前提：步驟 8～9 已完成（信已寄出或 NEW_TRADES=0 略過）。
Switch-ToCharts          # Verified 啟用 → Click-At 557 33（Charts）→ 游標移 (960,400) 收 hover 選單
→ Capture-Screen dbg_charts → view 驗證 Charts 選中、無選單、無對話框。
主頁籤列 y≈33（build 1993）：Monitor=262、Charts=557（2026-07-13 舊版 y≈52/336 已不適用）。
10.3 沒切成功：從全螢幕截圖重新定位（不要 Capture-Region——常拍到搶前景的 Claude 視窗）→ 重點一次 → 再驗證；第二次失敗放棄，log non-fatal。
10.4 安全限制：只點 Charts 頁籤本身；不點頁內元件/下拉選單項目；不改 symbol/週期；不點下單按鈕。

【步驟 11】寫 log（append logs\daily_run.log）
OK (CSV/OCR) / PARTIAL（gate 擋下或 push 失敗）/ FAIL 格式照舊；每行註明寄信結果
（emailed N new trades to juliahsu13 (<message-id>) / no new trades, no email / SEND FAILED: <原因>）。

【步驟 12】最終回報
表格：Date / Source / CSV / 驗算 / 本機寫入 / 新成交 N 筆（0=未寄信）/ Email / Subject / git push / Charts 還原。

失敗處理通則
debug 截圖 → FAILED 主旨草稿（只建不寄）→ log FAIL → write_sync_state --status failed（不帶 --date）→ 回報。
例外：步驟 10 非致命。不嘗試：computer-use 搜尋、AutoIt/SikuliX、重啟 TOS、輸入密碼、寄 9.2 以外的任何信。

禁止事項
TOS：不點 Buy/Sell/Place Order；不切 Account Statement 以外頁面（例外：步驟 10）；不重啟不輸密碼。
Gmail：唯一可寄=9.2 NEW_TRADES≥1 寄 juliahsu13；無新成交不寄；FAILED 草稿不寄；不寄其他收件人/舊草稿。
C:\TradeReview\ 可寫：_inbox\*.csv（TOS 匯出/OCR）、trades_all.xlsx（writer --trades-only）、
sync_state.json（write_sync_state.py）、verify_shots\TOS_verify_*.png（9.0b）。
絕不可寫：CS交易紀錄.xlsx、CS交易紀錄_dump.txt、offset_state.json、notes\。
其他：不手編 xlsx；不改 .py；writer 必帶 --trades-only；不 git add -A；不改數字。

實戰經驗
✅ 有效（2026-10-16 版）：Activate-TOSWindowVerified（單純 Activate 會被 Claude 搶回）；
Open-ExportDialog + Confirm-ExportSave + Move-Item 改名；Switch-ToCharts；截圖+view 判座標。
❌ 坑：改 days back 打字會 Unexpected error → N≥5 不動；滑鼠點三條線選單項目不觸發 → 用鍵盤；
Save dialog 打字被注音吃掉 → 接受預設檔名再改名；Capture-Region 常拍到 Claude 視窗 → 從全螢幕截圖裁切；
y 差 3-5px 就 miss；hover 選單要 SetCursorPos 移開；PS5.1 不用 C# discard；無 BOM UTF-8 .ps1 當 ANSI（_fix_ps1_bom.py）。
⚠️ OCR 風險：只截得到 Cash Balance；overflow 分段合併；逐列驗算；最易漏 BAL/DOI/JRN。

修訂紀錄
| 日期 | 變更 |
|---|---|
| 2026-10-16 | 舊內容校正：座標改 build 1993 實測（主頁籤 y≈33、Monitor 262、Charts 557）；新增「Claude 搶前景」坑與 Activate-TOSWindowVerified；步驟 4 改 N≥5 不動；步驟 5 改 Open-ExportDialog／Confirm-ExportSave（鍵盤選單、不打檔名再改名）；步驟 9.5 推送清單加 verify_shots\、repo 卡 rebase 時不跑 push；步驟 10 改 Switch-ToCharts |
| 2026-10-16 | 新增 9.0b 驗證截圖本機留存：交易日擷取三區塊存 logs\ 與 verify_shots\；不附信件、不動日期範圍欄位、不放 _inbox\ |
| 2026-10-16 | 美股交易日 06:00：排程改台灣週二～週六 06:00；Claude-TOS-Collector-Launch 05:45 喚醒；新增步驟 0（us_trading_day.py） |
| 2026-10-16 | 改為有新成交才寄：check_new_trades.py 判斷，有才建草稿直接寄 juliahsu13，寄出後 --mark；停用 Apps Script → sunbeamichelle |
| 2026-10-16 | 步驟 9 新增 STATEMENT TABLE 區塊：render_statement.py 純文字對帳表取代截圖 |
| 2026-08-11 | 整併為單一權威檔（兩份規則互打事故）；補 ExecutionPolicy 繞過寫法 |
| 2026-08-10 | 改回 CSV 為主；OCR 抓每一列；新增 6.5 與 9.5；可寫/不可寫清單 |
| 2026-07-17 | 曾改為 OCR-only（已解除） |
| 2026-07-13 | 實測步驟 10 切回 Charts |
| 2026-05-13 | 跑通全流程 |
