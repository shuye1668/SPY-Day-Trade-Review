# 舊機整台退役：30 個 Windows 工作排程重建總表

> **來源機**：`WINDOWS-AUC64IA\ADMIN`（Windows 11，Python 3.12.9）
> **盤點日**：2026-09-24。所有數值來自 `Get-ScheduledTask` 實際註冊狀態，**不是照描述文字抄的**
> —— 有 5 處描述與實際不符，見 §7。
> **舊機現況**：30 個任務**全部已 `Enabled=False`**。還原指令見 §9。
> **重建工具**：`rebuild_all_tasks.ps1`（同資料夾，已通過語法檢查與路徑 dry-run）
> **TradeReview 三個任務的深入說明**：另見 `MIGRATION_WINDOWS_TASKS.md`

---

## §0 三步走

```powershell
# 1) 先確認相依（§1）都通了，尤其是網路磁碟
Test-Path '\\SHUYE07\fileserver\書業(新)\Julia\0更新捷徑'

# 2) 乾跑，確認 30 個任務的腳本路徑都存在（不需提權）
.\rebuild_all_tasks.ps1 -WhatIf -DRoot 'C:\fileserver_D'

# 3) 管理員 PowerShell 實際註冊
.\rebuild_all_tasks.ps1 -DRoot 'C:\fileserver_D'
```

`-DRoot` 是舊機 `D:\fileserver_D` 在新機的新位置。原本就在 C 槽的 7 個專案路徑不變。
腳本會自動偵測 pythonw、自動用目前登入使用者當 Principal，並在腳本不存在時擋下該任務而不是硬建。

**跑完一定要做 §6 的看門狗修正**（舊機的 `_dashboard_watchdog.ps1` 有一個指向已不存在任務的死條目）。

---

## §1 相依：這條管線的源頭是一個網路磁碟

```
\\SHUYE07\fileserver\書業(新)\Julia\0更新捷徑        ← 唯一資料源（網路share）
        │  台股DATABASE2026.xlsb / 季年報DATABASE.xlsb
        │  daliy主力券商database.xlsb / XQ.xlsb
        ↓  CMoney_Hub_Sync（06:00 / 11:00 / 13:20）
C:\CMoney_DataHub\                                   ← 中央 Hub，下游一律讀這裡
        ├─→ TWD2885_Daily_Update      → C:\元大金儀表板\2885_state.json
        ├─→ TAIEX_Daily_Update        → C:\TWSE_Dashboard\taiex_state.json
        ├─→ Rising_Daily_Update       → C:\起漲點台股快篩\
        ├─→ FlatBottom_Daily_Update   → {D}\202607篩股方法重構\
        ├─→ SOP_Daily_Update          → C:\操作策略SOP_Dashboard\（讀 操作策略sop.xlsb）
        └─→ SHUYE_Cache_Snapshot      → {D}\DataCacheSync\*.csv → Google Drive → MCP/Claude
```

**`CMoney_Hub_Sync` 沒跑成，下游七個專案全部吃到舊資料，而且不會報錯。**
新機第一件事就是手動跑它一次，確認 `C:\CMoney_DataHub` 出現那 4 個 xlsb。

### 其他外部相依

| 相依 | 誰需要 | 沒有會怎樣 |
|---|---|---|
| `\\SHUYE07\fileserver\...\0更新捷徑`（網路share，舊機可達） | `CMoney_Hub_Sync` | 整條管線斷源 |
| **Bloomberg Terminal 裝在同一台** | `AIBubble_Daily_Update`（`--source blp`） | 該任務直接失敗 |
| Google Drive 桌面同步（掛在 `{D}\DataCacheSync`） | `SHUYE_Cache_*` | CSV 出得來但上不去雲端，Claude/MCP 讀不到 |
| CMoney 用戶端（產生 xlsb 的來源端） | 上游，不在本機 | — |
| GitHub 憑證（Windows 認證管理員，要能免密碼 push） | `TradeReview_SyncPush/Pull`、`Dashboards_PagesPublish`、`SpreadPagesPublish` | 排程無人看管時跳密碼視窗＝當天靜默失效 |
| `操作策略sop.xlsb`（人工維護的來源檔） | `SOP_Daily_Update` | 篩選結果不更新 |

### 涉及的 GitHub repo

| Repo | 誰推 | 性質 |
|---|---|---|
| `shuye1668/SPY-Day-Trade-Review` | `TradeReview_SyncPush_p502` | private，雙機帳本同步 |
| `shuye1668/market-dashboards` | `Dashboards_PagesPublish` | Pages，儀表板總表快照 |
| `shuye1668/taifex-spread-monitor` | `SpreadPagesPublish` | 台指期價差 |
| `shuye1668/spy-intraday-chart` | （TradeReview 側） | Pages，**`history_minute.xlsx` 唯一的異地備份** |

---

## §2 資料夾搬遷對照

| 舊路徑 | 檔數 | git | 新路徑 |
|---|---|---|---|
| `D:\fileserver_D\TradeReview` | — | ✔ private repo | `{DRoot}\TradeReview` |
| `D:\fileserver_D\ai_bubble_monitor` | 76 | — | `{DRoot}\ai_bubble_monitor` |
| `D:\fileserver_D\202607篩股方法重構` | 532 | — | `{DRoot}\202607篩股方法重構` |
| `D:\fileserver_D\台幣美金匯率預測模型` | 1,464 | — | `{DRoot}\台幣美金匯率預測模型` |
| `D:\fileserver_D\DataCacheSync` | 436 | — | `{DRoot}\DataCacheSync` |
| `D:\fileserver_D\PagesPublish` | 213 | ✔（remote 由腳本執行時 add） | `{DRoot}\PagesPublish` |
| `C:\CMoney_DataHub` | 11 | — | **路徑不變** |
| `C:\操作策略SOP_Dashboard` | 505 | — | **路徑不變** |
| `C:\起漲點台股快篩` | 3,917 | — | **路徑不變** |
| `C:\TWSE_Dashboard` | 79 | — | **路徑不變** |
| `C:\元大金儀表板` | 15 | — | **路徑不變** |
| `C:\台指期監控` | 738 | ✔ | **路徑不變** |
| `C:\XlsbUpdater` | 3,833 | — | **路徑不變** |
| `C:\202607_AI泡沫` | 103 | — | 已淘汰，可不搬 |

> 只有 2 個資料夾有 git、其餘 12 個**全靠檔案複製**。漏掉就是永久遺失。
> 特別注意 `{D}\TradeReview\history_minute.xlsx`（4.9 MB，刻意不進 git）——
> yfinance 只回得了最近約 30 天，救不回來；唯一異地備份是 `spy-intraday-chart` Pages。

---

## §3 重建順序（相依決定的，不要跳）

1. **環境**：Python 3.12+、各專案的套件、git 憑證、Google Drive 桌面版、Bloomberg Terminal。
2. **網路磁碟**：確認 `\\SHUYE07\fileserver\書業(新)\Julia\0更新捷徑` 可達（可能要先掛認證）。
3. **手動跑 `CMoney_Hub_Sync`**，確認 Hub 那 4 個 xlsb 到位，看 `C:\CMoney_DataHub\logs\hub_sync.log`。
4. **`.\rebuild_all_tasks.ps1`** 一次註冊 28 個現役任務（2 個 legacy 預設不建）。
5. **修 `_dashboard_watchdog.ps1`**（§6），否則看門狗會對錯的埠與不存在的任務動作。
6. **§10 驗證**。

---

## §4 30 個任務總表

分層看：**資料源 → 每日更新 → 快取轉存 → 常駐儀表板 → 發布同步 → 看門狗**。
`TimeLimit=0` 與 `Multi=IgnoreNew` 是常駐服務的兩個關鍵設定，所有儀表板都一樣。

### 第 1 層 資料源（3）

| 任務 | 觸發 | 執行 | TimeLimit |
|---|---|---|---|
| `CMoney_Hub_Sync` | 每日 06:00 / 11:00 / 13:20 | `pythonw C:\CMoney_DataHub\sync_cmoney_hub.py` | 1h |
| `SHUYE_watch` | 09:25 起每 2 分、持續 12h | `ps C:\XlsbUpdater\shuye_watch.ps1` | 72h |
| `ClaudeGuardWatchdog` | 11:40 起每 30 分（無限） | `ps C:\XlsbUpdater\claude_guard.ps1 -Mode Watchdog` | 72h |

### 第 2 層 每日更新（8）

| 任務 | 觸發 | 執行 | TimeLimit |
|---|---|---|---|
| `TWD2885_Daily_Update` | 每日 06:10 | `pythonw C:\元大金儀表板\auto_update_2885.py` | 1h |
| `TAIEX_Daily_Update` | 每日 06:20 | `pythonw C:\TWSE_Dashboard\auto_update_taiex.py` | 30m |
| `SOP_Daily_Update` | 每日 06:00 / 13:30 / 20:00 | `pythonw C:\操作策略SOP_Dashboard\run_daily.py` | 1h |
| `Rising_Daily_Update` | 每日 11:30 / 14:00 | `pythonw C:\起漲點台股快篩\auto_update_rising.py` | 2h |
| `FlatBottom_Daily_Update` | 每日 11:40 / 14:10 | `pythonw {D}\202607篩股方法重構\auto_update_flatbottom.py` | 2h |
| `AIBubble_Daily_Update` | **週一~五** 12:00 | `pythonw {D}\ai_bubble_monitor\run_daily.py --source blp` | 1h |
| `TWD_Step1_Pre` | 每日 07:30 | `pythonw {D}\台幣美金匯率預測模型\twd_step1_pre.py` | 2h |
| `TWD_Step2_Post` | 每日 10:50 | `pythonw {D}\台幣美金匯率預測模型\twd_step2_post.py` | 2h |

> `TWD_Step1` / `Step2` 中間夾著人工更新：07:30 抓 API → 人工匯入 → 10:50 建特徵並預測。

### 第 3 層 快取轉存（2）

| 任務 | 觸發 | 執行 | TimeLimit |
|---|---|---|---|
| `SHUYE_Cache_Snapshot` | 每日 08:00 / 09:45 / 10:00 / 14:40 / 23:55 | `pythonw {D}\DataCacheSync\snapshot_cache.py` | 2h |
| `SHUYE_Cache_台股盤後更新表` | 每日 14:45 / 14:55 | 同上 + 參數 `台股盤後更新表` | 30m |

> 兩個刻意獨立。都是「有變才重寫」→ Google Drive 只重傳變動檔。

### 第 4 層 常駐儀表板（9）

全部：`LogonTrigger` + 每 10 分重複（無限）、`TimeLimit=0`、`Multi=IgnoreNew`、失敗重啟 3 次／間隔 1 分。
「每 10 分重複 + IgnoreNew」＝零成本看門狗：活著就忽略新實例，死了就拉起來。

| 任務 | 埠 | 執行 |
|---|---|---|
| `SOP_Dashboard_Server` | 8801 | `pythonw C:\操作策略SOP_Dashboard\操作策略SOP_dashboard_server.py --serve-only` |
| `SOP_History_Dashboard` | 8802 | `pythonw C:\操作策略SOP_Dashboard\history_dashboard_server.py` |
| `TAIEX_Dashboard` | 8600 | `pythonw C:\TWSE_Dashboard\taiex_dashboard.py` |
| `TWD2885_Dashboard` | 8602 | `pythonw C:\元大金儀表板\dashboard_2885.py` |
| `Rising_Dashboard` | 8888 | `pythonw C:\起漲點台股快篩\rising_screener_v5_3.py --serve-only --no-browser --port 8888` |
| `FlatBottom_Dashboard` | 8808 | `pythonw {D}\202607篩股方法重構\serve_dashboard.py` |
| `AIBubble_Dashboard` | 8809 | `pythonw {D}\ai_bubble_monitor\dashboard_server.py` |
| `TWD_Dashboard` | **8606**（見 §7） | `pythonw {D}\台幣美金匯率預測模型\twd_dashboard.py` |
| `TradeReview_Dashboard` | 5500 | `pythonw {D}\TradeReview\trade_review_app.py` |

### 第 5 層 發布與同步（4）

| 任務 | 觸發 | 執行 | TimeLimit |
|---|---|---|---|
| `Dashboards_PagesPublish` | Logon + 每 30 分 | `pythonw {D}\PagesPublish\publish_dashboards.py` | 15m |
| `SpreadPagesPublish` | 週一~五 08:30 起每 1 分共 5h15m；15:00 起每 1 分共 14h | `pythonw C:\台指期監控\pages_publish.py` | 10m |
| `TradeReview_SyncPull_p502` | 每日 09:30 / 11:30 | `ps {D}\TradeReview\sync_pull.ps1 -Owner p502 -Quiet` | 20m |
| `TradeReview_SyncPush_p502` | 每日 10:00 / 12:00 / 18:00 | `ps {D}\TradeReview\sync_push.ps1 -Owner p502 -Quiet` | 20m |

> `SpreadPagesPublish` 的兩段刻意避開死時段（05:00-08:30、13:45-15:00 期貨不跳動）。
> TradeReview 兩支的順序刻意是先拉再推，且 `sync_push` **絕不可帶 `-Force`**（回歸閘門）。

### 第 6 層 看門狗（2）

| 任務 | 觸發 | 執行 | TimeLimit |
|---|---|---|---|
| `Dashboard_Watchdog` | Logon + 00:01 起每 15 分共 1 天 | `ps51 C:\操作策略SOP_Dashboard\_dashboard_watchdog.ps1` | 10m |
| `SnapshotPanel_KeepAlive` | 每日 09:00 + Logon | `ps51 {D}\DataCacheSync\panel_guard.ps1` | 5m |

### 已淘汰（2）：預設不重建

| 任務 | 為什麼不建 |
|---|---|
| `AI泡沫監測` | 舊版（`cmd /c cd /d C:\202607_AI泡沫\...`），已被 `AIBubble_Daily_Update` 取代，舊機早已停用 |
| `XlsbUpdater_Afternoon_TWPostmkt` | 搬遷前就已停用 |

要建就加 `-IncludeLegacy`，腳本會建完立刻 `Disable`（維持舊機原狀）。

### 我沒有動、也不用搬的

`Adobe Acrobat Update Task`、`Bloomberg Updater`（`C:\blp\Wintrv\clientratermgr.exe`）、
`NVIDIA App SelfUpdate`、`OneDrive` ×3，以及 Microsoft / Google / ASUS / SoftLanding
目錄下的系統任務 —— 都是廠商與系統自己的更新，新機裝好對應軟體就會自己有。

---

## §5 執行體與參數的三種樣板

腳本裡用 `Exec` 欄位區分，手動建時照抄：

| 型別 | Execute | Argument |
|---|---|---|
| `py` | `<pythonw.exe 絕對路徑>` | `"<script.py>" [額外參數]` |
| `ps` | `powershell.exe` | `-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "<script.ps1>" [額外參數]` |
| `ps51` | `C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe` | 同上 |
| `cmd` | `cmd` | `/c cd /d <dir> && python ... >> log 2>&1` |

共通：`Principal` = 目前登入使用者、`LogonType Interactive`、`RunLevel Limited`、
`MultipleInstances IgnoreNew`、`StartWhenAvailable`（`SHUYE_watch`、`ClaudeGuardWatchdog`、
`AI泡沫監測` 除外）、電池不限制。

**不要沿用舊機的 SID** `S-1-5-21-2060264650-2152385818-2097199698-1000`。

---

## §6 埠號對照與看門狗（重建後必須修）

| 埠 | 服務 | 誰看門 |
|---|---|---|
| 5500 | TradeReview 複盤 app | 自己的 10 分觸發器（watchdog 那條是死的，見下） |
| 8600 | TWSE 大盤 | `Dashboard_Watchdog` |
| 8602 | 元大金 2885 | `Dashboard_Watchdog` |
| **8606** | 美台匯率（watchdog 卻在探 8601） | 需修 |
| 8790 | Snapshot 控制台 | `SnapshotPanel_KeepAlive` |
| 8801 | SOP 完整版 | `Dashboard_Watchdog` |
| 8802 | SOP 歷史比較 | `Dashboard_Watchdog` |
| 8808 | 底部平躺 | `Dashboard_Watchdog` |
| 8809 | AI 泡沫監測 | 只有自己的 10 分觸發器 |
| 8888 | 起漲點快篩 | `Dashboard_Watchdog` |

`C:\操作策略SOP_Dashboard\_dashboard_watchdog.ps1` 裡的對照表有兩處要改：

```powershell
@{Port=5500; Task='SPY_ChartWeb';  Name='SPY當沖繪圖'}   # ← 這個任務名不存在了
@{Port=8601; Task='TWD_Dashboard'; Name='美台匯率'}      # ← 埠號對不上
```

- `SPY_ChartWeb` 是舊任務名，現在叫 `TradeReview_Dashboard`。腳本用
  `Get-ScheduledTask -ErrorAction SilentlyContinue`，所以它**靜默失效**、什麼都不做。
  5500 目前實際上沒有 watchdog 覆蓋（靠 app 自己的 10 分觸發器撐著，所以沒人發現）。
- `twd_dashboard.py` 是 `PORT = int(os.environ.get("TWD_DASH_PORT", "8606"))`。
  排程沒有設這個環境變數，所以它聽 **8606**，watchdog 卻在探 8601 ——
  永遠探不到、每 15 分鐘對一個還活著的任務下 `Start-ScheduledTask`（`IgnoreNew`
  讓它無害，但等於 TWD 也沒有實質看門）。二選一：把 watchdog 改成 8606，
  或在任務裡設 `TWD_DASH_PORT=8601`。

`AIBubble_Dashboard`（8809）從來就不在 watchdog 表裡，順手補上會更一致。

---

## §7 描述與實際不符的 5 處（重建一律以實際值為準）

| 任務 | 描述寫的 | 實際註冊的 |
|---|---|---|
| `Rising_Daily_Update` | 每天 09:40 / 14:00 | **11:30** / 14:00 |
| `SHUYE_Cache_台股盤後更新表` | 每天 15:00 與 15:30 | **14:45 / 14:55** |
| `SOP_Daily_Update` | 每天 09:00 | **06:00 / 13:30 / 20:00** |
| `TradeReview_SyncPush_p502` | `CLAUDE.md` 表：10:00、12:00 | 10:00 / 12:00 / **18:00** |
| `Dashboard_Watchdog` | 「五個儀表板 8801/8802/8602/8600/8888」 | 腳本裡其實有 **8 條**（多 8808 / 5500 / 8601） |

`rebuild_all_tasks.ps1` 用的全是右欄。描述文字我照原樣搬過去（含不符的部分），
是為了讓你在新機比對時看得出來哪些是歷史遺留；要順手修正描述就直接改腳本的 `Desc`。

---

## §8 舊機環境基準

| 項目 | 值 |
|---|---|
| Python | 3.12.9，`C:\Users\ADMIN\AppData\Local\Programs\Python\Python312\pythonw.exe` |
| 主要套件 | Flask 3.1.3、pandas 3.0.2、numpy 2.4.4、openpyxl 3.1.5、requests 2.34.2、pytz 2026.1、yfinance 1.3.0（另有 `pyxlsb` 供 SOP 讀 xlsb） |
| Git | 2.54.0.windows.1 |
| 使用者 | `WINDOWS-AUC64IA\ADMIN`（SID 尾碼 -1000） |

**兩條動到 `.ps1` 才會遇上的陷阱**（各專案通用）：

1. PowerShell 5.1 讀沒有 BOM 的 UTF-8 `.ps1` 會當 ANSI，中文註解直接炸語法。
   TradeReview 底下改過任何 `.ps1` 就跑 `python _fix_ps1_bom.py`。
2. 排程一律用 `powershell -ExecutionPolicy Bypass -File`，不要假設新機的 ExecutionPolicy 允許直接跑 `.ps1`。

---

## §9 舊機還原（萬一新機出狀況）

30 個任務目前全是 `Enabled=False`，定義都還在，一行就能全開：

```powershell
Get-ScheduledTask -TaskPath '\' | Where-Object { $_.TaskName -match '^(AIBubble|FlatBottom|Rising|SOP_|TAIEX|TWD|TradeReview|SHUYE|Snapshot|Dashboard|CMoney|Spread|ClaudeGuard)' } | Enable-ScheduledTask
```

只要開其中一個專案就把名字換掉即可。注意 `AI泡沫監測` 與
`XlsbUpdater_Afternoon_TWPostmkt` **在搬遷前就已經是停用狀態**，不要一併開起來。

⚠️ **新機的 `TradeReview_SyncPush_p502` 一旦開始跑，舊機的就絕對不能再開** ——
兩台同時以 `p502` 身分推同一個 private repo，xlsx 是 binary、git 無法 merge，會互相覆蓋。

---

## §10 驗證清單

```powershell
# 1) 28 個現役任務都在且啟用（legacy 那 2 個若建了應為 Disabled）
Get-ScheduledTask -TaskPath '\' | Where-Object { $_.Settings.Enabled } |
  Select-Object TaskName, State | Sort-Object TaskName | Format-Table -AutoSize

# 2) 常駐儀表板的兩個關鍵設定
'SOP_Dashboard_Server','SOP_History_Dashboard','TAIEX_Dashboard','TWD2885_Dashboard',
'Rising_Dashboard','FlatBottom_Dashboard','AIBubble_Dashboard','TWD_Dashboard',
'TradeReview_Dashboard' | ForEach-Object {
  $t = Get-ScheduledTask $_
  '{0,-24} TimeLimit={1,-6} Multi={2}' -f $_, $t.Settings.ExecutionTimeLimit, $t.Settings.MultipleInstances
}
# TimeLimit 必須是 PT0S 或空白、Multi 必須 IgnoreNew

# 3) 資料源真的通了
Start-ScheduledTask CMoney_Hub_Sync; Start-Sleep 60
Get-Content C:\CMoney_DataHub\logs\hub_sync.log -Tail 5
Get-ChildItem C:\CMoney_DataHub\*.xlsb | Select-Object Name, Length, LastWriteTime

# 4) 十個埠都有人聽（逐一啟動後）
Get-ScheduledTask -TaskPath '\' | Where-Object { $_.TaskName -like '*Dashboard*' } | Start-ScheduledTask
Start-Sleep 20
5500,8600,8602,8606,8790,8801,8802,8808,8809,8888 | ForEach-Object {
  '{0} {1}' -f $_, $(if (Get-NetTCPConnection -LocalPort $_ -State Listen -EA SilentlyContinue) {'OK'} else {'--'})
}

# 5) 同步／發布兩支看 log，不要只看回傳碼
$D = 'C:\fileserver_D'      # 換成你的 -DRoot
Get-Content "$D\TradeReview\_logs\sync_pull.log" -Tail 3   # 出現 "SKIP dirty" 就是失敗
Get-Content "$D\TradeReview\_logs\sync_push.log" -Tail 3   # "nothing-to-push" 是正常
```

第 4 步的 `8606` 若沒人聽，就是 §6 的埠號問題還沒處理。
