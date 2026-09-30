<#
  舊機整台退役 → 新機一次重建全部 Windows 工作排程
  ══════════════════════════════════════════════════════════════════════
  來源機：WINDOWS-AUC64IA\ADMIN（2026-09-24 盤點，30 個任務已全部停用）
  搭配文件：MIGRATION_ALL_TASKS.md（先讀它的 §1 依賴與 §3 順序再跑這支）

  ⚠️ 必須以「系統管理員」執行 —— 寫入工作排程器根目錄需要提權。

      .\rebuild_all_tasks.ps1 -WhatIf                     # 只列出要建什麼，不寫入
      .\rebuild_all_tasks.ps1 -DRoot 'C:\fileserver_D'    # 實際註冊（預設值）
      .\rebuild_all_tasks.ps1 -Only TradeReview_*         # 只建某幾個
      .\rebuild_all_tasks.ps1 -IncludeLegacy              # 連兩個已淘汰的也建

  -DRoot 是舊機 D:\fileserver_D 在新機的新位置。原本就在 C 槽的專案
  （CMoney_DataHub、操作策略SOP_Dashboard、起漲點台股快篩、TWSE_Dashboard、
   元大金儀表板、台指期監控、XlsbUpdater）路徑不變，不受 -DRoot 影響。

  所有設定值都是從舊機 Get-ScheduledTask 實際讀出來的，不是照描述文字猜的
  —— 有好幾個任務的 Description 與實際觸發時間不一致，詳見文件 §7。
#>
[CmdletBinding()]
param(
    [string]$DRoot = 'C:\fileserver_D',
    [string]$PythonW,
    [string[]]$Only,
    [switch]$IncludeLegacy,
    [switch]$WhatIf
)

$ErrorActionPreference = 'Stop'

# ── 提權檢查 ──────────────────────────────────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal] `
            [Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and -not $WhatIf) {
    Write-Host "`n  [X] 需要系統管理員權限（-WhatIf 模式可以不提權）`n" -ForegroundColor Red
    exit 1
}

# ── pythonw 偵測 ──────────────────────────────────────────────────────
if (-not $PythonW) {
    $cand = @()
    $c = Get-Command pythonw.exe -ErrorAction SilentlyContinue
    if ($c) { $cand += $c.Source }
    $cand += Get-ChildItem "$env:LOCALAPPDATA\Programs\Python" -Filter pythonw.exe `
                -Recurse -ErrorAction SilentlyContinue |
             Sort-Object FullName -Descending | Select-Object -ExpandProperty FullName
    $PythonW = $cand | Select-Object -First 1
}
if (-not $PythonW -or -not (Test-Path $PythonW)) {
    throw "找不到 pythonw.exe，請用 -PythonW 明確指定"
}
$PS51 = 'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe'

Write-Host ""
Write-Host "  pythonw : $PythonW"
Write-Host "  D 槽新根: $DRoot"
Write-Host ""

# ── 觸發器小工具 ──────────────────────────────────────────────────────
function T-Logon { New-ScheduledTaskTrigger -AtLogOn }

# 每 N 分鐘無限重複（搭配 MultipleInstances=IgnoreNew 就是零成本看門狗）
function T-Every([int]$Min) {
    New-ScheduledTaskTrigger -Once -At (Get-Date).Date `
        -RepetitionInterval (New-TimeSpan -Minutes $Min)
}

function T-Daily([string[]]$At) {
    $At | ForEach-Object { New-ScheduledTaskTrigger -Daily -At $_ }
}

# 週一~五；$RepMin 有值時掛上重複（New-ScheduledTaskTrigger 的 -Weekly 不吃
# -RepetitionInterval，只能借一個 Once 觸發器的 Repetition 過來）
#
# ⚠️ 持續時間一律用「分鐘」。New-TimeSpan -Hours 吃的是整數，傳 5.25 會被截成
#    5 小時 —— SpreadPagesPublish 的日盤窗口就會短 15 分鐘（13:30 就停而不是 13:45）。
function T-Weekdays([string]$At, [int]$RepMin = 0, [int]$DurMin = 0) {
    $w = New-ScheduledTaskTrigger -Weekly `
            -DaysOfWeek Monday, Tuesday, Wednesday, Thursday, Friday -At $At
    if ($RepMin -gt 0) {
        $tmp = New-ScheduledTaskTrigger -Once -At (Get-Date) `
                  -RepetitionInterval (New-TimeSpan -Minutes $RepMin) `
                  -RepetitionDuration (New-TimeSpan -Minutes $DurMin)
        $w.Repetition = $tmp.Repetition
    }
    $w
}

# 指定時間起，每 $Min 分鐘重複、只跑 $DurMin 分鐘（不是無限）
function T-EveryAtFor([string]$At, [int]$Min, [int]$DurMin) {
    New-ScheduledTaskTrigger -Once -At $At `
        -RepetitionInterval (New-TimeSpan -Minutes $Min) `
        -RepetitionDuration (New-TimeSpan -Minutes $DurMin)
}

# ── 任務清單（值全部來自舊機實際註冊狀態）────────────────────────────
# Exec : py = pythonw.exe / ps = powershell.exe / ps51 = 完整路徑的 PS 5.1 / cmd
# Limit: [timespan] 或 0（0 = 不限時，常駐服務必須是 0）
# Rst  : 失敗重啟次數 / 間隔分鐘
$D = $DRoot
$TASKS = @(
    # ── 第 1 層：資料源。整條管線的源頭，一定要先能跑 ──────────────
    @{ Name='CMoney_Hub_Sync'; Exec='py'
       Script="C:\CMoney_DataHub\sync_cmoney_hub.py"; WorkDir='C:\CMoney_DataHub\'
       Trig={ T-Daily @('06:00','11:00','13:20') }
       Limit=(New-TimeSpan -Hours 1); Rst=0; Swa=$true
       Desc='CMoney 中央 Hub 單一同步點（腳本移至 C:\CMoney_DataHub\）' }

    @{ Name='SHUYE_watch'; Exec='ps'
       Script="C:\XlsbUpdater\shuye_watch.ps1"; WorkDir=''
       Trig={ T-EveryAtFor '09:25' 2 720 }
       Limit=(New-TimeSpan -Hours 72); Rst=0; Swa=$false
       Desc='' }

    @{ Name='ClaudeGuardWatchdog'; Exec='ps'
       Script="C:\XlsbUpdater\claude_guard.ps1"; Extra='-Mode Watchdog'; WorkDir=''
       Trig={ New-ScheduledTaskTrigger -Once -At '11:40' `
                -RepetitionInterval (New-TimeSpan -Minutes 30) }
       Limit=(New-TimeSpan -Hours 72); Rst=0; Swa=$false
       Desc='' }

    # ── 第 2 層：每日更新（讀 Hub，產出各專案的 state/csv）─────────
    @{ Name='TWD2885_Daily_Update'; Exec='py'
       Script="C:\元大金儀表板\auto_update_2885.py"; WorkDir='C:\元大金儀表板'
       Trig={ T-Daily @('06:10') }
       Limit=(New-TimeSpan -Hours 1); Rst=0; Swa=$true
       Desc='每天06:10 從0更新捷徑複製4個CMoney檔→本機、重建2885_state.json' }

    @{ Name='TAIEX_Daily_Update'; Exec='py'
       Script="C:\TWSE_Dashboard\auto_update_taiex.py"; WorkDir='C:\TWSE_Dashboard'
       Trig={ T-Daily @('06:20') }
       Limit=(New-TimeSpan -Minutes 30); Rst=0; Swa=$true
       Desc='每天06:20 重建taiex_state.json（讀元大金已更新的台股DATABASE2026.xlsb + 抓TWSE/TPEx/Yahoo端點）' }

    @{ Name='SOP_Daily_Update'; Exec='py'
       Script="C:\操作策略SOP_Dashboard\run_daily.py"; WorkDir='C:\操作策略SOP_Dashboard'
       Trig={ T-Daily @('06:00','13:30','20:00') }
       Limit=(New-TimeSpan -Hours 1); Rst=0; Swa=$true
       Desc='複製操作策略sop.xlsb→本機、記錄篩選結果到SQLite/CSV、重建dashboard.html' }

    @{ Name='Rising_Daily_Update'; Exec='py'
       Script="C:\起漲點台股快篩\auto_update_rising.py"; WorkDir='C:\起漲點台股快篩'
       Trig={ T-Daily @('11:30','14:00') }
       Limit=(New-TimeSpan -Hours 2); Rst=0; Swa=$true
       Desc='複製CMoney檔→Cmoney_data_tw_potential、跑rising_screener_v5_3全市場分析' }

    @{ Name='FlatBottom_Daily_Update'; Exec='py'
       Script="$D\202607篩股方法重構\auto_update_flatbottom.py"
       WorkDir="$D\202607篩股方法重構"
       Trig={ T-Daily @('11:40','14:10') }
       Limit=(New-TimeSpan -Hours 2); Rst=0; Swa=$true
       Desc='每天11:40/14:10 底部平躺全市場重掃＋重生儀表板(讀CMoney_DataHub，pythonw靜默)' }

    @{ Name='AIBubble_Daily_Update'; Exec='py'
       Script="$D\ai_bubble_monitor\run_daily.py"; Extra='--source blp'
       WorkDir="$D\ai_bubble_monitor"
       Trig={ T-Weekdays '12:00' }
       Limit=(New-TimeSpan -Hours 1); Rst=0; Swa=$true
       Desc='平日12:00 AI泡沫監測 run_daily.py --source blp（需同機 Bloomberg Terminal）' }

    @{ Name='TWD_Step1_Pre'; Exec='py'
       Script="$D\台幣美金匯率預測模型\twd_step1_pre.py"
       WorkDir="$D\台幣美金匯率預測模型"
       Trig={ T-Daily @('07:30') }
       Limit=(New-TimeSpan -Hours 2); Rst=0; Swa=$true
       Desc='每日07:30 USD/TWD 第一段(人工更新前API抓取)' }

    @{ Name='TWD_Step2_Post'; Exec='py'
       Script="$D\台幣美金匯率預測模型\twd_step2_post.py"
       WorkDir="$D\台幣美金匯率預測模型"
       Trig={ T-Daily @('10:50') }
       Limit=(New-TimeSpan -Hours 2); Rst=0; Swa=$true
       Desc='每日10:50 USD/TWD 第二段(匯入→建特徵→預測 latest_full.json)' }

    # ── 第 3 層：快取轉存（餵 Google Drive / MCP）──────────────────
    @{ Name='SHUYE_Cache_Snapshot'; Exec='py'
       Script="$D\DataCacheSync\snapshot_cache.py"; WorkDir="$D\DataCacheSync"
       Trig={ T-Daily @('08:00','09:45','10:00','14:40','23:55') }
       Limit=(New-TimeSpan -Hours 2); Rst=0; Swa=$true
       Desc='把 CMoney Hub + SHUYE 的 xlsb/xlsm 快取值轉存CSV到 DataCacheSync（有變才重寫→Drive只重傳變動檔）' }

    @{ Name='SHUYE_Cache_台股盤後更新表'; Exec='py'
       Script="$D\DataCacheSync\snapshot_cache.py"; Extra='台股盤後更新表'
       WorkDir="$D\DataCacheSync"
       Trig={ T-Daily @('14:45','14:55') }
       Limit=(New-TimeSpan -Minutes 30); Rst=0; Swa=$true
       Desc='只把 台股盤後更新表.xlsx 的快取值轉存CSV。與 SHUYE_Cache_Snapshot 獨立。' }

    # ── 第 4 層：常駐儀表板（TimeLimit 必須 0，Multi 必須 IgnoreNew）─
    @{ Name='SOP_Dashboard_Server'; Exec='py'
       Script="C:\操作策略SOP_Dashboard\操作策略SOP_dashboard_server.py"
       Extra='--serve-only'; WorkDir='C:\操作策略SOP_Dashboard'
       Trig={ @((T-Logon), (T-Every 10)) }; Limit=0; Rst=3; Swa=$true; Desc='' }

    @{ Name='SOP_History_Dashboard'; Exec='py'
       Script="C:\操作策略SOP_Dashboard\history_dashboard_server.py"
       WorkDir='C:\操作策略SOP_Dashboard'
       Trig={ @((T-Logon), (T-Every 10)) }; Limit=0; Rst=3; Swa=$true; Desc='' }

    @{ Name='TAIEX_Dashboard'; Exec='py'
       Script="C:\TWSE_Dashboard\taiex_dashboard.py"; WorkDir='C:\TWSE_Dashboard'
       Trig={ @((T-Logon), (T-Every 10)) }; Limit=0; Rst=3; Swa=$true; Desc='' }

    @{ Name='TWD2885_Dashboard'; Exec='py'
       Script="C:\元大金儀表板\dashboard_2885.py"; WorkDir='C:\元大金儀表板'
       Trig={ @((T-Logon), (T-Every 10)) }; Limit=0; Rst=3; Swa=$true; Desc='' }

    @{ Name='Rising_Dashboard'; Exec='py'
       Script="C:\起漲點台股快篩\rising_screener_v5_3.py"
       Extra='--serve-only --no-browser --port 8888'; WorkDir='C:\起漲點台股快篩'
       Trig={ @((T-Logon), (T-Every 10)) }; Limit=0; Rst=3; Swa=$true; Desc='' }

    @{ Name='FlatBottom_Dashboard'; Exec='py'
       Script="$D\202607篩股方法重構\serve_dashboard.py"
       WorkDir="$D\202607篩股方法重構"
       Trig={ @((T-Logon), (T-Every 10)) }; Limit=0; Rst=3; Swa=$true; Desc='' }

    @{ Name='AIBubble_Dashboard'; Exec='py'
       Script="$D\ai_bubble_monitor\dashboard_server.py"; WorkDir="$D\ai_bubble_monitor"
       Trig={ @((T-Logon), (T-Every 10)) }; Limit=0; Rst=3; Swa=$true; Desc='' }

    @{ Name='TWD_Dashboard'; Exec='py'
       Script="$D\台幣美金匯率預測模型\twd_dashboard.py"
       WorkDir="$D\台幣美金匯率預測模型"
       Trig={ @((T-Logon), (T-Every 10)) }; Limit=0; Rst=3; Swa=$true; Desc='' }

    @{ Name='TradeReview_Dashboard'; Exec='py'
       Script="$D\TradeReview\trade_review_app.py"; WorkDir="$D\TradeReview"
       Trig={ @((T-Logon), (T-Every 10)) }; Limit=0; Rst=3; Swa=$true; Desc='' }

    # ── 第 5 層：發布與同步 ─────────────────────────────────────────
    @{ Name='Dashboards_PagesPublish'; Exec='py'
       Script="$D\PagesPublish\publish_dashboards.py"; WorkDir="$D\PagesPublish"
       Trig={ @((T-Logon), (T-Every 30)) }
       Limit=(New-TimeSpan -Minutes 15); Rst=0; Swa=$true; Desc='' }

    @{ Name='SpreadPagesPublish'; Exec='py'
       Script="C:\台指期監控\pages_publish.py"; WorkDir='C:\台指期監控'
       Trig={ @((T-Weekdays '08:30' 1 315), (T-Weekdays '15:00' 1 840)) }
       Limit=(New-TimeSpan -Minutes 10); Rst=0; Swa=$true
       Desc='台指期價差即時頁發布：日盤08:30-13:45、夜盤15:00-次日05:00 每分鐘；死時段不跑' }

    @{ Name='TradeReview_SyncPull_p502'; Exec='ps'
       Script="$D\TradeReview\sync_pull.ps1"; Extra='-Owner p502 -Quiet'
       WorkDir="$D\TradeReview"
       Trig={ T-Daily @('09:30','11:30') }
       Limit=(New-TimeSpan -Minutes 20); Rst=2; RstMin=5; Swa=$true; Desc='' }

    @{ Name='TradeReview_SyncPush_p502'; Exec='ps'
       Script="$D\TradeReview\sync_push.ps1"; Extra='-Owner p502 -Quiet'
       WorkDir="$D\TradeReview"
       Trig={ T-Daily @('10:00','12:00','18:00') }
       Limit=(New-TimeSpan -Minutes 20); Rst=2; RstMin=5; Swa=$true; Desc='' }

    # ── 第 6 層：看門狗（要最後建，被它監看的任務得先存在）──────────
    @{ Name='Dashboard_Watchdog'; Exec='ps51'
       Script="C:\操作策略SOP_Dashboard\_dashboard_watchdog.ps1"
       WorkDir='C:\操作策略SOP_Dashboard'
       Trig={ @((T-Logon), (T-EveryAtFor '00:01' 15 1440)) }
       Limit=(New-TimeSpan -Minutes 10); Rst=0; Swa=$true
       Desc='每15分鐘檢查儀表板 server，沒在聽就重啟對應工作' }

    @{ Name='SnapshotPanel_KeepAlive'; Exec='ps51'
       Script="$D\DataCacheSync\panel_guard.ps1"; WorkDir="$D\DataCacheSync"
       Trig={ @((T-Daily @('09:00')), (T-Logon)) }
       Limit=(New-TimeSpan -Minutes 5); Rst=0; Swa=$true
       Desc='Snapshot 控制台看門狗：檢查 127.0.0.1:8790，沒在聽就用 pythonw 背景啟動 snapshot_panel.py' }

    # ── 已淘汰：預設不建，要建請加 -IncludeLegacy ───────────────────
    @{ Name='AI泡沫監測'; Legacy=$true; Exec='cmd'
       Raw='/c cd /d C:\202607_AI泡沫\ai_bubble_monitor && python run_daily.py --source blp >> output\run.log 2>&1'
       WorkDir=''
       Trig={ T-Daily @('06:00') }
       Limit=(New-TimeSpan -Hours 72); Rst=0; Swa=$false
       Desc='舊版 AI 泡沫監測，已被 AIBubble_Daily_Update 取代（舊機已停用）' }

    @{ Name='XlsbUpdater_Afternoon_TWPostmkt'; Legacy=$true; Exec='ps'
       Script="C:\XlsbUpdater\files\tw_postmkt\run_guarded.ps1"; WorkDir='C:\XlsbUpdater'
       Trig={ T-Weekdays '14:20' 15 45 }
       Limit=(New-TimeSpan -Hours 2); Rst=0; Swa=$true
       Desc='舊機在搬遷前就已停用' }
)

# ── 註冊 ──────────────────────────────────────────────────────────────
$me = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$done = 0; $skipped = 0; $failed = @()

foreach ($t in $TASKS) {
    if ($t.Legacy -and -not $IncludeLegacy) {
        Write-Host ("  --   {0}  (legacy，用 -IncludeLegacy 才建)" -f $t.Name) -ForegroundColor DarkGray
        $skipped++; continue
    }
    if ($Only -and -not ($Only | Where-Object { $t.Name -like $_ })) { continue }

    # 執行體與參數
    switch ($t.Exec) {
        'py'   { $exe = $PythonW; $arg = '"' + $t.Script + '"' }
        'ps'   { $exe = 'powershell.exe'
                 $arg = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $t.Script + '"' }
        'ps51' { $exe = $PS51
                 $arg = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + $t.Script + '"' }
        'cmd'  { $exe = 'cmd'; $arg = $t.Raw }
    }
    if ($t.Extra) { $arg += ' ' + $t.Extra }

    # 腳本不存在就先擋下來，不要建一個註定失敗的任務
    if ($t.Script -and -not (Test-Path $t.Script)) {
        Write-Host ("  [!]  {0}  找不到腳本：{1}" -f $t.Name, $t.Script) -ForegroundColor Yellow
        $failed += ($t.Name + ' (腳本不存在)'); continue
    }

    if ($WhatIf) {
        Write-Host ("  WHATIF {0}" -f $t.Name) -ForegroundColor Cyan
        Write-Host ("         {0} {1}" -f $exe, $arg)
        continue
    }

    try {
        $actParams = @{ Execute = $exe; Argument = $arg }
        if ($t.WorkDir) { $actParams.WorkingDirectory = $t.WorkDir }
        $A = New-ScheduledTaskAction @actParams

        $T = @(& $t.Trig)

        $P = New-ScheduledTaskPrincipal -UserId $me -LogonType Interactive -RunLevel Limited

        $setParams = @{
            MultipleInstances        = 'IgnoreNew'
            AllowStartIfOnBatteries  = $true
            DontStopIfGoingOnBatteries = $true
        }
        if ($t.Limit -eq 0) { $setParams.ExecutionTimeLimit = 0 }
        else                { $setParams.ExecutionTimeLimit = $t.Limit }
        if ($t.Swa)  { $setParams.StartWhenAvailable = $true }
        if ($t.Rst -gt 0) {
            $setParams.RestartCount = $t.Rst
            $m = if ($t.RstMin) { $t.RstMin } else { 1 }
            $setParams.RestartInterval = (New-TimeSpan -Minutes $m)
        }
        $S = New-ScheduledTaskSettingsSet @setParams

        $reg = @{ TaskName = $t.Name; Action = $A; Trigger = $T
                  Principal = $P; Settings = $S; Force = $true }
        if ($t.Desc) { $reg.Description = $t.Desc }
        Register-ScheduledTask @reg | Out-Null

        if ($t.Legacy) { Disable-ScheduledTask -TaskName $t.Name | Out-Null }

        Write-Host ("  [OK] {0}" -f $t.Name) -ForegroundColor Green
        $done++
    }
    catch {
        Write-Host ("  [X]  {0} : {1}" -f $t.Name, $_.Exception.Message) -ForegroundColor Red
        $failed += $t.Name
    }
}

# ── 回報 ──────────────────────────────────────────────────────────────
Write-Host ""
Write-Host ("  已建立 {0}、跳過 {1}、失敗 {2}" -f $done, $skipped, $failed.Count)
if ($failed.Count) { $failed | ForEach-Object { Write-Host "      失敗：$_" -ForegroundColor Red } }
Write-Host ""
Write-Host "  接著做（見 MIGRATION_ALL_TASKS.md §10）：" -ForegroundColor Cyan
Write-Host "      1. 修 _dashboard_watchdog.ps1 的 SPY_ChartWeb 死條目與 TWD 埠號"
Write-Host "      2. 先手動跑 CMoney_Hub_Sync，確認能讀到 \\SHUYE07 的 0更新捷徑"
Write-Host "      3. Get-ScheduledTask -TaskPath '\' | ? {-not `$_.Settings.Enabled}  應該只剩 legacy"
Write-Host ""
