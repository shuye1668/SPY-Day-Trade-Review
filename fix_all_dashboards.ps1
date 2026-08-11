<#
  一次修好本機所有 dashboard 的「開機常駐 + 崩潰自動復原」
  ══════════════════════════════════════════════════════════════
  ⚠️ 需以「系統管理員」執行（註冊工作排程器根目錄要提權）

      powershell -ExecutionPolicy Bypass -File fix_all_dashboards.ps1
      powershell -ExecutionPolicy Bypass -File fix_all_dashboards.ps1 -WhatIf   # 只看要改什麼

  ── 為什麼要這支 ──────────────────────────────────────────────
  2026-08-11 實測診斷：
    * 8 支 dashboard 用 python.exe 全部正常啟動（AIBubble 除外，它要 Bloomberg）
    * 改用 pythonw.exe 也全部正常啟動 —— 所以程式碼沒有 bug
    * 所有 *_Daily_Update 資料產生任務都 rc=0 —— 資料是新的
    * 但所有 *_Dashboard 任務的 LastRun 都停在 08-07 08:26

  真正的原因是排程設定，不是程式：
    觸發器只有 AtLogOn、沒有任何重複觸發 → 行程一旦結束就永遠不會再起來，
    要等下次登入。機器沒登出入，就一路 DOWN 四天。

  ── 修法 ──────────────────────────────────────────────────────
  加一個每 N 分鐘的重複觸發當看門狗，並把 MultipleInstances 設為 IgnoreNew：
    還活著 → 新執行個體被忽略（不會開出第二個去搶同一個埠、寫壞同一份檔）
    已死掉 → 直接被拉起來
  這是零成本的健康檢查，不需要另外寫 watchdog 程式。
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [int]$WatchdogMinutes = 10,
    [string[]]$Only,                 # 只處理指定任務名（可省略＝全部）
    [switch]$IncludeTradeReview      # TradeReview_Dashboard 已單獨設過，預設跳過
)

$ErrorActionPreference = 'Stop'

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
# -WhatIf 只是預覽、不寫入任何東西，不該要求提權
if (-not $isAdmin -and -not $WhatIfPreference) {
    Write-Host "`n  [X] 需要系統管理員權限" -ForegroundColor Red
    Write-Host "      請改用「以系統管理員身分執行」的 PowerShell 再跑一次。"
    Write-Host "      （想先看會改哪些任務，可加 -WhatIf，不需要提權）`n"
    exit 1
}

$tasks = Get-ScheduledTask | Where-Object {
    $_.TaskName -match '_Dashboard$|_Dashboard_Server$|_Server$' -and
    $_.Actions.Execute -match 'python'
}
if (-not $IncludeTradeReview) {
    $tasks = $tasks | Where-Object { $_.TaskName -ne 'TradeReview_Dashboard' }
}
if ($Only) { $tasks = $tasks | Where-Object { $Only -contains $_.TaskName } }

if (-not $tasks) { Write-Host "  找不到符合的任務"; exit 0 }

Write-Host "`n  將為 $($tasks.Count) 個任務加上看門狗（每 $WatchdogMinutes 分鐘）`n" -ForegroundColor Cyan

$me = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
$ok = 0; $fail = 0

foreach ($t in $tasks) {
    $name = $t.TaskName
    $a = $t.Actions[0]
    # pythonw 才是真正的「silent 常駐」；有些任務可能被設成 python.exe（會冒出黑視窗）
    $exec = $a.Execute
    if ($exec -match '\\python\.exe$') {
        $pyw = $exec -replace '\\python\.exe$', '\pythonw.exe'
        if (Test-Path $pyw) { $exec = $pyw }
    }

    Write-Host ("  {0,-26}" -f $name) -NoNewline
    if (-not $PSCmdlet.ShouldProcess($name, "加看門狗重複觸發")) { Write-Host " (WhatIf)"; continue }

    try {
        $act = New-ScheduledTaskAction -Execute $exec -Argument $a.Arguments `
                                       -WorkingDirectory $a.WorkingDirectory
        $t1 = New-ScheduledTaskTrigger -AtLogOn
        $t2 = New-ScheduledTaskTrigger -Once -At (Get-Date).Date `
                -RepetitionInterval (New-TimeSpan -Minutes $WatchdogMinutes)
        $pr = New-ScheduledTaskPrincipal -UserId $me -LogonType Interactive -RunLevel Limited
        $st = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -ExecutionTimeLimit 0 `
                -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
                -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)

        Register-ScheduledTask -TaskName $name -Action $act -Trigger @($t1, $t2) `
                               -Principal $pr -Settings $st -Force | Out-Null
        Start-ScheduledTask -TaskName $name
        Write-Host " OK" -ForegroundColor Green
        $ok++
    }
    catch {
        Write-Host (" 失敗：" + $_.Exception.Message) -ForegroundColor Red
        $fail++
    }
}

Write-Host "`n  完成：成功 $ok，失敗 $fail" -ForegroundColor $(if ($fail) { 'Yellow' } else { 'Green' })
Write-Host "  等 15 秒後檢查各埠是否聽起來：`n"
Write-Host '    8600,8601,8602,8801,8802,8808,8888,5500 | ForEach-Object {'
Write-Host '      $c = Get-NetTCPConnection -LocalPort $_ -State Listen -ErrorAction SilentlyContinue'
Write-Host '      "{0}  {1}" -f $_, $(if($c){"LISTEN"}else{"DOWN"}) }'
Write-Host ""
