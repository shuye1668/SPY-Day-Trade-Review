# boss_screen_guard.ps1 — Boss PC 畫面守門（2026-10-22）
#
# 使用者要求：Boss 採集不管有沒有完成都要把自己（Claude Desktop）縮小；任務卡住也不要擋在螢幕上，
# 要還原螢幕原本的使用狀況。SKILL 正常跑完時由步驟 10.5 Restore-ScreenState 處理；這支是保險，
# 由 Windows 排程 Boss_Screen_Guard 週二～週六 06:10～09:30 每 10 分鐘以隱藏方式執行（boss_run_hidden.vbs）。
#
# 判斷（今天的 run_flag_<日期>.json，由 boss_screen_state.ps1 寫）：
#   - 使用者 3 分鐘內有動鍵盤/滑鼠 → 什麼都不做（人在用，不打擾）
#   - started 後還沒 finished：未滿 30 分鐘＝還在跑 → 不動；滿 30 分鐘＝卡住 →
#       Restore-ScreenState（TOS 原樣、Z 順序、前景、游標）＋縮小 Claude，記 guard_restored
#   - 已 finished → 補縮一次 Claude（防結束後又跳出來），每次 run 只做一次（guard_final）
#   - 沒有旗標（休市日步驟 0 就結束、或步驟 1.5 前就卡住）且已過 06:40 → 縮小 Claude（05:45 開機時打開的）
# 只做視窗管理：不點擊、不打字、不關視窗、不碰 TOS 內容。
#
# 測試：-DryRun（只印判斷，不動視窗不寫檔）；-LogDir／-Now 可指定測試用資料夾與時間。

param([switch]$DryRun, [string]$LogDir = 'C:\TradeReview_Boss\logs', [datetime]$Now = (Get-Date),
      [int]$StuckMin = 30, [int]$UserIdleSec = 180)

$ErrorActionPreference = 'Continue'
$script:BossLogDir = $LogDir
. (Join-Path $PSScriptRoot 'boss_screen_state.ps1')

$day = $Now.ToString('yyyy-MM-dd')
$log = Join-Path $LogDir 'screen_guard.log'
function GLog($m) {
    $line = "[{0}] {1}{2}" -f $Now.ToString('yyyy-MM-dd HH:mm:ss'), $(if ($DryRun) { 'DRYRUN ' } else { '' }), $m
    Write-Host $line
    if (-not $DryRun) { try { New-Item -ItemType Directory -Force -Path $LogDir | Out-Null; Add-Content -LiteralPath $log -Value $line -Encoding UTF8 } catch {} }
}
function Mark($field) { if (-not $DryRun) { Set-RunFlag $field ($Now.ToString('s')) } }

$idle = Get-IdleSeconds
if ($idle -lt $UserIdleSec) { GLog "user active (idle ${idle}s) → skip"; exit 0 }

$flag = Read-RunFlag $day
$ss = Join-Path $LogDir "screen_state_$day.json"
$started  = if ($flag -and $flag.started)  { [datetime]$flag.started }  else { $null }
$finished = if ($flag -and $flag.finished) { [datetime]$flag.finished } else { $null }

if ($started -and (-not $finished -or $finished -lt $started)) {
    $age = [math]::Round(($Now - $started).TotalMinutes, 1)
    if ($age -lt $StuckMin) { GLog "run in progress (${age} min since started) → wait"; exit 0 }
    if ($flag.guard_restored -and [datetime]$flag.guard_restored -ge $started) {
        # 已還原過；只確保 Claude 還是縮著
        $c = if ($DryRun) { Minimize-ClaudeWindows -WhatIf } else { Minimize-ClaudeWindows }
        GLog "stuck run already restored; claude minimized $c"; exit 0
    }
    if ($DryRun) { GLog "STUCK ${age} min → would Restore-ScreenState + minimize Claude (claude $(Minimize-ClaudeWindows -WhatIf))"; exit 0 }
    $r = Restore-ScreenState -Path $ss -NoReassert -ForegroundOnly:$false
    # Restore-ScreenState 會記 finished；改記 guard_restored 並把 finished 清回原值，免得看起來像正常完成
    Set-RunFlag finished $(if ($finished) { $finished.ToString('s') } else { '' })
    Mark guard_restored
    GLog ("STUCK ${age} min since started → restored: tos={0} zorder={1} foreground={2} claude_minimized={3}" -f $r.tos, $r.zorder, $r.foreground, $r.claude)
    exit 0
}

if ($started -and $finished) {
    if ($flag.guard_final -and [datetime]$flag.guard_final -ge $finished) { exit 0 }
    $c = if ($DryRun) { Minimize-ClaudeWindows -WhatIf } else { Minimize-ClaudeWindows }
    Mark guard_final
    GLog "run finished $($finished.ToString('HH:mm')) → claude minimized $c"
    exit 0
}

# 沒有旗標：休市日（步驟 0 就結束）或步驟 1.5 之前就卡住
if ($Now.TimeOfDay -ge [timespan]'06:40') {
    if ($flag -and $flag.guard_final) { exit 0 }
    $c = if ($DryRun) { Minimize-ClaudeWindows -WhatIf } else { Minimize-ClaudeWindows }
    Mark guard_final
    GLog "no run flag today → claude minimized $c"
} else { GLog "no run flag yet → wait" }
exit 0
