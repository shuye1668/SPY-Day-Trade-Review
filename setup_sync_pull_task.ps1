<#
  註冊「502 每日自動接收 Boss 產出」排程
  ══════════════════════════════════════════════════════════════
  ⚠️ 需以「系統管理員」執行

      powershell -ExecutionPolicy Bypass -File setup_sync_pull_task.ps1
      powershell -ExecutionPolicy Bypass -File setup_sync_pull_task.ps1 -At 08:00,11:30

  每天從 GitHub 取回 Boss 的產出：trades_all.xlsx / _inbox\ / sync_state.json

  為什麼需要這支（2026-08-19 補）：
  502 原本只有「推」（TradeReview_SyncPush_p502）沒有「拉」，Boss 就算準時跑完並
  推上 GitHub，502 也要等人手動執行才拿得到 —— 這正是每天得手動按
  「每日一鍵複盤.bat」的其中一個原因。

  預設 09:30 與 11:30：Boss 的採集 routine 約 06:47 起跑、跑完才推，
  09:30 先取一次，11:30 補一次（涵蓋 Boss 晚跑或補跑的情況）；
  兩個時間都排在 SyncPush 的 10:00 / 12:00 之前，順序是「先拉再推」。

  安全性：sync_pull.ps1 -Owner p502 只會丟棄「Boss 所有物」的本機改動
  （trades_all / _inbox / sync_state —— 常駐 app 的配對回寫就屬於這類，
  同步後會重取），CS交易紀錄.xlsx、offset_state.json、notes\ 完全不碰。
  若有本機未推送的自有改動，它會直接跳過並回報 exit 2，不會覆蓋你的人工修正。
#>
[CmdletBinding()]
param(
    [string[]]$At = @('09:30', '11:30'),
    [string]$TaskName = 'TradeReview_SyncPull_p502'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host "`n  [X] 需要系統管理員權限" -ForegroundColor Red
    Write-Host "      請改用「以系統管理員身分執行」的 PowerShell 再跑一次。`n"
    exit 1
}

$script = Join-Path $root 'sync_pull.ps1'
if (-not (Test-Path $script)) { throw "找不到 $script" }

# -WindowStyle Hidden：不要在畫面上彈出藍色視窗
# -ExecutionPolicy Bypass：這台的 ExecutionPolicy 未必允許直接跑 .ps1
$argline = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden ' +
           '-File "' + $script + '" -Owner p502 -Quiet'

$A = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argline -WorkingDirectory $root
$T = @($At | ForEach-Object { New-ScheduledTaskTrigger -Daily -At $_ })
$P = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
        -LogonType Interactive -RunLevel Limited
# StartWhenAvailable：當時電腦沒開的話，開機後補跑（否則整天就漏掉了）
$S = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 20) -StartWhenAvailable `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 5)

Register-ScheduledTask -TaskName $TaskName -Action $A -Trigger $T `
                       -Principal $P -Settings $S -Force | Out-Null

$t = Get-ScheduledTask -TaskName $TaskName
Write-Host ""
Write-Host "  [OK] 已註冊：$TaskName" -ForegroundColor Green
Write-Host ("       每天 " + ($At -join '、') + " 各執行一次；電腦當時沒開會在開機後補跑")
Write-Host ("       觸發器 : " + (($t.Triggers | ForEach-Object {
                ([datetime]$_.StartBoundary).ToString('HH:mm') }) -join '、'))
Write-Host ("       Log  : " + (Join-Path $root '_logs\sync_pull.log'))
Write-Host ""
Write-Host "  立即測試：" -ForegroundColor Cyan
Write-Host "      Start-ScheduledTask -TaskName $TaskName"
Write-Host "      Get-Content '$root\_logs\sync_pull.log' -Tail 3"
Write-Host ""
