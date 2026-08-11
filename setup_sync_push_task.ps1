<#
  註冊「502 每日自動推送」排程
  ══════════════════════════════════════════════════════════════
  ⚠️ 需以「系統管理員」執行

      powershell -ExecutionPolicy Bypass -File setup_sync_push_task.ps1
      powershell -ExecutionPolicy Bypass -File setup_sync_push_task.ps1 -At 12:00

  每天中午把 502 擁有的檔案推上 GitHub：
      CS交易紀錄.xlsx / CS交易紀錄_dump.txt / offset_state.json / notes\ / 程式碼

  為什麼排 12:00：做帳與人工修正通常在早上完成，中午推一次剛好涵蓋當天成果。
  下午之後才改的東西會等到隔天中午 —— 想立刻推就手動跑：
      powershell -ExecutionPolicy Bypass -File sync_push.ps1 -Owner p502

  注意：這支只推「502 的所有物」。Boss 的產物（trades_all / _inbox /
  sync_state）由 Boss 自己推，這裡不會碰，也不會誤推。
#>
[CmdletBinding()]
param(
    [string]$At = '12:00',
    [string]$TaskName = 'TradeReview_SyncPush_p502'
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

$script = Join-Path $root 'sync_push.ps1'
if (-not (Test-Path $script)) { throw "找不到 $script" }

# -WindowStyle Hidden：中午不要在畫面上彈出藍色視窗
# -ExecutionPolicy Bypass：這台的 ExecutionPolicy 未必允許直接跑 .ps1
$argline = '-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden ' +
           '-File "' + $script + '" -Owner p502 -Quiet'

$A = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument $argline -WorkingDirectory $root
$T = New-ScheduledTaskTrigger -Daily -At $At
$P = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
        -LogonType Interactive -RunLevel Limited
# StartWhenAvailable：中午電腦沒開的話，開機後補跑（否則整天就漏掉了）
$S = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 20) -StartWhenAvailable `
        -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries `
        -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 5)

Register-ScheduledTask -TaskName $TaskName -Action $A -Trigger $T `
                       -Principal $P -Settings $S -Force | Out-Null

$t = Get-ScheduledTask -TaskName $TaskName
Write-Host ""
Write-Host "  [OK] 已註冊：$TaskName" -ForegroundColor Green
Write-Host ("       每天 " + $At + " 執行；電腦當時沒開會在開機後補跑")
Write-Host ("       Exec : " + $t.Actions[0].Execute)
Write-Host ("       Args : " + $t.Actions[0].Arguments)
Write-Host ("       Log  : " + (Join-Path $root '_logs\sync_push.log'))
Write-Host ""
Write-Host "  立即測試：" -ForegroundColor Cyan
Write-Host "      Start-ScheduledTask -TaskName $TaskName"
Write-Host "      Get-Content '$root\_logs\sync_push.log' -Tail 3"
Write-Host ""
