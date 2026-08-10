<#
  TradeReview — 註冊「開機自動 silent 常駐」工作排程
  ────────────────────────────────────────────────────────────
  兩台機器共用這一支（502 與 Boss PC）。

  ⚠️ 必須用「系統管理員」身分執行 —— 寫入工作排程器根目錄需要提權。
     做法：開始功能表 → 打 powershell → 右鍵「以系統管理員身分執行」→
           cd 到本資料夾 → .\setup_scheduled_task.ps1

  參數都可省略，會自動偵測本機的 pythonw 與本腳本所在資料夾。
      .\setup_scheduled_task.ps1
      .\setup_scheduled_task.ps1 -PythonW "C:\Users\TS USER\AppData\Local\Programs\Python\Python313\pythonw.exe"

  設定值刻意對齊這台既有的 AIBubble_Dashboard / FlatBottom_Dashboard：
      pythonw.exe            → 無主控台視窗，完全 silent
      AtLogOn                → 每次登入自動啟動
      ExecutionTimeLimit 0   → 不限執行時間（預設 72 小時會把常駐服務砍掉）
      MultipleInstances Ignore→ 已在跑就不重複開
#>
[CmdletBinding()]
param(
    [string]$PythonW,
    [string]$AppPath,
    [string]$TaskName = 'TradeReview_Dashboard'
)

$ErrorActionPreference = 'Stop'

# ── 提權檢查（先擋，否則錯到最後才發現）────────────────────────────
$isAdmin = ([Security.Principal.WindowsPrincipal] `
            [Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin) {
    Write-Host ""
    Write-Host "  [X] 需要系統管理員權限" -ForegroundColor Red
    Write-Host "      註冊工作排程器根目錄的任務必須提權。"
    Write-Host "      請關掉這個視窗，改用「以系統管理員身分執行」的 PowerShell 再跑一次。"
    Write-Host ""
    exit 1
}

# ── 路徑解析 ───────────────────────────────────────────────────────
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $AppPath) { $AppPath = Join-Path $root 'trade_review_app.py' }

if (-not $PythonW) {
    # 依序找：PATH 上的 pythonw → 目前使用者的 Python 安裝目錄
    $cand = @()
    $c = Get-Command pythonw.exe -ErrorAction SilentlyContinue
    if ($c) { $cand += $c.Source }
    $cand += Get-ChildItem "$env:LOCALAPPDATA\Programs\Python" -Filter pythonw.exe `
                -Recurse -ErrorAction SilentlyContinue |
             Sort-Object FullName -Descending | Select-Object -ExpandProperty FullName
    $PythonW = $cand | Select-Object -First 1
}

if (-not $PythonW -or -not (Test-Path $PythonW)) {
    throw "找不到 pythonw.exe。請用 -PythonW 明確指定，例如：
    .\setup_scheduled_task.ps1 -PythonW `"C:\Users\TS USER\AppData\Local\Programs\Python\Python313\pythonw.exe`""
}
if (-not (Test-Path $AppPath)) { throw "找不到 app：$AppPath" }

# ── 註冊 ───────────────────────────────────────────────────────────
# 路徑可能含空白（例如 "C:\Users\TS USER\..."），Argument 一定要自帶引號
$A = New-ScheduledTaskAction -Execute $PythonW -Argument "`"$AppPath`"" -WorkingDirectory $root
$T = New-ScheduledTaskTrigger -AtLogOn
$P = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) `
                                -LogonType Interactive -RunLevel Limited
$S = New-ScheduledTaskSettingsSet -MultipleInstances IgnoreNew -ExecutionTimeLimit 0 `
                                  -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName $TaskName -Action $A -Trigger $T `
                       -Principal $P -Settings $S -Force | Out-Null

# ── 回報 ───────────────────────────────────────────────────────────
$t = Get-ScheduledTask -TaskName $TaskName
Write-Host ""
Write-Host "  [OK] 已註冊：$TaskName" -ForegroundColor Green
Write-Host ("       Exec     : " + $t.Actions[0].Execute)
Write-Host ("       Args     : " + $t.Actions[0].Arguments)
Write-Host ("       WorkDir  : " + $t.Actions[0].WorkingDirectory)
Write-Host ("       User     : " + $t.Principal.UserId)
Write-Host ("       Trigger  : " + $t.Triggers[0].CimClass.CimClassName)
Write-Host ("       TimeLimit: " + $t.Settings.ExecutionTimeLimit + "  (應為空白/PT0S = 不限時)")
Write-Host ("       MultiInst: " + $t.Settings.MultipleInstances)
Write-Host ""
Write-Host "  馬上測試（不必重開機）：" -ForegroundColor Cyan
Write-Host "      Start-ScheduledTask -TaskName $TaskName"
Write-Host "      Start-Sleep 6; (Invoke-WebRequest http://localhost:5500 -UseBasicParsing).StatusCode"
Write-Host "  應回 200。log 在 _logs\trade_review_app.log"
Write-Host ""
