<#
  Boss PC（交易機）一次性建置腳本
  ═══════════════════════════════════════════════════════════════
  在 Boss PC 上以「系統管理員」身分執行一次即可：

      powershell -ExecutionPolicy Bypass -File bootstrap_boss.ps1

  它會做：
      1. 檢查 git / python / port 5500
      2. clone private repo 到 C:\TradeReview
      3. 安裝 Python 套件
      4. 註冊 pythonw 開機常駐排程
      5. 檢查 history_minute.xlsx（K 線資料，不走 git，需手動帶）

  ⚠️ 這台機器只做 CLAUDE.md 的「第一段」：對帳單 → trades_all.xlsx。
     CS交易紀錄.xlsx / offset_state.json 是 502 的所有物，
     Boss 上的採集流程一律用 spy_daytrade_writer.py --trades-only，
     絕不可跑不帶該旗標的完整 writer。
#>
[CmdletBinding()]
param(
    [string]$Root      = 'C:\TradeReview',
    [string]$RepoUrl   = 'https://github.com/shuye1668/SPY-Day-Trade-Review.git',
    [string]$PythonExe = 'C:\Users\TS USER\AppData\Local\Programs\Python\Python313\python.exe',
    [string]$PythonWExe= 'C:\Users\TS USER\AppData\Local\Programs\Python\Python313\pythonw.exe',
    [switch]$SkipTask
)

$ErrorActionPreference = 'Stop'
$fail = 0
function Step($n, $t) { Write-Host "`n[$n] $t" -ForegroundColor Cyan }
function Ok($m)   { Write-Host "    [OK] $m" -ForegroundColor Green }
function Warn($m) { Write-Host "    [!]  $m" -ForegroundColor Yellow }
function Bad($m)  { Write-Host "    [X]  $m" -ForegroundColor Red; $script:fail++ }

Write-Host "`n  Boss PC 建置：$Root" -ForegroundColor White
Write-Host "  Repo：$RepoUrl"

# ── 1. 前置檢查 ───────────────────────────────────────────────────
Step 1 "前置檢查"

if (Get-Command git.exe -ErrorAction SilentlyContinue) {
    Ok "git：$((& git.exe --version))"
} else {
    Bad "找不到 git。請先安裝 https://git-scm.com/download/win"
}

if (Test-Path $PythonExe)  { Ok "python：$PythonExe" }  else { Bad "找不到 python：$PythonExe" }
if (Test-Path $PythonWExe) { Ok "pythonw：$PythonWExe" } else { Bad "找不到 pythonw：$PythonWExe" }

$busy = Get-NetTCPConnection -LocalPort 5500 -State Listen -ErrorAction SilentlyContinue
if ($busy) {
    $p = Get-Process -Id $busy[0].OwningProcess -ErrorAction SilentlyContinue
    Warn "port 5500 已被 $($p.ProcessName) (PID $($p.Id)) 佔用 —— app 會起不來"
} else { Ok "port 5500 未被佔用" }

$isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()
           ).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $isAdmin -and -not $SkipTask) {
    Warn "非系統管理員 —— 步驟 4（註冊排程）會失敗。可先跑完其餘步驟，之後再提權執行 setup_scheduled_task.ps1"
}

if ($fail) { Write-Host "`n  前置檢查未通過，停止。`n" -ForegroundColor Red; exit 1 }

# ── 2. clone ──────────────────────────────────────────────────────
# 原生指令寫 stderr 時，搭配 2>&1 與 $ErrorActionPreference='Stop'，
# PowerShell 會把 git 的「警告」也當成終止錯誤丟出（例如 CRLF 提示）。
# 成敗只能看 exit code。這裡自帶包裝，不依賴 repo 內的 _git_helper.ps1
# —— clone 之前那個檔還不存在。
function Git-Quiet {
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try { & git.exe @args 2>&1 | Out-Null } finally { $ErrorActionPreference = $prev }
    return $LASTEXITCODE
}

Step 2 "取得 repo"
if (Test-Path (Join-Path $Root '.git')) {
    Ok "$Root 已是 git repo，改為更新"
    Push-Location $Root

    # Boss 不擁有這些檔，本機任何改動都不具權威性 → 直接丟棄，
    # 否則它們會讓 pull --rebase 失敗（實測 2026-08-10 就卡在這）
    foreach ($p in @('CS交易紀錄.xlsx','CS交易紀錄_dump.txt','offset_state.json','notes')) {
        if (Test-Path $p) { Git-Quiet checkout -- $p | Out-Null }
    }

    Git-Quiet fetch origin | Out-Null
    # --autostash：剩下的本機改動（例如常駐 app 對 trades_all 的回寫）
    # 自動收起再放回，不必人工介入
    if ((Git-Quiet pull --rebase --autostash origin main) -ne 0) {
        Warn "pull 失敗。請在 $Root 執行 git status 看是什麼擋住，"
        Warn "     多半是本機有未提交改動；確認不需要保留就用："
        Warn "     git checkout -- . ; git clean -fd"
    } else {
        Ok "已更新到最新版"
    }
    Pop-Location
} else {
    if ((Test-Path $Root) -and (Get-ChildItem $Root -Force | Measure-Object).Count -gt 0) {
        Bad "$Root 已存在且非空，且不是 git repo。請先清空或改用 -Root 指定其他路徑"
        exit 1
    }
    # 第一次 clone 會要求登入：用瀏覽器授權（git credential manager）或 PAT
    & git.exe clone $RepoUrl $Root
    if ($LASTEXITCODE -ne 0) { Bad "clone 失敗（私有 repo 需要認證）"; exit 1 }
    Ok "已 clone 到 $Root"
}

Push-Location $Root
Git-Quiet config user.name  "boss-pc"             | Out-Null
Git-Quiet config user.email "shuye1668@gmail.com" | Out-Null
Ok "git 身分已設定"

# ── 3. Python 套件 ────────────────────────────────────────────────
Step 3 "安裝 Python 套件"
$prevEap = $ErrorActionPreference; $ErrorActionPreference = 'Continue'
& $PythonExe -m pip install --quiet --upgrade pip 2>&1 | Out-Null
& $PythonExe -m pip install --quiet flask pandas numpy openpyxl yfinance 2>&1 | Out-Null
$pipRc = $LASTEXITCODE
$ErrorActionPreference = $prevEap
if ($pipRc -eq 0) { Ok "flask / pandas / numpy / openpyxl / yfinance" }
else { Bad "pip 安裝失敗，請手動執行： & `"$PythonExe`" -m pip install flask pandas numpy openpyxl yfinance" }

# ── 4. 開機常駐排程 ───────────────────────────────────────────────
Step 4 "註冊開機常駐排程"
if ($SkipTask) {
    Warn "已指定 -SkipTask，略過"
} elseif (-not $isAdmin) {
    Warn "非管理員，略過。之後請以管理員執行： .\setup_scheduled_task.ps1 -PythonW `"$PythonWExe`""
} else {
    & (Join-Path $Root 'setup_scheduled_task.ps1') -PythonW $PythonWExe
    if ($LASTEXITCODE -eq 0) { Ok "TradeReview_Dashboard 已註冊" } else { Bad "排程註冊失敗" }
}

# ── 5. K 線資料 ───────────────────────────────────────────────────
Step 5 "檢查 K 線資料"
$hist = Join-Path $Root 'history_minute.xlsx'
if (Test-Path $hist) {
    Ok "history_minute.xlsx 已就位（$([math]::Round((Get-Item $hist).Length/1MB,1)) MB）"
} else {
    Warn "缺 history_minute.xlsx（4.2 MB，刻意不進 git —— 每天被 app 追加，"
    Warn "     進 git 一年會撐到 1 GB）。請從 502 用隨身碟/共享複製過來，"
    Warn "     放到 $Root\。不放也能跑，但 app 首次要靠 yfinance 補 318 天，很慢。"
}

Pop-Location

# ── 收尾 ──────────────────────────────────────────────────────────
Write-Host "`n═══════════════════════════════════════════════" -ForegroundColor White
if ($fail) {
    Write-Host "  有 $fail 項失敗，請看上面紅字" -ForegroundColor Red
} else {
    Write-Host "  建置完成" -ForegroundColor Green
    Write-Host ""
    Write-Host "  驗收："
    Write-Host "    Start-ScheduledTask -TaskName TradeReview_Dashboard"
    Write-Host "    Start-Sleep 8; (Invoke-WebRequest http://localhost:5500 -UseBasicParsing).StatusCode   # 應為 200"
    Write-Host ""
    Write-Host "  每日採集流程接上本機寫入（採集到 CSV 之後）："
    Write-Host "    cd $Root"
    Write-Host "    & `"$PythonExe`" spy_daytrade_engine.py _inbox\<date>.csv          # 先乾跑，看是否全綠"
    Write-Host "    & `"$PythonExe`" spy_daytrade_writer.py _inbox\<date>.csv --commit --trades-only"
    Write-Host "    .\sync_push.ps1 -Owner boss"
    Write-Host ""
    Write-Host "  ⚠️ writer 一定要帶 --trades-only。CS交易紀錄/offset_state 是 502 的所有物。" -ForegroundColor Yellow
}
Write-Host ""
exit ([int]($fail -gt 0))
