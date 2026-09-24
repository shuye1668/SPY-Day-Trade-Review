<#
  推送本機產出到 GitHub（兩台都用，靠 -Owner 決定推哪些檔）
  ────────────────────────────────────────────────────────────
      .\sync_push.ps1 -Owner boss   # Boss PC：trades_all / _inbox / sync_state
      .\sync_push.ps1 -Owner p502   # 502：CS交易紀錄 / offset_state / notes / 程式碼

  分區所有權是這套同步不會衝突的關鍵：xlsx 是 binary，git 無法 merge，
  兩台若改同一個 .xlsx 就是死結。各自只推自己的檔 → rebase 永遠成功。

  exit 0=完成或無變更  1=失敗
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][ValidateSet('boss', 'p502')][string]$Owner,
    [string]$Message,
    [switch]$Quiet,
    # 回歸閘門攔下後，人工確認無誤時用這個繞過（排程一律不帶此參數）
    [switch]$Force
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root
. (Join-Path $root '_git_helper.ps1')

function Say($m) { if (-not $Quiet) { Write-Host $m } }
$log = Join-Path $root '_logs\sync_push.log'
New-Item -ItemType Directory -Force -Path (Split-Path $log) | Out-Null
function Log($m) { "$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')  [$Owner] $m" | Add-Content $log -Encoding utf8 }

# 只 stage 自己擁有的路徑，絕不 git add -A
if ($Owner -eq 'boss') {
    $paths = @('trades_all.xlsx', '_inbox', 'sync_state.json')
} else {
    $paths = @('CS交易紀錄.xlsx', 'CS交易紀錄_dump.txt', 'offset_state.json',
               'notes', '*.py', '*.md', '*.bat', '*.ps1', '.gitignore', '.gitattributes')
    # 502 推之前先更新 CS 文字側寫，git diff 才看得到帳本改了哪一格
    if (Test-Path (Join-Path $root 'CS交易紀錄.xlsx')) {
        try { & python (Join-Path $root 'dump_cs_text.py') | Out-Null } catch { }
    }
}

try {
    Invoke-Git fetch origin | Out-Null

    # 先 stage 自己的改動，pull --rebase 才不會被 unstaged changes 擋掉
    foreach ($p in $paths) { Try-Git add -- $p | Out-Null }

    $staged = @(Invoke-Git diff --cached --name-only)
    if ($staged.Count -eq 0) {
        Say "  沒有要推送的變更"
        Log "nothing-to-push"
        exit 0
    }

    # ── 回歸閘門（2026-09-24 加，勿移除）────────────────────────────
    # 2026-09-18 18:00 的自動推送，把一批 2026-07 的舊檔（有人把搬家前的
    # C:\TradeReview 備份複製進來）當成正常變更提交，一次砍掉 trade_review_app.py
    # 805 行、CS 帳本縮水 43%，完全沒有警示。無人看管的推送絕不該有能力靜默大倒退。
    #
    # 帳本大小一律走 git blob SHA 比對，不碰檔名：PowerShell 5.1 以 ANSI(cp950)
    # 解讀 git 的 UTF-8 輸出，「CS交易紀錄.xlsx」會被讀成亂碼，用檔名去 Get-Item /
    # cat-file 必定查無此檔、大小算成 0，閘門就對最該保護的那本帳靜默失效
    # （2026-09-24 實測揪出，第一版正是這樣壞的）。SHA 是純 ASCII，不受影響。
    [Console]::OutputEncoding = [Text.Encoding]::UTF8   # 讓訊息裡的中文檔名可讀
    $ins = 0; $del = 0; $shrunk = @()
    foreach ($line in @(Invoke-Git -c core.quotepath=false diff --cached --numstat)) {
        $f = ($line -split "`t")
        if ($f.Count -lt 3) { continue }
        if ($f[0] -ne '-' -and $f[1] -ne '-') { $ins += [int]$f[0]; $del += [int]$f[1] }
    }
    foreach ($line in @(Invoke-Git -c core.quotepath=false diff --cached --raw --abbrev=40)) {
        $parts = ($line -split "`t")
        if ($parts.Count -lt 2) { continue }
        $path = $parts[1]
        if ($path -notmatch '\.xlsx"?$') { continue }        # 只看帳本，避免誤擋
        $meta = ($parts[0] -split ' ')
        if ($meta.Count -lt 5) { continue }
        $oldSha = $meta[2]; $newSha = $meta[3]
        if ($oldSha -match '^0+$' -or $newSha -match '^0+$') { continue }   # 新增或刪除
        $os = 0; $ns = 0
        try { $os = [int]((Invoke-Git cat-file -s $oldSha | Out-String).Trim()) } catch { }
        try { $ns = [int]((Invoke-Git cat-file -s $newSha | Out-String).Trim()) } catch { }
        if ($os -gt 0 -and $ns -gt 0) {
            if ($ns -lt ($os * 0.75)) { $shrunk += ("{0}（{1} → {2} bytes）" -f $path, $os, $ns) }
        } else {
            Say ("  [?] 無法比對帳本大小，未做縮水檢查：$path")   # 量不到就要出聲
            Log "size-check-skipped $path old=$os new=$ns"
        }
    }
    $massDelete = ($del -ge 300 -and $del -gt ($ins * 5))
    if ((-not $Force) -and ($massDelete -or $shrunk.Count -gt 0)) {
        Say ""
        Say "  [!!] 偵測到疑似大規模倒退，已中止推送（沒有 commit、沒有 push）"
        if ($massDelete) { Say ("       文字檔：新增 $ins 行 / 刪除 $del 行") }
        foreach ($x in $shrunk) { Say ("       帳本縮水：$x") }
        Say "       這通常代表舊備份被複製進資料夾覆蓋了現況。"
        Say "       請先確認：  git diff --cached --stat"
        Say "       確定無誤要強制推送：  .\sync_push.ps1 -Owner $Owner -Force"
        Log "BLOCKED regression: ins=$ins del=$del shrunk=$($shrunk -join '; ')"
        Try-Git reset | Out-Null
        exit 3
    }

    if (-not $Message) { $Message = "$Owner sync $(Get-Date -f 'yyyy-MM-dd HH:mm')" }
    Invoke-Git commit -m $Message | Out-Null

    # --autostash：working tree 可能還有「對方所有物」的未 staged 改動
    # （例如 502 上出現了 Boss 才該寫的 sync_state.json），那會讓 rebase
    # 直接失敗。autostash 會自動收起再放回，不必人工介入。
    Invoke-Git pull --rebase --autostash origin main | Out-Null

    # push 失敗要重試：無人看管下靜默失敗＝對方拿到舊資料卻不知道
    $ok = $false
    for ($i = 1; $i -le 3; $i++) {
        if ((Try-Git push origin main) -eq 0) { $ok = $true; break }
        Say "  push 第 $i 次失敗，重試..."
        Start-Sleep -Seconds (5 * $i)
        Try-Git pull --rebase --autostash origin main | Out-Null
    }

    if ($ok) {
        $head = (Invoke-Git rev-parse --short HEAD | Out-String).Trim()
        Say "  已推送 $($staged.Count) 個檔案（HEAD $head）"
        $staged | ForEach-Object { Say "    $_" }
        Log "pushed files=$($staged.Count) head=$head"
        exit 0
    } else {
        Say "  [X] push 三次均失敗 —— 本機已 commit，資料沒掉，但對方拿不到"
        Log "FAIL push-after-3-retries"
        exit 1
    }
}
catch {
    Say "  [X] 失敗：$($_.Exception.Message)"
    Log "FAIL $($_.Exception.Message)"
    exit 1
}
