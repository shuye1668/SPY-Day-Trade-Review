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
    [switch]$Quiet
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

    if (-not $Message) { $Message = "$Owner sync $(Get-Date -f 'yyyy-MM-dd HH:mm')" }
    Invoke-Git commit -m $Message | Out-Null

    # commit 之後才 rebase：working tree 乾淨，rebase 一定能跑
    Invoke-Git pull --rebase origin main | Out-Null

    # push 失敗要重試：無人看管下靜默失敗＝對方拿到舊資料卻不知道
    $ok = $false
    for ($i = 1; $i -le 3; $i++) {
        if ((Try-Git push origin main) -eq 0) { $ok = $true; break }
        Say "  push 第 $i 次失敗，重試..."
        Start-Sleep -Seconds (5 * $i)
        Try-Git pull --rebase origin main | Out-Null
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
