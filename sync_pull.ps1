<#
  502 端：從 GitHub 拉取 Boss PC 產出的資料
  ────────────────────────────────────────────────────────────
  排程建議：開機時 + 每 30 分鐘。不需要提權。

  關鍵細節 —— 為什麼不能直接 git pull：
  常駐的 trade_review_app 會回寫 trades_all.xlsx（補配對欄位 Action/Status/
  Pair_Date/Pair_Time），而 trades_all 是 Boss 的所有物。這會讓 working tree
  變髒、pull 直接失敗。所以先精準丟棄「Boss 所有物」的本地改動，
  notes/ 與 CS交易紀錄.xlsx（502 的所有物）完全不碰。

  exit 0=完成  1=失敗  2=有本機未推送的改動而跳過
#>
[CmdletBinding()]
param(
    [ValidateSet('boss', 'p502')][string]$Owner = 'p502',
    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root
. (Join-Path $root '_git_helper.ps1')

function Say($m) { if (-not $Quiet) { Write-Host $m } }
$log = Join-Path $root '_logs\sync_pull.log'
New-Item -ItemType Directory -Force -Path (Split-Path $log) | Out-Null
function Log($m) { "$(Get-Date -f 'yyyy-MM-dd HH:mm:ss')  $m" | Add-Content $log -Encoding utf8 }

# 丟棄「對方所有物」的本地改動 —— 那些改動不具權威性。
# 對稱處理很重要：Boss 若誤跑了完整 writer（會寫 CS/offset_state），
# 那些改動不會被 -Owner boss 推出去，就會永遠卡住 Boss 的 pull。
if ($Owner -eq 'p502') {
    $foreign = @('trades_all.xlsx', 'sync_state.json', '_inbox')
} else {
    $foreign = @('CS交易紀錄.xlsx', 'CS交易紀錄_dump.txt', 'offset_state.json', 'notes')
}

try {
    foreach ($p in $foreign) {
        if (Test-Path $p) { Try-Git checkout -- $p | Out-Null }
    }

    # 還有殘留改動就停 —— 寧可停下來讓人看到，也不要 pull 失敗卻回報成功
    $dirty = Invoke-Git status --porcelain
    if ($dirty) {
        $names = (($dirty | ForEach-Object { $_.ToString().Substring(3) }) -join ', ')
        Say "  [!] 有未提交的本機改動，跳過 pull：$names"
        Say "      這些是本機自己的檔"
        Say "      → 請先執行： .\sync_push.ps1 -Owner $Owner"
        Log "SKIP dirty: $names"
        exit 2
    }

    $before = (Invoke-Git rev-parse HEAD | Out-String).Trim()
    Invoke-Git fetch origin | Out-Null
    Invoke-Git pull --rebase origin main | Out-Null
    $after = (Invoke-Git rev-parse HEAD | Out-String).Trim()

    if ($before -eq $after) {
        Say "  無新資料（HEAD $($after.Substring(0,7))）"
        Log "no-change $($after.Substring(0,7))"
    } else {
        $files = Invoke-Git diff --name-only $before $after
        Say "  已更新 $($before.Substring(0,7)) -> $($after.Substring(0,7))"
        $files | ForEach-Object { Say "    $_" }
        Log "updated $($before.Substring(0,7))->$($after.Substring(0,7)) files=$(@($files).Count)"
        # trade_review_app 每 4 秒偵測 trades_all mtime，會自己刷新，不必重啟
    }
    exit 0
}
catch {
    Say "  [X] pull 失敗：$($_.Exception.Message)"
    Log "FAIL $($_.Exception.Message)"
    exit 1
}
