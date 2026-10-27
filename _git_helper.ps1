<#
  兩支同步腳本共用的 git 呼叫包裝。

  這裡有兩個 Windows PowerShell 的坑，都會造成「明明失敗卻回報成功」或
  「明明成功卻報錯」，在無人看管的排程裡特別致命：

  1. 函式若命名為 Git，因為 PowerShell 大小寫不敏感，內部的 `& git`
     會呼叫到自己 → 呼叫深度溢位。故一律用 Invoke-Git + git.exe。

  2. $ErrorActionPreference='Stop' 搭配 2>&1 時，原生指令寫到 stderr 的
     「警告」（例如 git 的 LF/CRLF 提示）會被當成終止錯誤丟出。
     成敗只能看 exit code，不能看有沒有 stderr 輸出。
#>

function Invoke-Git {
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & git.exe @args 2>&1
    } finally {
        $ErrorActionPreference = $prev
    }
    if ($LASTEXITCODE -ne 0) {
        throw "git $($args -join ' ') 失敗 (exit $LASTEXITCODE)：$(($out | Out-String).Trim())"
    }
    return ($out | Where-Object { $_ -isnot [System.Management.Automation.ErrorRecord] })
}

# 允許失敗、只取 exit code 的版本（例如 git add 不存在的路徑）
function Try-Git {
    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        & git.exe @args 2>&1 | Out-Null
    } finally {
        $ErrorActionPreference = $prev
    }
    return $LASTEXITCODE
}

# repo 是否停在 rebase 中間（.git\rebase-merge 或 rebase-apply 存在）
function Test-RebaseInProgress {
    $gitDir = (Invoke-Git rev-parse --git-dir | Out-String).Trim()
    return ((Test-Path (Join-Path $gitDir 'rebase-merge')) -or (Test-Path (Join-Path $gitDir 'rebase-apply')))
}

# pull --rebase 失敗就一定 abort（2026-10-27 加，勿移除）。
# 原本 pull --rebase 衝突時直接 throw，repo 就停在 rebase 中間（detached HEAD），
# 之後每天的 push/pull 都失敗或在 detached HEAD 上堆孤兒 commit —— Boss 2026-08-11
# 卡兩個月、2026-10-23 又卡在 sync_state.json 衝突，都是這樣來的。
# abort 會回到 pull 前的分支與 commit，--autostash 收起的改動也會放回，資料不掉；下次執行再試。
# 回傳 @{ Ok; Aborted; Conflicts; Output }，由呼叫端決定怎麼記 log、回什麼 exit code。
function Invoke-PullRebase {
    param([switch]$AutoStash)
    $pullArgs = @('pull', '--rebase')
    if ($AutoStash) { $pullArgs += '--autostash' }
    $pullArgs += @('origin', 'main')

    $prev = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    try {
        $out = & git.exe @pullArgs 2>&1
    } finally {
        $ErrorActionPreference = $prev
    }
    $code = $LASTEXITCODE
    $result = @{ Ok = ($code -eq 0); Aborted = $false; Conflicts = @(); Output = ($out | Out-String).Trim() }
    if ($code -ne 0 -and (Test-RebaseInProgress)) {
        $prev = $ErrorActionPreference
        $ErrorActionPreference = 'Continue'
        try {
            $result.Conflicts = @(& git.exe -c core.quotepath=false diff --name-only --diff-filter=U 2>$null)
        } finally {
            $ErrorActionPreference = $prev
        }
        Try-Git rebase --abort | Out-Null
        $result.Aborted = -not (Test-RebaseInProgress)
    }
    return $result
}
