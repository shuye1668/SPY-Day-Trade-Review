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
