<#
  Boss PC 每日採集的開工前檢查（SHUYE07 維護，經 git 送達；2026-10-28）

  用法（SKILL 步驟 0.5）：
      powershell -ExecutionPolicy Bypass -File C:\TradeReview\boss_preflight.ps1

  做三件事，全部「只回報、不修」，任何一項失敗都不擋採集：
    1. sync_pull -Owner boss：開工前先拿到 SHUYE07 最新的程式與 trades_all
       （原本只在步驟 9.5 才 pull，writer 改版要晚一天才生效）。
    2. 檢查 repo 是否停在 rebase、守門排程 Boss_Screen_Guard 是否存在。
    3. 列出比 BOSS_SKILL_current.md 修訂紀錄最新日期還新的 BOSS_PATCH_*.md（＝尚未套用）。

  輸出一律是 PREFLIGHT_<項目>=<值> 一行一項，給 SKILL 照抄進步驟 11 log 與步驟 12 回報。
  exit 永遠 0（檢查本身壞掉也不擋採集，改印 PREFLIGHT_ERROR=...）。
#>
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $root

function Out-Pre($k, $v) { Write-Host "PREFLIGHT_$k=$v" }

# 1) 開工前 pull（丟棄 Boss 本機對 SHUYE07 所有物的改動 → 以 SHUYE07 版 trades_all 為準）
try {
    $out = & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $root 'sync_pull.ps1') -Owner boss 2>&1
    $code = $LASTEXITCODE
    $txt = (($out | Out-String).Trim() -replace '\s+', ' ')
    switch ($code) {
        0 { Out-Pre 'PULL' "OK $txt" }
        2 { Out-Pre 'PULL' "SKIPPED (本機有未推送的改動) $txt" }
        default { Out-Pre 'PULL' "FAILED exit=$code $txt" }
    }
} catch { Out-Pre 'PULL' "FAILED $($_.Exception.Message)" }

# 2) repo／守門排程狀態
try {
    . (Join-Path $root '_git_helper.ps1')
    Out-Pre 'REBASE' $(if (Test-RebaseInProgress) { 'STUCK（repo 停在 rebase 中間，不要自行修復，回報使用者）' } else { 'OK' })
    Out-Pre 'HEAD' ((Invoke-Git rev-parse --short HEAD | Out-String).Trim())
} catch { Out-Pre 'ERROR' "git 狀態檢查失敗：$($_.Exception.Message)" }

$guard = Get-ScheduledTask -TaskName 'Boss_Screen_Guard' -ErrorAction SilentlyContinue
Out-Pre 'GUARD' $(if ($guard) { "OK ($($guard.State))" } else { 'MISSING（步驟 10.5 第一次執行 Restore-ScreenState 會自動註冊）' })

# 3) 尚未套用的 Boss patch：檔名日期 > BOSS_SKILL_current.md 修訂紀錄最新一列的日期
try {
    $skill = Join-Path $root 'BOSS_SKILL_current.md'
    $latest = ''
    if (Test-Path $skill) {
        $rows = Select-String -Path $skill -Encoding UTF8 -Pattern '^\|\s*(\d{4}-\d{2}-\d{2})\s*\|'
        $dates = @($rows | ForEach-Object { $_.Matches[0].Groups[1].Value } | Sort-Object -Descending)
        if ($dates.Count) { $latest = $dates[0] -replace '-', '' }
    }
    $pending = @(Get-ChildItem -Path $root -Filter 'BOSS_PATCH_*.md' | Where-Object {
            $_.BaseName -match '_(\d{8})$' -and $Matches[1] -gt $latest } | Sort-Object Name | ForEach-Object { $_.Name })
    Out-Pre 'SKILL_REV' $(if ($latest) { $latest } else { 'unknown' })
    Out-Pre 'PENDING_PATCHES' $(if ($pending.Count) { $pending -join ', ' } else { 'none' })
} catch { Out-Pre 'ERROR' "patch 檢查失敗：$($_.Exception.Message)" }

exit 0
