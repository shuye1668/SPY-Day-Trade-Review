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
    # 2026-10-16 所有權定案（完成 2026-09-30 的未決事項）：trades_all.xlsx 與
    # sync_state.json 歸 502 維護（自 8/11 起實質如此），Boss 一律不推。
    # Boss 只推三樣：
    #   _inbox\              原始對帳單 CSV（源頭在 Boss）
    #   verify_shots\        每交易日 TOS 驗證截圖 PNG（502 端貼進 Julia 母檔供讀者核帳；
    #                        Gmail base64 附件實測超過 Boss Claude 可靠處理長度，git 是唯一自動通道）
    #   BOSS_SKILL_current.md  Boss 端排程 SKILL 的現行全文副本（502 端母版靠它刷新，
    #                          Boss 每次改自己的 SKILL 後要重新複製一份進 repo）
    $paths = @('_inbox', 'verify_shots', 'BOSS_SKILL_current.md')
} else {
    # 2026-09-24 補上 colors/dividends/events/.claude/*.png/*.txt：
    # 這些已追蹤檔原本不在任何一方的清單裡（「孤兒檔」）。sync_push 永遠不會
    # stage 它們，而 sync_pull 的 dirty 閘門看到已追蹤檔有改動就 exit 2，
    # 於是「請先執行 sync_push」變成死循環 —— 每天的自動拉取其實一直在失敗。
    # 2026-09-30 所有權變更：trades_all.xlsx 與 sync_state.json 改為 502 所有。
    # 原因：Boss PC 的採集腳本自 2026-08-11 起停跑（sync_state.json 自己的 note 就是
    # 這樣寫的），此後 trades_all 一直由 502 人工維護（帳本已寫到 2026-09-28）。
    # 若繼續當成 Boss 所有物，會同時踩到兩個坑：
    #   ① sync_push 不 stage 它們 → 永遠推不出去 → sync_pull 天天 exit 2 靜默停擺
    #   ② sync_pull 會 git checkout -- 把它們丟棄 → 人工補的帳直接被還原掉
    # ⚠️ 前提：Boss PC 不得再以 -Owner boss 推這兩個檔。若 Boss 日後復活，
    #    必須先決定由誰維護 trades_all，再把這裡與 sync_pull.ps1 的 $foreign 一起改回。
    $paths = @('CS交易紀錄.xlsx', 'CS交易紀錄_dump.txt', 'offset_state.json',
               'trades_all.xlsx', 'sync_state.json',
               'notes', '*.py', '*.md', '*.bat', '*.ps1', '.gitignore', '.gitattributes',
               'colors.xlsx', 'dividends.xlsx', 'events.xlsx', '.claude', '*.png', '*.txt')
    # 502 推之前先更新 CS 文字側寫，git diff 才看得到帳本改了哪一格
    if (Test-Path (Join-Path $root 'CS交易紀錄.xlsx')) {
        try { & python (Join-Path $root 'dump_cs_text.py') | Out-Null } catch { }
    }
}

function Report-PullFail($r) {
    if ($r.Aborted) {
        Say "  [X] pull --rebase 失敗，已自動 git rebase --abort（repo 回到 pull 前狀態，本機 commit 仍在）"
        if ($r.Conflicts.Count -gt 0) { Say "      衝突檔：$($r.Conflicts -join ', ')" }
        Log "FAIL pull-rebase ABORTED conflicts=$($r.Conflicts -join ', ') :: $($r.Output)"
    } else {
        Say "  [X] pull --rebase 失敗：$($r.Output)"
        Log "FAIL pull-rebase (abort=$($r.Aborted)) :: $($r.Output)"
    }
}

try {
    # 上一輪遺留的 rebase（2026-10-27 前的舊版會留下）：不自動處理——rebase 中途
    # 可能有 stage 了還沒 commit 的資料，abort 會丟掉。大聲停下等人工。
    if (Test-RebaseInProgress) {
        Say "  [X] repo 停在未完成的 rebase（.git\rebase-merge 存在），不推送；請人工處理後重跑"
        Log "FAIL rebase-in-progress (pre-existing)"
        exit 1
    }

    Invoke-Git fetch origin | Out-Null

    # 先 stage 自己的改動，pull --rebase 才不會被 unstaged changes 擋掉
    foreach ($p in $paths) { Try-Git add -- $p | Out-Null }

    # 孤兒偵測：已追蹤檔有改動卻沒被 stage，代表它不屬於任何一方的清單。
    # 這種檔會讓 sync_pull 每天 exit 2 卻沒人知道；一定要讓它出聲。
    $dirtyAll = @(Invoke-Git -c core.quotepath=false diff --name-only)
    $stagedNow = @(Invoke-Git -c core.quotepath=false diff --cached --name-only)
    # 2026-09-30：原本把 trades_all / sync_state 也排除在孤兒警告之外（因為當時它們
    # 屬於 Boss）。現在它們已納入 502 的 $paths，就不該再豁免 —— 否則哪天它們沒被
    # stage，這支又會安靜放過去。只保留 _inbox/（仍是 Boss 的原始對帳單）。
    # 2026-10-16：Boss 端 trades_all/sync_state 是「502 所有物被本機程序改動」的常態
    # （writer --trades-only 與 write_sync_state 仍在 Boss 本機跑），不是孤兒 —— 它們由
    # Boss 端 sync_pull 的 $foreign 清單丟棄，不該天天觸發孤兒警告。
    $orphanExempt = if ($Owner -eq 'boss') { '^(_inbox/|trades_all\.xlsx$|sync_state\.json$)' } else { '^_inbox/' }
    $orphan = @($dirtyAll | Where-Object { $stagedNow -notcontains $_ -and
                $_ -notmatch $orphanExempt })
    if ($orphan.Count -gt 0) {
        Say "  [?] 下列已追蹤檔有改動但不屬於任何一方，不會被推送："
        $orphan | ForEach-Object { Say "      $_" }
        Say "      它們會讓 sync_pull 每天卡住（exit 2）。請把它們加進 `$paths 或 .gitignore。"
        Log "ORPHAN dirty-but-unstaged: $($orphan -join ', ')"
    }

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
    # 失敗時 Invoke-PullRebase 會自動 rebase --abort（見 _git_helper.ps1）
    $r = Invoke-PullRebase -AutoStash
    if (-not $r.Ok) { Report-PullFail $r; exit 1 }

    # push 失敗要重試：無人看管下靜默失敗＝對方拿到舊資料卻不知道
    $ok = $false
    for ($i = 1; $i -le 3; $i++) {
        if ((Try-Git push origin main) -eq 0) { $ok = $true; break }
        Say "  push 第 $i 次失敗，重試..."
        Start-Sleep -Seconds (5 * $i)
        $r = Invoke-PullRebase -AutoStash
        if (-not $r.Ok) { Report-PullFail $r; break }
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
