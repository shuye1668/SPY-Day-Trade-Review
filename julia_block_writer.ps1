# julia_block_writer.ps1 — Julia 母檔「個別識別法」每日區塊自動寫入（2026-10-20）
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File julia_block_writer.ps1 -Csv <_inbox\YYYY-MM-DD.csv> -Date YYYY-MM-DD [-Target <xlsx>] [-DryRun]
#   截圖來源：先依 MIGRATION.md §4.1 轉成同格式 CSV，再用本腳本。
#
# 內容全部來自 julia_block_spec.py（＝spy_daytrade_engine 引擎輸出），版面照 2026.10.16 使用者範本：
#   標頭  W=日期  X=券商開盤 BAL  Y='損益'  Z='成交價' AA='成交時間' AB='多空'
#   每腿  W=淨額(買=數值；賣='=毛額-手續費')  X=餘額鏈  Y=損益(當沖出場=W進+W出；跨日平倉=W平+W原留倉列)
#         Z=成交價  AA=台灣成交時間  AB=多單/空單/留倉(多)/平倉(多)
#   合計 / 核驗（核驗=X末-X頭-留倉腿W-平倉腿W+平倉腿Y，必等於合計）
#   當沖彙總四行（有當沖才寫）、平倉彙總四行（有跨日平倉才寫）、累積留倉區塊（收盤仍有留倉才寫）
#   最下方總結行：MM/DD(周X)SPY當沖 N股 損益 ±X.XX ±X.XX%（平倉/留倉行依 §5）
# 格式：從最近一個有「多空」表頭的區塊逐列 PasteSpecial(格式)，鏈式繼承。
# 安全：絕不用 openpyxl（188 張圖）；~$ 鎖檔或唯讀 → 不寫；冪等（W 欄已有該日標頭 → 跳過）；
#       寫前備份；Calculate 回讀核驗（合計=核驗=引擎Σpnl、末腿X=收盤BAL）全過才 Save，否則不存。
# 結束碼：0=寫入或冪等跳過  3=引擎擋下  5=Excel 2010 不相容  6=檔案鎖住/唯讀  7=結構異常  8=核驗失敗
param(
    [Parameter(Mandatory = $true)][string]$Csv,
    [Parameter(Mandatory = $true)][string]$Date,
    [string]$Target = 'D:\fileserver\書業(新)\Julia\0更新捷徑\SPY INTRADAY\CS交易紀錄2026年起.xlsx',
    [switch]$DryRun,
    # 2026-10-30「不要擋，發現錯誤警示就好」：引擎原因／核驗不過都照寫，區塊下方加紅字「⚠️待修正」列，結束碼 9
    [switch]$Flag
)
$ErrorActionPreference = 'Stop'
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = 'C:\Users\server\AppData\Local\Python\pythoncore-3.14-64\python.exe'

[Console]::OutputEncoding = [Text.Encoding]::UTF8
$specArgs = @($Csv, '--date', $Date); if ($Flag) { $specArgs += '--flag' }
$json = (& $py (Join-Path $here 'julia_block_spec.py') @specArgs) | Out-String
$spec = $json | ConvertFrom-Json
if (-not $spec.ok) { Write-Output "🔴 引擎擋下，不寫 Julia：$($spec.reasons -join ' / ')"; exit 3 }
$legs = @($spec.legs)
if ($legs.Count -eq 0) { Write-Output "🔴 沒有可寫的腿：$($spec.reasons -join ' / ')"; exit 3 }
$flags = @($spec.flags | Where-Object { $_ })
Write-Output ("spec {0}: {1} 腿, 當沖 {2}, 平倉 {3}, 開 {4} 收 {5}" -f $spec.dot, $legs.Count, $spec.day_pnl, $spec.close_pnl, $spec.open_bal, $spec.close_bal)

$dir = Split-Path -Parent $Target; $leaf = Split-Path -Leaf $Target
if (Test-Path -LiteralPath (Join-Path $dir ('~$' + $leaf))) { Write-Output "LOCKED(~$)：$Target 有人開著，未寫入"; exit 6 }

. C:\XlsbUpdater\excel_env.ps1
$cx = Test-Excel2010Compat $Target
if ($cx.Skip) { Write-Output $cx.Line; exit 5 }

function Num($v) { if ($v -is [double] -or $v -is [int] -or $v -is [decimal]) { return [double]$v }; return $null }

$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
$code = 0
try {
    $wb = $app.Workbooks.Open($Target)
    if ($wb.ReadOnly) { Write-Output 'LOCKED：檔案唯讀（別人開著），未寫入'; $wb.Close($false); exit 6 }
    $ws = $wb.Worksheets.Item('個別識別法')

    # ── 定位 ───────────────────────────────────────────────────────────
    $last = 0
    for ($c = 23; $c -le 28; $c++) { $r = $ws.Cells($ws.Rows.Count, $c).End(-4162).Row; if ($r -gt $last) { $last = $r } }
    $scan0 = [math]::Max(1, $last - 600)
    $aW = $ws.Range("W${scan0}:W$last").Value2; $aX = $ws.Range("X${scan0}:X$last").Value2
    $aY = $ws.Range("Y${scan0}:Y$last").Value2; $aAB = $ws.Range("AB${scan0}:AB$last").Value2
    $n = $last - $scan0 + 1
    function V($arr, $row) { return $arr[($row - $scan0 + 1), 1] }

    for ($i = 1; $i -le $n; $i++) { if ([string]$aW[$i, 1] -eq $spec.dot) { Write-Output "SKIP：W$($scan0 + $i - 1) 已有 $($spec.dot)（冪等，不重寫）"; $wb.Close($false); exit 0 } }
    $found = $ws.Range("W1:W$last").Find($spec.dot, [Type]::Missing, -4163, 1)
    if ($found) { Write-Output "SKIP：$($found.Address()) 已有 $($spec.dot)（冪等，不重寫）"; $wb.Close($false); exit 0 }

    # 範本區塊：最近一個 AB='多空' 的標頭
    $tpl = @{}
    for ($r = $last; $r -ge $scan0; $r--) { if ([string](V $aAB $r) -eq '多空' -and [string](V $aY $r) -eq '損益') { $tpl.hdr = $r; break } }
    if (-not $tpl.hdr) { Write-Output 'ABORT：找不到範本區塊（AB=多空 的標頭）'; $wb.Close($false); exit 7 }
    for ($r = $tpl.hdr + 1; $r -le $last; $r++) {
        $w = [string](V $aW $r); $x = [string](V $aX $r); $ab = [string](V $aAB $r)
        if ($r -gt $tpl.hdr + 1 -and [string](V $aY $r) -eq '損益') { break }
        if ($ab -eq '多單' -and -not $tpl.day) { $tpl.day = $r }
        if ($ab -eq '留倉(多)' -and -not $tpl.hold) { $tpl.hold = $r }
        if ($x -eq '合計 ' -and -not $tpl.tot) { $tpl.tot = $r }
        if ($x -eq '核驗' -and -not $tpl.chk) { $tpl.chk = $r }
        if ($w.EndsWith('當沖') -and -not $tpl.sum) { $tpl.sum = $r }
        if ($w.StartsWith('累積留倉') -and -not $tpl.hh) { $tpl.hh = $r; $tpl.hr = $r + 1 }
    }
    # 範本區塊沒有的列型（例如純當沖日沒有留倉列）→ 往上找最近一次出現的同類列
    for ($r = $last; $r -ge $scan0 -and (-not $tpl.day -or -not $tpl.hold -or -not $tpl.hh -or -not $tpl.sum); $r--) {
        $ab = [string](V $aAB $r); $w = [string](V $aW $r)
        if (-not $tpl.day -and $ab -eq '多單') { $tpl.day = $r }
        if (-not $tpl.hold -and $ab -eq '留倉(多)') { $tpl.hold = $r }
        if (-not $tpl.hh -and $w.StartsWith('累積留倉')) { $tpl.hh = $r; $tpl.hr = $r + 1 }
        if (-not $tpl.sum -and $w.EndsWith('當沖')) { $tpl.sum = $r }
    }
    if (-not $tpl.day) { $tpl.day = $tpl.hold }; if (-not $tpl.hold) { $tpl.hold = $tpl.day }
    if (-not $tpl.day -or -not $tpl.tot -or -not $tpl.chk -or -not $tpl.sum) { Write-Output 'ABORT：範本列不齊（多單/合計/核驗/當沖彙總）'; $wb.Close($false); exit 7 }
    # 標籤樣式明確套用（2026.10.16 範本：多單=紅、留倉(多)=橘粗體）
    $labelStyle = @{ '多單' = @(255, $false); '空單' = @(32768, $false); '留倉(多)' = @(1137094, $true); '平倉(多)' = @(12611584, $true) }
    function SetLabel($row, $text) { $c = $ws.Cells($row, 28); $c.Value2 = $text; $c.Font.Color = $labelStyle[$text][0]; $c.Font.Bold = $labelStyle[$text][1]; $c.HorizontalAlignment = -4108 }
    Write-Output ("範本區塊 W{0}: day={1} hold={2} tot={3} chk={4} sum={5} 留倉={6}" -f $tpl.hdr, $tpl.day, $tpl.hold, $tpl.tot, $tpl.chk, $tpl.sum, $tpl.hh)

    function FindHoldRow($date, $net) {
        # 前面區塊裡 AB='留倉(多)' 且 W≈淨額、所屬區塊標頭=留倉日 的列
        $dot = $date.Replace('-', '.')
        $hdrDot = $null
        $cands = @()
        for ($r = $scan0; $r -le $last; $r++) {
            if ([string](V $aY $r) -eq '損益') { $hdrDot = [string](V $aW $r) }
            $wv = Num (V $aW $r)
            if ($null -eq $wv) { try { $wv = [double]$ws.Cells($r, 23).Value2 } catch { $wv = $null } }
            if ($hdrDot -eq $dot -and [string](V $aAB $r) -eq '留倉(多)' -and $null -ne $wv -and [math]::Abs($wv - $net) -lt 0.005) { $cands += $r }
        }
        if ($cands.Count -ge 1) { return $cands[0] }
        return $null
    }

    $base = $last + 3
    $hr = $base
    $copyFmt = { param($from, $to) $ws.Range("W${from}:AB${from}").Copy() | Out-Null; $ws.Range("W${to}:AB${to}").PasteSpecial(-4122) | Out-Null }

    # ── 標頭 ───────────────────────────────────────────────────────────
    & $copyFmt $tpl.hdr $hr
    $ws.Cells($hr, 23).NumberFormat = '@'; $ws.Cells($hr, 23).Value2 = $spec.dot
    $ws.Cells($hr, 24).Value2 = [double]$spec.open_bal
    $ws.Cells($hr, 25).Value2 = '損益'
    $ws.Cells($hr, 26).Value2 = '成交價'; $ws.Cells($hr, 27).Value2 = '成交時間'; $ws.Cells($hr, 28).Value2 = '多空'

    # ── 每腿 ───────────────────────────────────────────────────────────
    $rows = @(); $holdRows = @(); $closeRows = @(); $dayRows = @(); $entryRows = @()
    for ($i = 0; $i -lt $legs.Count; $i++) {
        $L = $legs[$i]; $r = $hr + 1 + $i; $rows += $r
        if ($L.kind -eq 'hold') { & $copyFmt $tpl.hold $r } else { & $copyFmt $tpl.day $r }
        if ($L.side -eq 'SOLD' -and [double]$L.fee -gt 0) { $ws.Cells($r, 23).Formula = ('={0}-{1}' -f $L.gross, $L.fee) }
        else { $ws.Cells($r, 23).Value2 = [double]$L.net }
        $prev = if ($i -eq 0) { $hr } else { $r - 1 }
        $ws.Cells($r, 24).Formula = "=+X$prev+W$r"
        $ws.Cells($r, 26).Value2 = [double]$L.price
        $ws.Cells($r, 27).NumberFormat = '@'; $ws.Cells($r, 27).Value2 = [string]$L.time_tw
        switch ($L.kind) {
            'day' {
                $dayRows += $r
                $lab = if ($L.type_tag -eq '空' -or ($null -eq $L.pnl -and $L.side -eq 'SOLD')) { '空單' } else { '多單' }
                if ($null -ne $L.pair) { $ws.Cells($r, 25).Formula = ('=W{0}+W{1}' -f ($hr + 1 + [int]$L.pair), $r) }
                SetLabel $r $lab
            }
            'hold' { $holdRows += $r; SetLabel $r '留倉(多)' }
            'close' {
                $closeRows += $r
                & $copyFmt $tpl.hold $r
                $ws.Cells($r, 27).NumberFormat = '@'
                if ($L.side -eq 'SOLD' -and [double]$L.fee -gt 0) { $ws.Cells($r, 23).Formula = ('={0}-{1}' -f $L.gross, $L.fee) } else { $ws.Cells($r, 23).Value2 = [double]$L.net }
                $ws.Cells($r, 24).Formula = "=+X$prev+W$r"; $ws.Cells($r, 26).Value2 = [double]$L.price; $ws.Cells($r, 27).Value2 = [string]$L.time_tw
                $er = FindHoldRow $L.entry.date ([double]$L.entry.net)
                if ($er) { $ws.Cells($r, 25).Formula = "=+W$r+W$er"; $entryRows += "W$er" }
                else { $ws.Cells($r, 25).Formula = ('=+W{0}{1:+0.00;-0.00}' -f $r, [double]$L.entry.net); $entryRows += ('{0}' -f [double]$L.entry.net) }
                SetLabel $r '平倉(多)'
            }
        }
    }
    $r1 = $rows[0]; $rN = $rows[-1]
    $ws.Range("Z${r1}:AA${rN}").HorizontalAlignment = -4108   # 成交價／成交時間置中（2026-10-20 使用者要求）

    # ── 合計 / 核驗 ────────────────────────────────────────────────────
    $tot = $rN + 1; $chk = $rN + 2
    & $copyFmt $tpl.tot $tot; & $copyFmt $tpl.chk $chk
    $ws.Cells($tot, 24).Value2 = '合計 '; $ws.Cells($tot, 25).Formula = "=SUM(Y${r1}:Y${rN})"
    $chkF = "=+X$rN-X$hr"
    foreach ($h in $holdRows) { $chkF += "-W$h" }
    foreach ($c in $closeRows) { $chkF += "-W$c+Y$c" }
    $ws.Cells($chk, 24).Value2 = '核驗'; $ws.Cells($chk, 25).Formula = $chkF
    $cur = $chk + 2

    function Contig($list) { if ($list.Count -le 1) { return $true }; for ($k = 1; $k -lt $list.Count; $k++) { if ($list[$k] -ne $list[$k - 1] + 1) { return $false } }; return $true }

    # ── 當沖彙總 ───────────────────────────────────────────────────────
    $daySum = $null
    $dayExits = @($legs | Where-Object { $_.kind -eq 'day' -and $null -ne $_.pnl })
    if ($dayExits.Count -gt 0) {
        $daySum = $cur
        for ($k = 0; $k -lt 4; $k++) { & $copyFmt ($tpl.sum + $k) ($cur + $k) }
        $ws.Cells($cur, 25).NumberFormat = $ws.Cells($cur + 1, 25).NumberFormat   # 買列也用千分位（2026.10.19 使用者範本）
        if (Contig $dayRows) {
            $a = $dayRows[0]; $b = $dayRows[-1]
            $buyF = "=SUMIF(W${a}:W${b},`"<0`")"; $sellF = "=SUMIF(W${a}:W${b},`">0`")"; $cntF = "=`"SPY `" & COUNT(Y${a}:Y${b}) * 100 & `"股`""
        } else {
            $neg = @(); $pos = @()
            for ($i = 0; $i -lt $legs.Count; $i++) { if ($legs[$i].kind -eq 'day') { if ([double]$legs[$i].net -lt 0) { $neg += "W$($rows[$i])" } else { $pos += "W$($rows[$i])" } } }
            $buyF = '=SUM(' + ($neg -join ',') + ')'; $sellF = '=SUM(' + ($pos -join ',') + ')'
            $cntF = "=`"SPY `" & COUNT(" + (($dayRows | ForEach-Object { "Y$_" }) -join ',') + ") * 100 & `"股`""
        }
        $ws.Cells($cur, 23).NumberFormat = '@'; $ws.Cells($cur, 23).Value2 = "$($spec.dot)當沖"; $ws.Cells($cur, 24).Value2 = '買'; $ws.Cells($cur, 25).Formula = $buyF
        $ws.Cells($cur + 1, 23).Formula = $cntF; $ws.Cells($cur + 1, 24).Value2 = '賣'; $ws.Cells($cur + 1, 25).Formula = $sellF
        $ws.Cells($cur + 2, 24).Value2 = '當日損益'; $ws.Cells($cur + 2, 25).Formula = "=+Y$cur+Y$($cur + 1)"
        $ws.Cells($cur + 3, 24).Value2 = '當日損益 %'; $ws.Cells($cur + 3, 25).Formula = "=Y$($cur + 2)/ABS(Y$cur)"
        $cur += 5
    }

    # ── 平倉彙總（2026.10.19 使用者範本）─────────────────────────────────
    $closeSum = $null
    if ($closeRows.Count -gt 0) {
        $closeSum = $cur
        for ($k = 0; $k -lt 4; $k++) { & $copyFmt ($tpl.sum + $k) ($cur + $k) }
        $ws.Cells($cur, 25).NumberFormat = $ws.Cells($cur + 1, 25).NumberFormat   # 買列也用千分位（2026.10.19 使用者範本）
        $ws.Cells($cur, 23).NumberFormat = '@'; $ws.Cells($cur, 23).Value2 = "$($spec.dot) 平倉"; $ws.Cells($cur, 24).Value2 = '買'
        $ws.Cells($cur, 25).Formula = '=+' + ($entryRows -join '+')
        $ws.Cells($cur + 1, 23).NumberFormat = '@'; $ws.Cells($cur + 1, 23).Value2 = "SPY $($closeRows.Count * 100)股"; $ws.Cells($cur + 1, 24).Value2 = '賣'
        $ws.Cells($cur + 1, 25).Formula = '=+' + (($closeRows | ForEach-Object { "W$_" }) -join '+')
        $ws.Cells($cur + 2, 24).Value2 = '當日損益'; $ws.Cells($cur + 2, 25).Formula = "=+Y$cur+Y$($cur + 1)"
        $ws.Cells($cur + 3, 24).Value2 = '當日損益 %'; $ws.Cells($cur + 3, 25).Formula = "=Y$($cur + 2)/ABS(Y$cur)"
        $cur += 5
    }

    # ── 累積留倉 ───────────────────────────────────────────────────────
    $holds = @($spec.holds_after)
    if ($holds.Count -gt 0) {
        $tot100 = ($holds | Measure-Object -Property qty -Sum).Sum
        if ($tpl.hh) { & $copyFmt $tpl.hh $cur }
        $ws.Cells($cur, 23).NumberFormat = '@'; $ws.Cells($cur, 23).Value2 = "累積留倉   $($tot100)股"
        $k = 1
        foreach ($h in $holds) {
            $rr = $cur + $k
            if ($tpl.hr) { & $copyFmt $tpl.hr $rr }
            $ws.Cells($rr, 23).NumberFormat = '@'; $ws.Cells($rr, 23).Value2 = ('{0} 留倉' -f $h.date.Replace('-', '.'))
            $ws.Cells($rr, 24).Value2 = "買 $($h.qty)股"
            $src = $null
            for ($i = 0; $i -lt $legs.Count; $i++) { if ($legs[$i].kind -eq 'hold' -and $legs[$i].time_edt -eq $h.time -and $h.date -eq $spec.date) { $src = $rows[$i] } }
            if (-not $src) { $src = FindHoldRow $h.date ([double]$h.net) }
            if ($src) { $ws.Cells($rr, 25).Formula = "=+W$src" } else { $ws.Cells($rr, 25).Value2 = [double]$h.net }
            $k++
        }
        $cur += $k + 1
    }

    # ── 總結行 ─────────────────────────────────────────────────────────
    $sumLines = @($spec.summary_lines)
    for ($k = 0; $k -lt $sumLines.Count; $k++) { $ws.Cells($cur + $k, 23).NumberFormat = '@'; $ws.Cells($cur + $k, 23).Value2 = [string]$sumLines[$k] }

    if ($ws.Columns(26).ColumnWidth -lt 8) { $ws.Columns(26).ColumnWidth = 10 }

    # ── 回讀核驗 ───────────────────────────────────────────────────────
    $app.Calculate()
    $errs = @()
    $vt = [double]$ws.Cells($tot, 25).Value2; $vc = [double]$ws.Cells($chk, 25).Value2; $vx = [double]$ws.Cells($rN, 24).Value2
    if ([math]::Abs($vt - [double]$spec.total_pnl) -gt 0.005) { $errs += "合計 $vt ≠ 引擎Σpnl $($spec.total_pnl)" }
    if ([math]::Abs($vc - [double]$spec.total_pnl) -gt 0.005) { $errs += "核驗 $vc ≠ 引擎Σpnl $($spec.total_pnl)" }
    if ([math]::Abs($vx - [double]$spec.close_bal) -gt 0.005) { $errs += "末腿X $vx ≠ 收盤BAL $($spec.close_bal)" }
    if ($daySum) { $v = [double]$ws.Cells($daySum + 2, 25).Value2; if ([math]::Abs($v - [double]$spec.day_pnl) -gt 0.005) { $errs += "當沖彙總 $v ≠ $($spec.day_pnl)" } }
    if ($closeSum) { $v = [double]$ws.Cells($closeSum + 2, 25).Value2; if ([math]::Abs($v - [double]$spec.close_pnl) -gt 0.005) { $errs += "平倉彙總 $v ≠ $($spec.close_pnl)" } }
    for ($r = $hr; $r -lt $cur; $r++) { for ($c = 23; $c -le 28; $c++) { if ($ws.Cells($r, $c).Value2 -is [int]) { $errs += "錯誤值 $($ws.Cells($r,$c).Address(0,0))=$($ws.Cells($r,$c).Text)" } } }
    Write-Output ("寫入 W{0}:AB{1}｜合計={2} 核驗={3} 末腿X={4}" -f $hr, ($cur + $sumLines.Count - 1), $vt, $vc, $vx)
    $sumLines | ForEach-Object { Write-Output "  總結行：$_" }
    Write-Output ("RESULT_JSON " + (@{ header = $hr; legs = @($r1, $rN); day_summary = $daySum; close_summary = $closeSum; last = ($cur + $sumLines.Count - 1) } | ConvertTo-Json -Compress))
    if ($errs.Count -and -not $Flag) { Write-Output ("🔴 核驗失敗，不存檔：" + ($errs -join ' / ')); $wb.Close($false); exit 8 }
    $flags += $errs
    if ($flags.Count) {
        $fr = $cur + $sumLines.Count
        $c = $ws.Cells($fr, 23); $c.NumberFormat = '@'; $c.Value2 = '⚠️待修正：' + ($flags -join '；')
        $c.Font.Color = 255; $c.Font.Bold = $true
        Write-Output ("⚠️ 待修正（W$fr）：" + ($flags -join ' / '))
        $code = 9
    }

    if ($DryRun) { Write-Output 'DRY-RUN：核驗通過，未存檔'; $wb.Close($false); exit 0 }
    $bk = Join-Path $dir ('舊檔\' + [IO.Path]::GetFileNameWithoutExtension($leaf) + '_backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.xlsx')
    Copy-Item -LiteralPath $Target -Destination $bk
    Write-Output "備份：$bk"
    # 存檔前把視窗捲到新區塊，使用者一打開就看到今天的帳（2026-10-20 使用者回報「打開還是舊畫面」）
    try {
        $ws.Activate()
        $win = $wb.Windows.Item(1)
        $win.ScrollColumn = [math]::Max(1, $ws.Range('Q1').Column)
        $win.ScrollRow = [math]::Max(1, $hr - 3)
        $ws.Range("W$hr").Select() | Out-Null
    } catch { Write-Output "（捲動視窗失敗，不影響帳：$($_.Exception.Message)）" }
    $wb.Save()
    $wb.Close($false)
    Write-Output '✅ SAVED'
} finally {
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
exit $code
