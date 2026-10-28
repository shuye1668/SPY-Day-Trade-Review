# julia_paste_shot.ps1 — 把 Boss PC 經 git 送來的 TOS 對帳截圖貼進 Julia 母檔「個別識別法」（2026-10-22）
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File julia_paste_shot.ps1 -ShotDate YYYY-MM-DD [-Target <xlsx>] [-Realign] [-DryRun]
#   ShotDate = Boss 匯出日（verify_shots\TOS_verify_<ShotDate>.png，含 _b/_c 分段檔），通常＝交易日隔天。
#   -Realign：已貼過的同名圖不 SKIP，改為依下面規則重新對齊（修正舊版貼歪的圖）。
#
# 版面（2026-10-28 使用者要求「縮小到跟右邊的交易檔對齊」，截圖是用來對照驗證輸入的交易）：
#   對應區塊＝W 欄最後一個「YYYY.MM.DD／損益」標頭、日期早於 ShotDate 者（10/28 匯出 → 10/27 區塊）；
#   圖頂＝該區塊標頭列頂、A 欄起；等比例縮小到「區塊標頭列～區塊最後一列」的高度內、且不超過 W 欄左緣，不放大。
#   多張分段檔在區塊高度內直向依序排。（照使用者 10/15 Picture 224、10/19 Picture 227 的擺法）
# 安全：~$ 鎖檔或唯讀 → 不寫；冪等（已有同名圖 → SKIP，-Realign 除外）；寫前備份到 舊檔\；只動圖片，不碰儲存格。
# 結束碼：0=已貼或 SKIP（含找不到截圖）  6=檔案鎖住/唯讀  7=找不到對應的交易區塊（未貼）
param(
    [Parameter(Mandatory = $true)][string]$ShotDate,
    [string]$Target = 'D:\fileserver\書業(新)\Julia\0更新捷徑\SPY INTRADAY\CS交易紀錄2026年起.xlsx',
    [switch]$Realign,
    [switch]$DryRun
)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$here = Split-Path -Parent $MyInvocation.MyCommand.Path
$dir = Split-Path -Parent $Target; $leaf = Split-Path -Leaf $Target
$imgDir = Join-Path $dir '圖片檔'

$shots = @(Get-ChildItem -LiteralPath (Join-Path $here 'verify_shots') -Filter "TOS_verify_$ShotDate*.png" -ErrorAction SilentlyContinue | Sort-Object Name)
if (-not $shots.Count) { Write-Output "SKIP：無 TOS_verify_$ShotDate*.png（無當日驗證截圖）"; exit 0 }
if (Test-Path -LiteralPath (Join-Path $dir ('~$' + $leaf))) { Write-Output "LOCKED(~$)：$Target 有人開著，未貼圖"; exit 6 }

$app = New-Object -ComObject Excel.Application
$app.Visible = $false
$app.DisplayAlerts = $false
try {
    $wb = $app.Workbooks.Open($Target)
    if ($wb.ReadOnly) { Write-Output 'LOCKED：檔案唯讀（別人開著），未貼圖'; $wb.Close($false); exit 6 }
    $ws = $wb.Worksheets.Item('個別識別法')
    $existing = @{}; foreach ($s in $ws.Shapes) { $existing[$s.Name] = $s }
    $todo = @($shots | Where-Object { $Realign -or -not $existing.ContainsKey($_.BaseName) })
    if (-not $todo.Count) { Write-Output "SKIP：$($shots.BaseName -join ', ') 已在檔內（冪等）"; $wb.Close($false); exit 0 }

    # 對應區塊：W 欄最後一個日期早於 ShotDate 的「YYYY.MM.DD／損益」標頭
    $shotDot = $ShotDate.Replace('-', '.')
    $last = $ws.Cells($ws.Rows.Count, 23).End(-4162).Row
    $hdr = 0; $nextHdr = 0
    for ($r = $last; $r -ge 1 -and $r -gt $last - 2000; $r--) {
        $w = [string]$ws.Cells($r, 23).Text
        if ($w -match '^\d{4}\.\d{2}\.\d{2}$' -and [string]$ws.Cells($r, 25).Text -eq '損益') {
            if ($w -lt $shotDot) { $hdr = $r; break }
            $nextHdr = $r
        }
    }
    if (-not $hdr) { Write-Output "未貼：W 欄找不到 $ShotDate 之前的交易區塊標頭（先跑 julia_block_writer）"; $wb.Close($false); exit 7 }
    # 區塊最後一列：下一個區塊標頭（或 W 欄最末列）往上第一個 W:AB 有值的列
    $end = if ($nextHdr) { $nextHdr - 1 } else { $last }
    while ($end -gt $hdr -and $app.WorksheetFunction.CountA($ws.Range("W${end}:AB${end}")) -eq 0) { $end-- }
    $top = [double]$ws.Cells($hdr, 1).Top
    $span = [double]$ws.Cells($end + 1, 1).Top - $top
    $wLeft = [double]$ws.Range('W1').Left
    $maxW = $wLeft - [double]$ws.Range('A1').Left - 4
    Write-Output ("對應區塊 W{0}:AB{1}（{2}），高 {3:N0}pt" -f $hdr, $end, [string]$ws.Cells($hdr, 23).Text, $span)

    $gap = 4.0
    $hEach = ($span - $gap * ($todo.Count - 1)) / $todo.Count
    foreach ($f in $todo) {
        if ($existing.ContainsKey($f.BaseName)) {
            $sh = $existing[$f.BaseName]
            $sh.LockAspectRatio = -1
            $sh.ScaleHeight(1, -1); $sh.ScaleWidth(1, -1)    # 回原尺寸再縮（-1 = 相對原圖）
        } else {
            Copy-Item -LiteralPath $f.FullName -Destination $imgDir -Force
            $pic = Join-Path $imgDir $f.Name
            $sh = $ws.Shapes.AddPicture($pic, 0, -1, [single]0, [single]$top, -1, -1)
            $sh.LockAspectRatio = -1
            $sh.Name = $f.BaseName
        }
        $k = [math]::Min(1.0, [math]::Min($hEach / [double]$sh.Height, $maxW / [double]$sh.Width))
        $sh.Height = [single]([double]$sh.Height * $k)
        $sh.Left = [single]0; $sh.Top = [single]$top
        Write-Output ("{0} {1}：{2}:{3}  {4:N0}x{5:N0}（{6:P0}）" -f $(if ($existing.ContainsKey($f.BaseName)) { '重新對齊' } else { '貼圖' }),
            $f.BaseName, $sh.TopLeftCell.Address(0, 0), $sh.BottomRightCell.Address(0, 0), $sh.Width, $sh.Height, $k)
        $top = [double]$sh.Top + [double]$sh.Height + $gap
    }

    if ($DryRun) { Write-Output 'DRY-RUN：未存檔'; $wb.Close($false); exit 0 }
    $bk = Join-Path $dir ('舊檔\' + [IO.Path]::GetFileNameWithoutExtension($leaf) + '_backup_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '_shot.xlsx')
    Copy-Item -LiteralPath $Target -Destination $bk
    Write-Output "備份：$bk"
    $wb.Save()
    $wb.Close($false)
    Write-Output '✅ SAVED'
} finally {
    $app.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}
exit 0
