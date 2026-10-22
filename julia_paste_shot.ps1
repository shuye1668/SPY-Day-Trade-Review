# julia_paste_shot.ps1 — 把 Boss PC 經 git 送來的 TOS 對帳截圖貼進 Julia 母檔「個別識別法」（2026-10-22）
#
# 用法：
#   powershell -ExecutionPolicy Bypass -File julia_paste_shot.ps1 -ShotDate YYYY-MM-DD [-Target <xlsx>] [-DryRun]
#   ShotDate = Boss 匯出日（verify_shots\TOS_verify_<ShotDate>.png，含 _b/_c 分段檔），通常＝交易日隔天。
#
# 版面（照 2026-10-22 使用者貼的 Picture 227）：A 欄起、接在 A–V 區最低一張圖的正下方（間距 15pt），
#   多張分段檔直向依序排；寬度超過 W 欄左緣才等比例縮小，不放大。
# 安全：~$ 鎖檔或唯讀 → 不寫；冪等（已有同名圖 → SKIP）；寫前備份到 舊檔\；只動圖片，不碰儲存格。
# 結束碼：0=已貼或 SKIP（含找不到截圖）  6=檔案鎖住/唯讀
param(
    [Parameter(Mandatory = $true)][string]$ShotDate,
    [string]$Target = 'D:\fileserver\書業(新)\Julia\0更新捷徑\SPY INTRADAY\CS交易紀錄2026年起.xlsx',
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
    $names = @{}; foreach ($s in $ws.Shapes) { $names[$s.Name] = 1 }
    $todo = @($shots | Where-Object { -not $names.ContainsKey($_.BaseName) })
    if (-not $todo.Count) { Write-Output "SKIP：$($shots.BaseName -join ', ') 已在檔內（冪等）"; $wb.Close($false); exit 0 }

    $wLeft = [double]$ws.Range('W1').Left
    $maxW = $wLeft - [double]$ws.Range('A1').Left
    $bottom = 0.0
    foreach ($s in $ws.Shapes) { if ([double]$s.Left -lt $wLeft) { $b = [double]$s.Top + [double]$s.Height; if ($b -gt $bottom) { $bottom = $b } } }
    $top = $bottom + 15

    foreach ($f in $todo) {
        Copy-Item -LiteralPath $f.FullName -Destination $imgDir -Force
        $pic = Join-Path $imgDir $f.Name
        $sh = $ws.Shapes.AddPicture($pic, 0, -1, [single]0, [single]$top, -1, -1)
        $sh.LockAspectRatio = -1
        if ([double]$sh.Width -gt $maxW) { $sh.Width = [single]$maxW }
        $sh.Name = $f.BaseName
        Write-Output ("貼圖 {0}：{1}:{2}  {3:N0}x{4:N0}" -f $f.BaseName, $sh.TopLeftCell.Address(0, 0), $sh.BottomRightCell.Address(0, 0), $sh.Width, $sh.Height)
        $top = [double]$sh.Top + [double]$sh.Height + 15
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
