# boss_screen_state.ps1 — Boss PC 採集「隱形化」：開工前記住畫面、收尾原樣還原（2026-10-20）
#
# 使用者要求：Boss PC 截圖/匯出之後要回到原本的螢幕畫面，讓使用者看不出來有動過、不影響使用。
#
# 用法（Boss SKILL 每個 call 都要重新 dot-source；每個 PowerShell call 是新 process）：
#   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
#   . C:\TradeReview\boss_screen_state.ps1
#   Wait-UserIdle -MinIdleSec 120 -MaxWaitMin 30      # 使用者正在操作就先等（回傳 $true=已閒置）
#   Save-ScreenState  -Path C:\TradeReview_Boss\logs\screen_state_$today.json -TosTab Charts
#   ...（原本的採集流程）...
#   Restore-ScreenState -Path C:\TradeReview_Boss\logs\screen_state_$today.json
#
# 記住／還原的東西：
#   - 前景視窗（hwnd、標題、程式）→ 還原時搶回前景並驗證（AttachThreadInput + Alt 解鎖）
#   - 所有可見頂層視窗的 Z 順序 → 依原順序重排（不啟用、不改大小）
#   - TOS 視窗的 WINDOWPLACEMENT（最小化／最大化／一般＋位置）→ 原樣還原
#   - 滑鼠游標位置
#   - TOS 原本停的主頁籤名稱（由 Claude 從步驟 2 截圖判斷後傳入 -TosTab；切回頁籤仍由 SKILL 步驟 10 點擊）
# 還原後另起一個隱藏的 PowerShell，在 5 秒與 15 秒後再各還原一次前景
# （Claude Desktop 在工具呼叫結束後常會自己跳回前面）。
#
# 2026-10-22 加（使用者要求「不管有沒有完成都要把自己縮小，卡住也不要擋在螢幕上」）：
#   - Restore-ScreenState 一律把 Claude Desktop 視窗最小化（不搶前景），5/15 秒補還原時再縮一次；
#     畫面狀態檔不存在時也照縮。
#   - 執行旗標 C:\TradeReview_Boss\logs\run_flag_<日期>.json：Wait-UserIdle 結束時記 started，
#     Restore-ScreenState 記 finished。
#   - 守門排程 Boss_Screen_Guard（Wait-UserIdle／Restore-ScreenState 第一次在 Boss PC 跑時自動註冊，
#     只限本機使用者、互動式）：週二～週六 06:10～09:30 每 10 分鐘跑 boss_screen_guard.ps1——
#     started 超過 30 分鐘還沒 finished（任務卡住）→ 還原畫面＋縮小 Claude；使用者 3 分鐘內有操作就不動。
#   以上只在 Boss PC（有 C:\TradeReview_Boss）生效；Boss SKILL 文字不用改。
#
# 安全：只做視窗管理（不點擊、不打字、不關任何視窗）。所有函式失敗都不 throw，回傳狀態物件。

if (-not ('BossWin' -as [type])) {
Add-Type @'
using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
public static class BossWin {
    [StructLayout(LayoutKind.Sequential)] public struct POINT { public int X; public int Y; }
    [StructLayout(LayoutKind.Sequential)] public struct RECT { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)] public struct WINDOWPLACEMENT {
        public int length; public int flags; public int showCmd;
        public POINT ptMinPosition; public POINT ptMaxPosition; public RECT rcNormalPosition; }
    [StructLayout(LayoutKind.Sequential)] public struct LASTINPUTINFO { public uint cbSize; public uint dwTime; }
    public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
    [DllImport("user32.dll")] public static extern bool EnumWindows(EnumWindowsProc cb, IntPtr lParam);
    [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool BringWindowToTop(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindow(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
    [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
    [DllImport("user32.dll")] public static extern int GetWindowTextLength(IntPtr h);
    [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
    [DllImport("user32.dll")] public static extern bool GetWindowPlacement(IntPtr h, ref WINDOWPLACEMENT wp);
    [DllImport("user32.dll")] public static extern bool SetWindowPlacement(IntPtr h, ref WINDOWPLACEMENT wp);
    [DllImport("user32.dll")] public static extern bool SetWindowPos(IntPtr h, IntPtr after, int x, int y, int cx, int cy, uint flags);
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int cmd);
    [DllImport("user32.dll")] public static extern bool GetCursorPos(out POINT p);
    [DllImport("user32.dll")] public static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll")] public static extern bool GetLastInputInfo(ref LASTINPUTINFO lii);
    [DllImport("user32.dll")] public static extern bool AttachThreadInput(uint a, uint b, bool attach);
    [DllImport("kernel32.dll")] public static extern uint GetCurrentThreadId();
    [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
    [DllImport("user32.dll")] public static extern IntPtr GetWindow(IntPtr h, uint cmd);
    [DllImport("user32.dll")] public static extern int GetWindowLong(IntPtr h, int idx);

    public static string Title(IntPtr h) { int n = GetWindowTextLength(h); var sb = new StringBuilder(n + 1); GetWindowText(h, sb, sb.Capacity); return sb.ToString(); }
    public static uint IdleMs() { var l = new LASTINPUTINFO(); l.cbSize = (uint)Marshal.SizeOf(l); GetLastInputInfo(ref l); return (uint)Environment.TickCount - l.dwTime; }
    // 可見、有標題、非 tool window、無 owner 的頂層視窗，依 Z 順序（最上面在前）
    public static List<IntPtr> TopLevel() {
        var list = new List<IntPtr>();
        EnumWindows((h, p) => {
            if (!IsWindowVisible(h) || GetWindowTextLength(h) == 0) return true;
            if (GetWindow(h, 4) != IntPtr.Zero) return true;             // GW_OWNER
            int ex = GetWindowLong(h, -20);                               // GWL_EXSTYLE
            if ((ex & 0x80) != 0) return true;                            // WS_EX_TOOLWINDOW
            list.Add(h); return true; }, IntPtr.Zero);
        return list;
    }
    public static bool ForceForeground(IntPtr h) {
        if (!IsWindow(h)) return false;
        if (IsIconic(h)) ShowWindow(h, 9);                                // SW_RESTORE（原本就最小化的不會走到這）
        IntPtr fg = GetForegroundWindow();
        uint dummy; uint fgThread = GetWindowThreadProcessId(fg, out dummy);
        uint me = GetCurrentThreadId();
        bool attached = fgThread != me && AttachThreadInput(me, fgThread, true);
        keybd_event(0x12, 0, 0, UIntPtr.Zero); keybd_event(0x12, 0, 2, UIntPtr.Zero);   // Alt 按放：解除前景鎖
        BringWindowToTop(h); bool ok = SetForegroundWindow(h);
        if (attached) AttachThreadInput(me, fgThread, false);
        return GetForegroundWindow() == h;
    }
}
'@
}

function Get-IdleSeconds { [math]::Round([BossWin]::IdleMs() / 1000.0, 1) }

# ── 執行旗標／Claude 視窗／守門排程（2026-10-22）───────────────────────────
if (-not $script:BossLogDir) { $script:BossLogDir = 'C:\TradeReview_Boss\logs' }
$script:BossGuardTask = 'Boss_Screen_Guard'

function Test-IsBossPC { Test-Path -LiteralPath (Split-Path -Parent $script:BossLogDir) }

function Get-RunFlagPath { param([string]$Day = (Get-Date).ToString('yyyy-MM-dd')) Join-Path $script:BossLogDir "run_flag_$Day.json" }

function Read-RunFlag {
    param([string]$Day = (Get-Date).ToString('yyyy-MM-dd'))
    $p = Get-RunFlagPath $Day
    if (-not (Test-Path -LiteralPath $p)) { return $null }
    try { return Get-Content -LiteralPath $p -Raw -Encoding UTF8 | ConvertFrom-Json } catch { return $null }
}

function Set-RunFlag {
    # 把旗標的某個欄位（started / finished / guard_restored / guard_final）設成現在時間；不是 Boss PC 就不做
    param([Parameter(Mandatory = $true)][string]$Field, [string]$Value = (Get-Date).ToString('s'))
    if (-not (Test-IsBossPC)) { return }
    try {
        $f = Read-RunFlag
        if (-not $f) { $f = [pscustomobject]@{} }
        $f | Add-Member -NotePropertyName $Field -NotePropertyValue $Value -Force
        New-Item -ItemType Directory -Force -Path $script:BossLogDir | Out-Null
        $f | ConvertTo-Json | Set-Content -LiteralPath (Get-RunFlagPath) -Encoding UTF8
    } catch {}
}

function Get-ClaudeWindows {
    # Claude Desktop（程序名 claude）的可見頂層視窗
    $ids = @(Get-Process -Name claude -ErrorAction SilentlyContinue | ForEach-Object { [uint32]$_.Id })
    if (-not $ids.Count) { return @() }
    $out = @()
    foreach ($h in [BossWin]::TopLevel()) {
        $wpid = [uint32]0; [void][BossWin]::GetWindowThreadProcessId($h, [ref]$wpid)
        if ($ids -contains $wpid) { $out += $h }
    }
    return $out
}

function Minimize-ClaudeWindows {
    # 最小化 Claude Desktop 視窗（SW_SHOWMINNOACTIVE：不把焦點交給別的視窗，前景由呼叫端還原）
    param([switch]$WhatIf)
    $wins = @(Get-ClaudeWindows)
    $n = 0
    foreach ($h in $wins) {
        if (-not [BossWin]::IsIconic($h)) { if (-not $WhatIf) { [void][BossWin]::ShowWindow($h, 7) }; $n++ }
    }
    return "$n/$($wins.Count)"
}

function Install-BossScreenGuard {
    # 在 Boss PC 註冊守門排程（已存在就不動）。本機目前使用者、互動式、一般權限。
    param([string]$TaskName = $script:BossGuardTask, [string]$ScriptDir = 'C:\TradeReview', [switch]$Force)
    if (-not $Force -and -not (Test-IsBossPC)) { return 'not boss pc' }
    try {
        if (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue) { return 'exists' }
        $vbs = Join-Path $ScriptDir 'boss_run_hidden.vbs'
        $guard = Join-Path $ScriptDir 'boss_screen_guard.ps1'
        if (-not (Test-Path $vbs) -or -not (Test-Path $guard)) { return 'guard files missing' }
        $act = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument ('//B //Nologo "{0}" "{1}"' -f $vbs, $guard)
        $trg = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Tuesday, Wednesday, Thursday, Friday, Saturday -At '06:10'
        $trg.Repetition = (New-ScheduledTaskTrigger -Once -At '06:10' -RepetitionInterval (New-TimeSpan -Minutes 10) -RepetitionDuration (New-TimeSpan -Hours 3 -Minutes 20)).Repetition
        $pr = New-ScheduledTaskPrincipal -UserId ("{0}\{1}" -f $env:USERDOMAIN, $env:USERNAME) -LogonType Interactive -RunLevel Limited
        $st = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 5) -MultipleInstances IgnoreNew
        Register-ScheduledTask -TaskName $TaskName -Action $act -Trigger $trg -Principal $pr -Settings $st `
            -Description 'Boss 採集卡住時還原畫面＋縮小 Claude（boss_screen_state.ps1 自動註冊，2026-10-22）' | Out-Null
        return 'installed'
    } catch { return "install failed: $($_.Exception.Message)" }
}

function Wait-UserIdle {
    # 使用者最近 MinIdleSec 秒內有動鍵盤/滑鼠 → 每 15 秒再看一次，最多等 MaxWaitMin 分鐘。
    # 回傳 $true=已閒置可開工；$false=等到上限仍在使用（由 SKILL 決定：照跑或延後）。
    param([int]$MinIdleSec = 120, [int]$MaxWaitMin = 30)
    # 回傳前記執行旗標 started（守門排程用它判斷「跑太久＝卡住」），並確認守門排程已註冊
    $deadline = (Get-Date).AddMinutes($MaxWaitMin)
    while ($true) {
        $idle = Get-IdleSeconds
        if ($idle -ge $MinIdleSec) { Write-Host "user idle ${idle}s → OK"; Set-RunFlag started; [void](Install-BossScreenGuard); return $true }
        if ((Get-Date) -ge $deadline) { Write-Host "user still active (idle ${idle}s) after $MaxWaitMin min"; Set-RunFlag started; [void](Install-BossScreenGuard); return $false }
        Write-Host "user active (idle ${idle}s) → wait 15s"
        Start-Sleep -Seconds 15
    }
}

function Find-TosHwnd {
    foreach ($p in Get-Process -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -ne 0 }) {
        $t = $p.MainWindowTitle
        if ($p.ProcessName -match 'thinkorswim' -or $t -match 'thinkorswim' -or ($p.ProcessName -match '^javaw?$' -and $t -match 'Main@|thinkorswim')) { return $p.MainWindowHandle }
    }
    return [IntPtr]::Zero
}

function Get-Placement([IntPtr]$h) {
    $wp = New-Object BossWin+WINDOWPLACEMENT
    $wp.length = [Runtime.InteropServices.Marshal]::SizeOf($wp)
    [void][BossWin]::GetWindowPlacement($h, [ref]$wp)
    return $wp
}

function Save-ScreenState {
    param([Parameter(Mandatory = $true)][string]$Path, [string]$TosTab = '', [IntPtr]$TosHwnd = [IntPtr]::Zero)   # -TosHwnd 只供測試
    $fg = [BossWin]::GetForegroundWindow()
    $fgPid = [uint32]0; [void][BossWin]::GetWindowThreadProcessId($fg, [ref]$fgPid)
    $fgProc = try { (Get-Process -Id $fgPid -ErrorAction Stop).ProcessName } catch { '' }
    $pt = New-Object BossWin+POINT; [void][BossWin]::GetCursorPos([ref]$pt)
    $wins = @()
    foreach ($h in [BossWin]::TopLevel()) {
        $wp = Get-Placement $h
        $wins += [pscustomobject]@{ hwnd = [int64]$h; title = [BossWin]::Title($h); showCmd = $wp.showCmd }
    }
    $tos = if ($TosHwnd -ne [IntPtr]::Zero) { $TosHwnd } else { Find-TosHwnd }
    $tosState = $null
    if ($tos -ne [IntPtr]::Zero) {
        $wp = Get-Placement $tos
        $tosState = [pscustomobject]@{ hwnd = [int64]$tos; title = [BossWin]::Title($tos); showCmd = $wp.showCmd; flags = $wp.flags
            minX = $wp.ptMinPosition.X; minY = $wp.ptMinPosition.Y; maxX = $wp.ptMaxPosition.X; maxY = $wp.ptMaxPosition.Y
            left = $wp.rcNormalPosition.Left; top = $wp.rcNormalPosition.Top; right = $wp.rcNormalPosition.Right; bottom = $wp.rcNormalPosition.Bottom
            tab = $TosTab }
    }
    $state = [pscustomobject]@{
        saved_at = (Get-Date).ToString('s'); idle_sec = (Get-IdleSeconds)
        foreground = [pscustomobject]@{ hwnd = [int64]$fg; title = [BossWin]::Title($fg); process = $fgProc }
        cursor = [pscustomobject]@{ x = $pt.X; y = $pt.Y }
        zorder = $wins; tos = $tosState
    }
    New-Item -ItemType Directory -Force -Path (Split-Path -Parent $Path) | Out-Null
    $state | ConvertTo-Json -Depth 5 | Set-Content -Path $Path -Encoding UTF8
    $desc = if ($tosState) { "TOS showCmd=$($tosState.showCmd) tab='$TosTab'" } else { 'TOS not found' }
    Write-Host ("screen state saved: fg='{0}' ({1}), {2} windows, {3}, cursor=({4},{5})" -f $state.foreground.title, $fgProc, $wins.Count, $desc, $pt.X, $pt.Y)
    return $state
}

function Set-SavedTosTab {
    # 步驟 2 看完截圖後補記 TOS 原本停的主頁籤（Save 時 TOS 還沒被叫到前景，頁籤要從截圖判斷）
    param([Parameter(Mandatory = $true)][string]$Path, [Parameter(Mandatory = $true)][string]$Tab)
    $s = Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($s.tos) { $s.tos.tab = $Tab; $s | ConvertTo-Json -Depth 5 | Set-Content -Path $Path -Encoding UTF8; Write-Host "saved TOS tab = $Tab" }
}

function Get-SavedTosTab {
    param([Parameter(Mandatory = $true)][string]$Path)
    if (-not (Test-Path $Path)) { return '' }
    $s = Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($s.tos -and $s.tos.tab) { return [string]$s.tos.tab }
    return ''
}

function Restore-ScreenState {
    # -NoReassert：不啟動延遲補還原的背景程序（背景程序本身呼叫時用）
    param([Parameter(Mandatory = $true)][string]$Path, [switch]$NoReassert, [switch]$ForegroundOnly)
    $r = [ordered]@{ tos = 'skip'; zorder = 'skip'; foreground = $false; cursor = $false; claude = 'skip' }
    if (-not $ForegroundOnly) { Set-RunFlag finished; if (-not $NoReassert) { [void](Install-BossScreenGuard) } }
    if (-not (Test-Path $Path)) {
        # 沒記到原本畫面（例如步驟 1.5 前就失敗）：至少把 Claude 縮小，不擋在螢幕上
        $r.claude = Minimize-ClaudeWindows
        Write-Host "no state file: $Path → claude minimized $($r.claude)"
        return [pscustomobject]$r
    }
    $s = Get-Content -Path $Path -Raw -Encoding UTF8 | ConvertFrom-Json

    if (-not $ForegroundOnly) {
        # 1) TOS 視窗大小/位置/最小化狀態原樣還原
        if ($s.tos) {
            $h = [IntPtr][int64]$s.tos.hwnd
            if (-not [BossWin]::IsWindow($h)) { $h = Find-TosHwnd }
            if ($h -ne [IntPtr]::Zero) {
                $wp = Get-Placement $h
                $wp.flags = [int]$s.tos.flags
                $wp.ptMinPosition.X = [int]$s.tos.minX; $wp.ptMinPosition.Y = [int]$s.tos.minY
                $wp.ptMaxPosition.X = [int]$s.tos.maxX; $wp.ptMaxPosition.Y = [int]$s.tos.maxY
                $wp.rcNormalPosition.Left = [int]$s.tos.left; $wp.rcNormalPosition.Top = [int]$s.tos.top
                $wp.rcNormalPosition.Right = [int]$s.tos.right; $wp.rcNormalPosition.Bottom = [int]$s.tos.bottom
                $cmd = [int]$s.tos.showCmd
                if ($cmd -eq 2) { $cmd = 7 }          # SW_SHOWMINIMIZED → SW_SHOWMINNOACTIVE（還原時不搶前景）
                elseif ($cmd -eq 1) { $cmd = 4 }      # SW_SHOWNORMAL → SW_SHOWNOACTIVATE
                $wp.showCmd = $cmd
                $r.tos = if ([BossWin]::SetWindowPlacement($h, [ref]$wp)) { "restored showCmd=$($s.tos.showCmd)" } else { 'FAILED' }
            } else { $r.tos = 'TOS not found' }
        }
        # 2) Z 順序：由下往上把每個仍存在的視窗放到最上（SWP_NOMOVE|NOSIZE|NOACTIVATE），結果＝原順序
        $ok = 0
        $list = @($s.zorder)
        for ($i = $list.Count - 1; $i -ge 0; $i--) {
            $h = [IntPtr][int64]$list[$i].hwnd
            if ([BossWin]::IsWindow($h) -and [int]$list[$i].showCmd -ne 2) {
                if ([BossWin]::SetWindowPos($h, [IntPtr]::Zero, 0, 0, 0, 0, 0x0013)) { $ok++ }
            }
        }
        $r.zorder = "$ok/$($list.Count)"
    }

    # 2.5) Claude Desktop 一律縮小（使用者 2026-10-22 要求：不管有沒有完成都不要擋在螢幕上）。
    #      原前景視窗若就是 Claude 本身，縮小後就不再搶回它。
    $r.claude = Minimize-ClaudeWindows
    $claudeHw = @(Get-ClaudeWindows | ForEach-Object { [int64]$_ })

    # 3) 前景視窗搶回並驗證（最多 6 次）
    $fg = [IntPtr][int64]$s.foreground.hwnd
    $fgNote = ''
    if ($claudeHw -contains [int64]$fg -or $s.foreground.process -eq 'claude') {
        # 原前景就是 Claude（2026-10-28：使用者執行前正看著 Claude）：Claude 已縮小，
        # 改把原 Z 順序（最上層在前）中下一個未最小化的非 Claude 視窗設為前景，不再回報 foreground=False
        $fg = [IntPtr]::Zero
        foreach ($z in @($s.zorder)) {
            $h = [IntPtr][int64]$z.hwnd
            if ([int64]$z.hwnd -ne [int64]$s.foreground.hwnd -and $claudeHw -notcontains [int64]$z.hwnd `
                    -and [int]$z.showCmd -ne 2 -and [BossWin]::IsWindow($h) -and [BossWin]::IsWindowVisible($h)) {
                $fg = $h; $fgNote = " -> '$([BossWin]::Title($h))'"; break
            }
        }
    }
    if ([BossWin]::IsWindow($fg)) {
        for ($k = 0; $k -lt 6 -and -not $r.foreground; $k++) { $r.foreground = [BossWin]::ForceForeground($fg); if (-not $r.foreground) { Start-Sleep -Milliseconds 300 } }
    }
    # 4) 游標
    $r.cursor = [BossWin]::SetCursorPos([int]$s.cursor.x, [int]$s.cursor.y)

    if (-not $NoReassert) {
        # Claude Desktop 在工具呼叫結束後常會跳回前面 → 5 秒、15 秒後各補還原前景一次（隱藏背景程序）
        $me = $PSCommandPath; if (-not $me) { $me = 'C:\TradeReview\boss_screen_state.ps1' }
        $cmd = "Start-Sleep 5; . '$me'; Restore-ScreenState -Path '$Path' -ForegroundOnly -NoReassert | Out-Null; Start-Sleep 10; Restore-ScreenState -Path '$Path' -ForegroundOnly -NoReassert | Out-Null"
        Start-Process powershell.exe -WindowStyle Hidden -ArgumentList @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-Command', $cmd) | Out-Null
    }
    $o = [pscustomobject]$r
    Write-Host ("screen state restored: tos={0} zorder={1} foreground={2} ('{3}'{6}) cursor={4} claude_minimized={5}" -f $o.tos, $o.zorder, $o.foreground, $s.foreground.title, $o.cursor, $o.claude, $fgNote)
    return $o
}
