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

function Wait-UserIdle {
    # 使用者最近 MinIdleSec 秒內有動鍵盤/滑鼠 → 每 15 秒再看一次，最多等 MaxWaitMin 分鐘。
    # 回傳 $true=已閒置可開工；$false=等到上限仍在使用（由 SKILL 決定：照跑或延後）。
    param([int]$MinIdleSec = 120, [int]$MaxWaitMin = 30)
    $deadline = (Get-Date).AddMinutes($MaxWaitMin)
    while ($true) {
        $idle = Get-IdleSeconds
        if ($idle -ge $MinIdleSec) { Write-Host "user idle ${idle}s → OK"; return $true }
        if ((Get-Date) -ge $deadline) { Write-Host "user still active (idle ${idle}s) after $MaxWaitMin min"; return $false }
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
    $r = [ordered]@{ tos = 'skip'; zorder = 'skip'; foreground = $false; cursor = $false }
    if (-not (Test-Path $Path)) { Write-Host "no state file: $Path"; return [pscustomobject]$r }
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

    # 3) 前景視窗搶回並驗證（最多 6 次）
    $fg = [IntPtr][int64]$s.foreground.hwnd
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
    Write-Host ("screen state restored: tos={0} zorder={1} foreground={2} ('{3}') cursor={4}" -f $o.tos, $o.zorder, $o.foreground, $s.foreground.title, $o.cursor)
    return $o
}
