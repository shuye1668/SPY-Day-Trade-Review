' boss_run_hidden.vbs — run a PowerShell script with no window (Boss_Screen_Guard task, 2026-10-22).
' A visible console would itself steal the foreground, which is exactly what the guard is fixing.
' Usage: wscript.exe //B //Nologo boss_run_hidden.vbs "C:\TradeReview\boss_screen_guard.ps1"
If WScript.Arguments.Count < 1 Then WScript.Quit 1
CreateObject("WScript.Shell").Run "powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & WScript.Arguments(0) & """", 0, False
