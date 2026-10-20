# Boss SKILL 修改指示：採集後還原原本畫面（2026-10-20，SHUYE07 撰寫）

> **給 Boss PC 上的 Claude**：使用者要求「截圖/匯出之後回到原本的螢幕畫面，讓使用者看不出來有截圖，不影響使用」。
> 請照下面四處修改 Boss 的排程 SKILL（tos-daily-collector），改完依 SKILL 檔頭規定同步
> `C:\TradeReview_Boss\CLAUDE.md` 與 `C:\TradeReview\BOSS_SKILL_current.md`，並在修訂紀錄加一列。
> 只改流程文字，不改 `.py`；新模組 `boss_screen_state.ps1` 由 502/SHUYE07 端維護、經 git 送達。

## 0. 前置：確認模組已到本機

```powershell
powershell -ExecutionPolicy Bypass -File C:\TradeReview\sync_pull.ps1 -Owner boss
Test-Path C:\TradeReview\boss_screen_state.ps1      # 必須 True
```

模組功能（已在 SHUYE07 實測：前景被搶後還原、最小化視窗被放大後還原成最小化、延遲補還原程序）：

| 函式 | 作用 |
|---|---|
| `Wait-UserIdle -MinIdleSec 120 -MaxWaitMin 9` | 使用者 2 分鐘內有動鍵盤/滑鼠就先等（每 15 秒看一次）；回傳 `$true`=可開工 |
| `Save-ScreenState -Path <json>` | 記住前景視窗、所有視窗 Z 順序、TOS 視窗大小/位置/最小化狀態、游標位置 |
| `Set-SavedTosTab -Path <json> -Tab <名稱>` / `Get-SavedTosTab` | 補記/讀回 TOS 原本停的主頁籤 |
| `Restore-ScreenState -Path <json>` | TOS 視窗原樣還原（原本最小化就縮回去）、Z 順序重排、原前景視窗搶回並驗證、游標歸位；並在 5 秒與 15 秒後背景再補還原前景一次（防 Claude Desktop 事後跳回前面） |

模組只做視窗管理：不點擊、不打字、不關視窗；失敗不 throw。

## 1. 新增【步驟 1.5】記住畫面（放在步驟 1 之後、步驟 2 之前）

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview\boss_screen_state.ps1
$today = (Get-Date).ToString("yyyy-MM-dd")
$ss = "C:\TradeReview_Boss\logs\screen_state_$today.json"
$idle = Wait-UserIdle -MinIdleSec 120 -MaxWaitMin 9      # 這個 call 的 timeout 設 600000
if (-not (Test-Path $ss)) { Save-ScreenState -Path $ss | Out-Null }   # 補跑/重跑時不覆蓋第一次記下的原始畫面
```

- `Wait-UserIdle` 回 `$false`（使用者一直在用）→ 再呼叫最多 2 次（共約 27 分鐘）；仍在用就照跑，log 註記 `user active, ran anyway`。
- **一定要在步驟 2 叫出 TOS 之前存**，存到的才是使用者原本的畫面。

## 2. 步驟 2 補一行：記下 TOS 原本的頁籤

步驟 2 view 截圖確認 TOS 狀態時，順便判斷**主頁籤**目前停在哪（Monitor／Trade／Analyze／Scan／MarketWatch／Charts／…），然後：

```powershell
. C:\TradeReview\boss_screen_state.ps1
Set-SavedTosTab -Path "C:\TradeReview_Boss\logs\screen_state_$today.json" -Tab Charts   # 換成實際看到的頁籤
```

## 3. 步驟 10 改為「切回原頁籤 ＋ 還原畫面」

- **10.1** 原本固定切 Charts → 改切回 `Get-SavedTosTab` 記下的頁籤（空白就照舊 Charts）。
  Charts＝`Switch-ToCharts`；Monitor＝`Click-At 262 33`；其他頁籤從步驟 2 的全螢幕截圖找座標點擊。
  點完一樣 `SetCursorPos` 移開（hover 選單），10.2/10.3/10.4 的驗證與安全限制不變。
  若原頁籤就是 Monitor 且子頁籤原本不是 Account Statement，也點回原子頁籤（座標從步驟 2 截圖取）。
- **10.5（新增，最後一個動作）**：

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
. C:\TradeReview\boss_screen_state.ps1
Restore-ScreenState -Path "C:\TradeReview_Boss\logs\screen_state_$today.json"
```

  - 不要再接 `Capture-Screen`／`Activate-TOSWindowVerified` 或任何會把 TOS/Claude 叫到前面的動作——還原後就結束。
  - 驗證只看輸出那一行：`foreground=True` 才算還原成功；`False` → log 註記 `screen restore partial`（非致命）。
  - 步驟 11 的 log、步驟 12 的回報、寄信等都要在 10.5 **之前**做完（Gmail 工具不會搶前景，可以先做）。
    建議順序：9（寄信）→ 9.5（push）→ 11（log）→ 12（回報表格先在對話寫好）→ 10（切頁籤＋還原）。

## 4. 失敗處理通則補一條

舊規則「失敗時保持畫面在失敗當下（方便人工檢查）」改為：**先** `Capture-Screen` 存 debug 截圖（人工檢查看截圖就好），
**再**照步驟 10 切回原頁籤並 `Restore-ScreenState`，不讓使用者看到停在錯誤畫面的 TOS。
例外：TOS 跳出錯誤對話框（如 Unexpected error）時不要亂點，只做 `Restore-ScreenState`（對話框留著，log 註記）。

## 5. 修訂紀錄要加的一列

| 2026-10-20 | **採集隱形化**：新增步驟 1.5（`Wait-UserIdle`＋`Save-ScreenState`）、步驟 2 記原頁籤、步驟 10 切回原頁籤＋`Restore-ScreenState`（TOS 視窗大小/最小化、Z 順序、前景、游標原樣還原，5/15 秒背景補還原）；失敗時也還原（debug 截圖照存）。模組 `C:\TradeReview\boss_screen_state.ps1` 由 502 端維護 |

## 做不到的部分（請照實告知使用者）

- 採集進行中（約 3～8 分鐘）TOS 會被叫到前面操作，這段期間畫面一定會動；`Wait-UserIdle` 只能確保「有人在用的時候不開工」。
- TOS 內部狀態只還原「主頁籤／子頁籤」；圖表 symbol、捲動位置等採集流程本來就不碰，不受影響。
