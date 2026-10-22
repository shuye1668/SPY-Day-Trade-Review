# Boss PC 畫面守門：Claude 一律縮小＋卡住也還原畫面（2026-10-22，SHUYE07 撰寫）

> **給 Boss PC 上的 Claude**：這次**不用改 SKILL 文字**，全部做在 502 端維護的模組裡，經 git 自動送達。
> 可以在 SKILL 修訂紀錄加下面那一列，並照檔頭規定同步 `BOSS_SKILL_current.md`（選做）。

## 使用者要求

Boss 採集不管有沒有完成，都要把自己（Claude Desktop）縮小；任務卡住也不要擋在螢幕上，
要還原螢幕原本的使用狀況。

## 改了什麼（都在 `C:\TradeReview\`）

| 檔案 | 內容 |
|---|---|
| `boss_screen_state.ps1` | `Restore-ScreenState`（步驟 10.5）現在一律把 Claude Desktop 視窗最小化，5/15 秒補還原時再縮一次；畫面狀態檔不存在時也照縮。`Wait-UserIdle`（步驟 1.5）結束時記執行旗標 `started`，`Restore-ScreenState` 記 `finished`（`C:\TradeReview_Boss\logs\run_flag_<日期>.json`）。兩者第一次在 Boss PC 執行時自動註冊守門排程 |
| `boss_screen_guard.ps1` | 守門：`started` 超過 30 分鐘還沒 `finished`＝卡住 → `Restore-ScreenState`（TOS 原樣、Z 順序、前景、游標）＋縮小 Claude；跑完後補縮一次 Claude；沒有旗標（休市日、或步驟 1.5 前就卡住）過 06:40 → 縮小 Claude。使用者 3 分鐘內有動鍵盤/滑鼠就不動 |
| `boss_run_hidden.vbs` | 以隱藏視窗執行守門（避免主控台視窗自己搶前景） |

守門排程 **Boss_Screen_Guard**：本機目前使用者、互動式、一般權限；週二～週六 06:10 起每 10 分鐘、到 09:30（涵蓋 07:30 補匯出）。
只做視窗管理：不點擊、不打字、不關視窗、不碰 TOS 內容。

## 生效時間

- 10/23 那一輪：步驟 9.5 push 時 `pull --rebase` 會帶進新模組 → 步驟 10.5 就會縮小 Claude 並註冊守門排程。
- 10/24 起：步驟 1.5 開始記旗標，卡住也有守門救。
- 若 10/23 在 9.5 之前就卡住，新模組還沒到：在 Boss PC 跑一次 `powershell -ExecutionPolicy Bypass -File C:\TradeReview\sync_pull.ps1 -Owner boss` 即可。

## 在 Boss PC 驗證

```powershell
Get-ScheduledTask -TaskName Boss_Screen_Guard | Select TaskName, State
Get-Content C:\TradeReview_Boss\logs\run_flag_$((Get-Date).ToString('yyyy-MM-dd')).json
Get-Content C:\TradeReview_Boss\logs\screen_guard.log -Tail 10
```

步驟 10.5 的輸出行會多一欄 `claude_minimized=N/M`。

## 修訂紀錄（選做，加在 SKILL 修訂紀錄表）

| 2026-10-22 | **畫面守門**：`Restore-ScreenState` 一律縮小 Claude Desktop；執行旗標 run_flag（1.5 記 started、10.5 記 finished）；守門排程 Boss_Screen_Guard（週二～六 06:10–09:30 每 10 分，卡住 30 分鐘→還原畫面＋縮小 Claude），由模組自動註冊。SKILL 流程文字不變 |

## 做不到的部分

- 採集進行中（約 3～8 分鐘）TOS 一定會被叫到前面操作。
- 卡住的判定要等 30 分鐘（正常採集＋寄信＋push 約 10～15 分鐘，留餘裕避免誤判）。
- 守門只把 Claude 縮小，不會替卡住的權限提示按允許或拒絕；任務本身仍卡著，要人工處理。
