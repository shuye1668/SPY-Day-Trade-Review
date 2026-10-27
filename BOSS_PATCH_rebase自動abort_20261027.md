# Boss PC：pull --rebase 失敗自動 abort＋更新 SKILL 過時 rebase 註記（2026-10-27，SHUYE07 撰寫）

> **給 Boss PC 上的 Claude**：程式部分已做在共用模組裡、經 git 自動送達，不用動。
> **要你做的只有一件事**：改 SKILL.md 步驟 9.5 那段過時的 rebase 註記（下面「SKILL 文字修改」），
> 再照檔頭規定複製成 `BOSS_SKILL_current.md` 推送。⚠️ **不要**請 502 端直接改 `BOSS_SKILL_current.md`——
> 那是 Boss 每次推送前覆蓋的檔，兩邊同時改會在 rebase 撞車（正是這次卡住的同一種死結）。

## 背景

2026-10-23 起 Boss repo 又卡在未完成的 rebase：未推出的 commit `4a84ebb` 含 502 所有的
`sync_state.json`，`sync_push.ps1` 的 `pull --rebase` 重放時衝突，腳本直接 throw → repo 停在 rebase 中間，
之後每天都推不出去（10/23、10/24、10/27 的 CSV 卡在本機）。2026-10-27 已由 Boss 端人工修復（HEAD `710a581`）。

## 改了什麼（都在 `C:\TradeReview\`，經 git 送達）

| 檔案 | 內容 |
|---|---|
| `_git_helper.ps1` | 新增 `Test-RebaseInProgress`、`Invoke-PullRebase`：`pull --rebase` 失敗且 repo 停在 rebase 中 → 記下衝突檔 → **自動 `git rebase --abort`**，回到 pull 前狀態（本機 commit 與 autostash 都保留） |
| `sync_push.ps1` | 兩處 `pull --rebase`（commit 後、push 重試時）改走 `Invoke-PullRebase`；失敗印 `[X] pull --rebase 失敗，已自動 git rebase --abort` ＋衝突檔，log `FAIL pull-rebase ABORTED conflicts=...`，exit 1。開頭若發現**上一輪遺留**的 rebase（舊版留下的），不自動處理、印 `[X] repo 停在未完成的 rebase`、exit 1（中途可能有 stage 未 commit 的資料，abort 會丟） |
| `sync_pull.ps1` | `pull --rebase` 同樣改走 `Invoke-PullRebase`，失敗自動 abort、exit 1 |

效果：之後衝突只會讓**當次**推送失敗（exit 1、log 有衝突檔名），repo 不會再卡在 rebase；下次執行自動再試。

## SKILL 文字修改（步驟 9.5）

把這三行：

```
⚠️ 2026-08-11 起 repo 卡在未完成的 interactive rebase（`.git\rebase-merge` 存在）：
存在時**不要跑 sync_push**（只會在 detached HEAD 上堆孤兒 commit），log 記 `PUSH SKIPPED (repo in rebase)`，
**不要自行修復**，等使用者手動處理。
```

換成：

```
rebase 防卡（2026-10-27）：sync_push／sync_pull 的 `pull --rebase` 失敗時會**自動 `git rebase --abort`**，
repo 不會再停在 rebase 中間；當次 exit 1、輸出 `[X] pull --rebase 失敗，已自動 git rebase --abort`＋衝突檔名
→ log 記 `PUSH FAILED (rebase conflict: <檔名>)`，寫進步驟 11 log 與回報，**不要自行解衝突**（下次執行會自動再試；
若連續兩天同一檔衝突，回報使用者）。
若 sync_push 輸出 `[X] repo 停在未完成的 rebase`（舊版遺留）→ 不推送、**不要自行修復**，回報使用者處理。
衝突最常見原因：commit 裡混進 502 所有物（`trades_all.xlsx`／`sync_state.json`）——一律只用 `sync_push.ps1 -Owner boss`，絕不 `git add -A`。
```

## 修訂紀錄（加在 SKILL 修訂紀錄表）

| 2026-10-27 | **rebase 防卡**：sync_push/sync_pull `pull --rebase` 失敗自動 `rebase --abort`（`_git_helper.ps1` `Invoke-PullRebase`）；9.5 過時的「2026-08-11 起卡 rebase」註記改為新處置；根因＝10/24 commit 混入 sync_state.json，10/27 人工修復（HEAD 710a581） |

## 修復安全網的清理

10/27 修復留下的備份分支 `backup-boss-20261027-*` 與 `stash@{0}`：**保留到 2026-10-29**（確認 10/28、10/29 兩輪推送都正常）再刪。
另外兩個修復前就存在的 autostash stash 一併在那時清掉。
