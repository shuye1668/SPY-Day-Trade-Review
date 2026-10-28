# 給 Boss PC 上的 Claude：套用待處理的 patch（SHUYE07 撰寫，2026-10-28）

> 使用者經 AnyDesk 遠端操作、無法貼上長文字，所以把指示放在這個檔。
> 使用者只會打一行：`run C:\TradeReview\sync_pull.ps1 -Owner boss, then follow C:\TradeReview\BOSS_APPLY.md`
> 使用者在場，這是**人工指示**，可以修改本機排程 SKILL。

## 要做的事（依序）

1. 確認 sync_pull 已成功（exit 0）。若 exit 2（本機有未推送改動）→ 先跑
   `powershell -ExecutionPolicy Bypass -File C:\TradeReview\sync_push.ps1 -Owner boss` 再 pull 一次；
   仍失敗就停下，把輸出回報給使用者。
2. 依序讀這兩份 patch，照各自「SKILL 文字修改」段落，**一次**改好 `tos-daily-collector` 排程的 SKILL.md：
   - `C:\TradeReview\BOSS_PATCH_特例日與SHUYE07正名_20261028.md`
   - `C:\TradeReview\BOSS_PATCH_每日順跑_20261028.md`
   兩份的修訂紀錄列都要加（都放在修訂紀錄表最上面）。
3. 照 SKILL 檔頭規定同步：整份複製回 `C:\TradeReview_Boss\CLAUDE.md`（開頭說明保留）、
   複製成 `C:\TradeReview\BOSS_SKILL_current.md`。
4. 推送：`powershell -ExecutionPolicy Bypass -File C:\TradeReview\sync_push.ps1 -Owner boss`
5. 驗收：`powershell -ExecutionPolicy Bypass -File C:\TradeReview\boss_preflight.ps1`
   - 預期 `PREFLIGHT_PULL=OK`、`PREFLIGHT_REBASE=OK`、`PREFLIGHT_GUARD=OK`、`PREFLIGHT_PENDING_PATCHES=none`。
6. 查本排程權限模式：`mcp__ccd_session_mgmt__get_session`（`session_id: "self"`）的 `permissionMode`；
   若是排程本身的設定查不到，就看 Claude Desktop 排程頁 `tos-daily-collector` 的權限設定。

## 回報給使用者（照這個格式，短）

```
patch 套用：✓ 兩份都已改進 SKILL（修訂紀錄 2026-10-28 兩列）／✗ <卡在哪>
推送：✓ <sha>／✗ <原因>
PREFLIGHT_*：<每行原樣貼上>
權限模式：<值>（不是 auto → 請到 Claude Desktop 排程設定把 tos-daily-collector 改成 Auto）
```

## 不要做的事

- 不要改任何 `.py`／`.ps1`（SHUYE07 維護，git 送達）。
- 不要 `git add -A`；只用 `sync_push.ps1 -Owner boss`。
- 不要動 TOS、不要跑採集流程（這次只套 patch）。
