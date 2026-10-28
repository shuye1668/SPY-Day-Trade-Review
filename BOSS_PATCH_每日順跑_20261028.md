# Boss PC：每日順跑——開工前先同步＋自我檢查（2026-10-28，SHUYE07 撰寫）

> **給 Boss PC 上的 Claude**：請**和同日的 `BOSS_PATCH_特例日與SHUYE07正名_20261028.md` 一起套用**（同一次改 SKILL）。
> 程式 `C:\TradeReview\boss_preflight.ps1` 已由 SHUYE07 寫好、經 git 送達，**不用動程式**。
> 改完照檔頭規定：① 整份複製回 `C:\TradeReview_Boss\CLAUDE.md`、② 複製成 `C:\TradeReview\BOSS_SKILL_current.md` 推送。
> 套用 patch 是**使用者在場時**做的事；每日無人看管的那一輪只負責「回報有沒有待套用的 patch」，不要自己改 SKILL。

## 為什麼要改（過去一個月 Boss 不順的根因）

| 症狀 | 根因 | 這次的處置 |
|---|---|---|
| SHUYE07 修好的程式要**晚一天**才生效（10/28 修的 writer，10/29 那輪仍會用舊版） | Boss 只在步驟 9.5 push 時才 `pull`，那時步驟 6.5 writer 早就跑完了 | **新增步驟 0.5：開工前先 pull** |
| 10/23、10/24 整天沒推出去，10/27 才發現 | repo 卡在 rebase，回報裡看不出來 | 0.5 每天檢查 `REBASE`，卡住就寫進回報第一行 |
| patch 寫好了卻沒人套用，Boss 照舊說法跑好幾天 | 沒有任何地方提醒 | 0.5 列出 `PENDING_PATCHES`，回報表多一列 |
| 排程卡在「允許」提示，半天沒動 | 排程權限模式不是 Auto 時，無人看管遇到新指令會等人按允許 | 0.5 查自己的權限模式，不是 auto 就寫進回報 |

## 新的程式（SHUYE07 維護，經 git 送達）

`C:\TradeReview\boss_preflight.ps1`：**只回報、不修**，任何一項失敗都不擋採集，exit 永遠 0。輸出：

```
PREFLIGHT_PULL=OK ／ SKIPPED (本機有未推送的改動) ／ FAILED ...   ← 跑 sync_pull.ps1 -Owner boss
PREFLIGHT_REBASE=OK ／ STUCK(...)
PREFLIGHT_HEAD=<sha>
PREFLIGHT_GUARD=OK (Ready) ／ MISSING(...)                           ← Boss_Screen_Guard 守門排程
PREFLIGHT_SKILL_REV=<BOSS_SKILL_current.md 修訂紀錄最新日期>
PREFLIGHT_PENDING_PATCHES=none ／ <檔名, 檔名>                        ← 檔名日期比修訂紀錄新的 BOSS_PATCH_*.md
```

已在 SHUYE07 以 Windows PowerShell 5.1 測過（scratch git repo＋假 sync_pull）：正常路徑、pull 跳過、rebase 卡住三種輸出都正確。

## SKILL 文字修改

### A. 新增【步驟 0.5】（放在步驟 0 與步驟 1 之間）

```
## 【步驟 0.5】開工前同步＋自我檢查（2026-10-28 新增）

只有步驟 0 `US_SESSION=YES`（或手動補跑）才做；休市日照舊直接結束。

    powershell -ExecutionPolicy Bypass -File C:\TradeReview\boss_preflight.ps1

- 檔案不存在（patch 剛套用、git 還沒送到）→ 改跑
  `powershell -ExecutionPolicy Bypass -File C:\TradeReview\sync_pull.ps1 -Owner boss`，結果記為 PREFLIGHT_PULL。
- 把每一行 `PREFLIGHT_*` 原樣記下，步驟 11 log、步驟 12 回報要用。
- **任何結果都照常繼續步驟 1**（這一步只回報、不修）：
  - `PULL=SKIPPED／FAILED` → 用本機現有程式照跑；9.5 推送時會再試一次。
  - `REBASE=STUCK` → 照常採集、寄信，但 9.5 會推不出去；**不要自行修復**，回報第一行寫「⚠ repo 卡在 rebase」。
  - `PENDING_PATCHES` 不是 none → 照現行 SKILL 跑，**不要在無人看管時自己套 patch**，回報裡列出檔名。
- 同時查本排程自己的權限模式：呼叫 `mcp__ccd_session_mgmt__get_session`（`session_id: "self"`）看 `permissionMode`
  （工具不存在或呼叫失敗就記 `unknown`，不要為此停下）。不是 `auto` → 回報寫
  「⚠ 排程權限模式=<值>，無人看管時可能卡在允許提示；請到 Claude Desktop 的排程設定改成 Auto」。
```

### B. 步驟 9.5

開頭補一句：「步驟 0.5 已先 pull 過；這裡的 `pull --rebase` 只是推送前再同步一次，照舊。」

### C. 步驟 11（log）

每一行 log 行尾加：`preflight pull=<OK|SKIPPED|FAILED> rebase=<OK|STUCK> patches=<none|N 個> perm=<auto|default|unknown|...>`

### D. 步驟 12（回報表）新增三列，放在表格最上面

```
| 開工前同步 | ✓ pull OK（HEAD <sha>） ／ ⚠ SKIPPED／FAILED：<原因> ／ ⚠ repo 卡在 rebase |
| 待套用 patch | none ／ ⚠ <檔名>（請使用者在場時要 Boss 套用） |
| 排程權限模式 | auto ／ ⚠ <值>（請改 Auto） ／ unknown |
```

回報第一行規則：上面三列全正常、且採集／寄信／推送都成功 → 第一行寫 `🟢 Boss 採集全部完成`；
任一列有 ⚠ → 第一行寫 `🟡 Boss 採集完成，但有 N 項要處理`，接著逐項列出。

### E. 修訂紀錄（加在 SKILL 修訂紀錄表最上面；和正名 patch 的那列並列）

| 2026-10-28 | **每日順跑**：新增步驟 0.5 `boss_preflight.ps1`（開工前 `sync_pull -Owner boss`，SHUYE07 改的程式當天生效；檢查 rebase 卡住、守門排程、待套用 BOSS_PATCH、排程權限模式；只回報不修、不擋採集）；步驟 11 log 行尾加 preflight 摘要；步驟 12 表格加三列＋第一行 🟢／🟡 規則 |

## 套用後怎麼驗收

1. 套用當下（使用者在場）：手動跑一次 `boss_preflight.ps1`，應看到 `PENDING_PATCHES=none`
   （因為修訂紀錄已有 2026-10-28 兩列）、`PULL=OK`、`REBASE=OK`、`GUARD=OK`。
2. 下一輪 06:00：回報表最上面三列都正常；本機寫入不再出現 10/22–10/27 的 gate 擋下（同日正名 patch 的 writer 修正此時已生效）。
3. 若排程權限模式不是 auto：請使用者在 Claude Desktop → 排程 → `tos-daily-collector` 改成 Auto。
   這是把「卡在允許提示」從根本解掉的唯一方法；Boss 自己不能改這個設定。
