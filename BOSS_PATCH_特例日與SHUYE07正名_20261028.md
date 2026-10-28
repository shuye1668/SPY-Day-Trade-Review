# Boss PC：已寫入日不再擋 gate＋「502」正名 SHUYE07＋畫面還原前景（2026-10-28，SHUYE07 撰寫）

> **給 Boss PC 上的 Claude**：程式部分（`spy_daytrade_writer.py`、`boss_screen_state.ps1`）已由 SHUYE07 改好、
> 經 git 自動送達 `C:\TradeReview\`，**不用動程式**。
> **要你做的**：照下面「SKILL 文字修改」改 `tos-daily-collector` 的 SKILL.md，改完照檔頭規定
> ① 整份複製回 `C:\TradeReview_Boss\CLAUDE.md`、② 複製成 `C:\TradeReview\BOSS_SKILL_current.md` 推送。
> ⚠️ 不要請 SHUYE07 直接改 `BOSS_SKILL_current.md`（Boss 每次推送前會覆蓋它，兩邊同改會 rebase 撞車）。

## 背景（2026-10-28 06:14 那一輪）

1. **「502」稱呼過時**：舊 502（WINDOWS-AUC64IA）早已退役，第二段由 **SHUYE07**（主機 508，D 槽 000AI專區）處理。
   Boss 回報與寄給 Julia 的採集信仍寫「交 502 人工確認」。
2. **整份擋下一整週**：Boss 每天匯出 7 天，`writer --trades-only` 舊版先把 CSV 裡**每個** session 過 gate，才看哪天已寫過。
   10/22、10/23 收盤留倉、10/27 跨日平倉 → 整份 exit 2，Boss 本機 trades_all 停在 10/21，
   且這幾天滾出 7 天範圍之前每天都會報「gate 擋下」。
   但這三天 **SHUYE07 已於 10/28 依核定範本自動寫入**（writer §3b；10/23「留倉中又加碼留倉」範本 10/28 新增），
   trades_all 已到 10/27，經 git 送回 Boss（sync_pull 以 SHUYE07 版覆蓋本機 trades_all）。
3. **畫面還原 foreground=False**：執行前前景視窗就是 Claude 本身。`Restore-ScreenState` 依 10/22 要求把 Claude 縮小、
   不搶回它，但也沒有把前景交給其他視窗 → 永遠回報 False（partial）。

## 改了什麼（SHUYE07 維護、經 git 送達）

| 檔案 | 內容 |
|---|---|
| `spy_daytrade_writer.py` | `--trades-only`：**已在 trades_all 的交易日先排除、印「冪等跳過」，只有未寫入的 session 才過 gate**。全部已寫 → exit 0。未寫入的留倉／平倉日照舊擋（exit 2），gate 本身一條都沒放寬 |
| `boss_screen_state.ps1` | `Restore-ScreenState`：原前景是 Claude 時，Claude 照樣縮小，改把**原 Z 順序中下一個未最小化的非 Claude 視窗**設為前景；log 行多一段 `('<原前景>' -> '<改前景視窗>')` |

已在 SHUYE07 測過（scratch 複本）：trades_all 已含 10/27 + 10/28 CSV → 4 天全冪等跳過、exit 0、檔案不變；
trades_all 停在 10/22 → 10/23、10/27 照擋 exit 2、檔案不變；未寫入的純當沖日（10/21）照常寫入。
畫面模組只做了語法檢查（SHUYE07 沒有 TOS／不在 Boss 桌面，無法實跑），**請以 10/29 那一輪的回報驗收**。

## SKILL 文字修改

### A. 「502」正名（全檔）

指**機器／處理端**的「502」一律改「SHUYE07」，例如「502 端母版」「歸 502」「502 的所有物」「502 的 App」「交由 502 人工處理」。
**不要改**識別字：`sync_push.ps1 -Owner p502`、commit 訊息裡的 `p502`、log 裡的 `[p502]`。

另改架構圖右欄標題（步驟前「這台機器在整套系統裡的角色」）：
```
502（D:\fileserver_D\TradeReview）
```
→
```
SHUYE07（D:\fileserver\書業(新)\000AI專區\1. 程式檔(勿動勿移)\SPYTradeReview）
```
同圖最後一行 `→ CS交易紀錄.xlsx（人工修正在這）` → `→ CS交易紀錄.xlsx（範本內特例日 writer 自動寫，範本外才人工）`。

### B. 步驟 6.5

`write_sync_state.py --status blocked` 那行的 note 改成：
```powershell
    & $py write_sync_state.py --status blocked --note "Boss 本機 gate 擋下（未寫入日），交 SHUYE07"
```
上面註解「記錄狀態供 502 的 App 顯示新鮮度」→「本機乾跑用；sync_state 歸 SHUYE07，不推送」。

把「### 被 gate 擋下時：什麼都不要改」整節換成：

```
### 被 gate 擋下時：什麼都不要改

`--trades-only` 只把**還沒在 trades_all 的交易日**拿去過 gate（2026-10-28 起；已寫過的日子直接冪等跳過，
不再因 7 天匯出範圍內有舊的留倉／平倉日而整份擋下）。被擋時 trades_all 停在前一日**是正確的，不是壞掉**。
照常繼續步驟 8，把被擋原因寫進正文，交由 **SHUYE07** 處理：
- 只有「收盤留倉」及／或「跨日平倉」→ 正文寫「留倉／跨日平倉特例日，SHUYE07 依核定範本自動寫入」，
  **不要寫「需人工確認」**（SHUYE07 的 writer 會自動寫：留倉建立日、平倉日、留倉中又加碼留倉）；
- 其他原因（`BALANCE 對不上`、`非交易現金異動（DOI/JRN）`、`session 開盤/收盤餘額不明`、非 100 股）
  → 正文寫「SHUYE07 需人工確認：<原因>」。

隔天 sync_pull 會帶回 SHUYE07 已寫好的 trades_all，那幾天之後就自動冪等跳過、不會再擋。
```

### C. 步驟 11／12（回報與 log）

- 步驟 12 表格「本機寫入」列改成：
  `| 本機寫入 | ✓ 已寫入 trades_all ／ — 全部已存在（冪等跳過） ／ ⚠ 被 gate 擋下：<原因>（留倉/平倉 → SHUYE07 自動寫；其他 → SHUYE07 人工） |`
- 「畫面還原」列改成：
  `| 畫面還原 | 原頁籤 <tab>；Restore-ScreenState foreground=True（原前景是 Claude 時註明改前景到哪個視窗）／ partial（non-fatal） |`
- 步驟 11 log「本機寫入被 gate 擋下」那行不變；新增一種：全部冪等跳過時寫 `local write SKIPPED (all days already in trades_all)`。

### D. 修訂紀錄（加在 SKILL 修訂紀錄表最上面）

| 2026-10-28 | **已寫入日不再擋 gate**：`--trades-only` 先排除 trades_all 已有的交易日再過 gate（10/28 一輪因 10/22–10/27 留倉/平倉在 7 天範圍內整份擋下）；被擋時留倉/跨日平倉寫「SHUYE07 依範本自動寫入」，不再寫需人工；**「502」全檔正名 SHUYE07**（識別字 `p502` 不變）；畫面還原：原前景＝Claude 時改把前景交給 Z 順序下一個視窗（`boss_screen_state.ps1`） |

## 驗收（10/29 06:00 那一輪）

- 本機寫入：若 10/28 美股無新留倉/平倉 → 應是「✓ 已寫入」或「— 全部已存在」，**不應**再出現 10/22–10/27 的 gate 擋下。
- 回報與 Julia 採集信：不再出現「502」。
- 畫面還原：`foreground=True`；若原前景是 Claude，log 行有 `-> '<視窗標題>'`。
  仍是 False → 把 `screen state restored:` 那整行原文寫進回報，交 SHUYE07 查。
