# 每日 TOS 採集（排程用）

> **這個檔要放到 Boss PC 的：**
> `C:\Users\TS USER\.claude\scheduled-tasks\tos-daily-collector\SKILL.md`（整檔取代）
>
> 最後更新：2026-08-11

---

## 唯一權威來源

**完全依照 `C:\TradeReview_Boss\CLAUDE.md` 執行。**

先讀完那份檔，再開始動作。它定義了模式選擇、每一個步驟、安全閘門與禁止事項。

⚠️ **本檔刻意不重複任何規則。**
2026-08-11 發現的問題：這個 SKILL.md 原本內嵌了一整份 2026-07-17 的 OCR-only 流程，
與後來更新的 `CLAUDE.md` 直接衝突 —— SKILL.md 說「不再操作三條線匯出 CSV」「不操作
`C:\TradeReview\` 下任何檔案」，而 `CLAUDE.md` 說「CSV 為主」「步驟 6.5 寫入
trades_all」「步驟 9.5 git push」。直接指令會蓋過被引用的檔，結果是本機寫入與
同步從來沒被執行過，每天都要人工補。

**規則只能有一份。要改流程請改 `C:\TradeReview_Boss\CLAUDE.md`，不要寫在這裡。**

---

## 執行環境注意事項（這些不屬於流程規則，故留在本檔）

- 此 session **不會有** desktop computer-use MCP。
  不要 tool_search 找 computer use、不要用 `Claude_in_Chrome__computer`
  （那個要 browser tabId，控制不了桌面程式）。一律走 PowerShell + Win32。

- **ExecutionPolicy 會擋 `.ps1`**（本機全 scope 為 Undefined，實際等同 Restricted）。
  每個需要 dot-source 的 PowerShell call 開頭都要先：

  ```powershell
  Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass -Force
  . C:\TradeReview_Boss\tos_automation.ps1
  ```

  執行獨立腳本則用 `powershell -ExecutionPolicy Bypass -File <path>`。

- **每次 PowerShell tool call 都是新 process**，變數不跨 call 保留。
  所以每個 call 都要重新 dot-source，重要狀態要落檔。

- 每個關鍵步驟後 `Capture-Screen` 存 PNG 到 `C:\TradeReview_Boss\logs\`，
  再用 view tool 讀圖確認畫面狀態，確認後才進下一步。

---

## 開始前的自我檢查

1. 已讀取並理解 `C:\TradeReview_Boss\CLAUDE.md`
2. 該檔的〈修訂紀錄〉最後一筆是 2026-08-10 或更新
   —— 若不是，代表 Boss PC 沒 pull 到最新版，先執行：

   ```powershell
   cd C:\TradeReview
   git pull --rebase --autostash origin main
   Copy-Item C:\TradeReview\BOSS_CLAUDE.md C:\TradeReview_Boss\CLAUDE.md -Force
   ```

3. 確認流程包含「步驟 6.5 本機寫入」與「步驟 9.5 推送」
   —— 沒有的話就是讀到舊版，停下來回報，不要用舊流程跑。
