# 每日 TOS 採集 routine（Boss PC）—— 指向說明

> **現行全文請看同資料夾的 `BOSS_SKILL_current.md`**（位元組級正本，Boss 每次 sync_push
> 自動更新；2026-10-16 git 修復後啟用這個機制）。權威版在 Boss 本機
> `C:\Users\TS USER\.claude\scheduled-tasks\tos-daily-collector\SKILL.md`，兩者由
> Boss 端的「改 SKILL 後複製一份進 repo」步驟保持一致。
>
> 本檔不再保存全文副本——歷史教訓（2026-08-11、2026-10-16 兩次）：同一套規則存兩份
> 手工維護的副本必然漂移、互打。要查 Boss 流程一律讀 `BOSS_SKILL_current.md`。
>
> 502 端需要記住的流程要點（詳見 memory `boss-mail-flow-20261016`）：
> - Boss 台灣週二～週六 06:00 跑（05:45 喚醒）；美股休市日步驟 0 直接結束。
> - **有「沒寄過的新成交」才建草稿並直接寄 juliahsu13@gmail.com**（寄出後在寄件備份，
>   不在草稿匣）；無新成交＝不寄信（正常），但照樣匯出 CSV 並 sync_push。
> - 交易日另存 `verify_shots\TOS_verify_*.png`（Cash Balance＋Order History＋Trade History
>   驗證截圖），502 端貼進 Julia 母檔。
> - Boss sync_push 只推 `_inbox\`、`verify_shots\`、`BOSS_SKILL_current.md` 三樣；
>   trades_all.xlsx 與 sync_state.json 歸 502（2026-10-16 所有權定案）。
> - 修復歷程：2026-08-11 卡 rebase 兩個月 → 2026-10-16 使用者授權修復
>   （備份分支 backup-boss-20261016-head/-main、檔案備份 git_repair_stash_20261016\、
>   重置到 origin/main、推 923b67f）。Gmail 信件自此退為備援通道。
