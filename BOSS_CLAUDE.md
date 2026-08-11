# Boss PC 採集 routine —— 規則已移出本檔

> 本檔曾經是每日 TOS 採集流程的完整定義。**2026-08-11 起不再是。**

## 權威來源

**`BOSS_tos-daily-collector_SKILL.md`**（本 repo 內）
→ 部署位置：`C:\Users\TS USER\.claude\scheduled-tasks\tos-daily-collector\SKILL.md`

排程執行的就是那一份，它是自足的完整流程。**要改流程請改那一份。**

## 為什麼變成這樣

2026-08-11 查出的事故：

同一套流程當時存在**兩份**且互相矛盾 ——

| | 排程實際執行的 SKILL.md | 本檔（`CLAUDE.md`） |
|---|---|---|
| 取料 | ❌ 明文禁止 CSV 匯出 | ✅ CSV 為主、OCR 備援 |
| OCR 範圍 | 只抓 `TRD` | 每一列（含 `BAL`/`DOI`/`JRN`） |
| 本機寫入 | 無，且**禁止操作 `C:\TradeReview\`** | 步驟 6.5 `--trades-only` |
| git push | 無 | 步驟 9.5 `sync_push` |

SKILL.md 第一行雖寫「依照 `C:\TradeReview_Boss\CLAUDE.md` 的規則執行」，
但後面立刻用明確禁令推翻它。**直接指令優先於被引用的檔**，所以：

- 本機寫入（步驟 6.5）與 git 推送（步驟 9.5）**從未執行過**
- `trades_all.xlsx` 每天都要人工補
- 2026-08-07 之後三個交易日的資料全是手動進來的

把 `CLAUDE.md` 更新得再完整都沒用 —— 排程根本沒讀那些內容。

## 結論

**規則只能有一份。** 本檔保留為指標，不再放任何流程規則；
若你在本檔看到任何具體步驟，那就是又分岔了，請刪掉並回到權威檔。

`C:\TradeReview_Boss\CLAUDE.md` 這個檔案本身可以留著（內容同本檔），
也可以刪除 —— 排程已不依賴它。
