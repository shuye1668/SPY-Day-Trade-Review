# Spy daytrade auto csv — SKILL（引擎化 SOP）

本檔補足 CLAUDE.md 引用的 `SKILL.md §0`，並定義以 `spy_daytrade_engine.py` 為核心的每日流程。
CLAUDE.md 仍是欄位/區塊結構的權威定義；本檔定義「怎麼跑」與「§0 攔截原則」。

## §0（最高優先）無人看管自動跑的攔截原則

本 routine 每交易日 06:47 無人看管自動跑。凡遇任何「本該停下問使用者」的情境，一律：
把相關儲存格填 `⚠️需人工確認`、對話首行加 `🔴 今日 N 處需人工確認`、
受牽連下游（餘額鏈／合計／核驗／SUMIF／累積損益check）留空；
**絕不為了跑完流程而靜默寫入估算／推算／猜值。**

觸發 `⚠️需人工確認` 的情境（引擎已自動偵測）：
1. **BALANCE 對不上**：逐筆用 `開盤 + Σnet_cash` 重建餘額，與對帳單 BALANCE 欄不符（漏單／多單／費用錯）。
2. **Cash Balance 不完整**：session 缺 BAL 開盤列、或有非交易現金異動（DEP/JRN，如疑似入金的 3 萬落差）。
3. **留倉／跨日平倉**：收盤仍有未平部位，或次日平掉隔夜倉 → offset 會跳動。
4. **合併後非 100 股**：奇數量或零股，需人工判斷。
5. **Day Header B / offset 對不上**：頭尾雙 offset 不一致。

這是 2026-07 多筆 silent 錯誤的共同根因：舊流程沒有 BALANCE 交叉驗證閘門，錯誤靜默流過。

## 每日流程

1. 取得當日券商 Account Statement（CSV 匯出優先；截圖 OCR 為備援，但 OCR 必過 BALANCE 交叉驗證才可信）。
2. **先跑 dry-run**（不寫檔）：
   ```
   python spy_daytrade_engine.py <statement.csv> --date YYYY-MM-DD
   ```
   檢查：`Σcol1 == Σpnl == 收−開`、無 `BALANCE 對不上`、無 `⚠️需人工確認`。
3. 若有任何 `⚠️` → 依 §0 處理，**不寫入**，回報使用者。
4. 全綠才寫入（寫入層 2026-07-24 已接上並驗證）：
   ```
   python spy_daytrade_writer.py <statement.csv> --date YYYY-MM-DD --commit
   ```
   writer 內含 gate（有任何 needs_review/留倉/BALANCE 對不上 → 全部不寫、回報）、
   §3 自動備份、§4 寫後全核驗、失敗自動還原。截圖來源請先依 MIGRATION.md §4.1 轉成 CSV。

## 引擎保證（對齊 CLAUDE.md 鐵則）

- §1 費用一律取對帳單 `Misc Fees` 實欄；按 REF# 合併拆單（股數加權均價、費用加總）。
- §2 寫入位置用值掃描 `last_nonempty_row`，**嚴禁 `ws.max_row`**。
- §3 寫入前備份 `*_backup_YYYYMMDD.xlsx`。
- §4 寫入後跑全部核驗，任一失敗還原備份。
- §5 寫入前檢查 `~$` 鎖定檔。
- CST→EDT 用 `zoneinfo` 自動換算（夏令-12h／冬令-13h），不手算。
- LIFO 配對複刻 `trade_review_app._analyse_all_trades` 當日慣例，與 review app 一致。

## 淘汰對象

`process_*.py`、`write_*.py`、`fix_and_rerun.py` 等每日拋棄式手抄腳本一律停用，改由本引擎產出。
歷史腳本保留供追溯，勿再新增同類。
