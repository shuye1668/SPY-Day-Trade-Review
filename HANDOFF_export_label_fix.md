# Trade Review App — Export 損益標籤重複位移修正（接手脈絡）

> 貼到新對話即可延續。最後更新：2026-07-03（第 12 節新增 7/2 堆柱問題修正）

## 1. 專案背景
- 檔案：`C:\TradeReview\trade_review_app.py`（Flask web app，約 2800 行；主體是內嵌在 Python raw string 裡的一大段 HTML/JS，用 canvas 畫圖）。
- 功能：Bloomberg 風格分鐘 K 線「交易檢討」工具。`draw()` 在 canvas 上畫 K 線、交易線、損益(PnL)標籤。
- 匯出：`exportPNG`（約原 2280+ 行）把同一天的圖**按價格切成上下多張 PNG（價格帶分頁）**下載。用途是**印出來手動上下拼貼**成一張完整長圖。
- 每頁價格帶：`pageRange = step*PAGE_GRIDS`（`PAGE_GRIDS=9`）。相鄰分頁的價格帶**刻意重疊**（拼貼不漏帶）。

## 2. 問題
匯出多分頁時，同一筆交易的損益標籤（如 `-54.95`、`-61.02`、`-283.27`）在相鄰兩頁都出現，但**落在不同位置**。拼貼後同一筆損益出現在兩個地方，無法判讀、破壞版面。

## 3. 需求澄清（關鍵，別搞反）
使用者要的**不是**「只出現一次」。正確需求：
> 重疊帶的標籤**兩頁都要出現**，但必須畫在**K 線價格軸的同一個絕對座標**。這樣兩頁疊起來完全重合，看起來就是一個；不管從重疊帶哪裡剪開拼貼，標籤都完整不被裁掉。

## 4. 根源
`draw()` 的 pass-2（PnL 標籤避讓排版）是**每一頁各自重算**：
- per-page 的 `placedLabels`、`obstacles`、`prevLabelCenterX` 單調游標、以及最後 `chosen.y` 對該頁 `[M.t, H-M.b]` 的夾擠都不同 → 同一標籤在不同頁被擺到不同的 (x, y)。
- 另外舊的 `labelBelongsOnPage(t)` 用 `margin=(px-pn)*0.15` 上下對稱放寬，使邊界附近標籤**同時通過上頁與下頁**判斷 → 兩頁都畫、且位置不同。

## 5. 解法：「算一次、各頁重播」
多分頁匯出時，改成在**跨整段價格的絕對座標系**跑一次排版，把每筆標籤位置存成 `(x, priceForY)`；之後每頁只用 `yOf(priceForY, pn, px)` 重畫。因為位置只跟「價格」有關、跟分頁無關，重疊帶標籤在兩頁會落在**完全相同的價格軸座標**。重播時裁切到繪圖區，邊界標籤在頁緣乾淨切斷、在隔壁頁完整出現。

**嚴格限制**：只在 `exportMode && exportPageTops && exportPageTops.length > 1` 生效。螢幕互動畫面、單頁匯出（不設 `exportPageTops`）**完全走原路徑、零變動**。

## 6. 實際改動（已套用在 trade_review_app.py）
四處，全部只影響 export：

1. 新增全域變數（在 `exportPrevClose=null;` 宣告之後）：
   ```js
   let exportLabelCache=null; // 多分頁匯出：損益標籤位置只算一次(絕對價格座標)，各頁重播
   ```
2. `exportPNG` 開頭，`exportPrevClose=prevCloseInfo;` 之後：
   ```js
   exportLabelCache=null; // 每次匯出重新計算
   ```
3. 還原段（`exportMode=false; ... exportPrevClose=null;` 那一行結尾）加上：
   ```js
   exportLabelCache=null;
   ```
4. pass-2 在 `if(exportMode&&di!==focusIdx){ return; }` 之後，插入約 150 行的多分頁分支：
   - 條件 `if(exportMode&&exportPageTops&&exportPageTops.length>1){ ... return; }`
   - `exportLabelCache===null` 時：用 `yAbs(p)=M.t+(topEdge-p)*pxPerDollar` 建絕對座標障礙物（K 線 wick + 分隔線 + 跨日線），跑與原本相同的避讓搜尋（候選點、單調 X、collision/overlap fallback），把結果存進 `out=[{x, priceForY, lines, isHold, isCross, pnl, lblW, lblH}]`。
   - 重播：`ctx.save()` 裁切 `[M.l,M.t,cW,cH]`，每筆 `yTop=yOf(priceForY,pn,px)`，可見性判斷 `if(yTop+lblH<M.t||yTop>H-M.b)continue;` 後畫字，`ctx.restore()`。

## 7. 關鍵座標數學（給接手者驗算）
- `yOf(p,pn,px) = M.t + (1-(p-pn)/(px-pn))*cH`，其中 `cH=H-M.t-M.b`。
- 每頁 `px-pn = pageRange`（各頁相同）。`pxPerDollar = cH/pageRange`（各頁相同）。
- `topEdge = exportPageTops[0]`；`bottomEdge = exportPageTops[last]-pageRange`。
- 存：`priceForY = topEdge - (chosen.y - M.t)/pxPerDollar`（`yAbs` 的精確反函數）。
- 重播：`yTop = yOf(priceForY,pn,px) = M.t + (pageTop - priceForY)*pxPerDollar`。
- 結論：標籤實體位置只由 `priceForY` 決定，與分頁無關 → 兩頁同座標。✓

## 8. 已完成的驗證
- 新插入 JS 區塊 `node --check` 隔離語法檢查：**通過**。
- 座標換算為精確反函數；`holdCol`、`T.pW`/`T.pL`、`cW` 都在 `draw()` scope 內（`holdCol` 定義於約 1207 行）。
- 邏輯走查：螢幕/單頁匯出路徑不受影響；每筆標籤至少出現在一頁（不漏畫）；重疊帶標籤兩頁同座標。

## 9. 尚未完成（需使用者本機實測）
- 實際跑 Flask + Bloomberg + 瀏覽器 canvas，匯出 **2026-06-26**（會切成多頁）目視確認：相鄰兩頁重疊帶損益是否落在同一高度、可直接拼貼、無漏畫。
- 注意：Cowork 沙盒的 bash mount 是 session 開始的**舊/不完整快照**，看不到檔案編輯，因此**無法用 bash 跑整檔驗證或啟動 app**；`Read`/`Edit` 檔案工具層才是真實檔案。要驗證只能在使用者本機。

## 10. 若要繼續微調
- 某筆仍對不齊：檢查 `priceForY` 換算、`yOf` 是否用到當頁 `pn/px`。
- 某筆漏畫：放寬可見性門檻或檢查 `chosen.y` 的 `virtTop/virtBot` 夾擠。
- 邊界標籤被切一半不美觀：可改「完整落在頁內才畫」規則 `yTop>=M.t && yTop+lblH<=H-M.b`（代價是該筆只會出現在完整容納它的那一頁）。

## 11. 備份
原檔已備份：`C:\TradeReview\trade_review_app_backup_20260629_055707.py`

---

## 12. 2026-07-03 追加修正：7/2 匯出「標籤堆成一柱 + 壓註解」

### 症狀（7/2 匯出圖）
中段約 10 筆標籤（+115.44、-621.19、+59.33…-64.24）全部堆在 x≈12:15 的同一垂直柱、遠離各自交易線；+115.44 壓在市場概述文字上。跨頁同座標（第 5 節機制）本身正常。

### 根因（已逐筆用像素驗算吻合）
1. **prevCX 毒化 + 備援瞬移**：分支內排序沿用 entryTime，但 -621.19 是 10:45 進場、13:50 出場的全日單，anchor（線段中點 midX）落在 ~12:15。它先被排 → `prevCX` 跳到 ~12:15 → 之後每筆中段交易的所有候選點（midX±120px）都在 prevCX 左邊被 `centerX<prevCX` 全數跳過 → `chosen`/`bestFallback` 皆 null → 最後備援 `fx=max(...,prevCX-lblW/2)` 把標籤瞬移到同一 X、各自 baseAbove 高度 = 垂直柱。
2. **長對角線 anchor 錯位**：`baseAbove=min(y1,y2)` 對多小時對角線是「進場端頂點價位」→ -621.19 飄在離線 3 美元的半空。
3. **壓註解**：多分頁分支沒把重播的標籤填入 `pnlLabelsForNotes` → 註解排版不知道標籤存在。

### 修正（5 處，全部只在多分頁分支內；螢幕/單頁匯出仍零變動）
1. 排序改用「標籤 anchor X」（`anchorXOf`：一般=線段中點、hold=進場x、cross=出場x），specials 仍排最後 → prevCX 天然遞增，可證明每筆至少有一個候選點通過單調檢查（prevCX ≤ prevMidX+searchXR ≤ midX+searchXR）。
2. 一般交易線也加入 aObs 障礙物（同 crossDay 的分段小矩形，pad=5）→ 標籤永不壓任何交易線。
3. 一般交易 anchor Y 改為線段在 midX 的高度 `(y1+y2)/2` → 配合 2，dy-walk 會讓標籤貼著自己的線。
4. 最後備援移除 prevCX 項（只留 `max(leftBound,midX-lblW/2)`）→ 永不瞬移。
5. 重播迴圈把當頁可見標籤 push 進 `pnlLabelsForNotes` → 註解自動避開。

### 驗證狀態
- 整段 `<script>` JS `node --check` 通過；整檔 Python `py_compile` 通過（用 mount + Read 尾端拼接檢查）。
- 跨頁同座標機制未動（cache/replay 原樣）。
- **待本機實測**：重新匯出 2026-07-02 目視確認：(a) 中段標籤各自貼近交易線、無垂直柱 (b) -621.19 貼著大對角線 (c) 註解不再被壓 (d) 重疊帶標籤兩頁同高度。
- 註：本次 session 發現 bash mount「大致即時但尾端截斷」（2806/2868 行）——編輯會同步進 mount，但檔尾約 62 行缺失；整檔驗證需拼接 Read 到的真實尾端。

### 備份
`C:\TradeReview\trade_review_app_backup_20260703.py`（本次修改前）

---

## 13. 2026-07-03 後續：螢幕路徑同款修正（已在檔案內、已驗證）

### 內容
原 per-page pass-2（**螢幕互動 + 單頁匯出共用路徑**）也套用了與第 12 節一一對應的四項修正（此部分非本 session 所改，係使用者側/另一對話套用；本 session 負責審閱與驗證）：
1. `scrAnchorXOf` anchor-X 排序（specials 仍排最後）
2. 一般交易線加入 `obstacles`（ei3/xi3 命名、pad=5，避開與 crossDay 區塊 ei2/xi2 撞名）
3. 一般交易 anchor Y 改 `(y1+y2)/2`
4. 最後備援移除 `prevLabelCenterX` 瞬移項

### 驗證（本 session）
- 整檔 `py_compile` + 整段 `<script>` `node --check`：**通過**（現行 2904 行狀態）。
- 逐項審閱：與多分頁分支邏輯一致；`finalRect`→`placedLabels`/`pnlLabelsForNotes`、游標更新、繪字皆原樣。

### 螢幕動態 vs 匯出的差異盤點（設計上保留、未互相汙染）
- 多分頁匯出走 cache/replay 分支（第 12 節），螢幕路徑改動對它**零影響**；反之亦然。
- 螢幕=每幀重排（pan/zoom/hover 都會重算）：排序與 anchor 純由幾何決定 → 拖曳縮放時順序穩定；ties 落回 entryTime/ti，決定性。
- per-day `placedLabels`/`prevLabelCenterX` 每日重置 → 多日並排視圖互不干擾。
- 差異參數保留：`boundPad=exportMode?12:2`、字級 16 vs `sz(11)`、`labelBelongsOnPage` 只在 exportYOverride（單頁匯出）過濾。
- K 線障礙物只收 viewport 內（原設計），新的交易線障礙物無 viewport 過濾 → 拖曳時更穩定。
- 效能：每日約 +100~200 個線段障礙矩形；候選點迴圈通常提早命中，正常無感。若日後大量交易日拖曳變卡，優化方向=按 (day,pn,px,zoom,panX) 快取障礙物。

### mount 行為更正（覆蓋第 12 節註記）
Cowork bash mount 讀取被「session 起始時的檔案大小」截斷（本次 127,680 bytes），**截斷點之前的內容是即時的**；檔案變大後尾端讀不到。因此：`trade_review_app_backup_20260703.py` 是**完整**的修改前狀態（cp 當下檔案尚未變大）；整檔驗證用「mount 前段 + Read 真實尾端」拼接。

### 待本機實測（螢幕）
開 2026-07-02：標籤各自貼近交易線、-621.19 貼大對角線、無垂直柱；拖曳/縮放重排順暢不閃跳；多日縮小視圖正常；單頁匯出版面同步改善。
