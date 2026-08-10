# -*- coding: utf-8 -*-
"""
Process 2026-06-15 SPY day trades into trades_all.xlsx and CS交易紀錄.xlsx
"""
import openpyxl
from copy import copy

# ============================================================
# 1. trades_all.xlsx
# ============================================================
wb_trades = openpyxl.load_workbook('trades_all.xlsx')
ws_trades = wb_trades.active

# Update 4 hold rows Status '1' -> '2' and set Pair_Date/Pair_Time
# LIFO matching (stack top first):
#   Row 689: 2026-06-11 12:33:06 BOT 100 @727.9459 -> closed by SOLD 09:30:08
#   Row 636: 2026-06-09 10:12:37 BOT 100 @742.915  -> closed by SOLD 09:30:33
#   Row 617: 2026-06-05 11:14:45 BOT 100 @748.45   -> closed by SOLD 09:31:07
#   Row 616: 2026-06-05 10:44:16 BOT 100 @749.37   -> closed by SOLD 09:33:45

hold_updates = [
    (689, '2026-06-15', '09:30:08'),
    (636, '2026-06-15', '09:30:33'),
    (617, '2026-06-15', '09:31:07'),
    (616, '2026-06-15', '09:33:45'),
]
for row, pair_date, pair_time in hold_updates:
    ws_trades.cell(row, 9).value = '2'       # Status
    ws_trades.cell(row, 10).value = pair_date # Pair_Date
    ws_trades.cell(row, 11).value = pair_time # Pair_Time

# Append 8 new rows for 6/15
# Columns: Date, Exec Time(EDT), Symbol, Price, Type, PnL, Shares, Action, Status, Pair_Date, Pair_Time
new_rows = [
    # 4 平倉 sells (closing prior-day holds)
    ('2026-06-15', '09:30:08', 'SPY', 751.945,   '平', '+2398.34', 100, 'S', '2', '2026-06-11', '12:33:06'),
    ('2026-06-15', '09:30:33', 'SPY', 752.255,   '平', '+932.43',  100, 'S', '2', '2026-06-09', '10:12:37'),
    ('2026-06-15', '09:31:07', 'SPY', 752.5902,  '平', '+412.45',  100, 'S', '2', '2026-06-05', '11:14:45'),
    ('2026-06-15', '09:33:45', 'SPY', 752.9301,  '平', '+354.44',  100, 'S', '2', '2026-06-05', '10:44:16'),
    # 2 day-trade round trips
    ('2026-06-15', '10:13:40', 'SPY', 752.965,   None,  None,      100, 'B', '0', '2026-06-15', '10:20:08'),
    ('2026-06-15', '10:20:08', 'SPY', 753.2101,  '多', '+22.94',   100, 'S', '0', '2026-06-15', '10:13:40'),
    ('2026-06-15', '12:40:39', 'SPY', 755.9399,  None,  None,      100, 'B', '0', '2026-06-15', '13:02:01'),
    ('2026-06-15', '13:02:01', 'SPY', 755.9284,  '多', '-2.73',    100, 'S', '0', '2026-06-15', '12:40:39'),
]

start_row = ws_trades.max_row + 1  # 706
for i, row_data in enumerate(new_rows):
    r = start_row + i
    for c, val in enumerate(row_data, 1):
        ws_trades.cell(r, c).value = val

wb_trades.save('trades_all.xlsx')
print(f"trades_all.xlsx updated: rows 616,617,636,689 Status->2, rows {start_row}-{start_row+7} added")

# ============================================================
# 2. CS交易紀錄.xlsx - 個別識別法 sheet
# ============================================================
wb_cs = openpyxl.load_workbook('CS交易紀錄.xlsx')
ws_cs = wb_cs['個別識別法']

# Starting row: after blank row following last content (row 1378)
# Row 1379 = blank, Row 1380 = blank (separator)
# Actually, 6/12 block ends at 1378, need blank row then Day Header
R = 1380  # Day Header row

# --- Day Header ---
ws_cs.cell(R, 1).value = '2026.06.15'
ws_cs.cell(R, 2).value = -115990.27    # 期初餘額 = -116164.69 + 174.42
ws_cs.cell(R, 3).value = '損益'

# --- 8 Detail rows (all TRDs in EDT order) ---
# col1 = AMOUNT - abs(Misc Fees)
# For sells (closing holds): positive col1
# For buys: negative col1
# col C = PnL for closing side of day-trades only (平倉 PnL goes to 平倉區)
detail_start = R + 1  # 1381

details = [
    # (col_A=col1, col_C=PnL_or_None)
    # 09:30:08 SOLD 100 @751.945 (平倉 exit for 6/11 hold) -> col1 = 75192.93, no C
    (75192.93, None),
    # 09:30:33 SOLD 100 @752.255 (平倉 exit for 6/09 hold) -> col1 = 75223.93, no C
    (75223.93, None),
    # 09:31:07 SOLD 100 @752.5902 (平倉 exit for 6/05 hold) -> col1 = 75257.45, no C
    (75257.45, None),
    # 09:33:45 SOLD 100 @752.9301 (平倉 exit for 6/05 hold) -> col1 = 75291.44, no C
    (75291.44, None),
    # 10:13:40 BOT 100 @752.965 (day-trade buy) -> col1 = -75296.50, no C
    (-75296.50, None),
    # 10:20:08 SOLD 100 @753.2101 (day-trade sell, closes above buy) -> col1 = 75319.44, C = +22.94
    (75319.44, 22.94),
    # 12:40:39 BOT 100 @755.9399 (day-trade buy) -> col1 = -75593.99, no C
    (-75593.99, None),
    # 13:02:01 SOLD 100 @755.9284 (day-trade sell, closes above buy) -> col1 = 75591.26, C = -2.73
    (75591.26, -2.73),
]

for i, (col_a, col_c) in enumerate(details):
    r = detail_start + i
    ws_cs.cell(r, 1).value = col_a
    if i == 0:
        ws_cs.cell(r, 2).value = f'=+B{R}+A{r}'
    else:
        ws_cs.cell(r, 2).value = f'=+B{r-1}+A{r}'
    ws_cs.cell(r, 3).value = col_c

detail_end = detail_start + 7  # 1388

# --- 合計 row ---
r_total = detail_end + 1  # 1389
ws_cs.cell(r_total, 2).value = '合計 '
ws_cs.cell(r_total, 3).value = f'=SUM(C{detail_start}:C{detail_end})'

# --- 核驗 row ---
r_check = r_total + 1  # 1390
ws_cs.cell(r_check, 2).value = '核驗'
ws_cs.cell(r_check, 3).value = f'=+B{detail_end}-B{R}'

# --- blank row ---
r_blank1 = r_check + 1  # 1391

# --- 當沖總結 (4 rows) ---
r_ds1 = r_blank1 + 1  # 1392 - 當沖 買
r_ds2 = r_ds1 + 1     # 1393 - 當沖 賣
r_ds3 = r_ds2 + 1     # 1394 - 當沖損益總額
r_ds4 = r_ds3 + 1     # 1395 - 當沖損益總額 %

ws_cs.cell(r_ds1, 1).value = '2026.06.15當沖'
ws_cs.cell(r_ds1, 2).value = '買'
# 買 side SUMIF: sum of negative col1 values (buys)
# No hold entries to subtract from buy side
ws_cs.cell(r_ds1, 3).value = f'=SUMIF(A{detail_start}:A{detail_end},"<0")'

ws_cs.cell(r_ds2, 1).value = f'="SPY " & COUNT(C{detail_start}:C{detail_end}) * 100 & "股"'
ws_cs.cell(r_ds2, 2).value = '賣'
# 賣 side SUMIF: sum of positive col1 values minus the 4 平倉 exit rows
# The 4 平倉 exits are detail rows 1381, 1382, 1383, 1384
ws_cs.cell(r_ds2, 3).value = f'=SUMIF(A{detail_start}:A{detail_end},">0")-A{detail_start}-A{detail_start+1}-A{detail_start+2}-A{detail_start+3}'

ws_cs.cell(r_ds3, 2).value = '當沖損益總額'
ws_cs.cell(r_ds3, 3).value = f'=+C{r_ds1}+C{r_ds2}'

ws_cs.cell(r_ds4, 2).value = '當沖損益總額 %'
ws_cs.cell(r_ds4, 3).value = f'=C{r_ds3}/ABS(C{r_ds1})'

# --- blank row ---
r_blank2 = r_ds4 + 1  # 1396

# --- 平倉區 (4 groups, LIFO order: 6/11, 6/09, 6/05@748.45, 6/05@749.37) ---
# Each group: 買入 / 平倉 / 損益 / % + blank
# The 6/12 留倉 references: row 1378=6/11, 1377=6/09, 1376=6/05@748.45, 1375=6/05@749.37

pc_groups = [
    # (label, 留倉_ref_row, exit_detail_row, hold_date_label)
    ('2026.06.11 留倉 → 2026.06.15 平倉', 1378, detail_start + 0, '2026.06.11 留倉'),
    ('2026.06.09 留倉 → 2026.06.15 平倉', 1377, detail_start + 1, '2026.06.09 留倉'),
    ('2026.06.05 留倉 → 2026.06.15 平倉', 1376, detail_start + 2, '2026.06.05 留倉(2)'),
    ('2026.06.05 留倉 → 2026.06.15 平倉', 1375, detail_start + 3, '2026.06.05 留倉(1)'),
]

cur = r_blank2 + 1  # 1397
for label, ref_row, exit_row, hd_label in pc_groups:
    # 買入 row
    ws_cs.cell(cur, 1).value = label
    ws_cs.cell(cur, 2).value = '買入 100股'
    ws_cs.cell(cur, 3).value = f'=C{ref_row}'  # reference 留倉 cost
    cur += 1
    # 平倉 row
    ws_cs.cell(cur, 2).value = '平倉 100股'
    ws_cs.cell(cur, 3).value = f'=A{exit_row}'  # exit amount from detail
    cur += 1
    # 損益 row
    ws_cs.cell(cur, 2).value = '損益'
    ws_cs.cell(cur, 3).value = f'=C{cur-2}+C{cur-1}'  # buy_cost + sell_amount
    cur += 1
    # % row
    ws_cs.cell(cur, 2).value = '損益 %'
    ws_cs.cell(cur, 3).value = f'=C{cur-1}/ABS(C{cur-3})'
    cur += 1
    # blank row
    cur += 1

# No 留倉區 (stack is empty after closing all 4 holds, and 2 day-trades are flat)

print(f"CS交易紀錄 6/15 block: rows {R}-{cur-2}")

# ============================================================
# 3. 累積損益check sheet
# ============================================================
ws_cum = wb_cs['累積損益check']
cum_row = ws_cum.max_row + 1  # 27

# Total PnL for 6/15:
# 平倉 PnL: 2398.34 + 932.43 + 412.45 + 354.44 = 4097.66
# 當沖 PnL: 22.94 + (-2.73) = 20.21
# Total: 4097.66 + 20.21 = 4117.87

from datetime import datetime
ws_cum.cell(cum_row, 1).value = datetime(2026, 6, 15)
ws_cum.cell(cum_row, 2).value = 4117.87
ws_cum.cell(cum_row, 3).value = f'=C{cum_row-1}+B{cum_row}'

print(f"累積損益check: row {cum_row} added (PnL=4117.87)")

wb_cs.save('CS交易紀錄.xlsx')
print("CS交易紀錄.xlsx saved")
print("\nAll writes complete.")
