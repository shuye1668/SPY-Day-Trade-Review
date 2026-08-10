import sys
sys.stdout.reconfigure(encoding='utf-8')

trades_raw = [
    ('6/4/26', '22:14:11', 'BOT', 100, 'SPY', 753.175,   0,    -75317.50, '1006606840770'),
    ('6/4/26', '22:23:39', 'SOLD', 100, 'SPY', 753.86,   1.57,  75386.00, '1006607291593'),
    ('6/4/26', '22:29:30', 'BOT', 100, 'SPY', 754.0544,  0,    -75405.44, '1006607743204'),
    ('6/4/26', '22:52:51', 'SOLD', 100, 'SPY', 753.7901,  1.57,  75379.01, '1006608921970'),
    ('6/4/26', '23:06:12', 'BOT', 100, 'SPY', 754.7288,  0,    -75472.88, '1006609793178'),
    ('6/4/26', '23:07:09', 'SOLD', 100, 'SPY', 754.81,   1.57,  75481.00, '1006609793427'),
    ('6/4/26', '23:10:15', 'BOT',  60, 'SPY', 755.17,    0,    -45310.20, '1006610240223'),
    ('6/4/26', '23:10:15', 'BOT',  40, 'SPY', 755.17,    0,    -30206.80, '1006610240223'),
    ('6/4/26', '23:11:02', 'SOLD', 100, 'SPY', 755.4101,  1.58,  75541.01, '1006610240406'),
    ('6/4/26', '23:13:42', 'BOT', 100, 'SPY', 755.459,   0,    -75545.90, '1006610241023'),
    ('6/4/26', '23:15:18', 'SOLD', 100, 'SPY', 755.61,   1.58,  75561.00, '1006610241459'),
    ('6/5/26', '00:05:34', 'BOT', 100, 'SPY', 755.54,    0,    -75554.00, '1006612633335'),
    ('6/5/26', '00:23:42', 'SOLD', 100, 'SPY', 756.60,   1.58,  75660.00, '1006613109913'),
]

merged = []
i = 0
while i < len(trades_raw):
    t = trades_raw[i]
    date, time, action, qty, sym, price, misc, amount, ref = t
    total_qty = qty
    total_amount = amount
    total_misc = misc
    j = i + 1
    while j < len(trades_raw) and trades_raw[j][8] == ref:
        total_qty += trades_raw[j][3]
        total_amount += trades_raw[j][7]
        total_misc += trades_raw[j][5]
        j += 1
    col1 = total_amount - abs(total_misc)
    act = 'B' if action == 'BOT' else 'S'
    merged.append({'action': act, 'price': price, 'col1': col1, 'qty': total_qty, 'ref': ref})
    if j > i + 1:
        print(f'MERGE idx {i}-{j-1}: qty={total_qty} amt={total_amount} misc={total_misc} col1={col1}')
    i = j

print()
for idx, m in enumerate(merged):
    act = m['action']
    pr = m['price']
    c1 = m['col1']
    print(f'{idx}: {act} qty={m["qty"]} price={pr} col1={c1:.2f}')

print()
stack = []
for idx in range(len(merged)):
    t = merged[idx]
    if not stack:
        stack.append(idx)
    elif (t['action'] == 'S' and merged[stack[-1]]['action'] == 'B') or \
         (t['action'] == 'B' and merged[stack[-1]]['action'] == 'S'):
        entry_idx = stack.pop()
        entry = merged[entry_idx]
        pnl = t['col1'] + entry['col1']
        ea = entry['action']
        ep = entry['price']
        ec = entry['col1']
        ta = t['action']
        tp = t['price']
        tc = t['col1']
        print(f'Pair: {entry_idx}({ea}@{ep}) -> {idx}({ta}@{tp}) | entry_col1={ec:.2f} exit_col1={tc:.2f} | PnL={pnl:.2f}')
    else:
        stack.append(idx)
