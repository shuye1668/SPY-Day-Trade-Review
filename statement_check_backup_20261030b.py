#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
statement_check.py  —  對帳單完整性檢查（2026-10-22 新增，SHUYE07 維護、經 git 送 Boss PC）

起因：2026-10-22 匯出檔的 Cash Balance 漏列 10/21 22:16:02 SELL -100 SPY @768.95，
但 Order/Trade History 都有、Total Cash 也已入帳。Boss 只看 Cash Balance → 回報「新成交 1 筆」，
SHUYE07 引擎則把當天的 BOT 誤判成留倉。

兩個檢查：
  A. Trade History 每一筆 fill 都要在 Cash Balance 找到對應 TRD 列
     （同方向、同股數、價差 < 0.00005、時間 ±10 秒；TOS 兩表時間偶爾差 1 秒）。
  B. Cash Balance 最後一列 BALANCE 對 Total Cash —— 只當註記、不擋（券商 BAL 重編常讓兩者不等）；
     A 有缺漏時，用差額佐證缺漏成交是否已入 Total Cash。

回測：_inbox 30 份歷史檔只有 2026-10-22.csv 被 A 擋下，其餘 0 誤報。
只檢查、不改 CSV、不補列（絕不寫推算值）。

對帳期間外成交（2026-10-30）：Boss 匯出的 Trade History 會比 Cash Balance 多帶前幾天的成交
（10/30 檔 Cash Balance 從 10/23 13:00 BAL 起，Trade History 卻含 10/22 四筆 BUY）。
早於 Cash Balance 第一列的成交，現金已含在期初 BAL 裡，本來就不會出現在 Cash Balance：
若 trades_all 已有該美東日期、同價的列 → 只當 ℹ️ 註記、不擋；trades_all 沒有 → 照樣算缺漏擋下
（那一天根本還沒入帳，不能靠這份檔放行）。

用法：
  python statement_check.py <csv>
結束碼：0 = 完整（STATEMENT_CHECK=OK）
        4 = 不完整（STATEMENT_CHECK=INCOMPLETE，stdout 列出缺漏，可原樣貼進信件正文）
        2 = 解析失敗（STATEMENT_CHECK=ERROR）
"""

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from spy_daytrade_engine import (parse_statement, parse_total_cash, parse_trade_history,  # noqa: E402
                                  _to_edt)

PRICE_TOL = 0.00005
TIME_TOL_S = 10
CASH_TOL = 0.02
TRADES_PATH_DEFAULT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "trades_all.xlsx")


def _trades_all_keys(trades_path):
    """trades_all → {(美東日期 'YYYY-MM-DD', 價格四捨五入 4 位)}；讀不到回 None（=不放行期間外成交）。"""
    if not trades_path or not os.path.exists(trades_path):
        return None
    try:
        import openpyxl
        wb = openpyxl.load_workbook(trades_path, read_only=True, data_only=True)
    except Exception:  # noqa: BLE001
        return None
    keys = set()
    for r in wb.active.iter_rows(min_row=2, values_only=True):
        if not r or r[0] in (None, "") or r[3] in (None, ""):
            continue
        date = r[0].strftime("%Y-%m-%d") if hasattr(r[0], "strftime") else str(r[0]).strip()[:10]
        try:
            keys.add((date, round(float(r[3]), 4)))
        except (TypeError, ValueError):
            continue
    wb.close()
    return keys


def check(text, trades_path=TRADES_PATH_DEFAULT):
    """回傳 (ok, lines)。"""
    fills = parse_statement(text)
    history = parse_trade_history(text)
    total_cash = parse_total_cash(text)
    lines, ok = [], True

    # 對帳期間起點＝Cash Balance 第一列；更早的成交現金已含在期初 BAL（見檔頭說明）
    start = min((f.cst_dt for f in fills if f.cst_dt is not None), default=None)
    prior, keys = [], None
    if start is not None and any(h[0] < start for h in history):
        keys = _trades_all_keys(trades_path)
        if keys is not None:
            prior = [h for h in history
                     if h[0] < start and (_to_edt(h[0]).strftime("%Y-%m-%d"), round(h[4], 4)) in keys]
            history = [h for h in history if h not in prior]

    # ── A. Trade History ↔ Cash Balance TRD ──
    trd = [f for f in fills if f.typ == "TRD" and f.side]
    used = set()
    missing = []
    for h_dt, h_side, h_qty, h_pe, h_px in history:
        want = "BOT" if h_side == "BUY" else "SOLD"
        hit = None
        for k, f in enumerate(trd):
            if k in used or f.side != want or abs(f.qty) != abs(h_qty):
                continue
            if abs(f.price - h_px) < PRICE_TOL and abs((f.cst_dt - h_dt).total_seconds()) <= TIME_TOL_S:
                hit = k
                break
        if hit is None:
            missing.append((h_dt, h_side, h_qty, h_pe, h_px))
        else:
            used.add(hit)
    if prior:
        days = sorted({_to_edt(h[0]).strftime("%m/%d") for h in prior})
        lines.append(f"A. ℹ️ Trade History 另有 {len(prior)} 筆早於對帳期間（{start:%m/%d %H:%M} 台北）、"
                     f"已在 trades_all（美東 {'、'.join(days)}）→ 不擋")
    if not history and not prior:
        lines.append("A. Trade History：表不存在或無資料 → 未比對（不擋）")
    elif not history:
        pass
    elif missing:
        ok = False
        lines.append(f"A. ❌ Trade History 有 {len(missing)} 筆成交不在 Cash Balance：")
        for h_dt, h_side, h_qty, h_pe, h_px in missing:
            lines.append(f"   - {h_dt:%m/%d/%y %H:%M:%S}（台北）{h_side} {h_qty:+d} SPY @{h_px} {h_pe}")
    else:
        lines.append(f"A. ✓ Trade History {len(history)} 筆全部在 Cash Balance")
    if missing and start is not None and any(m[0] < start for m in missing):
        lines.append("   ↳ 缺漏含對帳期間前的成交，且 trades_all " +
                     ("讀不到" if keys is None else "沒有對應日期/價格") + " → 該日可能尚未入帳")

    # ── B. 最後 BALANCE ↔ Total Cash ──
    # 只當註記、不擋：回測 _inbox 30 份檔，券商 BAL 重編（-153k/-229k）與小額現金差（+47.66）
    # 都會讓兩者不等，但當天帳皆正確。只有 A 抓到缺漏時，用差額佐證缺漏成交是否已入帳。
    bal_rows = [f for f in fills if f.balance is not None]
    last_bal = bal_rows[-1].balance if bal_rows else None
    if total_cash is None or last_bal is None:
        lines.append("B. ℹ️ Total Cash 或 BALANCE 讀不到 → 未比對")
    else:
        diff = round(total_cash - last_bal, 2)
        if abs(diff) <= CASH_TOL:
            lines.append(f"B. ✓ 最後 BALANCE {last_bal:,.2f} = Total Cash {total_cash:,.2f}")
        else:
            lines.append(f"B. ℹ️ 最後 BALANCE {last_bal:,.2f} vs Total Cash {total_cash:,.2f}"
                         f"（差 {diff:+,.2f}；券商重編常見，僅供參考）")
            if missing:
                gross = round(sum(-q * p for _, _, q, _, p in missing), 2)  # SELL -100 → +毛額
                if abs(diff - gross) <= 5.0:   # 容許手續費
                    lines.append(f"   ↳ 差額 ≈ 缺漏成交毛額 {gross:+,.2f}（隱含費用 {diff - gross:+.2f}）"
                                 f" → 缺漏成交已入 Total Cash，只是 Cash Balance 未列出")

    if not ok:
        lines.append("→ Cash Balance 不完整：信照寄，主旨加【CASH BALANCE 不完整】，本段原樣貼進正文；"
                     "不要自行補列，交 SHUYE07 人工判斷。")
    return ok, lines


def main():
    if len(sys.argv) != 2:
        print(__doc__)
        sys.exit(2)
    try:
        with open(sys.argv[1], encoding="utf-8-sig") as fh:
            text = fh.read()
        ok, lines = check(text)
    except Exception as e:  # noqa: BLE001
        print(f"STATEMENT_CHECK=ERROR {type(e).__name__}: {e}")
        sys.exit(2)
    print("STATEMENT_CHECK=" + ("OK" if ok else "INCOMPLETE"))
    for ln in lines:
        print(ln)
    sys.exit(0 if ok else 4)


if __name__ == "__main__":
    main()
