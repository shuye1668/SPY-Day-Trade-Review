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
  python statement_check.py <csv> [--repair <out.csv>]
結束碼：0 = 完整（STATEMENT_CHECK=OK）
        4 = 不完整（STATEMENT_CHECK=INCOMPLETE，stdout 列出缺漏，可原樣貼進信件正文）
        3 = 不完整但已自動補列寫到 out.csv（--repair，2026-10-30 起排程用；後續步驟改用 out.csv 並標警示）
        2 = 解析失敗（STATEMENT_CHECK=ERROR）
"""

import csv
import io
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from spy_daytrade_engine import (parse_statement, parse_total_cash, parse_trade_history,  # noqa: E402
                                  _to_edt, _parse_cst_dt)

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


FEE_RATE_EST = 0.0000206   # 賣方 SEC+TAF 約 1.58/76,900；Total Cash 推不出時才用


def _missing_fills(text, trades_path):
    """check() 的 A 段缺漏清單（不含已在 trades_all 的期間外成交）。"""
    fills = parse_statement(text)
    history = parse_trade_history(text)
    start = min((f.cst_dt for f in fills if f.cst_dt is not None), default=None)
    keys = _trades_all_keys(trades_path) if start is not None else None
    if keys is not None:
        history = [h for h in history if not (h[0] < start and
                   (_to_edt(h[0]).strftime("%Y-%m-%d"), round(h[4], 4)) in keys)]
    trd = [f for f in fills if f.typ == "TRD" and f.side]
    used, missing = set(), []
    for h in history:
        want = "BOT" if h[1] == "BUY" else "SOLD"
        hit = next((k for k, f in enumerate(trd) if k not in used and f.side == want
                    and abs(f.qty) == abs(h[2]) and abs(f.price - h[4]) < PRICE_TOL
                    and abs((f.cst_dt - h[0]).total_seconds()) <= TIME_TOL_S), None)
        if hit is None:
            missing.append(h)
        else:
            used.add(hit)
    return fills, missing


def repair(text, trades_path=TRADES_PATH_DEFAULT):
    """把 Trade History 有、Cash Balance 漏列的成交補進 Cash Balance（2026-10-30，取代人工補列）。

    每筆補一列 TRD（REF=AUTO…，同時間同價同方向的部分成交共用 REF 讓引擎合併）：
    AMOUNT=毛額；賣方費用＝(Total Cash − 最後 BALANCE − Σ缺漏毛額) 依毛額比例分攤，
    推不出或不合理時用 FEE_RATE_EST 估；插在時間序位置，其後到下一個 BAL 之前的 BALANCE 一併加上。
    回傳 (new_text, lines)；沒有缺漏回傳 (None, lines)。"""
    fills, missing = _missing_fills(text, trades_path)
    if not missing:
        return None, ["無缺漏，不需補列"]
    total_cash = parse_total_cash(text)
    bal_rows = [f for f in fills if f.balance is not None]
    gross = [round(-q * p, 2) for _, _, q, _, p in missing]          # BUY +100 → 負；SELL -100 → 正
    sell_gross = sum(g for g in gross if g > 0)
    fee_total = None
    if total_cash is not None and bal_rows and sell_gross:
        f = round(total_cash - bal_rows[-1].balance - sum(gross), 2)
        if -sell_gross * 0.0002 <= f <= 0:
            fee_total = f
    fees = []
    for g in gross:
        if g <= 0:
            fees.append(0.0)
        elif fee_total is not None:
            fees.append(round(fee_total * g / sell_gross, 2))
        else:
            fees.append(-round(g * FEE_RATE_EST, 2))
    if fee_total is not None and sell_gross:
        k = max(i for i, g in enumerate(gross) if g > 0)
        fees[k] = round(fees[k] + fee_total - sum(fees), 2)          # 分攤尾差歸最後一筆賣出

    lines_in = text.splitlines()
    hdr = next(i for i, ln in enumerate(lines_in) if ln.startswith("DATE,TIME,TYPE") or
               ("Exec Time" in ln and "Balance" in ln))
    end = hdr + 1
    while end < len(lines_in) and lines_in[end][:1].isdigit():
        end += 1
    rows = []
    for ln in lines_in[hdr + 1:end]:
        c = next(csv.reader(io.StringIO(ln)))
        rows.append({"c": c, "dt": _parse_cst_dt(c[0], c[1]), "typ": c[2].strip(), "auto": False})
    refs = {}
    for (h_dt, h_side, h_qty, _, h_px), g, fee in zip(missing, gross, fees):
        key = (h_side, round(h_px, 4), h_dt.strftime("%Y%m%d%H%M"))
        ref = refs.setdefault(key, f"AUTO{h_dt:%Y%m%d%H%M%S}")
        side = "BOT" if h_side == "BUY" else "SOLD"
        c = [f"{h_dt:%m/%d/%y}", f"{h_dt:%H:%M:%S}", "TRD", f'="{ref}"',
             f"{side} {h_qty:+d} SPY @{h_px:g}", f"{fee:.2f}" if fee else "", "", f"{g:,.2f}", ""]
        pos = len(rows)
        for i, r in enumerate(rows):
            if r["dt"] > h_dt:
                pos = i
                break
        rows.insert(pos, {"c": c, "dt": h_dt, "typ": "TRD", "auto": True, "delta": round(g + fee, 2)})
    bal = None
    shift = 0.0
    for r in rows:
        c = r["c"]
        if r["typ"] == "BAL":
            shift = 0.0
            bal = _num_csv(c[8])
            continue
        if r["auto"]:
            shift = round(shift + r["delta"], 2)
            prev = bal if bal is not None else 0.0
            bal = round(prev + r["delta"], 2)
            c[8] = f"{bal:,.2f}"
        else:
            v = _num_csv(c[8])
            if v is not None:
                bal = round(v + shift, 2)
                c[8] = f"{bal:,.2f}"
    out_rows = []
    for r in rows:
        buf = io.StringIO()
        csv.writer(buf, lineterminator="").writerow(r["c"])
        out_rows.append(buf.getvalue().replace('"=""', '="').replace('""",', '",'))
    new_text = "\n".join(lines_in[:hdr + 1] + out_rows + lines_in[end:])
    info = [f"已自動補列 {len(missing)} 筆（REF AUTO…；費用 "
            + ("由 Total Cash 推算" if fee_total is not None else "依費率估計") + "）："]
    for (h_dt, h_side, h_qty, h_pe, h_px), g, fee in zip(missing, gross, fees):
        info.append(f"   + {h_dt:%m/%d/%y %H:%M:%S}（台北）{h_side} {h_qty:+d} SPY @{h_px} 毛額 {g:+,.2f} 費用 {fee:+.2f}")
    return new_text, info


def _num_csv(s):
    s = (s or "").strip().replace(",", "").replace("$", "")
    if not s:
        return None
    neg = s.startswith("(") and s.endswith(")")
    v = float(s.strip("()"))
    return -v if neg else v


def main():
    args = sys.argv[1:]
    out = None
    if "--repair" in args:
        i = args.index("--repair")
        out = args[i + 1] if i + 1 < len(args) else None
        args = args[:i] + args[i + 2:]
    if len(args) != 1 or ("--repair" in sys.argv and not out):
        print(__doc__)
        sys.exit(2)
    try:
        with open(args[0], encoding="utf-8-sig") as fh:
            text = fh.read()
        ok, lines = check(text)
        if not ok and out:
            new_text, info = repair(text)
            if new_text is not None:
                ok2, lines2 = check(new_text)
                with open(out, "w", encoding="utf-8") as fh:
                    fh.write(new_text + "\n")
                print("STATEMENT_CHECK=REPAIRED" if ok2 else "STATEMENT_CHECK=REPAIRED_STILL_INCOMPLETE")
                for ln in lines + info + ["補列後重檢："] + lines2:
                    print(ln)
                print(f"→ 已寫 {out}，後續步驟改用這份（writer/Julia 會標 ⚠️待修正）")
                sys.exit(3)
    except Exception as e:  # noqa: BLE001
        print(f"STATEMENT_CHECK=ERROR {type(e).__name__}: {e}")
        sys.exit(2)
    print("STATEMENT_CHECK=" + ("OK" if ok else "INCOMPLETE"))
    for ln in lines:
        print(ln)
    sys.exit(0 if ok else 4)


if __name__ == "__main__":
    main()
