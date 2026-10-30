#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
julia_block_spec.py — 給 julia_block_writer.ps1 用的「Julia 母檔個別識別法區塊規格」產生器（2026-10-20）

讀對帳單 CSV（截圖來源先依 MIGRATION.md §4.1 轉成同格式 CSV），跑 spy_daytrade_engine，
把指定交易日的每一腿輸出成 JSON（UTF-8，印到 stdout）。所有數字都取自引擎，不推算。

腿的種類 kind：
  day   當沖（進場或出場；出場腿帶 pair = 同區塊內配對進場腿的索引、pnl）
  hold  收盤留倉建立腿（Status 1）
  close 跨日平倉腿（Status 2；帶 entry = 原留倉腿的日期/時間/淨額、pnl）

ok=false 時 reasons 列出擋下原因（BALANCE 對不上、非交易現金、非 100 股、其他 ⚠️），
writer 一律不寫。留倉/跨日平倉本身不擋（§3b 已核定）。
"""
import argparse
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from spy_daytrade_engine import (parse_statement, group_sessions, process_all,
                                 TRADES_PATH_DEFAULT)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("csv_path")
    ap.add_argument("--date", required=True, help="EDT 交易日 YYYY-MM-DD")
    ap.add_argument("--trades", default=TRADES_PATH_DEFAULT)
    a = ap.parse_args()

    text = open(a.csv_path, encoding="utf-8-sig").read()
    fills = parse_statement(text)
    sessions = group_sessions(fills)
    stats = process_all(sessions, fills, text, a.trades)
    idx = next((i for i, s in enumerate(sessions) if s.edt_date == a.date), None)
    if idx is None:
        print(json.dumps({"ok": False, "reasons": [f"CSV 內找不到 {a.date} 的 session"]},
                         ensure_ascii=False))
        return
    s, st = sessions[idx], stats[idx]

    reasons = []
    if s.open_bal is None or s.close_bal is None:
        reasons.append("開盤/收盤餘額不明")
    for w in s.warnings:
        if w.startswith("留倉：") or w.startswith("跨日平倉："):
            continue
        reasons.append(w)
    for t in s.trades:
        if t.qty != 100:
            reasons.append(f"{t.edt_time} 合併後 {t.qty} 股（非 100）")
        if t.needs_review and not (t.review_reason.startswith("收盤留倉")
                                   or t.review_reason.startswith("跨日平倉（Status=2）")):
            reasons.append(f"{t.edt_time} {t.review_reason}")

    legs = []
    pos = {id(t): i for i, t in enumerate(s.trades)}
    for t in s.trades:
        if t.type_tag == "平":
            kind = "close"
        elif t.status == "1" or (t.status == "2" and not t.type_tag and t.edt_date == s.edt_date
                                 and t.pair_date != s.edt_date):
            kind = "hold"
        else:
            kind = "day"
        leg = {
            "time_edt": t.edt_time,
            "time_tw": t.cst_dt.strftime("%H:%M:%S") if t.cst_dt else "",
            "side": t.side, "price": t.price, "net": t.net_cash,
            "fee": round(-t.misc_fee, 2),                      # 正數
            "gross": round(t.net_cash - t.misc_fee, 2),        # 賣腿 = 對帳單 AMOUNT
            "kind": kind, "type_tag": t.type_tag, "pnl": t.pnl,
        }
        if kind == "day" and t.pnl is not None:
            entry = next((e for e in s.trades if e is not t and e.edt_time == t.pair_time
                          and e.edt_date == t.pair_date), None)
            leg["pair"] = pos.get(id(entry)) if entry else None
            if leg["pair"] is None:
                reasons.append(f"{t.edt_time} 當沖出場找不到同區塊進場腿")
        if kind == "close":
            leg["entry"] = {"date": t.pair_date, "time": t.pair_time}
            en_net = round(t.pnl - t.net_cash, 2) if t.pnl is not None else None
            leg["entry"]["net"] = en_net
        legs.append(leg)

    holds_after = [{"date": l["trade"].edt_date, "time": l["trade"].edt_time,
                    "net": l["trade"].net_cash, "qty": l["qty_rem"]}
                   for l in st["carry"] if l["trade"].side == "BOT"]
    day_pnl = round(sum(l["pnl"] for l in legs if l["kind"] == "day" and l["pnl"] is not None), 2)
    close_pnl = round(sum(l["pnl"] for l in legs if l["kind"] == "close" and l["pnl"] is not None), 2)
    # 總結行（§5 格式、與明細信主旨同款）：只第一行帶日期；周X 用交易日實際星期
    y, m, d = (int(x) for x in s.edt_date.split("-"))
    import datetime as _dt
    wd = "一二三四五六日"[_dt.date(y, m, d).weekday()]
    head = f"{m:02d}/{d:02d}(周{wd})"
    fmt = lambda v: f"{v:+.2f}"
    lines = []
    day_exits = [l for l in legs if l["kind"] == "day" and l["pnl"] is not None]
    if day_exits:
        buy = sum(abs(legs[l["pair"]]["net"]) if legs[l["pair"]]["side"] == "BOT" else abs(l["net"])
                  for l in day_exits)
        lines.append(f"SPY當沖 {len(day_exits) * 100}股 損益 {fmt(day_pnl)} {fmt(day_pnl / buy * 100)}%")
    closes = [l for l in legs if l["kind"] == "close" and l["pnl"] is not None]
    if closes:
        cost = sum(abs(l["entry"]["net"]) for l in closes)
        lines.append(f"SPY平倉 {len(closes) * 100}股 損益 {fmt(close_pnl)} {fmt(close_pnl / cost * 100)}%")
    for l in legs:
        if l["kind"] == "hold":
            lines.append(f"SPY留倉 100股 * {l['price']:.4f}".rstrip("0").rstrip("."))
    if lines:
        lines[0] = head + lines[0]

    out = {
        "ok": not reasons, "reasons": reasons,
        "summary_lines": lines,
        "date": s.edt_date, "dot": s.edt_date.replace("-", "."),
        "open_bal": s.open_bal, "close_bal": s.close_bal,
        "legs": legs, "day_pnl": day_pnl, "close_pnl": close_pnl,
        "total_pnl": round(day_pnl + close_pnl, 2),
        "holds_after": holds_after,
        "notes": s.notes,
    }
    print(json.dumps(out, ensure_ascii=False))


if __name__ == "__main__":
    main()
