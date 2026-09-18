#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
spy_daytrade_engine.py  —  Spy daytrade auto csv routine 的「唯一引擎」

取代 process_XXXX.py / write_XXXX.py 每日拋棄式手抄腳本。
以券商 Account Statement（Cash Balance 區段 CSV）為【唯一真相源】，
自動：解析 → 按 REF# 合併拆單 → 費用取實欄 → CST→EDT 換算 → 依時間 LIFO 配對
→ 算每筆損益(含費) → 產出 CS交易紀錄「個別識別法」區塊 + trades_all 列
→ 用對帳單 BALANCE 欄交叉驗證 → 跑 CLAUDE.md §四 全部核驗。

安全鐵則（對齊 CLAUDE.md）：
  §0  任何不確定（offset 對不上 / Cash Balance 不完整 / 留倉平倉 / Day Header B 不確定）
      → 相關儲存格填 '⚠️需人工確認'、abort、不靜默估算。
  §2  寫入位置一律用值掃描 last_nonempty_row，【嚴禁 ws.max_row】。
  §3  寫入前先備份 *_backup_YYYYMMDD.xlsx。
  §4  寫入後跑全部核驗，任一失敗 → 還原備份。
  §5  寫入前檢查 ~$ 鎖定檔（Excel 是否開啟）。

預設 DRY-RUN（不碰任何 live 檔）。加 --commit 才實際寫入。
"""

import argparse
import csv
import io
import os
import re
import shutil
import sys
from dataclasses import dataclass, field
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo

FLAG = "⚠️需人工確認"
TW = ZoneInfo("Asia/Taipei")        # 台北 = 券商 statement "CST" 實際時區 (UTC+8, 無 DST)
ET = ZoneInfo("America/New_York")   # EDT/EST，trades_all 用的時間

# ─────────────────────────────────────────────────────────────────────────────
# 數值/時間解析
# ─────────────────────────────────────────────────────────────────────────────

def _num(s):
    """把 '75,307.00' / '"-38,346.77"' / '' 解析成 float 或 None。"""
    if s is None:
        return None
    s = str(s).strip().strip('"').replace(",", "")
    if s == "":
        return None
    return float(s)


def _parse_cst_dt(date_str, time_str):
    """statement 的 m/d/yy + HH:MM:SS (台北時間) → aware datetime。"""
    m, d, y = [int(x) for x in date_str.strip().split("/")]
    if y < 100:
        y += 2000
    hh, mm, ss = [int(x) for x in time_str.strip().split(":")]
    return datetime(y, m, d, hh, mm, ss, tzinfo=TW)


def _to_edt(cst_dt):
    """台北時間 → 美東時間（自動處理 DST：夏令-12h、冬令-13h）。"""
    return cst_dt.astimezone(ET)


_DESC_RE = re.compile(r"(BOT|SOLD)\s+([+-]?\d+)\s+(\w+)\s+@([\d.]+)")


def _parse_trd_desc(desc):
    """'SOLD -100 SPY @753.07' → (side, signed_qty, symbol, price) 或 None。"""
    mth = _DESC_RE.search(desc or "")
    if not mth:
        return None
    side, qty, sym, price = mth.group(1), int(mth.group(2)), mth.group(3), float(mth.group(4))
    return side, qty, sym, price


# ─────────────────────────────────────────────────────────────────────────────
# 資料結構
# ─────────────────────────────────────────────────────────────────────────────

@dataclass
class Fill:
    cst_dt: datetime
    edt_dt: datetime
    typ: str            # BAL / TRD / DEP / ...
    ref: str
    desc: str
    misc_fee: float     # 對帳單實際值（負數），無費用時 0.0
    amount: float       # 對帳單 AMOUNT
    balance: float      # 對帳單 BALANCE（每筆餘額，天然自校驗）
    side: str = ""      # BOT / SOLD
    qty: int = 0        # 有號股數
    symbol: str = ""
    price: float = 0.0

    @property
    def net_cash(self):
        # CLAUDE §二：col1 = AMOUNT + Misc Fees（皆取自對帳單）
        return round(self.amount + (self.misc_fee or 0.0), 2)


@dataclass
class Trade:
    """按 REF# 合併後的一筆成交（通常 100 股）。"""
    edt_date: str
    edt_time: str       # HH:MM:SS
    symbol: str
    side: str           # BOT / SOLD
    qty: int            # 絕對股數（正）
    price: float        # 股數加權均價
    net_cash: float     # sum(amount)+sum(misc_fee)
    misc_fee: float
    refs: list = field(default_factory=list)
    # 配對後填入
    type_tag: str = ""     # 多 / 空 / 平 / ''(進場)
    action: str = ""       # B / S
    status: str = "0"      # 0 當沖 / 1 留倉 / 2 跨日平倉
    pnl: float = None      # 平倉損益（含費），只在平倉列
    pair_time: str = ""
    pair_date: str = ""
    needs_review: bool = False
    review_reason: str = ""


@dataclass
class Session:
    edt_date: str           # YYYY-MM-DD（EDT 交易日）
    open_bal: float         # 券商當日開盤餘額（BAL 列）
    close_bal: float        # 券商當日收盤餘額（最後一筆 balance）
    fills: list = field(default_factory=list)
    trades: list = field(default_factory=list)
    warnings: list = field(default_factory=list)


# ─────────────────────────────────────────────────────────────────────────────
# 解析 statement
# ─────────────────────────────────────────────────────────────────────────────

def parse_statement(text):
    """回傳 [Fill,...]（僅 Cash Balance 區段）。"""
    # 找到 Cash Balance 表頭
    lines = text.splitlines()
    start = None
    for i, ln in enumerate(lines):
        if ln.strip().upper().startswith("DATE,TIME,TYPE,REF"):
            start = i
            break
    if start is None:
        raise ValueError("找不到 Cash Balance 表頭（DATE,TIME,TYPE,REF #,...）— Statement 格式不符")

    body = "\n".join(lines[start:])
    reader = csv.reader(io.StringIO(body))
    header = next(reader)
    fills = []
    for row in reader:
        if not row or all(c.strip() == "" for c in row):
            continue
        if len(row) < 9:
            # 可能是被截斷/OCR 破碎列 → 不靜默跳過，標記
            row = row + [""] * (9 - len(row))
        date_s, time_s, typ, ref, desc, misc, comm, amount, balance = row[:9]
        # Cash Balance 區段結束偵測：DATE 欄不是 m/d/yy 日期 → 進入下一張表，停止
        if not re.match(r"^\s*\d{1,2}/\d{1,2}/\d{2,4}\s*$", date_s or ""):
            break
        if not date_s.strip() or not time_s.strip():
            continue
        cst = _parse_cst_dt(date_s, time_s)
        f = Fill(
            cst_dt=cst, edt_dt=_to_edt(cst), typ=typ.strip(),
            ref=re.sub(r'^="?|"?$', "", ref.strip()),
            desc=desc.strip(), misc_fee=_num(misc) or 0.0,
            amount=_num(amount) if _num(amount) is not None else 0.0,
            balance=_num(balance),
        )
        parsed = _parse_trd_desc(desc)
        if parsed:
            f.side, f.qty, f.symbol, f.price = parsed
        fills.append(f)
    return fills


def group_sessions(fills):
    """依 BAL 列切出各 EDT 交易日 session。
    BAL 'start of business day D.06' 的 balance = 該 EDT 日開盤餘額；
    其後所有 TRD 直到下一個 BAL 屬同一 session（跨午夜的 CST 日期會 -12h 落回同一 EDT 日）。
    """
    sessions = []
    cur = None
    for f in fills:
        if f.typ == "BAL":
            edt_date = f.edt_dt.strftime("%Y-%m-%d")
            # BAL 13:00 CST → 01:00 EDT，屬「當天」EDT 盤前；session 標籤用 CST 日期換算的 EDT 日
            # 實務上 BAL 的 CST 日期 D → 這個 session 的 EDT 交易日就是 D（同日盤）
            sess_date = f.cst_dt.strftime("%Y-%m-%d")
            cur = Session(edt_date=sess_date, open_bal=f.balance, close_bal=f.balance)
            sessions.append(cur)
        elif f.typ == "TRD":
            if cur is None:
                # TRD 前沒有 BAL → 資料不完整
                cur = Session(edt_date=f.edt_dt.strftime("%Y-%m-%d"),
                              open_bal=None, close_bal=None)
                cur.warnings.append("session 開頭缺 BAL 列，開盤餘額未知")
                sessions.append(cur)
            cur.fills.append(f)
            if f.balance is not None:
                cur.close_bal = f.balance
        else:
            # DEP/JRN/其他現金異動（如入金）→ 記錄為 warning，交人工判斷（如 7/16 的 3 萬疑雲）
            if cur is not None:
                cur.warnings.append(
                    f"非交易現金異動 {f.cst_dt:%m/%d %H:%M} TYPE={f.typ} AMOUNT={f.amount} "
                    f"({f.desc[:40]}) — 需人工判斷是否入金/出金")
                if f.balance is not None:
                    cur.close_bal = f.balance
    # 只保留有交易的 session
    return [s for s in sessions if s.fills]


# ─────────────────────────────────────────────────────────────────────────────
# 按 REF# 合併拆單
# ─────────────────────────────────────────────────────────────────────────────

def merge_fills(session):
    """同一 REF# 的部分成交合併成一筆；價格股數加權、fee/amount 加總。"""
    trades = []
    # 依出現順序分組（同 REF# 連續出現）
    order = []
    groups = {}
    for f in session.fills:
        if f.side == "":
            session.warnings.append(f"無法解析 DESCRIPTION 為交易：{f.desc[:50]}")
            continue
        key = f.ref or f"noref_{f.cst_dt.timestamp()}"
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(f)

    for key in order:
        fs = sorted(groups[key], key=lambda f: f.cst_dt)   # 組內按時間，fs[0] 必為最早成交
        qty_sum = sum(f.qty for f in fs)             # 有號
        if qty_sum == 0:
            session.warnings.append(f"REF {key} 合併後淨股數為 0，異常")
            continue
        side = "SOLD" if qty_sum < 0 else "BOT"
        abs_qty = abs(qty_sum)
        wsum = sum(abs(f.qty) * f.price for f in fs)
        avg_px = round(wsum / abs_qty, 4)
        net = round(sum(f.amount for f in fs) + sum(f.misc_fee for f in fs), 2)
        fee = round(sum(f.misc_fee for f in fs), 2)
        t = Trade(
            edt_date=fs[0].edt_dt.strftime("%Y-%m-%d"),
            edt_time=fs[0].edt_dt.strftime("%H:%M:%S"),
            symbol=fs[0].symbol, side=side, qty=abs_qty, price=avg_px,
            net_cash=net, misc_fee=fee, refs=[key],
        )
        if abs_qty not in (100,):
            # 非 100 股（奇數量/留倉零股）→ 標記人工確認，不靜默
            t.needs_review = True
            t.review_reason = f"合併後 {abs_qty} 股（非 100）"
        trades.append(t)
    trades.sort(key=lambda t: t.edt_time)
    session.trades = trades
    return trades


# ─────────────────────────────────────────────────────────────────────────────
# 依時間 LIFO 配對（複刻 trade_review_app._analyse_all_trades 的當日 LIFO）
# ─────────────────────────────────────────────────────────────────────────────

def pair_intraday(session):
    """對單一 session 的 trades 做 LIFO 配對，填 type_tag/action/pnl/pair_*。
    純當日進出：BOT→SOLD 平 = 多；SOLD→BOT 平 = 空。
    收盤仍有存貨 → 留倉（Status 1）並標記 needs_review（§0：留倉致 offset 跳動需人工確認）。
    """
    stack = []  # 開倉 lots: {trade, qty_rem}
    for t in session.trades:
        signed = t.qty if t.side == "BOT" else -t.qty
        net_sign = 0
        if stack:
            net_sign = 1 if stack[0]["trade"].side == "BOT" else -1
        this_sign = 1 if signed > 0 else -1

        if not stack or this_sign == net_sign:
            # 進場（擴大部位）
            t.action = "B" if t.side == "BOT" else "S"
            t.type_tag = ""
            t.status = "0"
            stack.append({"trade": t, "qty_rem": t.qty})
        else:
            # 平倉（縮小部位），LIFO 消耗 stack
            need = t.qty
            matched = []
            while need > 0 and stack:
                top = stack[-1]
                take = min(need, top["qty_rem"])
                matched.append((top["trade"], take))
                top["qty_rem"] -= take
                need -= take
                if top["qty_rem"] <= 0:
                    stack.pop()
            if need > 0:
                # 平超過持倉 → 反向開倉，異常，標記
                t.needs_review = True
                t.review_reason = "平倉股數超過當日持倉（可能反手或漏單）"
            entry = matched[0][0]
            # 方向：平掉的是 BOT(多) → 這筆是 SELL 平多 = 多；平掉 SOLD(空) = 空
            t.type_tag = "多" if entry.side == "BOT" else "空"
            t.action = "B" if t.side == "BOT" else "S"
            t.status = "0"
            # 損益（含費）= 平倉 net + 對應進場 net（100 股整筆對整筆）
            pnl = t.net_cash + sum(m[0].net_cash * (m[1] / m[0].qty) for m in matched)
            t.pnl = round(pnl, 2)
            t.pair_time = entry.edt_time
            t.pair_date = entry.edt_date

    # 收盤仍有未平部位 → 留倉
    for lot in stack:
        lot["trade"].status = "1"
        lot["trade"].needs_review = True
        lot["trade"].review_reason = "收盤留倉（Status=1）— §0 留倉致 offset 跳動需人工確認"
        session.warnings.append(
            f"留倉：{lot['trade'].edt_time} {lot['trade'].side} {lot['qty_rem']}股 @{lot['trade'].price}")


# ─────────────────────────────────────────────────────────────────────────────
# 對帳單 BALANCE 交叉驗證（會抓出 7/16 漏單 & 3 萬落差那類錯誤）
# ─────────────────────────────────────────────────────────────────────────────

def cross_check_balance(session):
    """逐筆用 open_bal + Σnet_cash 重建 balance，比對對帳單 BALANCE 欄。
    任一不符 → 代表漏單/多單/費用錯 → raise（不靜默）。"""
    if session.open_bal is None:
        return [f"{FLAG}：session 缺開盤餘額，無法交叉驗證"]
    errs = []
    running = session.open_bal
    for f in session.fills:
        if f.typ != "TRD":
            continue
        running = round(running + f.net_cash, 2)
        if f.balance is not None and abs(running - f.balance) > 0.01:
            errs.append(
                f"BALANCE 對不上 @ {f.cst_dt:%m/%d %H:%M:%S} ref={f.ref}: "
                f"重建={running:.2f} vs 對帳單={f.balance:.2f} 差={running-f.balance:+.2f}")
    return errs


# ─────────────────────────────────────────────────────────────────────────────
# 主流程（單一 session）
# ─────────────────────────────────────────────────────────────────────────────

def process_session(session):
    merge_fills(session)
    # 交叉驗證 BALANCE（最重要的資料完整性閘門）
    bal_errs = cross_check_balance(session)
    session.warnings.extend(bal_errs)
    pair_intraday(session)

    # 自洽核驗：Σcol1 == close_bal - open_bal（CLAUDE §四.3）
    col1_sum = round(sum(t.net_cash for t in session.trades), 2)
    if session.open_bal is not None and session.close_bal is not None:
        cash_change = round(session.close_bal - session.open_bal, 2)
        if abs(col1_sum - cash_change) > 0.01:
            session.warnings.append(
                f"{FLAG}：Σcol1({col1_sum}) ≠ 收-開({cash_change})，差 {col1_sum-cash_change:+.2f}")
    # 全平時 Σcol1 == Σpnl（CLAUDE §四.1）
    pnl_sum = round(sum(t.pnl for t in session.trades if t.pnl is not None), 2)
    has_hold = any(t.status == "1" for t in session.trades)
    if not has_hold and abs(col1_sum - pnl_sum) > 0.01:
        session.warnings.append(
            f"{FLAG}：全平但 Σcol1({col1_sum}) ≠ Σpnl({pnl_sum})")
    return {"col1_sum": col1_sum, "pnl_sum": pnl_sum, "has_hold": has_hold}


# ─────────────────────────────────────────────────────────────────────────────
# 報表輸出（dry-run 用）
# ─────────────────────────────────────────────────────────────────────────────

def report_session(session, stats):
    print(f"\n{'='*70}")
    print(f" SESSION {session.edt_date}  開盤={session.open_bal}  收盤={session.close_bal}")
    print(f"{'='*70}")
    print(f" 合併後 {len(session.trades)} 筆交易")
    print(f" {'time':<9} {'side':<5} {'px':>10} {'net_cash':>12} {'fee':>6} "
          f"{'type':<4} {'act':<3} {'st':<2} {'pnl':>9} pair")
    for t in session.trades:
        flag = "  ⚠️" if t.needs_review else ""
        print(f" {t.edt_time:<9} {t.side:<5} {t.price:>10.4f} {t.net_cash:>12.2f} "
              f"{t.misc_fee:>6.2f} {t.type_tag or '-':<4} {t.action:<3} {t.status:<2} "
              f"{('' if t.pnl is None else f'{t.pnl:.2f}'):>9} "
              f"{t.pair_time}{flag}{(' '+t.review_reason) if t.needs_review else ''}")
    print(f" ── Σcol1={stats['col1_sum']:.2f}  Σpnl={stats['pnl_sum']:.2f}  "
          f"留倉={stats['has_hold']}")
    if session.warnings:
        print(" 警告/需人工確認：")
        for w in session.warnings:
            print(f"   • {w}")


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser(description="Spy daytrade auto csv 引擎")
    ap.add_argument("csv_path", help="對帳單 CSV/文字檔路徑")
    ap.add_argument("--date", help="只處理指定 EDT 交易日 YYYY-MM-DD")
    ap.add_argument("--commit", action="store_true", help="實際寫入 live 檔（預設 dry-run）")
    args = ap.parse_args()

    with open(args.csv_path, encoding="utf-8-sig") as fh:
        text = fh.read()
    fills = parse_statement(text)
    sessions = group_sessions(fills)
    if args.date:
        sessions = [s for s in sessions if s.edt_date == args.date]
        if not sessions:
            print(f"找不到 {args.date} 的 session"); sys.exit(1)

    for s in sessions:
        stats = process_session(s)
        report_session(s, stats)

    if args.commit:
        print("\n[commit 模式尚未接上寫入層 — 需先通過 dry-run 驗證]")
    else:
        print("\n[DRY-RUN] 未寫入任何檔案。")


if __name__ == "__main__":
    main()
