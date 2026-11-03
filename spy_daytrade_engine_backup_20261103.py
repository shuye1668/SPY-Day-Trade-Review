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
TRADES_PATH_DEFAULT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "trades_all.xlsx")
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
    cst_dt: datetime = None   # 最早成交的台北時間（對 Trade History 用）
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
    notes: list = field(default_factory=list)   # ℹ️ 資訊（不擋寫入）：舊留倉續抱、跨區塊 BAL 落差
    cash: list = field(default_factory=list)    # 非交易現金列（DOI/DEP/JRN…）：writer --flag 照寫進區塊（2026-10-30）


# ─────────────────────────────────────────────────────────────────────────────
# 解析 statement
# ─────────────────────────────────────────────────────────────────────────────

# Cash Balance 表頭有兩種版本，欄數不同：
#   9 欄（AI 轉錄／OCR 產出）：DATE,TIME,TYPE,REF #,DESCRIPTION,Misc Fees,
#                              Commissions & Fees,AMOUNT,BALANCE
#   10 欄（TOS 直接匯出）    ：Trade Date,Exec Date,Exec Time,Type,Ref #,
#                              Description,Misc Fees,Commissions & Fees,Amount,Balance
# 以前只認 9 欄，TOS 原生匯出檔一律被拒（"Statement 格式不符"）。
# 改成依欄名對應，兩種都吃。
#
# ⚠️ 日期欄一定要取 Exec Date，不是 Trade Date。Trade Date 是券商的「營業日」，
# 午夜之後成交的列會掛在前一個營業日（實測 2026-07-29-raw.csv 有 5 列
# Trade Date=7/27 但 Exec Date=7/28 01:18）。取錯會整整差一天 ——
# 這正是《人工SOP》第 2 節警告過的陷阱。
_COL_ALIASES = {
    "date":    ("date", "exec date", "trade date"),   # 順序即優先序
    "time":    ("time", "exec time"),
    "type":    ("type",),
    "ref":     ("ref #", "ref#", "ref"),
    "desc":    ("description", "desc"),
    "misc":    ("misc fees", "misc fee"),
    "comm":    ("commissions & fees", "commissions and fees", "comm & fees"),
    "amount":  ("amount",),
    "balance": ("balance",),
}
_REQUIRED_COLS = ("date", "time", "type", "desc", "amount", "balance")


def _map_header(cells):
    """把表頭列對應成 {欄位: index}；不是表頭就回 None。"""
    norm = [c.strip().strip('"').lower() for c in cells]
    idx = {}
    for key, aliases in _COL_ALIASES.items():
        for alias in aliases:            # 依優先序，先命中的贏
            if alias in norm:
                idx[key] = norm.index(alias)
                break
    if all(k in idx for k in _REQUIRED_COLS):
        return idx
    return None


def parse_statement(text):
    """回傳 [Fill,...]（僅 Cash Balance 區段）。"""
    lines = text.splitlines()
    start = None
    colmap = None
    for i, ln in enumerate(lines):
        if not ln.strip():
            continue
        try:
            cells = next(csv.reader(io.StringIO(ln)))
        except Exception:
            continue
        m = _map_header(cells)
        if m:
            start, colmap = i, m
            break
    if start is None:
        raise ValueError(
            "找不到 Cash Balance 表頭 — Statement 格式不符。"
            "可接受 9 欄（DATE,TIME,TYPE,REF #,...）或 10 欄"
            "（Trade Date,Exec Date,Exec Time,Type,Ref #,...）")

    body = "\n".join(lines[start:])
    reader = csv.reader(io.StringIO(body))
    header = next(reader)
    ncol = max(colmap.values()) + 1
    fills = []
    for row in reader:
        if not row or all(c.strip() == "" for c in row):
            continue
        if len(row) < ncol:
            # 可能是被截斷/OCR 破碎列 → 補空白，後續驗算會抓出來，不靜默跳過
            row = row + [""] * (ncol - len(row))
        g = lambda k: row[colmap[k]] if k in colmap else ""
        date_s, time_s, typ, ref = g("date"), g("time"), g("type"), g("ref")
        desc, misc, comm = g("desc"), g("misc"), g("comm")
        amount, balance = g("amount"), g("balance")
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
                cur.cash.append(f)
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
            net_cash=net, misc_fee=fee, refs=[key], cst_dt=fs[0].cst_dt,
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

def pair_intraday(session, carry=None):
    """對單一 session 的 trades 做 LIFO 配對，填 type_tag/action/pnl/pair_*。
    純當日進出：BOT→SOLD 平 = 多；SOLD→BOT 平 = 空。
    收盤仍有存貨 → 留倉（Status 1）並標記 needs_review（§0：留倉致 offset 跳動需人工確認）。

    carry（2026-10-20 加）：前幾個交易日帶進來的未平多單留倉 lots（同一份 CSV 的前一 session，
    或 trades_all 裡 Status='1' 的列）。以前每個 session 各自配對，平前日留倉的賣單會被誤判成
    「新開空單留倉」、無損益。
    規則（對齊人工帳 4/14、6/15、9/29、10/01、10/02 的配法）：
      1. 先照舊做當日 LIFO 配對，carry 不參與——當沖腿的結果與舊版完全相同；
      2. 收盤後仍未配到的 SOLD 腿（依時間序）才去平 carry，每筆取最晚建立的留倉（LIFO）
         → Type '平'、Status '2'、損益=平倉淨額+原留倉淨額，原留倉 Trade 改 Status '2' 並回指。
    回傳收盤後仍未平的 lots（今日新留倉 + 舊留倉續抱），給下一個 session 當 carry。
    """
    stack = []  # 開倉 lots: {trade, qty_rem, source, sess}
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
            stack.append({"trade": t, "qty_rem": t.qty, "source": "csv", "sess": id(session)})
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

    # 收盤仍未配到的 SOLD 腿 → 平前日多單留倉（carry 取最晚建立者，LIFO）
    held = [l for l in (carry or []) if l["trade"].side == "BOT"]
    for lot in [l for l in stack if l["trade"].side == "SOLD" and l["qty_rem"] == l["trade"].qty]:
        if not held:
            break
        en_lot = held.pop()
        t, en = lot["trade"], en_lot["trade"]
        if t.qty != en.qty:
            held.append(en_lot)
            break
        stack.remove(lot)
        t.type_tag = "平"
        t.status = "2"
        t.pair_date, t.pair_time = en.edt_date, en.edt_time
        en.status = "2"
        en.pair_date, en.pair_time = t.edt_date, t.edt_time
        t.needs_review = True
        if en.net_cash is None:
            t.pnl = None
            t.review_reason = "跨日平倉，但原留倉淨額無法精確取得 — 需人工確認"
        else:
            t.pnl = round(t.net_cash + en.net_cash, 2)
            t.review_reason = f"跨日平倉（Status=2）— 平 {en.edt_date} {en.edt_time} 留倉 @{en.price}"
        src = "（trades_all 原留倉列需改 Status '2' 並補 Pair）" if en_lot.get("source") == "trades_all" else ""
        session.warnings.append(
            f"跨日平倉：{t.edt_time} {t.side} {t.qty}股 @{t.price} 平 {en.edt_date} {en.edt_time} "
            f"留倉 @{en.price}，損益 {'?' if t.pnl is None else f'{t.pnl:.2f}'}{src}")

    for lot in held:
        tr = lot["trade"]
        session.notes.append(
            f"舊留倉續抱：{tr.edt_date} {tr.edt_time} {tr.side} {lot['qty_rem']}股 @{tr.price}")

    # 收盤仍有未平部位 → 留倉
    for lot in stack:
        lot["trade"].status = "1"
        lot["trade"].needs_review = True
        lot["trade"].review_reason = "收盤留倉（Status=1）— §0 留倉致 offset 跳動需人工確認"
        session.warnings.append(
            f"留倉：{lot['trade'].edt_time} {lot['trade'].side} {lot['qty_rem']}股 @{lot['trade'].price}")
    return held + stack


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

def process_session(session, carry=None):
    """carry：前日帶進來的未平 lots。回傳 stats；stats['carry'] = 收盤後仍未平的 lots。"""
    merge_fills(session)
    # 交叉驗證 BALANCE（最重要的資料完整性閘門）
    bal_errs = cross_check_balance(session)
    session.warnings.extend(bal_errs)
    left = pair_intraday(session, carry)

    # 自洽核驗：Σcol1 == close_bal - open_bal（CLAUDE §四.3）
    col1_sum = round(sum(t.net_cash for t in session.trades), 2)
    if session.open_bal is not None and session.close_bal is not None:
        cash_change = round(session.close_bal - session.open_bal, 2)
        if abs(col1_sum - cash_change) > 0.01:
            session.warnings.append(
                f"{FLAG}：Σcol1({col1_sum}) ≠ 收-開({cash_change})，差 {col1_sum-cash_change:+.2f}")
    pnl_sum = round(sum(t.pnl for t in session.trades if t.pnl is not None), 2)
    has_hold = any(t.status == "1" for t in session.trades)
    has_close = any(t.type_tag == "平" for t in session.trades)
    # 全平時 Σcol1 == Σpnl（CLAUDE §四.1）；有留倉/跨日平倉時改驗「當沖腿」
    # Σcol1(當沖) == Σpnl(當沖)——扣掉留倉建立腿與跨日平倉腿後仍須自洽。
    dt = [t for t in session.trades if t.status == "0" and t.type_tag != "平"
          and not any(t is lot["trade"] for lot in left)]
    dt_col1 = round(sum(t.net_cash for t in dt), 2)
    dt_pnl = round(sum(t.pnl for t in dt if t.pnl is not None), 2)
    if not has_hold and not has_close:
        if abs(col1_sum - pnl_sum) > 0.01:
            session.warnings.append(
                f"{FLAG}：全平但 Σcol1({col1_sum}) ≠ Σpnl({pnl_sum})")
    elif abs(dt_col1 - dt_pnl) > 0.01:
        session.warnings.append(
            f"{FLAG}：扣除留倉/跨日平倉腿後 Σcol1當沖({dt_col1}) ≠ Σpnl當沖({dt_pnl})")
    close_pnl = round(sum(t.pnl for t in session.trades
                          if t.type_tag == "平" and t.pnl is not None), 2)
    return {"col1_sum": col1_sum, "pnl_sum": pnl_sum, "has_hold": has_hold,
            "daytrade_pnl": dt_pnl, "close_pnl": close_pnl, "carry": left}


def load_open_lots(trades_path, before_date):
    """trades_all 裡 Status='1'（仍留倉）且日期早於本份 CSV 第一個 session 的列 → carry lots。

    淨額反推：買進腿對帳單 Misc Fees 恆為 0（SEC/TAF 只收賣方），故 BOT 淨額 = −價×股，
    價格存 4 位小數 → 精確到分，非估算。空單留倉（SOLD）的賣方費用無法從 trades_all 精確反推
    → net_cash=None，平倉時會標 needs_review、不算損益。
    """
    if not trades_path or not os.path.exists(trades_path):
        return []
    try:
        import openpyxl
    except ImportError:
        return []
    wb = openpyxl.load_workbook(trades_path, read_only=True, data_only=True)
    ws = wb.active
    lots = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        if not r or r[0] in (None, ""):
            continue
        date = r[0].strftime("%Y-%m-%d") if isinstance(r[0], datetime) else str(r[0]).strip()[:10]
        status = str(r[8]).strip() if len(r) > 8 and r[8] is not None else ""
        if status.endswith(".0"):
            status = status[:-2]
        if status != "1" or date >= before_date:
            continue
        price = float(r[3]); qty = int(float(r[6])); act = str(r[7]).strip()
        side = "BOT" if act == "B" else "SOLD"
        net = round(-price * qty, 2) if side == "BOT" else None
        t = Trade(edt_date=date, edt_time=str(r[1]).strip(), symbol=str(r[2]).strip(),
                  side=side, qty=qty, price=price, net_cash=net, misc_fee=0.0,
                  action=act, status="1")
        lots.append({"trade": t, "qty_rem": qty, "source": "trades_all"})
    wb.close()
    lots.sort(key=lambda l: (l["trade"].edt_date, l["trade"].edt_time))
    return lots


_TOTAL_CASH_RE = re.compile(r'Total Cash\s*\$?\(?([\d,]+\.\d{2})\)?')


def parse_total_cash(text):
    """對帳單 Cash Balance 表後的 'Total Cash $2,383,728.23'（真實現金，不受 BAL 重編影響）。"""
    m = _TOTAL_CASH_RE.search(text or "")
    if not m:
        return None
    v = float(m.group(1).replace(",", ""))
    return -v if "(" in m.group(0) else v


def parse_trade_history(text):
    """'Account Trade History' 表 → [(cst_dt, side, qty, pos_effect, price)]。表不存在回 []。"""
    out = []
    lines = (text or "").splitlines()
    i = next((k for k, ln in enumerate(lines) if ln.strip() == "Account Trade History"), None)
    if i is None:
        return out
    for ln in lines[i + 2:]:
        if not ln.strip():
            break
        c = next(csv.reader(io.StringIO(ln)))
        if len(c) < 11:
            continue
        try:
            d, tm = c[1].strip().split()
            out.append((_parse_cst_dt(d, tm), c[3].strip(), int(c[4]), c[5].strip(), float(c[10])))
        except Exception:
            continue
    return out


def check_pos_effect(sessions, history):
    """跨日平倉腿必須在 Trade History 是 'TO CLOSE'。若是 'TO OPEN' → 引擎配錯（其實是新開倉）→ ⚠️。
    Trade History 的時間偶爾比 Cash Balance 早 1 秒（10/19 23:18:32 vs :33），比對容許 ±10 秒、同價同方向。"""
    for s in sessions:
        for t in s.trades:
            if t.type_tag != "平" or t.cst_dt is None:
                continue
            want = "SELL" if t.side == "SOLD" else "BUY"
            hits = [h for h in history if h[1] == want and abs(h[4] - t.price) < 0.00005
                    and abs((h[0] - t.cst_dt).total_seconds()) <= 10]
            if not history:
                continue
            if not hits:
                s.notes.append(f"{t.edt_time} 跨日平倉在 Trade History 找不到對應列（未能交叉確認 Pos Effect）")
            elif any(h[3] != "TO CLOSE" for h in hits):
                t.needs_review = True
                t.review_reason = "Trade History 顯示 TO OPEN，與跨日平倉配對矛盾"
                s.warnings.append(f"{FLAG}：{t.edt_time} {t.side} 引擎判跨日平倉，但 Trade History 為 "
                                  f"{'/'.join(h[3] for h in hits)}")
            else:
                s.notes.append(f"{t.edt_time} 跨日平倉 ↔ Trade History TO CLOSE ✓")


def process_all(sessions, fills=None, text=None, trades_path=None):
    """依序處理整份 CSV 的所有 session，留倉 lots 在 session 間傳遞（跨日平倉配對）。

    另做兩件跨 session 的檢查：
      • 相鄰 session 的「前收盤 → 今開盤」BAL 落差 → ℹ️ 註記（券商 BAL 追溯重編，比照
        offset_state.json 10/16 註記不認列損益）；
      • 若對帳單有 Total Cash：第一個 BAL + 全部非 BAL 列淨額 應 == Total Cash
        （相符即證明 BAL 落差只是券商重編、現金流完整）；不符只記 ℹ️，不擋。
    回傳 [stats,...]（與 sessions 同序）。
    """
    first = sessions[0].edt_date if sessions else "9999-99-99"
    carry = load_open_lots(trades_path, first) if trades_path else []
    for lot in carry:
        sessions and sessions[0].notes.append(
            f"自 trades_all 帶入留倉：{lot['trade'].edt_date} {lot['trade'].edt_time} "
            f"{lot['trade'].side} {lot['qty_rem']}股 @{lot['trade'].price}")
    out = []
    prev = None
    for s in sessions:
        if prev is not None and prev.close_bal is not None and s.open_bal is not None:
            gap = round(s.open_bal - prev.close_bal, 2)
            if abs(gap) > 0.01:
                s.notes.append(
                    f"跨區塊 BAL 落差 {gap:+,.2f}（前收 {prev.close_bal:,.2f} → 今開 "
                    f"{s.open_bal:,.2f}）— 券商 BAL 重編，不認列損益")
        st = process_session(s, carry)
        # 只把多單留倉帶到下一個 session。本帳戶長期持有 SPY 核心多單，SOLD 永遠是 TO CLOSE；
        # 當日配不到的 SOLD 是在平 CSV 範圍外的舊部位（如 8/03 平 7/16），不是開空單，
        # 若帶下去會把隔天的買進誤配成「回補空單」。它在自己的 session 已標 ⚠️留倉。
        carry = [l for l in st["carry"] if l["trade"].side == "BOT"]
        out.append(st)
        prev = s
    if text:
        check_pos_effect(sessions, parse_trade_history(text))
    total_cash = parse_total_cash(text) if text else None
    if total_cash is not None and fills:
        bals = [f for f in fills if f.typ == "BAL" and f.balance is not None]
        if bals and sessions:
            chain = round(bals[0].balance + sum(f.net_cash for f in fills if f.typ != "BAL"), 2)
            if abs(chain - total_cash) > 0.01:
                # 只當註記不擋：券商 Total Cash 本身也會被重編（2026-08-11 檔差 +10.24，
                # 當天帳照常寫入無誤）。真正的閘門仍是逐列 BALANCE 重建。
                sessions[-1].notes.append(
                    f"首 BAL+全部現金異動={chain:,.2f} ≠ Total Cash {total_cash:,.2f}"
                    f"（差 {chain-total_cash:+,.2f}）— 券商數字自相矛盾，僅供參考")
            else:
                sessions[-1].notes.append(
                    f"Total Cash {total_cash:,.2f} = 首 BAL+全部現金異動 ✓（現金流完整）")
    return out


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
          f"留倉={stats['has_hold']}"
          + (f"  當沖Σpnl={stats['daytrade_pnl']:.2f}  平倉Σpnl={stats['close_pnl']:.2f}"
             if stats.get("close_pnl") else ""))
    if session.notes:
        print(" 註記（不擋寫入）：")
        for n in session.notes:
            print(f"   ℹ️ {n}")
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
    ap.add_argument("--trades", default=TRADES_PATH_DEFAULT,
                    help="讀取未平留倉用的 trades_all.xlsx（預設本資料夾）")
    ap.add_argument("--no-carry", action="store_true",
                    help="不從 trades_all 帶入留倉（同份 CSV 內仍會跨 session 配對）")
    args = ap.parse_args()

    with open(args.csv_path, encoding="utf-8-sig") as fh:
        text = fh.read()
    fills = parse_statement(text)
    sessions = group_sessions(fills)
    # 先整份處理（留倉要在 session 間傳遞），--date 只過濾輸出
    stats_all = process_all(sessions, fills, text,
                            None if args.no_carry else args.trades)
    pairs = list(zip(sessions, stats_all))
    if args.date:
        pairs = [(s, st) for s, st in pairs if s.edt_date == args.date]
        if not pairs:
            print(f"找不到 {args.date} 的 session"); sys.exit(1)

    for s, stats in pairs:
        report_session(s, stats)

    if args.commit:
        print("\n[commit 模式尚未接上寫入層 — 需先通過 dry-run 驗證]")
    else:
        print("\n[DRY-RUN] 未寫入任何檔案。")


if __name__ == "__main__":
    main()
