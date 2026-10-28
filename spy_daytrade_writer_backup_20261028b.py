#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
spy_daytrade_writer.py — spy_daytrade_engine.py 的 --commit 寫入層（2026-07-24 接上）

把引擎分析結果寫入 live 財務檔，嚴格對齊 CLAUDE.md §一/§二/§三/§四。
處理兩種日子：
  • 標準日：純當日進出、offset 不變。
  • 特例日（2026-10-22 起，使用者 10/19 核定）：引擎警告只有「收盤留倉」及/或「跨日平倉」、
    每腿合併後 100 股 → 依核定範本寫兩本帳（10/16 留倉、10/19 平倉、9/29 單筆平倉、7/09；
    2026-10-28 加：留倉中又加碼留倉，10/23 樣式）。
    數字全取自引擎；範本沒涵蓋的版面（同日平倉＋新留倉、部分平倉、
    留倉區塊含配息列、當沖腿與留倉腿交錯）→ 一律不寫、回報人工。
其餘特例（非100股/BALANCE 對不上/非交易現金/offset 不確定/Cash Balance 不完整）→
由引擎的 needs_review/FLAG 或 statement_check 攔下，writer 直接 abort，
不寫入、回報需人工確認。這對齊 §0：絕不為跑完流程而靜默寫猜值。

安全：§3 寫前備份、§5 檢查 ~$ 鎖檔、§4 寫後全核驗、失敗還原備份。

offset 模型（CLAUDE.md §二）：Day Header B = 券商開盤餘額 + offset。
offset 是常數（2026-08-05 起為 0）。特例日寫後頭尾雙錨照驗，offset 數值不動，
只在 offset_state.json 的 note 前面補一段當日註記；offset 真的跳動 → 核驗不過 → 還原、轉人工。
"""
import contextlib
import json
import os
import re
import shutil
from datetime import datetime

import openpyxl

from spy_daytrade_engine import (
    parse_statement, group_sessions, process_all, FLAG,
)
from statement_check import check as statement_check

# 路徑相對本腳本所在目錄 → 整個資料夾可原樣搬到任一機器/路徑而不必改碼（遷移用）。
_BASE = os.path.dirname(os.path.abspath(__file__))
CS_PATH_DEFAULT = os.path.join(_BASE, "CS交易紀錄.xlsx")
TRADES_PATH_DEFAULT = os.path.join(_BASE, "trades_all.xlsx")
OFFSET_STATE = os.path.join(_BASE, "offset_state.json")
SHEET_DETAIL = "個別識別法"

FMT_NUM = r"0.00_);[Red]\(0.00\)"
FMT_PCT = "0.00%"


# ── helpers ──────────────────────────────────────────────────────────────────
def last_nonempty_row(ws, ncols=5):
    for r in range(ws.max_row, 0, -1):
        if any(ws.cell(r, c).value not in (None, "") for c in range(1, ncols + 1)):
            return r
    return 0


def check_lock(path):
    d, base = os.path.split(path)
    if os.path.exists(os.path.join(d, "~$" + base)):
        raise RuntimeError(f"§5 鎖檔存在（Excel 開著？）：{path}，請關閉後重跑")


# ── 與常駐 app 的互斥 ────────────────────────────────────────────────────────
WRITER_LOCK = os.path.join(_BASE, ".writer.lock")
STALE_LOCK_SECONDS = 600   # 10 分鐘；writer 正常執行是秒級，超過必定是殘留

@contextlib.contextmanager
def writer_lock():
    """寫入期間宣告互斥，擋掉常駐 app 的 trades_all 回寫。

    trade_review_app 的 _analyse_all_trades() 會把整張 trades_all 回寫
    （見 trade_review_app.py 的 df.to_excel），而 app 是 pythonw 常駐、每 4 秒
    偵測 mtime。writer 正在 openpyxl save 時若撞上這個回寫，兩邊會互相蓋掉，
    而且無人看管的 Boss PC 上不會有人發現。

    殘留鎖處理：writer 若被強制中斷可能留下鎖檔，逾時後視為無效，
    否則 app 會永遠不再回寫。
    """
    try:
        with open(WRITER_LOCK, "w", encoding="utf-8") as f:
            f.write(f"{os.getpid()}\n{datetime.now().isoformat()}\n")
        yield
    finally:
        try:
            os.remove(WRITER_LOCK)
        except OSError:
            pass


def backup(path):
    dst = path.replace(".xlsx", f"_backup_{datetime.now():%Y%m%d}.xlsx") \
        if False else path[:-5] + f"_backup_{_today()}.xlsx"
    shutil.copy2(path, dst)
    return dst


def _today():
    # Date.now 在某些沙盒被禁；用檔案系統時間避免 import time 問題時可換。
    return datetime.now().strftime("%Y%m%d")


def load_offset():
    if not os.path.exists(OFFSET_STATE):
        raise RuntimeError(f"找不到 {OFFSET_STATE}（需先 seed offset 常數）")
    with open(OFFSET_STATE, encoding="utf-8") as f:
        return json.load(f)


def save_offset(state):
    with open(OFFSET_STATE, "w", encoding="utf-8") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)


# ── gate：判斷 session 是否可 auto-commit ────────────────────────────────────
def day_already_written(ws, dot):
    """§5.5.1 冪等：CS 個別識別法 已含該日 Day Header（A=='YYYY.MM.DD' 且 C=='損益'）。"""
    for r in range(1, last_nonempty_row(ws) + 1):
        if ws.cell(r, 1).value == dot and ws.cell(r, 3).value == "損益":
            return True
    return False


def day_already_in_trades(ws, edt_date):
    """--trades-only 模式的冪等判斷。

    trades_only 模式沒有 CS 帳可查（Boss PC 上根本不存在該檔），改查
    trades_all 的 Date 欄是否已有該交易日。Date 欄是字串 'YYYY-MM-DD'
    （CLAUDE.md §一：嚴禁寫入 datetime 物件），但歷史列若曾被其他工具
    寫成 datetime 也要能認得，故兩種型別都比對。
    """
    for r in range(2, ws.max_row + 1):     # 第 1 列是表頭
        v = ws.cell(r, 1).value
        if v is None or v == "":
            continue
        if isinstance(v, datetime):
            v = v.strftime("%Y-%m-%d")
        if str(v).strip()[:10] == edt_date:
            return True
    return False


def leg_kind(t):
    """'close' 跨日平倉腿／'hold' 收盤留倉建立腿（含同份 CSV 稍後已被平掉的）／'day' 當沖腿。"""
    if t.type_tag == "平":
        return "close"
    if t.type_tag == "" and t.status in ("1", "2"):
        return "hold"
    return "day"


def is_special(session):
    return any(leg_kind(t) != "day" for t in session.trades)


def _special_leg_ok(t):
    # review_reason 會被 pair_intraday 蓋掉（「合併後 N 股」→「收盤留倉…」），股數另驗。
    if t.qty != 100:
        return False
    if t.review_reason.startswith("收盤留倉") and t.side == "BOT":
        return True
    return t.review_reason.startswith("跨日平倉（Status=2）") and t.pnl is not None


def _special_warning_ok(w):
    if w.startswith("留倉：") and " BOT " in w:
        return True
    return w.startswith("跨日平倉：") and "損益 ?" not in w


def gate(session, allow_special=False):
    """allow_special=False：舊行為（任何留倉/跨日平倉都擋）——--trades-only 仍用這個。
    allow_special=True：警告只有多單收盤留倉／跨日平倉（有損益）、每腿 100 股 → 放行，
    但特例日的任何其他警告一律擋（比標準日更嚴，不只看關鍵字）。"""
    reasons = []
    if session.open_bal is None or session.close_bal is None:
        reasons.append("session 開盤/收盤餘額不明")
    for t in session.trades:
        if t.needs_review and not (allow_special and _special_leg_ok(t)):
            reasons.append(f"{t.edt_time} {t.review_reason}")
    for w in session.warnings:
        if allow_special and _special_warning_ok(w):
            continue
        if FLAG in w or "BALANCE 對不上" in w or "留倉" in w or "非交易現金" in w:
            reasons.append(w)
    if allow_special and is_special(session):
        for t in session.trades:
            if t.qty != 100:
                reasons.append(f"{t.edt_time} 合併後 {t.qty} 股（非 100）")
        for w in session.warnings:
            if not _special_warning_ok(w) and w not in reasons:
                reasons.append(w)
    return reasons


# ── 反向補 pair_time（引擎只在平倉端設 pair_time，進場端要補回指）────────────
def backlink_pairs(session):
    exits = [t for t in session.trades if t.pair_time]
    for ex in exits:
        for en in session.trades:
            if en is ex:
                continue
            # 比對日期：跨日平倉腿的 pair_time 是前日留倉時間，不可誤指到同時刻的當日進場腿
            if (not en.pair_time and en.type_tag == "" and en.edt_time == ex.pair_time
                    and en.edt_date == ex.pair_date):
                en.pair_time = ex.edt_time
                en.pair_date = ex.edt_date
                break


# ── CS 個別識別法 寫入 ───────────────────────────────────────────────────────
def _detect_tail_hold_section(ws, last):
    """若檔尾是『累積留倉』區段，回傳 (start_row, rows_values)；否則 (None, [])。
    rows_values 為從『累積留倉』標頭列到 last 的 [(A,B,C),...]（含公式字串）。"""
    hold_start = None
    for r in range(last, 0, -1):
        a = ws.cell(r, 1).value
        if isinstance(a, str) and a.startswith("累積留倉"):
            hold_start = r
            break
        # 遇到 Day Header（A=YYYY.MM.DD, C=損益）就停：尾端沒有留倉區
        if isinstance(a, str) and ws.cell(r, 3).value == "損益":
            break
    if hold_start is None:
        return None, []
    rows = [(ws.cell(r, 1).value, ws.cell(r, 2).value, ws.cell(r, 3).value)
            for r in range(hold_start, last + 1)]
    return hold_start, rows


def write_cs_block(ws, session, header_B):
    """在值掃描定位處寫入一個日區塊。回傳 (header_row, detail_start, detail_end)。
    detail 只放當沖 TRD（純當日進出日無留倉/平倉列，故核驗=+B{dN}-B{hr} 不需扣除）。"""
    dot = session.edt_date.replace("-", ".")
    last = last_nonempty_row(ws)
    hr = last + 3  # §2：前留 2 空行
    ws.cell(hr, 1, dot); ws.cell(hr, 2, round(header_B, 2)); ws.cell(hr, 3, "損益")
    d0 = hr + 1
    trades = session.trades
    for i, t in enumerate(trades):
        r = d0 + i
        ws.cell(r, 1, t.net_cash)
        ws.cell(r, 2, f"=+B{hr}+A{r}" if i == 0 else f"=+B{r-1}+A{r}")
        if t.pnl is not None:
            ws.cell(r, 3, t.pnl)
    dN = d0 + len(trades) - 1
    tot = dN + 1
    ws.cell(tot, 2, "合計 "); ws.cell(tot, 3, f"=SUM(C{d0}:C{dN})")
    ws.cell(tot, 3).number_format = FMT_NUM
    chk = tot + 1
    ws.cell(chk, 2, "核驗"); ws.cell(chk, 3, f"=+B{dN}-B{hr}")
    ws.cell(chk, 3).number_format = FMT_NUM
    # 1 空行後當沖摘要
    buy = chk + 2
    ws.cell(buy, 1, f"{dot}當沖"); ws.cell(buy, 2, "買"); ws.cell(buy, 3, f'=SUMIF(A{d0}:A{dN},"<0")')
    ws.cell(buy, 3).number_format = FMT_NUM
    sell = buy + 1
    ws.cell(sell, 1, f'="SPY " & COUNT(C{d0}:C{dN}) * 100 & "股"'); ws.cell(sell, 2, "賣")
    ws.cell(sell, 3, f'=SUMIF(A{d0}:A{dN},">0")'); ws.cell(sell, 3).number_format = FMT_NUM
    pnl = sell + 1
    ws.cell(pnl, 2, "當日損益"); ws.cell(pnl, 3, f"=+C{buy}+C{sell}"); ws.cell(pnl, 3).number_format = FMT_NUM
    pct = pnl + 1
    ws.cell(pct, 2, "當日損益 %"); ws.cell(pct, 3, f"=C{pnl}/ABS(C{buy})"); ws.cell(pct, 3).number_format = FMT_PCT
    return hr, d0, dN


_REF_RE = re.compile(r"(?<![A-Za-z$])([A-Z]{1,3})(\d+)(?!\d)")


def _shift_refs(formula, lo, hi, delta):
    """公式中列號落在 [lo, hi] 的參照位移 delta 列（雙引號內的字串不動）。"""
    def sub(m):
        row = int(m.group(2))
        return f"{m.group(1)}{row + delta}" if lo <= row <= hi else m.group(0)
    parts = formula.split('"')
    parts[::2] = [_REF_RE.sub(sub, p) for p in parts[::2]]
    return '"'.join(parts)


def append_hold_section(ws, rows, src_start=None):
    """把抬起的累積留倉區段貼回檔尾。src_start＝原起始列：區段內互指的公式（平倉損益
    '=+C2616+C2617'）要跟著位移，否則貼回後會指到別的列（2026-10-15/10-21 兩次因此壞掉，
    10/19 人工修 C2618/C2619）。指到區段外的參照（'=+A2597' 平倉明細）不動。"""
    ws.append([None, None, None])  # 1 空行
    for (a, b, c) in rows:
        ws.append([a, b, c])
    if src_start is None:
        return
    end = last_nonempty_row(ws)
    start = end - len(rows) + 1
    delta = start - src_start
    if delta == 0:
        return
    for r in range(start, end + 1):
        for c in range(1, 4):
            v = ws.cell(r, c).value
            if isinstance(v, str) and v.startswith("="):
                ws.cell(r, c).value = _shift_refs(v, src_start, src_start + len(rows) - 1, delta)


def write_cumulative(wb, session, day_pnl):
    cs = wb["累積損益check"]  # 選 sheet 用名稱，避免夾層 Sheet1 造成 index 位移（2026-08-10 修）
    last = last_nonempty_row(cs, ncols=3)
    r = last + 1
    y, m, d = [int(x) for x in session.edt_date.split("-")]
    cs.cell(r, 1, datetime(y, m, d)); cs.cell(r, 2, round(day_pnl, 2))
    cs.cell(r, 3, f"=C{last}+B{r}")


def write_trades_all(ws, session):
    # 損益 欄：與既有歷史列一致用 float（CLAUDE.md §一 舊述「字串」與實檔不符；
    # app 以 dtype=str 讀取，float/str 皆可，取 float 對齊既有欄型）。
    for t in session.trades:
        if t.status == "1":
            # 收盤留倉腿：尚無配對（10/16 範本 J/K 留空），平倉日再由 flip_held_rows 補
            pair_date, pair_time = None, None
        else:
            pair_date, pair_time = t.pair_date or session.edt_date, t.pair_time
        ws.append([session.edt_date, t.edt_time, t.symbol, t.price, t.type_tag or None,
                   t.pnl, t.qty, t.action, t.status, pair_date, pair_time])


def _norm_date(v):
    if isinstance(v, datetime):
        return v.strftime("%Y-%m-%d")
    return str(v).strip()[:10] if v not in (None, "") else ""


def _norm_status(v):
    s = "" if v is None else str(v).strip()
    return s[:-2] if s.endswith(".0") else s


def flip_held_rows(ws, session):
    """跨日平倉：把原留倉列 Status '1'→'2'、補 Pair_Date/Pair_Time（10/19 範本 1442-1444）。
    同一份 CSV 內才剛寫入的留倉列，引擎已直接給 Status '2'＋配對 → 確認一致即可。"""
    for t in session.trades:
        if leg_kind(t) != "close":
            continue
        hits = [r for r in range(2, ws.max_row + 1)
                if _norm_date(ws.cell(r, 1).value) == t.pair_date
                and str(ws.cell(r, 2).value).strip() == t.pair_time
                and str(ws.cell(r, 8).value).strip() == "B"]
        if len(hits) != 1:
            raise Blocked(f"trades_all 找不到唯一的原留倉列 {t.pair_date} {t.pair_time}（找到 {len(hits)} 列）")
        r = hits[0]
        st = _norm_status(ws.cell(r, 9).value)
        if st == "1":
            ws.cell(r, 9).value = "2"
            ws.cell(r, 10).value = t.edt_date
            ws.cell(r, 11).value = t.edt_time
        elif not (st == "2" and _norm_date(ws.cell(r, 10).value) == t.edt_date
                  and str(ws.cell(r, 11).value).strip() == t.edt_time):
            raise Blocked(f"trades_all 原留倉列 {r} Status={st!r} 已配給別筆，與 {t.edt_time} 平倉矛盾")


# ── 特例日（收盤留倉／跨日平倉）CS 寫入 ─────────────────────────────────────
# 範本（使用者核定，逐格對照用）：
#   留倉日 2026.10.16（CS 2622-2639）：明細全腿、當沖腿才有 C、核驗扣各留倉腿 A、
#     當沖彙總 SUMIF 只含當沖腿、尾接「累積留倉 N股」＋每筆留倉；檔尾已全平的舊留倉區段原地不動。
#   平倉日 2026.10.19（CS 2637-2652）：平倉腿無 C（合計=當沖損益）、核驗扣各平倉腿 A；
#     抬起檔尾留倉區段，貼回時標頭改「0股」、每筆留倉後接「平倉」列（=+A平倉明細，依留倉順序）、
#     平倉損益/％；單筆時照 9/29 修正後樣式（2615-2619）。
#   加碼留倉日 2026.10.23（2026-10-28 使用者要求加範本）：前日留倉續抱中又新增留倉 →
#     抬起檔尾留倉區段，區塊照留倉日寫法；貼回時標頭改合計股數、舊留倉列原值照抄、新留倉接在後面。
#   累積損益check：當日損益＝當沖損益＋平倉損益（10/16 13.29、10/19 1132.54）。

class Blocked(Exception):
    """寫入當下才看得出的範本外情況（不是程式錯誤）→ 不寫、回報人工。"""


_LOT_RE = re.compile(r"^\d{4}\.\d{2}\.\d{2} 留倉$")
_CLOSE_RE = re.compile(r"^\d{4}\.\d{2}\.\d{2} 平倉$")


def _hold_section_parts(rows):
    """累積留倉區段 rows（含標頭）→ (lots, n_close, others)；lots=[(A, B, C)]。"""
    lots, n_close, others = [], 0, []
    for a, b, c in rows[1:]:
        if isinstance(a, str) and _LOT_RE.match(a):
            lots.append((a, b, c))
        elif isinstance(a, str) and _CLOSE_RE.match(a):
            n_close += 1
        elif a is None and b in ("平倉損益", "平倉損益 %"):
            pass
        else:
            others.append((a, b, c))
    return lots, n_close, others


def plan_special(ws, session):
    """確認這個特例日落在核定範本內，回傳寫法；不在範本內 → raise Blocked。"""
    legs = session.trades
    kinds = [leg_kind(t) for t in legs]
    holds = [t for t, k in zip(legs, kinds) if k == "hold"]
    closes = [t for t, k in zip(legs, kinds) if k == "close"]
    day_idx = [i for i, k in enumerate(kinds) if k == "day"]
    if holds and closes:
        raise Blocked("同日既有跨日平倉又有新留倉——無核定範本")
    if day_idx and day_idx != list(range(day_idx[0], day_idx[-1] + 1)):
        raise Blocked("當沖腿與留倉/平倉腿交錯（當沖彙總 SUMIF 無法只含當沖腿）——無核定範本")
    carry_out = getattr(session, "carry_out", None)
    if carry_out is None:
        raise Blocked("缺引擎收盤留倉結果（carry_out）")
    hold_start, hold_rows = _detect_tail_hold_section(ws, last_nonempty_row(ws))

    if holds:
        hold_ids = {id(t) for t in holds}
        if not hold_ids <= {id(l["trade"]) for l in carry_out}:
            raise Blocked("引擎收盤留倉結果不含今日全部留倉腿——兩本帳不一致")
        prior = sorted((l for l in carry_out if id(l["trade"]) not in hold_ids),
                       key=lambda l: (l["trade"].edt_date, l["trade"].edt_time))
        if not prior:
            if hold_start is not None:
                lots, n_close, _ = _hold_section_parts(hold_rows)
                if len(lots) != n_close:
                    raise Blocked(f"CS 檔尾累積留倉區段仍有 {len(lots) - n_close} 筆未平，"
                                  "但引擎沒有前日留倉——兩本帳不一致")
            return {"mode": "hold", "holds": holds}
        # 加碼留倉日：檔尾留倉區段必須恰好是引擎續抱的前日留倉（全數未平、逐筆對得上）
        if hold_start is None:
            raise Blocked("引擎有續抱的前日留倉，但 CS 檔尾沒有累積留倉區段——兩本帳不一致")
        lots, n_close, others = _hold_section_parts(hold_rows)
        if n_close or others:
            raise Blocked("CS 累積留倉區段已有平倉／配息等其他列——無核定範本")
        if len(lots) != len(prior):
            raise Blocked(f"CS 累積留倉 {len(lots)} 筆 ≠ 引擎續抱前日留倉 {len(prior)} 筆")
        for (a, b, c), l in zip(lots, prior):
            t = l["trade"]
            if l["qty_rem"] != t.qty or t.qty != 100:
                raise Blocked(f"續抱留倉 {t.edt_date} {t.edt_time} 剩 {l['qty_rem']}/{t.qty} 股——無核定範本")
            if (a != f"{t.edt_date.replace('-', '.')} 留倉" or b != f"買 {t.qty}股"
                    or not isinstance(c, (int, float)) or t.net_cash is None
                    or abs(c - t.net_cash) > 0.005):
                raise Blocked(f"CS 留倉列 {a!r} {b!r} {c!r} 對不上引擎續抱的留倉 "
                              f"{t.edt_date} {t.edt_time} 淨額 {t.net_cash}")
        return {"mode": "hold", "holds": holds, "hold_start": hold_start, "lots": lots,
                "prior": [l["trade"] for l in prior]}

    if carry_out:
        raise Blocked("部分平倉（收盤仍有前日留倉續抱）——無核定範本")
    if hold_start is None:
        raise Blocked("CS 檔尾沒有累積留倉區段可對應今日平倉")
    lots, n_close, others = _hold_section_parts(hold_rows)
    if n_close or others:
        raise Blocked("CS 累積留倉區段已有平倉／配息等其他列——無核定範本")
    if len(lots) != len(closes):
        raise Blocked(f"CS 累積留倉 {len(lots)} 筆 ≠ 今日跨日平倉 {len(closes)} 筆")
    by_lot = sorted(closes, key=lambda t: (t.pair_date, t.pair_time))
    for (a, b, c), t in zip(lots, by_lot):
        en_net = round(t.pnl - t.net_cash, 2)   # 引擎 pnl = 平倉淨額 + 原留倉淨額
        if (a != f"{t.pair_date.replace('-', '.')} 留倉" or b != f"買 {t.qty}股"
                or not isinstance(c, (int, float)) or abs(c - en_net) > 0.005):
            raise Blocked(f"CS 留倉列 {a!r} {b!r} {c!r} 對不上引擎配對的留倉 "
                          f"{t.pair_date} {t.pair_time} 淨額 {en_net}")
    return {"mode": "close", "hold_start": hold_start, "lots": lots, "by_lot": by_lot}


def write_special_block(ws, session, header_B, plan):
    """依 plan 寫特例日區塊＋累積留倉區段。回傳 layout（核驗用的列號）。"""
    dot = session.edt_date.replace("-", ".")
    if "hold_start" in plan:    # 平倉日／加碼留倉日：抬起檔尾留倉區段，寫完區塊後重寫
        last = last_nonempty_row(ws)
        ws.delete_rows(plan["hold_start"], last - plan["hold_start"] + 1)
        while last_nonempty_row(ws) < ws.max_row:
            ws.delete_rows(ws.max_row, 1)
    legs = session.trades
    hr = last_nonempty_row(ws) + 3  # §2：前留 2 空行
    ws.cell(hr, 1, dot); ws.cell(hr, 2, round(header_B, 2)); ws.cell(hr, 3, "損益")
    d0 = hr + 1
    row_of = {}
    for i, t in enumerate(legs):
        r = d0 + i
        ws.cell(r, 1, t.net_cash)
        ws.cell(r, 2, f"=+B{hr}+A{r}" if i == 0 else f"=+B{r-1}+A{r}")
        if leg_kind(t) == "day" and t.pnl is not None:
            ws.cell(r, 3, t.pnl)
        row_of[id(t)] = r
    dN = d0 + len(legs) - 1
    tot = dN + 1
    ws.cell(tot, 2, "合計 "); ws.cell(tot, 3, f"=SUM(C{d0}:C{dN})")
    ws.cell(tot, 3).number_format = FMT_NUM
    chk = tot + 1
    special_rows = [row_of[id(t)] for t in legs if leg_kind(t) != "day"]
    ws.cell(chk, 2, "核驗")
    ws.cell(chk, 3, f"=+B{dN}-B{hr}" + "".join(f"-A{r}" for r in special_rows))
    ws.cell(chk, 3).number_format = FMT_NUM
    lay = {"hr": hr, "d0": d0, "dN": dN, "tot": tot, "chk": chk, "row_of": row_of,
           "mode": plan["mode"], "pnl_row": None, "day_pnl_row": None,
           "lifted": "hold_start" in plan}
    end = chk
    day_rows = [row_of[id(t)] for t in legs if leg_kind(t) == "day"]
    if day_rows:
        a0, a1 = day_rows[0], day_rows[-1]
        buy = chk + 2
        ws.cell(buy, 1, f"{dot}當沖"); ws.cell(buy, 2, "買"); ws.cell(buy, 3, f'=SUMIF(A{a0}:A{a1},"<0")')
        ws.cell(buy, 3).number_format = FMT_NUM
        sell = buy + 1
        ws.cell(sell, 1, f'="SPY " & COUNT(C{a0}:C{a1}) * 100 & "股"'); ws.cell(sell, 2, "賣")
        ws.cell(sell, 3, f'=SUMIF(A{a0}:A{a1},">0")'); ws.cell(sell, 3).number_format = FMT_NUM
        pnl = sell + 1
        ws.cell(pnl, 2, "當日損益"); ws.cell(pnl, 3, f"=+C{buy}+C{sell}"); ws.cell(pnl, 3).number_format = FMT_NUM
        pct = pnl + 1
        ws.cell(pct, 2, "當日損益 %"); ws.cell(pct, 3, f"=C{pnl}/ABS(C{buy})"); ws.cell(pct, 3).number_format = FMT_PCT
        lay["day_pnl_row"] = pnl
        end = pct

    h0 = end + 2  # 1 空行後接累積留倉
    if plan["mode"] == "hold":
        holds = plan["holds"]
        lots, prior = plan.get("lots", []), plan.get("prior", [])
        ws.cell(h0, 1, f"累積留倉   {sum(t.qty for t in prior + holds)}股")
        for i, (a, b, c) in enumerate(lots):     # 加碼留倉日：續抱的舊留倉列原值照抄
            ws.cell(h0 + 1 + i, 1, a); ws.cell(h0 + 1 + i, 2, b); ws.cell(h0 + 1 + i, 3, c)
        for i, t in enumerate(holds):
            r = h0 + 1 + len(lots) + i
            ws.cell(r, 1, f"{dot} 留倉"); ws.cell(r, 2, f"買 {t.qty}股"); ws.cell(r, 3, t.net_cash)
        n = len(lots) + len(holds)
        lay["lot_rows"] = list(range(h0 + 1, h0 + 1 + n))
        lay["lot_expect"] = [t.net_cash for t in prior + holds]
        lay["hold_end"] = h0 + n
    else:
        lots, by_lot = plan["lots"], plan["by_lot"]
        ws.cell(h0, 1, "累積留倉   0股")
        lot_rows = []
        for i, (a, b, c) in enumerate(lots):
            r = h0 + 1 + i
            ws.cell(r, 1, a); ws.cell(r, 2, b); ws.cell(r, 3, c)
            lot_rows.append(r)
        close_rows = []
        for i, t in enumerate(by_lot):
            r = h0 + 1 + len(lots) + i
            ws.cell(r, 1, f"{dot} 平倉"); ws.cell(r, 2, f"賣{t.qty}股"); ws.cell(r, 3, f"=+A{row_of[id(t)]}")
            close_rows.append(r)
        pr = close_rows[-1] + 1
        if len(lots) == 1:      # 9/29 修正後樣式（2618/2619）
            ws.cell(pr, 3, f"=+C{lot_rows[0]}+C{close_rows[0]}")
            pct_f = f"=C{pr}/ABS(C{lot_rows[0]})"
        else:                   # 10/19 樣式（2651/2652）
            ws.cell(pr, 3, "=" + "+".join(f"C{r}" for r in lot_rows + close_rows))
            pct_f = f"=C{pr}/ABS(" + "+".join(f"C{r}" for r in lot_rows) + ")"
        ws.cell(pr, 2, "平倉損益")
        ws.cell(pr + 1, 2, "平倉損益 %"); ws.cell(pr + 1, 3, pct_f)
        lay["lot_rows"], lay["pnl_row"], lay["hold_end"] = lot_rows, pr, pr + 1
    lay["h0"] = h0
    return lay


# ── 簡易公式求值（只認本檔用到的 +−*/、SUM(欄n:欄m)、ABS）——核驗用 ─────────
def _cell_num(ws, coord, depth=0):
    v = ws[coord].value
    if isinstance(v, str) and v.startswith("="):
        return _eval_formula(ws, v[1:], depth + 1)
    if v is None:
        return 0.0
    if isinstance(v, bool) or not isinstance(v, (int, float)):
        raise ValueError(f"{coord} 不是數值：{v!r}")
    return float(v)


def _eval_formula(ws, f, depth=0):
    if depth > 30:
        raise ValueError("公式遞迴過深")

    def rng(m):
        col, r1, col2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        if col != col2:
            raise ValueError(f"不支援跨欄 SUM：{m.group(0)}")
        tot = 0.0
        for r in range(r1, r2 + 1):
            v = ws[f"{col}{r}"].value
            if isinstance(v, str) and v.startswith("="):
                tot += _eval_formula(ws, v[1:], depth + 1)
            elif isinstance(v, (int, float)) and not isinstance(v, bool):
                tot += v
        return repr(tot)

    def sumif(m):
        col, r1, col2, r2, op = m.group(1), int(m.group(2)), m.group(3), int(m.group(4)), m.group(5)
        if col != col2:
            raise ValueError(f"不支援跨欄 SUMIF：{m.group(0)}")
        vals = [_cell_num(ws, f"{col}{r}", depth) for r in range(r1, r2 + 1)]
        return repr(sum(v for v in vals if (v < 0 if op == "<" else v > 0)))

    f = re.sub(r'SUMIF\(([A-Z]+)(\d+):([A-Z]+)(\d+),"([<>])0"\)', sumif, f)
    expr = re.sub(r"SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)", rng, f)
    expr = _REF_RE.sub(lambda m: repr(_cell_num(ws, m.group(0), depth)), expr)
    expr = expr.replace("ABS(", "abs(")
    if not re.fullmatch(r"[0-9.eE+\-*/() abs]*", expr):
        raise ValueError(f"不支援的公式：={f}")
    return float(eval(expr, {"__builtins__": {}}, {"abs": abs}))


def verify_special(ws, lay, session, offset, check_hold=True):
    """特例日寫後核驗（公式實際求值，不是只看寫了什麼）。回傳 (errs, 當日損益)。
    check_hold=False：同一輪稍後的平倉日已把這個留倉區段抬走重寫，列號已失效，只驗區塊本體。"""
    errs = []
    hr, d0, dN = lay["hr"], lay["d0"], lay["dN"]
    legs = session.trades
    day = [t for t in legs if leg_kind(t) == "day"]
    closes = [t for t in legs if leg_kind(t) == "close"]
    day_pnl = round(sum(t.pnl for t in day if t.pnl is not None), 2)
    close_pnl = round(sum(t.pnl for t in closes), 2)
    try:
        hb = _cell_num(ws, f"B{hr}")
        tail = round(_cell_num(ws, f"B{dN}"), 2)
        total = round(_cell_num(ws, f"C{lay['tot']}"), 2)
        check = round(_cell_num(ws, f"C{lay['chk']}"), 2)
        s1 = round(sum(_cell_num(ws, f"A{r}") for r in range(d0, dN + 1)), 2)
        # §四.2 頭尾雙 offset／§四.3 Σcol1 == 收−開
        if abs(round(hb - session.open_bal, 2) - offset) > 0.01 or \
                abs(round(tail - session.close_bal, 2) - offset) > 0.01:
            errs.append(f"§4.2 頭尾offset head={hb - session.open_bal:.2f} "
                        f"tail={tail - session.close_bal:.2f} 期望={offset}")
        if abs(s1 - round(session.close_bal - session.open_bal, 2)) > 0.01:
            errs.append(f"§4.3 Σcol1({s1})≠收-開({round(session.close_bal - session.open_bal, 2)})")
        # 合計 == 核驗 == 引擎當沖 Σpnl（留倉/平倉腿不進合計，核驗扣掉它們的 A）
        if abs(total - day_pnl) > 0.01 or abs(check - day_pnl) > 0.01:
            errs.append(f"合計({total})/核驗({check}) ≠ 引擎當沖Σpnl({day_pnl})")
        if lay["day_pnl_row"] and abs(_cell_num(ws, f"C{lay['day_pnl_row']}") - day_pnl) > 0.01:
            errs.append(f"當沖彙總 當日損益 ≠ 引擎當沖Σpnl({day_pnl})")
        if not check_hold:
            pass
        elif lay["mode"] == "hold":
            want = lay["lot_expect"]
            got = [_cell_num(ws, f"C{r}") for r in lay["lot_rows"]]
            if len(got) != len(want) or any(abs(g - w) > 0.005 for g, w in zip(got, want)):
                errs.append(f"累積留倉列 {got} ≠ 引擎留倉淨額 {want}")
        else:
            pr = round(_cell_num(ws, f"C{lay['pnl_row']}"), 2)
            if abs(pr - close_pnl) > 0.01:
                errs.append(f"平倉損益({pr}) ≠ 引擎平倉Σpnl({close_pnl})")
            _cell_num(ws, f"C{lay['pnl_row'] + 1}")
    except (ValueError, ZeroDivisionError, SyntaxError) as e:
        errs.append(f"公式求值失敗：{e}")
    # §四.4 新 header 前空白列==2
    blanks = 0
    for r in range(hr - 1, 0, -1):
        if all(ws.cell(r, c).value in (None, "") for c in range(1, 6)):
            blanks += 1
        else:
            break
    if blanks != 2:
        errs.append(f"§4.4 header 前空白列={blanks}≠2")
    # §四.5 B 欄公式樣式
    for i, r in enumerate(range(d0, dN + 1)):
        want = f"=+B{hr}+A{r}" if i == 0 else f"=+B{r-1}+A{r}"
        if ws.cell(r, 2).value != want:
            errs.append(f"§4.5 B{r} 公式={ws.cell(r, 2).value!r}≠{want!r}"); break
    return errs, round(day_pnl + close_pnl, 2)


def verify_trades_special(ws, session):
    """trades_all 寫後核驗：每腿恰一列、留倉/平倉欄位與配對雙向一致、Date 欄為字串。"""
    errs = []
    rows = {}
    for r in range(2, ws.max_row + 1):
        if isinstance(ws.cell(r, 1).value, datetime):
            errs.append(f"§4.6 Date 欄型別錯誤（datetime）於列 {r}")
            break
        rows.setdefault((_norm_date(ws.cell(r, 1).value), str(ws.cell(r, 2).value).strip()), []).append(r)
    for t in session.trades:
        hit = rows.get((session.edt_date, t.edt_time), [])
        if len(hit) != 1:
            errs.append(f"{t.edt_time} 在 trades_all 有 {len(hit)} 列"); continue
        r = hit[0]
        st = _norm_status(ws.cell(r, 9).value)
        if st != t.status:
            errs.append(f"列 {r} Status={st!r}≠{t.status!r}")
        k = leg_kind(t)
        if k == "close":
            if ws.cell(r, 5).value != "平" or _norm_date(ws.cell(r, 10).value) != t.pair_date \
                    or str(ws.cell(r, 11).value).strip() != t.pair_time \
                    or not isinstance(ws.cell(r, 6).value, (int, float)) \
                    or abs(ws.cell(r, 6).value - t.pnl) > 0.005:
                errs.append(f"列 {r} 平倉欄位不符")
            en = rows.get((t.pair_date, t.pair_time), [])
            if len(en) != 1 or _norm_status(ws.cell(en[0], 9).value) != "2" \
                    or _norm_date(ws.cell(en[0], 10).value) != session.edt_date \
                    or str(ws.cell(en[0], 11).value).strip() != t.edt_time:
                errs.append(f"原留倉列 {t.pair_date} {t.pair_time} 未改 Status '2'／未回指")
        elif k == "hold" and t.status == "1":
            if any(ws.cell(r, c).value not in (None, "") for c in (5, 6, 10, 11)):
                errs.append(f"列 {r} 留倉腿不該有 Type/損益/Pair")
    return errs


def offset_note(session, offset, lay):
    """offset_state.json 當日註記（offset 數值不動，格式比照 10/16、10/19 人工註記）。"""
    legs = session.trades
    day = [t for t in legs if leg_kind(t) == "day" and t.pnl is not None]
    parts = [f"當沖{len(day)}組{sum(t.pnl for t in day):+.2f}" if day else "無當沖腿"]
    if lay["mode"] == "hold":
        holds = [t for t in legs if leg_kind(t) == "hold"]
        kind = "加碼留倉日" if len(lay["lot_expect"]) > len(holds) else "留倉建立日"
        parts.append(f"新增留倉 {sum(t.qty for t in holds)}股("
                     + "/".join(f"{t.price}" for t in holds) + ")")
        if kind == "加碼留倉日":
            parts.append(f"續抱前日留倉 {len(lay['lot_expect']) - len(holds)} 筆，累積留倉 "
                         f"{len(lay['lot_expect']) * 100}股")
    else:
        closes = sorted((t for t in legs if leg_kind(t) == "close"), key=lambda t: (t.pair_date, t.pair_time))
        kind = "平倉日"
        parts.append("LIFO 配對 " + "、".join(f"{t.pair_date[5:].replace('-', '/')} {t.pair_time}↔{t.edt_time} "
                                           f"{t.pnl:+.2f}" for t in closes)
                     + f"，平倉損益 {sum(t.pnl for t in closes):+.2f}")
    gaps = [n for n in session.notes if n.startswith("跨區塊 BAL 落差")]
    return (f"{session.edt_date} {kind}(writer 自動寫入，§3b)：offset 仍為 {offset:g}。"
            + "，".join(parts)
            + f"。header_B={session.open_bal + offset:.2f}，尾端B={session.close_bal + offset:.2f}(head=tail={offset:g})"
            + (f"。{'；'.join(gaps)}" if gaps else "") + "。")


# ── 寫後核驗（CLAUDE.md §四）────────────────────────────────────────────────
def verify_cs(ws, hr, d0, dN, session, offset):
    errs = []
    col1 = [ws.cell(r, 1).value for r in range(d0, dN + 1)]
    colC = [ws.cell(r, 3).value for r in range(d0, dN + 1) if isinstance(ws.cell(r, 3).value, (int, float))]
    s1 = round(sum(col1), 2); sC = round(sum(colC), 2)
    hb = ws.cell(hr, 2).value
    chain_close = round(hb + s1, 2)
    # §四.1 全平 Σcol1==Σpnl
    if abs(s1 - sC) > 0.01:
        errs.append(f"§4.1 Σcol1({s1})≠Σpnl({sC})")
    # §四.2 頭尾雙 offset
    head = round(hb - session.open_bal, 2)
    tail = round(chain_close - session.close_bal, 2)
    if not (abs(head - offset) < 0.01 and abs(tail - offset) < 0.01):
        errs.append(f"§4.2 頭尾offset head={head} tail={tail} 期望={offset}")
    # §四.3 獨立自洽
    if abs(s1 - round(session.close_bal - session.open_bal, 2)) > 0.01:
        errs.append(f"§4.3 Σcol1({s1})≠收-開({round(session.close_bal-session.open_bal,2)})")
    # §四.4 新 header 前空白列==2
    blanks = 0
    for r in range(hr - 1, 0, -1):
        if all(ws.cell(r, c).value in (None, "") for c in range(1, 6)):
            blanks += 1
        else:
            break
    if blanks != 2:
        errs.append(f"§4.4 header 前空白列={blanks}≠2")
    # §四.5 B 欄公式樣式
    for i, r in enumerate(range(d0, dN + 1)):
        want = f"=+B{hr}+A{r}" if i == 0 else f"=+B{r-1}+A{r}"
        if ws.cell(r, 2).value != want:
            errs.append(f"§4.5 B{r} 公式={ws.cell(r,2).value!r}≠{want!r}"); break
    return errs, sC


# ── 主 commit ────────────────────────────────────────────────────────────────
def commit_trades_only(sessions, trades_path=TRADES_PATH_DEFAULT):
    """第一段專用：只寫 trades_all，完全不碰 CS交易紀錄／offset_state。

    用於 Boss PC（交易機）—— 它只負責 CLAUDE.md 的第一段「對帳單→trades_all」，
    第二段「trades_all→CS」連同人工修正一律留在 502。Boss 上根本沒有 CS 檔，
    所以 header_B / verify_cs / write_cumulative / offset 這條線整條跳過。

    ⚠️ 安全閘門一條都沒放寬：gate() 檢查的 open/close 餘額明確性、每筆
    needs_review、以及 BALANCE 對不上／留倉／非交易現金，全都是 session
    層級的判斷，不依賴 CS。§1b 的逐列餘額重建也在引擎端就做完了。
    """
    blocked = {}
    for s in sessions:
        r = gate(s)
        if r:
            blocked[s.edt_date] = r
    if blocked:
        print("🔴 以下 session 需人工確認，全部不寫入：")
        for d, rs in blocked.items():
            print(f"  {d}:")
            for x in rs:
                print(f"    • {x}")
        return False

    check_lock(trades_path)
    b = backup(trades_path)
    print(f"§3 備份：{os.path.basename(b)}")

    try:
        wb_tr = openpyxl.load_workbook(trades_path)
        tr = wb_tr.active

        written = 0
        for s in sessions:
            if day_already_in_trades(tr, s.edt_date):
                print(f"— {s.edt_date} 已存在於 trades_all（冪等跳過，不重寫）")
                continue
            backlink_pairs(s)
            n_before = tr.max_row
            write_trades_all(tr, s)
            written += 1
            print(f"✔ {s.edt_date} 寫入 trades_all：{tr.max_row - n_before} 列")

        if written == 0:
            print("（無新資料可寫——所有 session 皆已存在，冪等跳過）")
            return True

        wb_tr.save(trades_path)

        # 寫後核驗（§四.6）：新列 Date 欄必須是 str，不能是 datetime
        wb_chk = openpyxl.load_workbook(trades_path)
        ws_chk = wb_chk.active
        bad = [r for r in range(2, ws_chk.max_row + 1)
               if isinstance(ws_chk.cell(r, 1).value, datetime)]
        if bad:
            raise RuntimeError(f"§4.6 Date 欄型別錯誤（datetime）於列 {bad[:5]}")
        for s in sessions:
            if not day_already_in_trades(ws_chk, s.edt_date):
                raise RuntimeError(f"§4 寫後核驗失敗：{s.edt_date} 未出現在 trades_all")

        print(f"✅ {written} 個交易日寫入 trades_all 完成並通過核驗")
        return True
    except Exception as e:
        shutil.copy2(b, trades_path)
        print(f"❌ 寫入失敗，已還原備份：{e}")
        raise


def commit(sessions, cs_path=CS_PATH_DEFAULT, trades_path=TRADES_PATH_DEFAULT,
           dry_report=True, text=None):
    st = load_offset()
    offset = float(st["offset"])
    # gate 全部 session（特例日：只容許多單收盤留倉／跨日平倉）
    blocked = {}
    for s in sessions:
        r = gate(s, allow_special=True)
        if r:
            blocked[s.edt_date] = r
    # 特例日必過 Cash Balance 完整性檢查：漏列一筆賣出，引擎會把當沖誤判成留倉（2026-10-22 事件）
    if not blocked and any(is_special(s) for s in sessions):
        ok, lines = statement_check(text) if text else (False, ["沒有對帳單原文可檢查"])
        if not ok:
            blocked["Cash Balance"] = ["CASH_BALANCE_INCOMPLETE：特例日不自動寫入"] + lines
    if blocked:
        print("🔴 以下 session 需人工確認，全部不寫入：")
        for d, rs in blocked.items():
            print(f"  {d}:")
            for x in rs:
                print(f"    • {x}")
        return False

    check_lock(cs_path); check_lock(trades_path)
    b1 = backup(cs_path); b2 = backup(trades_path)
    print(f"§3 備份：{os.path.basename(b1)} / {os.path.basename(b2)}")
    b3 = None
    cur = None

    try:
        wb_cs = openpyxl.load_workbook(cs_path)
        ws = wb_cs[SHEET_DETAIL]
        wb_tr = openpyxl.load_workbook(trades_path)
        tr = wb_tr.active

        written = 0
        specials = []   # (session, layout)：存檔後重驗＋offset 註記用
        for s in sessions:
            cur = s.edt_date
            dot = s.edt_date.replace("-", ".")
            if day_already_written(ws, dot):
                print(f"— {s.edt_date} 已存在於 CS（冪等跳過，不重寫）")
                continue
            backlink_pairs(s)
            header_B = round(s.open_bal + offset, 2)
            if is_special(s):
                plan = plan_special(ws, s)
                lay = write_special_block(ws, s, header_B, plan)
                errs, day_pnl = verify_special(ws, lay, s, offset)
                if errs:
                    raise RuntimeError(f"{s.edt_date} 核驗失敗：{errs}")
                write_trades_all(tr, s)
                flip_held_rows(tr, s)
                write_cumulative(wb_cs, s, day_pnl)
                specials.append((s, lay))
                written += 1
                kind = "留倉" if plan["mode"] == "hold" else "平倉"
                n = sum(t.qty for t in s.trades if leg_kind(t) == ("hold" if kind == "留倉" else "close"))
                print(f"✔ {s.edt_date} 特例日（{kind} {n}股）寫入：CS header row {lay['hr']}, "
                      f"{lay['dN']-lay['d0']+1} 明細, 累積留倉區段 {lay['h0']}-{lay['hold_end']}, "
                      f"當日損益 {day_pnl:+.2f}")
                print(f"SPECIAL_DAY={s.edt_date} {kind}")
                continue
            # 若檔尾有累積留倉區段 → 先抬起，寫完日區塊後再貼回檔尾
            last = last_nonempty_row(ws)
            hold_start, hold_rows = _detect_tail_hold_section(ws, last)
            if hold_start is not None:
                ws.delete_rows(hold_start, last - hold_start + 1)
                # 連同緊鄰的前導空行一併移除，避免累積空行
                while last_nonempty_row(ws) < ws.max_row:
                    ws.delete_rows(ws.max_row, 1)

            hr, d0, dN = write_cs_block(ws, s, header_B)
            errs, day_pnl = verify_cs(ws, hr, d0, dN, s, offset)
            if errs:
                raise RuntimeError(f"{s.edt_date} 核驗失敗：{errs}")

            if hold_rows:
                append_hold_section(ws, hold_rows, hold_start)

            write_trades_all(tr, s)
            write_cumulative(wb_cs, s, day_pnl)
            written += 1
            print(f"✔ {s.edt_date} 寫入：CS header row {hr}, {dN-d0+1} 明細, 當日損益 {day_pnl:+.2f}")

        if written == 0:
            print("（無新資料可寫——所有 session 皆已存在，冪等跳過）")
            return True
        cur = None
        wb_cs.save(cs_path)
        wb_tr.save(trades_path)
        if specials:
            # 存檔後重新讀檔再驗一次（確認落地的是驗過的內容），再補 offset 註記
            wb_chk = openpyxl.load_workbook(cs_path)
            tr_chk = openpyxl.load_workbook(trades_path).active
            for i, (s, lay) in enumerate(specials):
                moved = lay["mode"] == "hold" and any(l["lifted"] for _, l in specials[i + 1:])
                errs, _ = verify_special(wb_chk[SHEET_DETAIL], lay, s, offset, check_hold=not moved)
                errs += verify_trades_special(tr_chk, s)
                if errs:
                    raise RuntimeError(f"{s.edt_date} 存檔後核驗失敗：{errs}")
            b3 = OFFSET_STATE[:-5] + f"_backup_{_today()}.json"
            shutil.copy2(OFFSET_STATE, b3)
            for s, lay in specials:
                st["note"] = offset_note(s, offset, lay) + " || 前註：" + st.get("note", "")
                st["as_of"] = s.edt_date
            save_offset(st)
        print(f"✅ {written} 個交易日寫入完成並通過 §4 核驗")
        return True
    except Blocked as e:
        shutil.copy2(b1, cs_path); shutil.copy2(b2, trades_path)
        print("🔴 以下 session 需人工確認，全部不寫入：")
        print(f"  {cur}:")
        print(f"    • {e}")
        return False
    except Exception as e:
        shutil.copy2(b1, cs_path); shutil.copy2(b2, trades_path)
        if b3:
            shutil.copy2(b3, OFFSET_STATE)
        print(f"❌ 寫入失敗，已還原備份：{e}")
        raise


def main():
    import argparse
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("csv_path")
    ap.add_argument("--date")
    ap.add_argument("--cs", default=CS_PATH_DEFAULT)
    ap.add_argument("--trades", default=TRADES_PATH_DEFAULT)
    ap.add_argument("--commit", action="store_true")
    ap.add_argument("--trades-only", action="store_true",
                    help="只寫 trades_all，不碰 CS交易紀錄／offset_state（Boss PC 第一段用）")
    a = ap.parse_args()
    with open(a.csv_path, encoding="utf-8-sig") as fh:
        text = fh.read()
    fills = parse_statement(text)
    sessions = group_sessions(fills)
    # 整份一起處理：留倉在 session 間傳遞、從 trades_all 帶入未平留倉（2026-10-20）。
    # 特例日（留倉/跨日平倉）由 commit() 依核定範本寫入；--trades-only 照舊擋下。
    stats = process_all(sessions, fills, text, a.trades)
    for s, st in zip(sessions, stats):
        s.carry_out = st["carry"]   # 收盤後仍未平的 lots → plan_special 判斷是否落在範本內
    if a.date:
        sessions = [s for s in sessions if s.edt_date == a.date]
    if a.commit:
        # 鎖在呼叫端取得，一處涵蓋備份與兩條寫入路徑，
        # commit()/commit_trades_only() 內部完全不必改動。
        with writer_lock():
            if a.trades_only:
                ok = commit_trades_only(sessions, a.trades)
            else:
                ok = commit(sessions, a.cs, a.trades, text=text)
        sys.exit(0 if ok else 2)   # 被 gate 擋下時回非 0，讓呼叫端腳本知道別繼續
    else:
        mode = "trades-only" if a.trades_only else "full"
        print(f"[DRY-RUN] writer 未寫入（{mode}）。session:", [s.edt_date for s in sessions])


if __name__ == "__main__":
    main()
