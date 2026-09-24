#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
spy_daytrade_writer.py — spy_daytrade_engine.py 的 --commit 寫入層（2026-07-24 接上）

把引擎分析結果寫入 live 財務檔，嚴格對齊 CLAUDE.md §一/§二/§三/§四。
只處理「純當日進出、offset 不變」的標準日；任何特例（留倉/跨日平倉/非100股/
BALANCE 對不上/offset 不確定）→ 由引擎的 needs_review/FLAG 攔下，writer 直接 abort，
不寫入、回報需人工確認。這對齊 §0：絕不為跑完流程而靜默寫猜值。

安全：§3 寫前備份、§5 檢查 ~$ 鎖檔、§4 寫後全核驗、失敗還原備份。

offset 模型（CLAUDE.md §二）：Day Header B = 券商開盤餘額 + offset。
offset 是常數，只在留倉/平倉日跳動——而那些日子會被引擎攔下 abort，
所以 auto-commit 的日子 offset 不變。offset 值持久化在 offset_state.json，
人工處理留倉/平倉日後手動更新該檔。
"""
import contextlib
import json
import os
import shutil
from datetime import datetime

import openpyxl

from spy_daytrade_engine import (
    parse_statement, group_sessions, process_session, FLAG,
)

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


def gate(session):
    reasons = []
    if session.open_bal is None or session.close_bal is None:
        reasons.append("session 開盤/收盤餘額不明")
    for t in session.trades:
        if t.needs_review:
            reasons.append(f"{t.edt_time} {t.review_reason}")
    for w in session.warnings:
        if FLAG in w or "BALANCE 對不上" in w or "留倉" in w or "非交易現金" in w:
            reasons.append(w)
    return reasons


# ── 反向補 pair_time（引擎只在平倉端設 pair_time，進場端要補回指）────────────
def backlink_pairs(session):
    exits = [t for t in session.trades if t.pair_time]
    for ex in exits:
        for en in session.trades:
            if en is ex:
                continue
            if not en.pair_time and en.type_tag == "" and en.edt_time == ex.pair_time:
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


def append_hold_section(ws, rows):
    ws.append([None, None, None])  # 1 空行
    for (a, b, c) in rows:
        ws.append([a, b, c])


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
        ws.append([session.edt_date, t.edt_time, t.symbol, t.price, t.type_tag or None,
                   t.pnl, t.qty, t.action, t.status, t.pair_date or session.edt_date, t.pair_time])


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
           dry_report=True):
    st = load_offset()
    offset = float(st["offset"])
    # gate 全部 session
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

    check_lock(cs_path); check_lock(trades_path)
    b1 = backup(cs_path); b2 = backup(trades_path)
    print(f"§3 備份：{os.path.basename(b1)} / {os.path.basename(b2)}")

    try:
        wb_cs = openpyxl.load_workbook(cs_path)
        ws = wb_cs[SHEET_DETAIL]
        wb_tr = openpyxl.load_workbook(trades_path)
        tr = wb_tr.active

        written = 0
        for s in sessions:
            dot = s.edt_date.replace("-", ".")
            if day_already_written(ws, dot):
                print(f"— {s.edt_date} 已存在於 CS（冪等跳過，不重寫）")
                continue
            backlink_pairs(s)
            header_B = round(s.open_bal + offset, 2)
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
                append_hold_section(ws, hold_rows)

            write_trades_all(tr, s)
            write_cumulative(wb_cs, s, day_pnl)
            written += 1
            print(f"✔ {s.edt_date} 寫入：CS header row {hr}, {dN-d0+1} 明細, 當日損益 {day_pnl:+.2f}")

        if written == 0:
            print("（無新資料可寫——所有 session 皆已存在，冪等跳過）")
            return True
        wb_cs.save(cs_path)
        wb_tr.save(trades_path)
        print(f"✅ {written} 個交易日寫入完成並通過 §4 核驗")
        return True
    except Exception as e:
        shutil.copy2(b1, cs_path); shutil.copy2(b2, trades_path)
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
    sessions = group_sessions(parse_statement(text))
    if a.date:
        sessions = [s for s in sessions if s.edt_date == a.date]
    for s in sessions:
        process_session(s)
    if a.commit:
        # 鎖在呼叫端取得，一處涵蓋備份與兩條寫入路徑，
        # commit()/commit_trades_only() 內部完全不必改動。
        with writer_lock():
            if a.trades_only:
                ok = commit_trades_only(sessions, a.trades)
            else:
                ok = commit(sessions, a.cs, a.trades)
        sys.exit(0 if ok else 2)   # 被 gate 擋下時回非 0，讓呼叫端腳本知道別繼續
    else:
        mode = "trades-only" if a.trades_only else "full"
        print(f"[DRY-RUN] writer 未寫入（{mode}）。session:", [s.edt_date for s in sessions])


if __name__ == "__main__":
    main()
