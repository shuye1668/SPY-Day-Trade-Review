#!/usr/bin/env python3
"""Boss PC 寫入成功後更新 sync_state.json（資料新鮮度戳記）。

為什麼需要：git push 失敗是靜默的。502 那台 pull 不到新資料時，畫面看起來
和「今天還沒開盤」一模一樣 —— 你不會知道是沒資料還是同步壞了。這支腳本
記下「最後成功寫入的交易日 + 時間」，隨 git 一起同步，App 標頭就能顯示
「資料截至 X」，超時轉紅。這是對齊 CLAUDE.md §0「絕不靜默」的必要件。

用法（Boss PC，writer 成功之後）：
    python write_sync_state.py --date 2026-08-10
    python write_sync_state.py --date 2026-08-10 --status blocked --note "留倉需人工確認"
"""
import argparse
import json
import os
import socket
import sys
from datetime import datetime, timezone

_BASE = os.path.dirname(os.path.abspath(__file__))
STATE = os.path.join(_BASE, "sync_state.json")
TRADES = os.path.join(_BASE, "trades_all.xlsx")


def latest_date_in_trades():
    """從 trades_all 取最新交易日，當作 --date 未給時的來源。"""
    try:
        import openpyxl
        ws = openpyxl.load_workbook(TRADES, read_only=True).active
        best = ""
        for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
            v = row[0]
            if v is None:
                continue
            s = v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else str(v).strip()[:10]
            if len(s) == 10 and s > best:
                best = s
        return best or None
    except Exception:
        return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--date", help="本次處理的交易日 YYYY-MM-DD（省略則取 trades_all 最新日）")
    ap.add_argument("--status", default="ok", choices=["ok", "blocked", "failed"],
                    help="ok=已寫入；blocked=被 gate 擋下需人工；failed=流程失敗")
    ap.add_argument("--note", default="", help="狀態說明，會顯示在 App 上")
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    date = a.date or latest_date_in_trades()
    state = {
        "last_trade_date": date,
        "status": a.status,
        "note": a.note,
        "updated_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "host": socket.gethostname(),
        "latest_in_trades_all": latest_date_in_trades(),
    }
    with open(STATE, "w", encoding="utf-8", newline="\n") as f:
        json.dump(state, f, ensure_ascii=False, indent=2)
        f.write("\n")
    print(f"sync_state.json: {a.status}  交易日={date}  ({state['updated_utc']})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
