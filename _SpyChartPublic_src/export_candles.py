#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""把 SPY 歷史 1 分鐘 K 線匯出成靜態 JSON（給 GitHub Pages 用）。

刻意只輸出 candles，不含任何交易紀錄 —— 這是公開版與內網複盤版的唯一差別。
資料來源優先序：
  1. 502 本機的 history_minute.xlsx（一次性把歷史全匯）
  2. yfinance（GitHub Actions 每天增量更新最新交易日用；美股，海外抓得到）

用法：
  python export_candles.py --from-history <path>   # 從 xlsx 全匯（本機一次性）
  python export_candles.py --recent 5              # 用 yfinance 補最近 5 天（Actions 用）
  python export_candles.py --recent 5 --out ./site # 指定輸出目錄
"""
import argparse
import json
import os
import sys
from datetime import datetime, timedelta, timezone

HERE = os.path.dirname(os.path.abspath(__file__))


def log(m):
    try:
        print(m, flush=True)
    except Exception:
        pass


def _vol(v):
    """Volume 可能是純數字，也可能是 '2.597M' / '1.2K' 這種字串（來自不同來源）。"""
    try:
        return int(v)
    except (ValueError, TypeError):
        pass
    s = str(v).strip().replace(",", "")
    mult = 1
    if s and s[-1] in "KkMmBb":
        mult = {"k": 1e3, "m": 1e6, "b": 1e9}[s[-1].lower()]
        s = s[:-1]
    try:
        return int(float(s) * mult)
    except (ValueError, TypeError):
        return 0


def write_day(out_dir, date_str, candles):
    """寫單日 JSON；candles 為 [{t,o,h,l,c,v}, ...]。"""
    d = os.path.join(out_dir, "data")
    os.makedirs(d, exist_ok=True)
    payload = {"date": date_str, "candles": candles}
    with open(os.path.join(d, f"{date_str}.json"), "w", encoding="utf-8") as f:
        json.dump(payload, f, separators=(",", ":"), ensure_ascii=False)
    return len(candles)


def rebuild_index(out_dir):
    """掃描 data/ 重建 dates.json 與 meta.json（含最後更新時戳）。"""
    d = os.path.join(out_dir, "data")
    dates = sorted((f[:-5] for f in os.listdir(d) if f.endswith(".json")
                    and len(f) == 15 and f[4] == "-"), reverse=True)
    with open(os.path.join(out_dir, "dates.json"), "w", encoding="utf-8") as f:
        json.dump(dates, f, separators=(",", ":"))
    # UTC 時戳；前端據此顯示「資料更新於」並判斷新鮮度
    meta = {"updated_utc": datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "latest": dates[0] if dates else None,
            "count": len(dates)}
    with open(os.path.join(out_dir, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f, separators=(",", ":"))
    return len(dates), (dates[0] if dates else None)


def from_history(xlsx_path, out_dir):
    import pandas as pd
    log(f"讀取 {xlsx_path} ...")
    df = pd.read_excel(xlsx_path)
    df.columns = [str(c).strip() for c in df.columns]
    df["Date"] = pd.to_datetime(df["Date"])
    df["d"] = df["Date"].dt.strftime("%Y-%m-%d")
    total = 0
    for day, g in df.groupby("d"):
        candles = [{"t": r["Date"].strftime("%H:%M"),
                    "o": round(float(r["Open"]), 4), "h": round(float(r["High"]), 4),
                    "l": round(float(r["Low"]), 4), "c": round(float(r["Close"]), 4),
                    "v": _vol(r["Volume"])}
                   for _, r in g.iterrows()]
        write_day(out_dir, day, candles)
        total += len(candles)
    n, latest = rebuild_index(out_dir)
    log(f"完成：{n} 個交易日、{total:,} 根 K 線，最新 {latest}")


def from_yfinance(recent_days, out_dir):
    import yfinance as yf
    import pandas as pd
    # 1 分鐘資料 yfinance 只回最近約 30 天；抓一段再依日期切
    period = f"{min(recent_days + 2, 30)}d"
    log(f"yfinance SPY 1m period={period} ...")
    h = yf.Ticker("SPY").history(period=period, interval="1m", auto_adjust=False)
    if h.empty:
        log("yfinance 回傳空，未更新"); return
    # yfinance index 是 tz-aware(美東)，直接用當地時間切日與取 HH:MM
    h = h.tz_convert("America/New_York")
    h["d"] = h.index.strftime("%Y-%m-%d")
    days = sorted(h["d"].unique())[-recent_days:]
    total = 0
    for day in days:
        g = h[h["d"] == day]
        # 只取正常盤 09:30–16:00（排除盤前盤後，與本機 history 對齊）
        candles = []
        for ts, r in g.iterrows():
            hm = ts.strftime("%H:%M")
            if hm < "09:30" or hm > "16:00":
                continue
            candles.append({"t": hm,
                            "o": round(float(r["Open"]), 4), "h": round(float(r["High"]), 4),
                            "l": round(float(r["Low"]), 4), "c": round(float(r["Close"]), 4),
                            "v": _vol(r["Volume"])})
        if candles:
            write_day(out_dir, day, candles)
            total += len(candles)
            log(f"  {day}: {len(candles)} 根")
    n, latest = rebuild_index(out_dir)
    log(f"完成：更新 {len(days)} 天、共 {total:,} 根；目前 {n} 個交易日，最新 {latest}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--from-history", metavar="XLSX")
    ap.add_argument("--recent", type=int, metavar="N")
    ap.add_argument("--out", default=os.path.join(HERE, "site"))
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    os.makedirs(a.out, exist_ok=True)
    if a.from_history:
        from_history(a.from_history, a.out)
    elif a.recent:
        from_yfinance(a.recent, a.out)
    else:
        ap.error("需指定 --from-history 或 --recent")


if __name__ == "__main__":
    main()
