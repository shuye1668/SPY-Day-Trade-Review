"""
Bloomberg G Chart Trade Review — 本地 Python 版（不需要 BQNT）
================================================================
適用於：Shared Terminal / Non-BBA 用戶
需求：Bloomberg Terminal 在背景運行 + 本地 Python 3.8+

首次安裝（在 Windows CMD 或 PowerShell 執行一次即可）：

  pip install --index-url=https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi
  pip install pandas numpy matplotlib openpyxl

每日使用：
  1. 確認 Bloomberg Terminal 已登入
  2. 修改下方 CONFIG 區塊
  3. 執行本腳本：python trade_review.py
"""

# ============================================================
# CONFIG — 每天只改這裡
# ============================================================
import os as _os
_BASE = _os.path.dirname(_os.path.abspath(__file__))
TRADE_EXCEL_PATH = _os.path.join(_BASE, "data", "trades.xlsx")    # 你的交易紀錄 Excel
TRADE_DATE       = "2026-04-01"                          # 交易日期 YYYY-MM-DD
OUTPUT_PNG       = _os.path.join(_BASE, "output", "20260401.png")   # 輸出 PNG 路徑
TICKER           = "SPY US Equity"

# 市場概述文字框（留空 "" 則不顯示）
MARKET_NOTES = """Middle East ceasefire hopes drove SPX
+0.7% and NDX +1.2% for a second straight session after
Trump claimed Iran requested a truce. Memory/storage
semiconductors led — WDC +10%, MU +8.9%, INTC +8.8%
— as Brent crude fell 2.8% to ~$101. Goldman and
JPMorgan trading desks flagged the rally as primarily
short-covering, not genuine re-risking. ISM Prices
Paid surged to 78.3, the strongest input-cost signal
in months, currently masked by geopolitical optimism.
NKE cratered 15.5% on weak guidance. Weekend selloff
pattern from the war period remains a key near-term risk.

"""

# 文字框位置（0~1，可微調）
NOTES_X_FRAC = 0.72
NOTES_Y_FRAC = 0.55

# ============================================================
# 0. IMPORTS
# ============================================================
import datetime as dt
import os
import sys
import warnings
warnings.filterwarnings("ignore")

import pandas as pd
import numpy as np
import matplotlib
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
from matplotlib.patches import Rectangle
from matplotlib.lines import Line2D

# Windows 中文字型支援
matplotlib.rcParams["font.sans-serif"] = ["Microsoft JhengHei", "SimHei", "Arial"]
matplotlib.rcParams["axes.unicode_minus"] = False

# ============================================================
# 1. BLOOMBERG 色彩系統（從 G Chart 截圖精確取色）
# ============================================================
class BBG:
    BG           = "#000000"
    GRID         = "#1A1D22"
    GRID_MAJOR   = "#252830"
    TEXT         = "#8B8F98"
    TEXT_BRIGHT  = "#C0C4CC"
    CANDLE_UP    = "#FFFFFF"    # 上漲=白
    CANDLE_DOWN  = "#3D6FCC"    # 下跌=藍
    WICK_UP      = "#7BA3F0"
    WICK_DOWN    = "#3D6FCC"
    LONG_TRADE   = "#00CC44"    # 做多=綠
    SHORT_TRADE  = "#FF3333"    # 做空=紅
    PNL_WIN      = "#FFFFFF"    # 獲利=白
    PNL_LOSS     = "#FF4444"    # 虧損=紅
    TITLE_BAR    = "#FF8C00"    # Bloomberg 橘
    LAST_PRICE   = "#FFD700"    # 最新價=黃
    BORDER       = "#333640"
    NOTES_BG     = "#1A1D22"
    NOTES_BORDER = "#3A3E48"
    NOTES_TEXT   = "#E0E0E0"

# ============================================================
# 2. 資料拉取（blpapi Desktop API）
# ============================================================
def fetch_intraday_blpapi(ticker, date_str):
    """
    透過 blpapi Desktop API 拉取 1-min intraday bars。
    Bloomberg Terminal 必須在背景運行且已登入。
    Shared Terminal (Non-BBA) 完全支援 Desktop API。
    """
    import blpapi

    so = blpapi.SessionOptions()
    so.setServerHost("localhost")
    so.setServerPort(8194)

    sess = blpapi.Session(so)
    if not sess.start():
        raise ConnectionError(
            "無法連線到 Bloomberg Terminal！\n"
            "請確認：\n"
            "  1. Bloomberg Terminal 正在運行且已登入\n"
            "  2. 不是鎖定/登出狀態\n"
            "  3. 沒有其他程式佔用 port 8194"
        )

    if not sess.openService("//blp/refdata"):
        sess.stop()
        raise ConnectionError("無法開啟 //blp/refdata 服務")

    svc = sess.getService("//blp/refdata")
    req = svc.createRequest("IntradayBarRequest")
    req.set("security", ticker)
    req.set("eventType", "TRADE")
    req.set("interval", 1)  # 1 分鐘

    d = dt.datetime.strptime(date_str, "%Y-%m-%d")
    # Bloomberg API 使用 UTC：ET 09:30 = UTC 13:30, ET 16:00 = UTC 20:00
    req.set("startDateTime", dt.datetime(d.year, d.month, d.day, 13, 30))
    req.set("endDateTime",   dt.datetime(d.year, d.month, d.day, 20, 0))

    sess.sendRequest(req)

    bars = []
    while True:
        ev = sess.nextEvent(5000)
        for msg in ev:
            if msg.hasElement("barData"):
                bd = msg.getElement("barData").getElement("barTickData")
                for i in range(bd.numValues()):
                    b = bd.getValueAsElement(i)
                    bars.append({
                        "datetime": b.getElementAsDatetime("time"),
                        "Open":     b.getElementAsFloat("open"),
                        "High":     b.getElementAsFloat("high"),
                        "Low":      b.getElementAsFloat("low"),
                        "Close":    b.getElementAsFloat("close"),
                        "Volume":   b.getElementAsInteger("volume"),
                    })
        if ev.eventType() == blpapi.Event.RESPONSE:
            break

    sess.stop()

    if not bars:
        raise ValueError(
            f"未取得任何資料！\n"
            f"請確認：\n"
            f"  1. 日期 {date_str} 是交易日（非假日/週末）\n"
            f"  2. Intraday 資料約保留 140 天\n"
            f"  3. Ticker '{ticker}' 正確"
        )

    df = pd.DataFrame(bars)
    df["datetime"] = pd.to_datetime(df["datetime"])
    # UTC → 美東時間
    df["datetime"] = (df["datetime"]
                      .dt.tz_localize("UTC")
                      .dt.tz_convert("US/Eastern")
                      .dt.tz_localize(None))
    return df

# ============================================================
# 3. 交易紀錄讀取 & 配對
# ============================================================
def read_trades(path, date_str):
    if not os.path.exists(path):
        raise FileNotFoundError(
            f"找不到交易紀錄：{path}\n"
            f"請確認檔案路徑正確。"
        )
    df = pd.read_excel(path, dtype=str)
    df.columns = [c.strip() for c in df.columns]

    required = ["Exec Time(EDT)", "Price", "Shares"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise KeyError(
            f"Excel 缺少必要欄位：{missing}\n"
            f"需要的欄位：Exec Time(EDT), Symbol, Price, Type, 損益(AI辨識), Shares"
        )

    base = dt.datetime.strptime(date_str, "%Y-%m-%d").date()
    df["datetime"] = df["Exec Time(EDT)"].apply(
        lambda t: dt.datetime.combine(base, dt.datetime.strptime(t.strip(), "%H:%M:%S").time()))
    df["Price"] = df["Price"].astype(float)
    df["Shares"] = df["Shares"].astype(int)
    df["Type"] = df["Type"].fillna("").str.strip() if "Type" in df.columns else ""
    pnl_col = [c for c in df.columns if "損益" in c or "PnL" in c.lower()]
    df["PnL_raw"] = df[pnl_col[0]].fillna("").str.strip() if pnl_col else ""
    return df

def pair_trades(df):
    trades, entry = [], None
    for _, r in df.iterrows():
        is_exit = r["Type"] in ("多", "空") and r["PnL_raw"] != ""
        if not is_exit:
            entry = r
        elif entry is not None:
            pnl_s = r["PnL_raw"].replace("+", "").replace("＋", "").replace(",", "")
            try: pnl = float(pnl_s)
            except: pnl = 0.0
            trades.append({
                "entry_time": entry["datetime"], "entry_price": entry["Price"],
                "exit_time": r["datetime"], "exit_price": r["Price"],
                "direction": r["Type"], "pnl": pnl, "shares": r["Shares"],
            })
    return trades

# ============================================================
# 4. 繪圖引擎 — Bloomberg G Chart 精確還原
# ============================================================
def plot_gchart(candles, trades, date_str, market_notes="", output_path=None):
    df = candles.copy().sort_values("datetime").reset_index(drop=True)
    df["idx"] = range(len(df))

    def nearest_idx(target):
        return df.loc[(df["datetime"] - target).abs().idxmin(), "idx"]

    fig, ax = plt.subplots(figsize=(22, 8), facecolor=BBG.BG)
    ax.set_facecolor(BBG.BG)

    # 格線
    ax.yaxis.grid(True, color=BBG.GRID, linewidth=0.4, linestyle="--", alpha=0.5)
    ax.xaxis.grid(False)
    for _, row in df.iterrows():
        t = row["datetime"]
        if t.minute == 0:
            ax.axvline(row["idx"], color=BBG.GRID_MAJOR, linewidth=0.4, linestyle="--", alpha=0.4)
        elif t.minute == 30:
            ax.axvline(row["idx"], color=BBG.GRID, linewidth=0.3, linestyle="--", alpha=0.3)

    # K 線
    bar_width = 0.55
    for _, row in df.iterrows():
        i = row["idx"]
        o, h, l, c = row["Open"], row["High"], row["Low"], row["Close"]
        is_up = c >= o
        ax.plot([i, i], [l, h],
                color=BBG.WICK_UP if is_up else BBG.WICK_DOWN,
                linewidth=0.6, zorder=2, solid_capstyle="round")
        body_bottom = min(o, c)
        body_height = max(abs(c - o), 0.005)
        rect = Rectangle((i - bar_width/2, body_bottom), bar_width, body_height,
                          facecolor=BBG.CANDLE_UP if is_up else BBG.CANDLE_DOWN,
                          edgecolor=BBG.CANDLE_UP if is_up else BBG.CANDLE_DOWN,
                          linewidth=0.3, zorder=3)
        ax.add_patch(rect)

    # 交易標註
    total_pnl = 0.0
    y_range = df["High"].max() - df["Low"].min()
    used_positions = []

    for trade in trades:
        ei = nearest_idx(trade["entry_time"])
        xi = nearest_idx(trade["exit_time"])
        ep, xp = trade["entry_price"], trade["exit_price"]
        pnl = trade["pnl"]
        total_pnl += pnl
        line_color = BBG.LONG_TRADE if trade["direction"] == "多" else BBG.SHORT_TRADE

        # 圓點虛線
        ax.plot([ei, xi], [ep, xp], color=line_color, linewidth=3.0,
                linestyle=(0, (1, 1.5)), solid_capstyle="round", dash_capstyle="round",
                alpha=0.95, zorder=6)
        for px, py in [(ei, ep), (xi, xp)]:
            ax.plot(px, py, "o", color=line_color, markersize=4.5,
                    markeredgecolor=line_color, markeredgewidth=0.8, zorder=7)

        # 損益文字
        pnl_str = f"+{pnl:.2f}" if pnl >= 0 else f"{pnl:.2f}"
        pnl_color = BBG.PNL_WIN if pnl >= 0 else BBG.PNL_LOSS
        mid_x = (ei + xi) / 2
        text_y = max(ep, xp) + y_range * 0.012
        for ux, uy in used_positions:
            if abs(mid_x - ux) < 15 and abs(text_y - uy) < y_range * 0.03:
                text_y = uy + y_range * 0.025
        used_positions.append((mid_x, text_y))
        ax.text(mid_x, text_y, pnl_str, color=pnl_color,
                fontsize=10, fontweight="bold", ha="center", va="bottom",
                fontfamily="monospace", zorder=8)

    # Y 軸（右側）
    ax.yaxis.tick_right()
    ax.yaxis.set_label_position("right")
    ax.yaxis.set_major_formatter(mticker.FormatStrFormatter("%.2f"))
    ax.tick_params(axis="y", colors=BBG.TEXT, labelsize=9, length=0, pad=8)

    # 最新價黃色標籤
    last_price = df.iloc[-1]["Close"]
    ax.annotate(f"{last_price:.2f}",
                xy=(len(df) - 1, last_price), xytext=(len(df) + 3, last_price),
                fontsize=9, fontweight="bold", color="#000000", fontfamily="monospace",
                bbox=dict(boxstyle="round,pad=0.2", facecolor=BBG.LAST_PRICE, edgecolor="none", alpha=0.95),
                arrowprops=dict(arrowstyle="-", color=BBG.LAST_PRICE, lw=0.5),
                ha="left", va="center", zorder=10)

    ax.set_ylim(df["Low"].min() - y_range * 0.03, df["High"].max() + y_range * 0.08)
    ax.set_xlim(-3, len(df) + 12)

    # X 軸時間
    tick_pos, tick_lab = [], []
    for _, row in df.iterrows():
        t = row["datetime"]
        if t.minute in (0, 30) and t.second == 0:
            tick_pos.append(row["idx"])
            tick_lab.append(t.strftime("%H:%M"))
    ax.set_xticks(tick_pos)
    ax.set_xticklabels(tick_lab, fontsize=8, color=BBG.TEXT, fontfamily="monospace")
    ax.tick_params(axis="x", length=0, pad=6)

    # 日期標記
    d = dt.datetime.strptime(date_str, "%Y-%m-%d")
    try:    date_label = d.strftime("%-d %b %Y")
    except: date_label = d.strftime("%d %b %Y").lstrip("0")
    ax.text(0.5, -0.06, date_label, transform=ax.transAxes,
            fontsize=9, color=BBG.TEXT, ha="center", fontfamily="monospace")

    # 邊框
    for spine in ax.spines.values():
        spine.set_color(BBG.BORDER)
        spine.set_linewidth(0.5)

    # 頂部標題列
    fig.text(0.01, 0.97, "SPY US Equity", fontsize=12, fontweight="bold",
             color="#FFFFFF", fontfamily="monospace",
             bbox=dict(facecolor="#1A1D22", edgecolor="none", pad=3),
             transform=fig.transFigure, va="top")
    fig.text(0.99, 0.97, "G 30: Intraday Candle Chart", fontsize=10,
             color="#000000", fontfamily="monospace", fontweight="bold",
             bbox=dict(facecolor=BBG.TITLE_BAR, edgecolor="none", pad=3),
             transform=fig.transFigure, va="top", ha="right")
    fig.text(0.01, 0.93, f"Period  1  Range  1  09:30 - 16:00    1 Min",
             fontsize=8, color=BBG.TEXT, fontfamily="monospace",
             transform=fig.transFigure, va="top")

    # Day P&L
    total_str = f"+${total_pnl:.2f}" if total_pnl >= 0 else f"-${abs(total_pnl):.2f}"
    total_color = BBG.PNL_WIN if total_pnl >= 0 else BBG.PNL_LOSS
    fig.text(0.5, 0.97, f"Day P&L: {total_str}   ({len(trades)} trades)",
             fontsize=11, fontweight="bold", color=total_color,
             fontfamily="monospace", transform=fig.transFigure, va="top", ha="center")

    # 市場概述文字框
    if market_notes.strip():
        xlim, ylim = ax.get_xlim(), ax.get_ylim()
        nx = xlim[0] + (xlim[1] - xlim[0]) * NOTES_X_FRAC
        ny = ylim[0] + (ylim[1] - ylim[0]) * NOTES_Y_FRAC
        ax.text(nx, ny, market_notes.strip(),
                fontsize=8.5, color=BBG.NOTES_TEXT,
                fontfamily="monospace", linespacing=1.6,
                va="top", ha="left",
                bbox=dict(boxstyle="round,pad=0.6", facecolor=BBG.NOTES_BG,
                          edgecolor=BBG.NOTES_BORDER, linewidth=0.8, alpha=0.92),
                zorder=9)

    plt.subplots_adjust(left=0.02, right=0.94, top=0.90, bottom=0.08)

    # 輸出
    if output_path:
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        fig.savefig(output_path, dpi=200, bbox_inches="tight",
                    facecolor=BBG.BG, edgecolor="none")
        print(f"  [OK] PNG 已存檔 → {output_path}")

    plt.show()
    return fig

# ============================================================
# 5. MAIN
# ============================================================
def main():
    print()
    print(f"  ╔══════════════════════════════════════════╗")
    print(f"  ║  Bloomberg G Chart Trade Review          ║")
    print(f"  ║  {TRADE_DATE}                            ║")
    print(f"  ╚══════════════════════════════════════════╝")
    print()

    # Step 1：拉取 intraday 資料
    print(f"  [1/3] 拉取 {TICKER} 1-min intraday bars ...")
    try:
        candles = fetch_intraday_blpapi(TICKER, TRADE_DATE)
        print(f"        → {len(candles)} bars ({candles['datetime'].iloc[0].strftime('%H:%M')} ~ {candles['datetime'].iloc[-1].strftime('%H:%M')})")
    except ImportError:
        print()
        print("  ❌ 找不到 blpapi 套件！請先安裝：")
        print()
        print("     pip install --index-url=https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi")
        print()
        sys.exit(1)
    except ConnectionError as e:
        print(f"\n  ❌ {e}")
        sys.exit(1)
    except Exception as e:
        print(f"\n  ❌ 資料拉取失敗：{e}")
        sys.exit(1)

    # Step 2：讀取交易紀錄
    print(f"  [2/3] 讀取交易紀錄 ...")
    try:
        tdf = read_trades(TRADE_EXCEL_PATH, TRADE_DATE)
        trades = pair_trades(tdf)
    except Exception as e:
        print(f"\n  ❌ {e}")
        sys.exit(1)

    print(f"        → {len(trades)} 筆交易")
    total = 0
    for i, t in enumerate(trades, 1):
        a = "▲ 多" if t["direction"] == "多" else "▼ 空"
        s = "+" if t["pnl"] >= 0 else ""
        print(f"        [{i}] {a}  {t['entry_time'].strftime('%H:%M:%S')} @{t['entry_price']:.2f}"
              f" → {t['exit_time'].strftime('%H:%M:%S')} @{t['exit_price']:.2f}"
              f"  {s}{t['pnl']:.2f}  ({t['shares']}sh)")
        total += t["pnl"]
    print(f"        ─────────────────────────────────")
    print(f"        Day Total: {'+'if total>=0 else ''}{total:.2f}")

    # Step 3：繪圖
    print(f"  [3/3] 繪製 G Chart ...")
    fig = plot_gchart(candles, trades, TRADE_DATE, MARKET_NOTES, OUTPUT_PNG)
    print()
    print(f"  ✓ 完成！")
    if OUTPUT_PNG:
        print(f"  ✓ 圖片存檔：{OUTPUT_PNG}")
    print()
    return fig


if __name__ == "__main__":
    main()
