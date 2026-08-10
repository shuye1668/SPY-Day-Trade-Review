"""
Bloomberg G Chart Trade Review — Web App v3 (Virtual Timeline)
===============================================================
Cross-day seamless scrolling via virtual coordinate system.
Each day's data is stored independently; never merged.

Reads candles from Bloomberg Desktop API (blpapi).
Reads trades from a single consolidated Excel file `trades_all.xlsx`
with a Date column — same file shared with trade_review_app_free.py.

Folder layout:
  D:\\fileserver_D\\TradeReview\\
    trades_all.xlsx           ← all trades, with Date column (YYYY-MM-DD)
    notes\\
      2026-04-08.txt          ← per-day notes
    candles\\                  ← optional cache (used by free version)

pip install flask pandas numpy openpyxl
pip install --index-url=https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi
python trade_review_app.py → http://localhost:5500
"""
import os as _os
# 路徑一律相對「本腳本所在資料夾」——整個資料夾搬到任何機器/任何路徑都能跑，
# 資料夾改名也不會壞（2026-07-24 改；原本寫死 C:\TradeReview）。
ROOT_FOLDER=_os.path.dirname(_os.path.abspath(__file__))
TRADES_FILE=_os.path.join(ROOT_FOLDER,"trades_all.xlsx")
NOTES_FOLDER=_os.path.join(ROOT_FOLDER,"notes")
COLORS_FILE=_os.path.join(ROOT_FOLDER,"colors.xlsx")
DIVIDENDS_FILE=_os.path.join(ROOT_FOLDER,"dividends.xlsx")
TICKER="SPY US Equity"
PORT=5500

import datetime as dt,json,os,glob,warnings,sys,logging,time
warnings.filterwarnings("ignore")
import pandas as pd,numpy as np
from flask import Flask,jsonify,request

# 2026-07-31：pythonw.exe（背景靜默、無主控台）下 sys.stdout/stderr 為 None，
# 本檔多處 print() 與 Flask/werkzeug 請求日誌會丟例外 → 背景常駐秒退（表面像沒常駐）。
# 對策：無 console 時把 stdout/stderr 導向 _logs\trade_review_app.log，並降低 werkzeug 日誌量。
def _pythonw_safe_streams():
    if sys.stdout is None or sys.stderr is None:
        try:
            _ld=os.path.join(ROOT_FOLDER,"_logs"); os.makedirs(_ld,exist_ok=True)
            _fh=open(os.path.join(_ld,"trade_review_app.log"),"a",encoding="utf-8",errors="backslashreplace")
            if sys.stdout is None: sys.stdout=_fh
            if sys.stderr is None: sys.stderr=_fh
        except Exception:
            import io
            if sys.stdout is None: sys.stdout=io.StringIO()
            if sys.stderr is None: sys.stderr=io.StringIO()
_pythonw_safe_streams()
try: logging.getLogger("werkzeug").setLevel(logging.WARNING)
except Exception: pass

# Local history file (user-provided minute bars, EDT timestamps).
# Format: Excel with columns Date, Open, High, Low, Close, Volume
# Date is a datetime like "2025/4/29 9:30:00" (EDT). Takes priority over Bloomberg.
HISTORY_FILE = _os.path.join(ROOT_FOLDER, "history_minute.xlsx")

# Cached per-date history: {date_str: [bars]} and mtime for invalidation
_history_cache = {"mtime": None, "by_date": None}

def _load_history_by_date():
    """Load history_minute.xlsx and group by date.
    Returns dict {YYYY-MM-DD: [{t,o,h,l,c,v}, ...]} with HH:MM time strings (EDT).
    Uses vectorized pandas ops for speed. Reloads if file mtime changes."""
    if not os.path.exists(HISTORY_FILE):
        return {}
    try:
        mtime = os.path.getmtime(HISTORY_FILE)
    except:
        return {}
    if _history_cache["mtime"] == mtime and _history_cache["by_date"] is not None:
        return _history_cache["by_date"]
    print(f"[history] loading {HISTORY_FILE} ...")
    try:
        df = pd.read_excel(HISTORY_FILE)
        df.columns = [str(c).strip() for c in df.columns]
        # Find columns by flexible matching
        col_map = {}
        for want in ("Date", "Open", "High", "Low", "Close", "Volume"):
            for c in df.columns:
                if c.lower() == want.lower():
                    col_map[want] = c; break
        if "Date" not in col_map:
            print(f"[history] no Date column in {HISTORY_FILE}")
            return {}
        dc = col_map["Date"]
        df[dc] = pd.to_datetime(df[dc], errors="coerce")
        df = df.dropna(subset=[dc])
        # Vectorized: extract date string and time string
        df["_date"] = df[dc].dt.strftime("%Y-%m-%d")
        df["_time"] = df[dc].dt.strftime("%H:%M")
        # Coerce numeric columns
        for want in ("Open", "High", "Low", "Close", "Volume"):
            if want in col_map:
                df[col_map[want]] = pd.to_numeric(df[col_map[want]], errors="coerce").fillna(0)
        # Group by date, build bars
        by_date = {}
        oc = col_map.get("Open", "")
        hc = col_map.get("High", "")
        lc = col_map.get("Low", "")
        cc = col_map.get("Close", "")
        vc = col_map.get("Volume", "")
        for date_str, grp in df.groupby("_date"):
            bars = []
            for _, r in grp.iterrows():
                bars.append({
                    "t": r["_time"],
                    "o": float(r[oc]) if oc else 0.0,
                    "h": float(r[hc]) if hc else 0.0,
                    "l": float(r[lc]) if lc else 0.0,
                    "c": float(r[cc]) if cc else 0.0,
                    "v": int(r[vc]) if vc else 0,
                })
            bars.sort(key=lambda b: b["t"])
            by_date[date_str] = bars
        _history_cache["mtime"] = mtime
        _history_cache["by_date"] = by_date
        print(f"[history] loaded {len(by_date)} trading days ({len(df)} bars) from {HISTORY_FILE}")
        return by_date
    except Exception as e:
        import traceback
        print(f"[history] load failed: {e}")
        traceback.print_exc()
        return {}

_cache={}
_negative_cache=set()  # dates confirmed to have no data (weekend/holiday/future)

def fetch_intraday(ticker,date_str):
    k=f"{ticker}|{date_str}"
    if k in _cache:return _cache[k]
    if date_str in _negative_cache:return []
    # Quick skip: weekends and future dates
    try:
        d = dt.datetime.strptime(date_str, "%Y-%m-%d")
        if d.weekday() >= 5:  # Saturday=5, Sunday=6
            _negative_cache.add(date_str)
            return []
        if d.date() > dt.datetime.now().date():
            _negative_cache.add(date_str)
            return []
    except:
        return []
    # Extract base ticker (handle "SPY US Equity" → "SPY")
    base_ticker = ticker.split()[0] if " " in ticker else ticker
    # Try local history file first (user-provided, takes priority)
    hist = _load_history_by_date()
    if date_str in hist and hist[date_str]:
        bars = hist[date_str]
        _cache[k] = bars
        return bars
    # Fallback: yfinance (for dates not yet in history file)
    # Skip weekends, future dates, and already-known misses
    try:
        d_obj = dt.datetime.strptime(date_str, "%Y-%m-%d")
        if d_obj.weekday() >= 5:  # Saturday=5, Sunday=6
            _cache[k] = []
            return []
        if d_obj.date() > dt.datetime.now().date():
            _cache[k] = []
            return []
    except:
        pass
    bars = _fetch_yfinance_minute(base_ticker, date_str)
    if bars:
        _append_to_history(base_ticker, date_str, bars)
        _cache[k] = bars
        return bars
    # Cache the miss so we don't retry yfinance for this date
    _negative_cache.add(date_str)
    return []


def _fetch_yfinance_minute(ticker, date_str):
    """Fetch 1-minute bars for a single date from yfinance. Returns [{t,o,h,l,c,v}, ...]
    yfinance's 1m interval only goes back ~30 days."""
    try:
        import yfinance as yf
    except ImportError:
        print("[yfinance] yfinance not installed; run: pip install yfinance")
        return []
    try:
        d = dt.datetime.strptime(date_str, "%Y-%m-%d")
    except:
        return []
    start = d
    end = d + dt.timedelta(days=1)
    try:
        df = yf.download(ticker, start=start, end=end, interval="1m",
                         progress=False, auto_adjust=False, prepost=False)
    except Exception as e:
        print(f"[yfinance] fetch failed for {date_str}: {e}")
        return []
    if df is None or df.empty:
        return []
    # yfinance returns tz-aware index (typically US/Eastern for US stocks, or UTC)
    bars = []
    # Flatten multi-level columns if present (yfinance sometimes returns (Open, SPY) tuples)
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [c[0] if isinstance(c, tuple) else c for c in df.columns]
    try:
        idx = df.index
        if idx.tz is None:
            idx = idx.tz_localize("UTC")
        idx = idx.tz_convert("US/Eastern")
    except Exception as e:
        print(f"[yfinance] tz convert failed: {e}")
        return []
    for i in range(len(df)):
        t = idx[i]
        if t.strftime("%Y-%m-%d") != date_str:
            continue  # safety: skip bars not on target date
        try:
            bars.append({
                "t": t.strftime("%H:%M"),
                "o": float(df.iloc[i]["Open"]),
                "h": float(df.iloc[i]["High"]),
                "l": float(df.iloc[i]["Low"]),
                "c": float(df.iloc[i]["Close"]),
                "v": int(df.iloc[i]["Volume"] or 0),
            })
        except:
            continue
    return bars


def _atomic_to_excel(df, path):
    """Write df to path atomically: full temp file first, then os.replace().

    to_excel() writes in place, so a second process (or a crash) mid-write leaves a
    truncated zip — the file is then unrecoverable, not merely stale. That is exactly
    how history_minute.xlsx was destroyed on 2026-08-10: two app instances appended
    the same 4.4 MB workbook concurrently and shredded xl/worksheets/sheet1.xml.
    os.replace() is atomic on the same volume, so readers only ever see a complete file.
    """
    # 副檔名必須留 .xlsx —— pandas 由副檔名決定 writer engine
    tmp = f"{path}.{os.getpid()}.tmp.xlsx"
    try:
        df.to_excel(tmp, index=False)
        os.replace(tmp, path)
    finally:
        try:
            if os.path.exists(tmp): os.remove(tmp)
        except OSError:
            pass


def _file_lock(path, stale_seconds=300):
    """Best-effort cross-process lock via O_EXCL lock file. Returns fd or None.

    Guards whole-file xlsx rewrites against a second app instance (scheduled task +
    a manually started copy is a normal situation here, and both write these files).
    """
    lock = path + ".lock"
    try:
        age = time.time() - os.path.getmtime(lock)
        if age > stale_seconds:
            os.remove(lock)
    except OSError:
        pass
    try:
        return os.open(lock, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
    except OSError:
        return None


def _file_unlock(path, fd):
    try:
        os.close(fd)
    except OSError:
        pass
    try:
        os.remove(path + ".lock")
    except OSError:
        pass


def _append_to_history(ticker, date_str, bars):
    """Append new bars to history_minute.xlsx. Creates file if missing.
    Skips if date_str already exists in file (avoid duplicates)."""
    if not bars: return
    fd = _file_lock(HISTORY_FILE)
    if fd is None:
        print(f"[history] append skipped for {date_str}: another process is writing")
        return
    try:
        # Build rows for this date
        new_rows = []
        for b in bars:
            hh, mm = b["t"].split(":")
            d = dt.datetime.strptime(date_str, "%Y-%m-%d").replace(hour=int(hh), minute=int(mm))
            new_rows.append({"Date": d, "Open": b["o"], "High": b["h"], "Low": b["l"],
                             "Close": b["c"], "Volume": b["v"]})
        new_df = pd.DataFrame(new_rows)
        if os.path.exists(HISTORY_FILE):
            old_df = pd.read_excel(HISTORY_FILE)
            old_df.columns = [str(c).strip() for c in old_df.columns]
            if "Date" in old_df.columns:
                old_df["Date"] = pd.to_datetime(old_df["Date"])
                # Check if date_str already present
                existing_dates = set(old_df["Date"].dt.strftime("%Y-%m-%d").unique())
                if date_str in existing_dates:
                    return  # already have it
            combined = pd.concat([old_df, new_df], ignore_index=True)
            combined = combined.sort_values("Date").reset_index(drop=True)
        else:
            combined = new_df
            os.makedirs(os.path.dirname(HISTORY_FILE), exist_ok=True)
        _atomic_to_excel(combined, HISTORY_FILE)
        # Invalidate cache so next read picks up new data
        _history_cache["mtime"] = None
        print(f"[history] appended {len(new_rows)} bars for {date_str} → {HISTORY_FILE}")
    except PermissionError:
        print(f"[history] write FAILED: {HISTORY_FILE} is open in Excel")
    except Exception as e:
        print(f"[history] append failed for {date_str}: {e}")
    finally:
        _file_unlock(HISTORY_FILE, fd)

# Cache the loaded trades DataFrame; reload if file mtime changes
_trades_cache={"mtime":None,"df":None}

def load_trades_df():
    """Load the consolidated trades_all.xlsx, with cache invalidation on file change."""
    if not os.path.exists(TRADES_FILE):
        return None
    mtime=os.path.getmtime(TRADES_FILE)
    if _trades_cache["mtime"]==mtime and _trades_cache["df"] is not None:
        return _trades_cache["df"]
    try:
        df=pd.read_excel(TRADES_FILE,dtype=str)
        df.columns=[c.strip() for c in df.columns]
        _trades_cache["mtime"]=mtime
        _trades_cache["df"]=df
        return df
    except Exception as e:
        print(f"[load_trades_df] failed: {e}")
        return None

def _norm_date(v):
    """Normalize a date cell value to YYYY-MM-DD string."""
    if v is None or (isinstance(v,float) and pd.isna(v)):return ""
    s=str(v).strip()
    for fmt in ("%Y-%m-%d","%Y/%m/%d","%m/%d/%Y","%Y%m%d"):
        try:return dt.datetime.strptime(s,fmt).strftime("%Y-%m-%d")
        except:pass
    try:return pd.to_datetime(s).strftime("%Y-%m-%d")
    except:return s

def _load_dividends():
    """Return dict {YYYY-MM-DD: per_share_float} from dividends.xlsx."""
    if not os.path.exists(DIVIDENDS_FILE): return {}
    try:
        df=pd.read_excel(DIVIDENDS_FILE,dtype=str)
        df.columns=[c.strip() for c in df.columns]
        d={}
        for _,r in df.iterrows():
            nd=_norm_date(r.get("Date",""))
            try:ps=float(str(r.get("Dividend_Per_Share","")).strip())
            except:continue
            if nd:d[nd]=ps
        return d
    except Exception as e:
        print(f"[load_dividends] {e}")
        return {}

def _load_colors():
    """Return list of color versions from colors.xlsx, up to 4."""
    if not os.path.exists(COLORS_FILE):
        # Fallback defaults
        return [{"Version":"1","App_Long":"#5BA0FF","App_Short":"#B0B0B0","App_Hold":"#9B59B6",
                 "Export_Long":"#0040C0","Export_Short":"#000000","Export_Hold":"#B07CD8"}]
    try:
        df=pd.read_excel(COLORS_FILE,dtype=str)
        df.columns=[c.strip() for c in df.columns]
        out=[]
        all_keys=["Version","App_Long","App_Short","App_Hold","Export_Long","Export_Short","Export_Hold",
                   "App_Long_Border","App_Short_Border","App_Hold_Border",
                   "Export_Long_Border","Export_Short_Border","Export_Hold_Border"]
        for _,r in df.iterrows():
            out.append({k:str(r.get(k,"")).strip() for k in all_keys})
            if len(out)>=4:break
        return out if out else [{"Version":"1","App_Long":"#5BA0FF","App_Short":"#B0B0B0","App_Hold":"#9B59B6","Export_Long":"#0040C0","Export_Short":"#000000","Export_Hold":"#B07CD8"}]
    except Exception as e:
        print(f"[load_colors] {e}")
        return [{"Version":"1","App_Long":"#5BA0FF","App_Short":"#B0B0B0","App_Hold":"#9B59B6","Export_Long":"#0040C0","Export_Short":"#000000","Export_Hold":"#B07CD8"}]

_analysis_cache={"mtime":None,"by_date":None,"full_df":None}
def _merge_split_rows(df, pnl_col):
    """Merge consecutive rows that sum to 100 shares (a single executed trade).
    Rule: walk rows in order, accumulate into a buffer; when sum >= 100 shares,
    consider the buffer one completed trade and merge it into a single row.
    Price = weighted average by shares; PnL = whichever row has it (if any);
    Type = whichever row has it (多/空). If multiple rows in buffer have Type,
    they must agree (use the first non-empty).
    """
    if len(df) == 0: return df
    df = df.reset_index(drop=True).copy()
    df["_Date_norm"] = df["Date"].apply(_norm_date)
    df["_Time_norm"] = df["Exec Time(EDT)"].astype(str).str.strip()

    merged_rows = []
    buffer = []
    buf_shares = 0
    buf_date = None

    def _flush_buffer():
        """Merge the buffer rows into a single row dict; returns the merged row."""
        nonlocal buffer, buf_shares
        if not buffer:
            return None
        total_sh = 0
        wsum_px = 0.0
        pnl_val = None
        type_val = None
        for rr in buffer:
            try: sh = int(float(str(rr.get("Shares", "0"))))
            except: sh = 0
            try: px = float(str(rr.get("Price", "0")))
            except: px = 0.0
            total_sh += sh
            wsum_px += sh * px
            typ_raw = str(rr.get("Type", "")).strip()
            if typ_raw in ("多", "空", "平") and type_val is None:
                type_val = typ_raw
            pnl_raw = str(rr.get(pnl_col, "")).strip() if pnl_col else ""
            if pnl_raw and pnl_raw.lower() != "nan" and pnl_val is None:
                pnl_val = pnl_raw
        avg_px = wsum_px / total_sh if total_sh > 0 else 0.0
        # Start from LAST row so Pair_* / Action / Status / Symbol etc are inherited
        merged = dict(buffer[-1])
        merged["Shares"] = total_sh
        merged["Price"] = round(avg_px, 4)
        # Type: use first non-empty in buffer (or keep empty if none)
        if type_val is not None:
            merged["Type"] = type_val
        else:
            merged["Type"] = ""
        # PnL: use first non-empty
        if pnl_col:
            merged[pnl_col] = pnl_val if pnl_val is not None else ""
        buffer = []
        buf_shares = 0
        return merged

    for i in range(len(df)):
        r = df.iloc[i].to_dict()
        try: sh = int(float(str(r.get("Shares", "0"))))
        except: sh = 0
        rd = r.get("_Date_norm", "")
        # Day boundary: flush whatever's in buffer (shouldn't happen normally but safety)
        if buf_date is not None and rd != buf_date and buffer:
            merged_rows.append(_flush_buffer())
            buf_date = None
        buffer.append(r)
        buf_shares += sh
        buf_date = rd
        # If buffer hits 100 shares (or more, which shouldn't happen but handle)
        if buf_shares >= 100:
            merged_rows.append(_flush_buffer())
            buf_date = None

    # Flush any remainder (should be empty if data is consistent)
    if buffer:
        merged_rows.append(_flush_buffer())

    out = pd.DataFrame(merged_rows)
    for c in ("_Date_norm", "_Time_norm"):
        if c in out.columns: out = out.drop(columns=[c])
    # Ensure column order
    base_cols = ["Date", "Exec Time(EDT)", "Symbol", "Price", "Type",
                 pnl_col if pnl_col else "損益(AI辨識)",
                 "Shares", "Action", "Status", "Pair_Date", "Pair_Time"]
    keep_cols = [c for c in base_cols if c in out.columns]
    extras = [c for c in out.columns if c not in keep_cols]
    out = out[keep_cols + extras]
    return out


def _analyse_all_trades():
    """Read trades_all.xlsx, merge split rows into 100-share trades, run per-day LIFO.
    Writes merged DataFrame back to Excel with Action/Status/Pair_Date/Pair_Time filled.
    If Excel already has Action filled (B/S), that value is trusted and NOT overwritten.
    Returns dict: {date: [trades...]}"""
    if not os.path.exists(TRADES_FILE): return {}
    mtime = os.path.getmtime(TRADES_FILE)
    if _analysis_cache["mtime"] == mtime and _analysis_cache["by_date"] is not None:
        return _analysis_cache["by_date"]
    try: df = pd.read_excel(TRADES_FILE, dtype=str)
    except Exception as e:
        print(f"[analyse_all] read failed: {e}"); return {}
    df.columns = [c.strip() for c in df.columns]
    for col in ("Action", "Status", "Pair_Date", "Pair_Time"):
        if col not in df.columns: df[col] = ""
        # Force column to string dtype (pandas may have inferred float64 if column was all NaN)
        df[col] = df[col].astype(object).fillna("").astype(str).replace("nan", "")
    if "Date" not in df.columns:
        print("[analyse_all] missing Date column"); return {}
    # Self-heal Date column: pd.read_excel(dtype=str) turns a genuine Excel date cell
    # (e.g. from an upstream write that stored a datetime object instead of a string —
    # forbidden by CLAUDE.md) into a suffixed 'YYYY-MM-DD 00:00:00' string. Without this,
    # the write-back below would freeze that suffix in permanently as literal text.
    # _norm_date() is idempotent on already-clean 'YYYY-MM-DD' strings.
    _date_raw = df["Date"].tolist()
    df["Date"] = df["Date"].apply(_norm_date)
    _date_col_dirty = df["Date"].tolist() != _date_raw

    pc = [c for c in df.columns if "損益" in c or "PnL" in c.lower()]
    pnl_col = pc[0] if pc else None

    # Capture user-filled Action per (date,time,shares,price) key — merge may change rows,
    # so we preserve any row where user already wrote an Action.
    user_actions = {}  # {(date, time, shares_int): "B" or "S"}
    for _, r in df.iterrows():
        a = str(r.get("Action", "")).strip().upper()
        if a in ("B", "S"):
            nd = _norm_date(r.get("Date", ""))
            tm = str(r.get("Exec Time(EDT)", "")).strip()
            try: sh = int(float(str(r.get("Shares", "0"))))
            except: sh = 0
            user_actions[(nd, tm, sh)] = a

    # Step 1: merge split rows into 100-share single trades
    _orig_len = len(df)
    df = _merge_split_rows(df, pnl_col)
    df = df.reset_index(drop=True)
    _rows_merged = len(df) != _orig_len

    divs = _load_dividends()

    # Step 2: build sorted row list with type info
    rows = []
    for idx, r in df.iterrows():
        nd = _norm_date(r.get("Date", ""))
        if not nd: continue
        tm = str(r.get("Exec Time(EDT)", "")).strip()
        try: pr = float(r.get("Price", ""))
        except: continue
        try: sh = int(float(r.get("Shares", "0")))
        except: sh = 0
        typ = str(r.get("Type", "")).strip()
        pnl_raw = str(r.get(pnl_col, "")).strip() if pnl_col else ""
        # Check if user pre-filled Action (from original, preserved across merge)
        user_act = user_actions.get((nd, tm, sh))
        rows.append({"idx": idx, "date": nd, "time": tm, "price": pr, "shares": sh,
                     "type": typ, "pnl_raw": pnl_raw, "user_action": user_act})
    rows.sort(key=lambda r: (r["date"], r["time"], r["idx"]))

    # Step 3: classify each row and run per-day LIFO pairing.
    #   Type=多  → intraday long exit (Action=S)
    #   Type=空  → intraday short exit (Action=B)
    #   Type=平  → cross-day hold closure (matches against overnight holds LIFO)
    #   Type empty, no user Action → entry (direction inferred later)
    #   Type empty, user Action=S  → sell: if holds have B → close hold; else new short entry
    #   Type empty, user Action=B  → buy:  if holds have S → close hold; else new long entry
    holds = []           # overnight positions: [{idx, date, time, price, shares_rem, user_action}]
    today_stack = []     # today's open entries
    current_day = None
    trades_out = []
    trade_id_seq = 0
    row_status = {}      # {df_idx: (action, status, pair_date, pair_time, type_tag)}
    pnl_writeback = {}   # {df_idx: pnl_value} — auto-calculated PnL for Type=平 rows

    def close_day():
        nonlocal today_stack, trade_id_seq
        for e in today_stack:
            holds.append(e)
            act = e.get("user_action") or ""
            row_status[e["idx"]] = (act, "1", "", "", "")
            trade_id_seq += 1
            trades_out.append({
                "id": trade_id_seq,
                "entryDate": e["date"], "entryTime": e["time"][:5], "entryPrice": e["price"],
                "exitDate": None, "exitTime": None, "exitPrice": None,
                "dir": "", "shares": e["shares_rem"],
                "pnl": 0, "div_income": 0, "dividends": [],
                "crossDay": False, "isHold": True,
            })
        today_stack = []

    def do_lifo_match(r, exit_action, exit_type, force_holds_only=False):
        """LIFO match an exit row against today_stack (unless force_holds_only) then holds.
        Returns list of matched pieces."""
        nonlocal trade_id_seq
        need = r["shares"]
        matched = []

        # Phase 1: consume today_stack LIFO (skip if force_holds_only)
        if not force_holds_only:
            while need > 0 and today_stack:
                top = today_stack[-1]
                take = min(need, top["shares_rem"])
                matched.append({"entry_idx": top["idx"], "entry_date": top["date"],
                                "entry_time": top["time"], "entry_price": top["price"],
                                "shares": take, "from_hold": False,
                                "entry_user_action": top.get("user_action")})
                top["shares_rem"] -= take
                need -= take
                if top["shares_rem"] <= 0: today_stack.pop()

        # Phase 2: consume overnight holds LIFO
        while need > 0 and holds:
            top = holds[-1]
            take = min(need, top["shares_rem"])
            matched.append({"entry_idx": top["idx"], "entry_date": top["date"],
                            "entry_time": top["time"], "entry_price": top["price"],
                            "shares": take, "from_hold": True,
                            "entry_user_action": top.get("user_action")})
            top["shares_rem"] -= take
            need -= take
            if top["shares_rem"] <= 0: holds.pop()

        # PnL: use from pnl_raw if available; else auto-calculate for cross-day
        try: total_pnl = float(r["pnl_raw"].replace("+", "").replace(",", ""))
        except: total_pnl = 0.0

        any_cross = any(p["from_hold"] for p in matched)
        # Auto-calc PnL for Type=平 or cross-day with no PnL filled
        if (exit_type == "平" or any_cross) and total_pnl == 0.0 and matched:
            # Calculate: for long (entry=B, exit=S): (exit - entry) * shares
            # For short (entry=S, exit=B): (entry - exit) * shares
            # Determine direction from first matched entry's user_action or from exit_action
            for p in matched:
                ea = p.get("entry_user_action", "")
                if ea == "B":
                    p["_auto_pnl"] = round((r["price"] - p["entry_price"]) * p["shares"], 2)
                elif ea == "S":
                    p["_auto_pnl"] = round((p["entry_price"] - r["price"]) * p["shares"], 2)
                else:
                    # Default: assume long (buy entry, sell exit)
                    p["_auto_pnl"] = round((r["price"] - p["entry_price"]) * p["shares"], 2)
            total_pnl = sum(p.get("_auto_pnl", 0) for p in matched)
            # Write back PnL to the exit row in Excel
            pnl_writeback[r["idx"]] = round(total_pnl, 2)

        exit_status = "2" if any_cross else "0"
        pair_d = matched[0]["entry_date"] if matched else ""
        pair_t = matched[0]["entry_time"] if matched else ""
        row_status[r["idx"]] = (exit_action, exit_status, pair_d, pair_t, exit_type)

        # Infer entry direction
        if exit_type == "多":
            entry_action = "B"
        elif exit_type == "空":
            entry_action = "S"
        elif exit_type == "平":
            # For 平: entry direction comes from the entry's own user_action (B or S)
            entry_action = ""  # will be set per-piece below
        else:
            entry_action = ""

        for p in matched:
            if exit_type == "平":
                ea = p["entry_user_action"] if p["entry_user_action"] in ("B", "S") else "B"
            else:
                ea = p["entry_user_action"] if p["entry_user_action"] in ("B", "S") else entry_action
            st = "2" if p["from_hold"] else "0"
            row_status[p["entry_idx"]] = (ea, st, r["date"], r["time"], exit_type)

            trade_id_seq += 1
            divs_in = []
            if p["from_hold"]:
                for d_date, d_ps in divs.items():
                    if p["entry_date"] < d_date <= r["date"]:
                        divs_in.append({"date": d_date, "ps": d_ps})
                divs_in.sort(key=lambda x: x["date"])
            piece_pnl = p.get("_auto_pnl", total_pnl * (p["shares"] / r["shares"]) if r["shares"] > 0 else 0)
            div_income = sum(d["ps"] * 0.7 for d in divs_in) * p["shares"]
            trades_out.append({
                "id": trade_id_seq,
                "entryDate": p["entry_date"], "entryTime": p["entry_time"][:5],
                "entryPrice": p["entry_price"],
                "exitDate": r["date"], "exitTime": r["time"][:5], "exitPrice": r["price"],
                "dir": exit_type,
                "shares": p["shares"],
                "pnl": round(piece_pnl, 2),
                "div_income": round(div_income, 2),
                "dividends": divs_in,
                "crossDay": p["from_hold"],
            })

    for r in rows:
        if current_day is None: current_day = r["date"]
        elif r["date"] != current_day:
            close_day()
            current_day = r["date"]

        typ = r["type"]
        ua = r.get("user_action")

        if typ == "平":
            # Type=平 → cross-day hold closure. Match ONLY against holds (skip today_stack).
            exit_action = ua if ua in ("B", "S") else "S"  # default S (selling to close long hold)
            do_lifo_match(r, exit_action, "平", force_holds_only=True)

        elif typ in ("多", "空"):
            # Intraday exit: Type=多 → sell to close long; Type=空 → buy to close short
            exit_action = "S" if typ == "多" else "B"
            if ua in ("B", "S"): exit_action = ua
            do_lifo_match(r, exit_action, typ)

        else:
            # No Type (多/空/平) — always an entry row
            today_stack.append({"idx": r["idx"], "date": r["date"], "time": r["time"],
                                "price": r["price"], "shares_rem": r["shares"],
                                "user_action": ua})
            row_status[r["idx"]] = (ua or "", "", "", "", "")

    close_day()

    # Step 4: Apply row_status to df + write auto-calculated PnL for 平 rows.
    # Track if anything actually changed to avoid unnecessary rewrites
    _any_changed = False
    for idx, (act, st, pd_, pt_, typ) in row_status.items():
        if idx in df.index:
            existing_action = str(df.at[idx, "Action"]).strip().upper()
            if existing_action not in ("B", "S"):
                if str(df.at[idx, "Action"]) != act: _any_changed = True
                df.at[idx, "Action"] = act
            if str(df.at[idx, "Status"]) != st: _any_changed = True
            df.at[idx, "Status"] = st
            if str(df.at[idx, "Pair_Date"]) != pd_: _any_changed = True
            df.at[idx, "Pair_Date"] = pd_
            if str(df.at[idx, "Pair_Time"]) != pt_: _any_changed = True
            df.at[idx, "Pair_Time"] = pt_
    # Write auto-calculated PnL for cross-day closures
    for idx, pnl_val in pnl_writeback.items():
        if idx in df.index and pnl_col:
            existing_pnl = str(df.at[idx, pnl_col]).strip()
            if not existing_pnl or existing_pnl.lower() == "nan":
                sign = "+" if pnl_val >= 0 else ""
                df.at[idx, pnl_col] = f"{sign}{pnl_val}"
                _any_changed = True
            existing_type = str(df.at[idx, "Type"]).strip()
            if not existing_type or existing_type.lower() == "nan":
                df.at[idx, "Type"] = "平"
                _any_changed = True

    # Only write back if something changed (avoid unnecessary IO + mtime churn)
    # ...and never while spy_daytrade_writer holds the lock: openpyxl's save is not
    # atomic, so a concurrent write-back would clobber it (or be clobbered). Skipping
    # is safe — the writer's save bumps mtime, which re-triggers this analysis and the
    # write-back then happens with the lock released.
    if _writer_active() and (_any_changed or _rows_merged or _date_col_dirty):
        print("[analyse_all] write-back skipped: .writer.lock held by writer")
    elif _any_changed or _rows_merged or _date_col_dirty:
        fd = _file_lock(TRADES_FILE)
        if fd is None:
            print("[analyse_all] write-back skipped: another process is writing trades_all")
        else:
            try:
                for col in ("Action", "Status", "Pair_Date", "Pair_Time"):
                    df[col] = df[col].astype(object).fillna("")
                _atomic_to_excel(df, TRADES_FILE)
                print(f"[analyse_all] wrote merged DataFrame ({len(df)} rows) to {TRADES_FILE}")
            except PermissionError:
                print(f"[analyse_all] write-back FAILED: trades_all.xlsx is open in Excel — close it and refresh")
            except Exception as e:
                print(f"[analyse_all] write-back skipped: {type(e).__name__}: {e}")
            finally:
                _file_unlock(TRADES_FILE, fd)

    by_date = {}
    for t in trades_out:
        d = t["entryDate"] if t.get("isHold") else t["exitDate"]
        by_date.setdefault(d, []).append(t)

    # Cache using POST-write mtime (the write itself changes the file mtime)
    try: new_mtime = os.path.getmtime(TRADES_FILE)
    except: new_mtime = mtime
    _analysis_cache["mtime"] = new_mtime
    _analysis_cache["by_date"] = by_date
    return by_date


WRITER_LOCK = _os.path.join(ROOT_FOLDER, ".writer.lock")
_STALE_LOCK_SECONDS = 600   # writer 正常執行是秒級；逾時代表被強制中斷後殘留

def _writer_active():
    """spy_daytrade_writer 是否正持有寫入鎖。

    殘留鎖必須逾時失效，否則 writer 一次被強制中斷就會讓常駐 app 永遠不再回寫
    配對結果，而且完全沒有徵兆。
    """
    try:
        age = time.time() - os.path.getmtime(WRITER_LOCK)
    except OSError:
        return False
    if age > _STALE_LOCK_SECONDS:
        print(f"[analyse_all] ignoring stale .writer.lock ({age:.0f}s old)")
        return False
    return True


def read_and_pair_for_date(date_str):
    """Return trades for a given date using the cross-day LIFO analysis."""
    by_date = _analyse_all_trades()
    return by_date.get(date_str, [])

app=Flask(__name__)

# Preload history file at startup so first request doesn't wait
_load_history_by_date()

@app.after_request
def _add_no_cache(resp):
    # Force browsers to revalidate every request so external file edits show up
    # on refresh (previously JSON/HTML got heuristically cached → had to restart app).
    resp.headers["Cache-Control"]="no-store, no-cache, must-revalidate, max-age=0"
    resp.headers["Pragma"]="no-cache"
    resp.headers["Expires"]="0"
    return resp

SYNC_STATE_FILE=_os.path.join(ROOT_FOLDER,"sync_state.json")

@app.route("/api/version")
def api_version():
    """Return data-file mtimes so the client can auto-detect edits and reload.

    Also returns sync freshness: Boss PC writes sync_state.json after each run and
    it travels with the git sync. Without this, a failed push looks exactly like
    "market hasn't opened yet" — the UI must be able to tell those apart (CLAUDE §0).
    """
    def _mt(p):
        try:return os.path.getmtime(p)
        except:return 0
    sync=None
    try:
        with open(SYNC_STATE_FILE,encoding="utf-8") as f:sync=json.load(f)
    except Exception:pass
    return jsonify({"trades":_mt(TRADES_FILE),"sync":sync})

@app.route("/")
def index():return HTML

@app.route("/api/colors")
def api_colors():
    return jsonify(_load_colors())

@app.route("/api/dates")
def api_dates():
    """Return all selectable dates, and which of them have no K-line data.

    A handful of dates exist in trades_all but not in history_minute (e.g. 2026-06-12,
    2026-06-15 — too old for yfinance to backfill). They are still listed, because the
    trades are worth seeing, but the client marks them so picking one is never a
    surprise: jumpTo() silently walks backwards to the nearest day with candles.
    """
    tdates=set()
    df=load_trades_df()
    if df is not None and "Date" in df.columns:
        for v in df["Date"].dropna().unique():
            nd=_norm_date(v)
            if nd:tdates.add(nd)
    hist=_load_history_by_date()
    hdates=set(hist.keys())
    return jsonify({"dates":sorted(tdates|hdates,reverse=True),
                    "no_candles":sorted(tdates-hdates,reverse=True)})

@app.route("/api/data")
def api_data():
    ds=request.args.get("date","")
    if not ds:return jsonify({"error":"missing date"}),400
    notes_path=os.path.join(NOTES_FOLDER,f"{ds}.txt")
    notes=""
    if os.path.exists(notes_path):
        with open(notes_path,"r",encoding="utf-8") as f:notes=f.read()
    try:c=fetch_intraday(TICKER,ds)
    except:return jsonify({"date":ds,"candles":[],"trades":[],"notes":notes,"noTrading":True})
    if not c:return jsonify({"date":ds,"candles":[],"trades":[],"notes":notes,"noTrading":True})
    try:t=read_and_pair_for_date(ds)
    except Exception as e:
        print(f"[api_data] read_and_pair_for_date error: {e}")
        t=[]
    return jsonify({"date":ds,"candles":c,"trades":t,"notes":notes,"noTrades":len(t)==0})

@app.route("/api/notes",methods=["POST"])
def api_notes():
    d=request.get_json()
    ds=d.get("date","")
    if not ds:return jsonify({"error":"missing date"}),400
    os.makedirs(NOTES_FOLDER,exist_ok=True)
    with open(os.path.join(NOTES_FOLDER,f"{ds}.txt"),"w",encoding="utf-8") as f:
        f.write(d.get("notes",""))
    return jsonify({"ok":True})

HTML=r"""<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="UTF-8">
<title>Trade Review</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
/* flex column：不再用 calc(100vh - 56px - 105px) 那種寫死高度。
   工具列一改高度、或瀏覽器字級縮放，寫死的數字就會讓畫布跟容器對不齊。 */
body{background:#000;color:#C0C4CC;font-family:'Consolas','Courier New',monospace;overflow:hidden;height:100vh;display:flex;flex-direction:column}
#hdr{height:26px;flex:0 0 auto;display:flex;align-items:center;padding:0 12px;background:#0D0F13;border-bottom:1px solid #1A1D22}
#hdr .tk{color:#FFF;font-weight:bold;font-size:13px}
#hdr .lbl{margin-left:auto;background:#FF8C00;color:#000;font-weight:bold;font-size:11px;padding:2px 10px;border-radius:2px}
/* flex-shrink:0 + nowrap — without them the P&L / trade-count labels get squeezed
   to a fraction of their text width and visually overlap each other, and the
   Export button is clipped off the right edge, below ~950px wide. */
#tb{height:38px;flex:0 0 auto;display:flex;align-items:center;gap:8px;padding:0 12px;background:#0A0C0F;border-bottom:1px solid #1A1D22;font-size:12px;overflow-x:auto;overflow-y:hidden;scrollbar-width:none}
#tb::-webkit-scrollbar{height:0}
#tb>*{flex:0 0 auto;white-space:nowrap}
/* 對比：原本 #8B8F98 on #1A1D22 約 4.4:1、字 11px 又矮，長時間看很吃力。
   提到 #D4D9E0 約 11:1，字級與點擊區一併放大。 */
#tb button{background:#20242B;color:#D4D9E0;border:1px solid #3D434E;padding:5px 13px;cursor:pointer;font-family:inherit;font-size:12px;font-weight:bold;border-radius:3px;line-height:1.2}
#tb button:hover{background:#2E343D;border-color:#5A6270;color:#FFF}
#tb button:active{background:#3A414C}
#tb button:focus-visible{outline:2px solid #FF8C00;outline-offset:1px}
#tb button:disabled{opacity:.4;cursor:default}
#tb button:disabled:hover{background:#20242B;color:#D4D9E0;border-color:#3D434E}
/* 年月日三格 —— 一般日期輸入的習慣：分段、可各自打數字、滿位自動跳下一格 */
#dbox{display:flex;align-items:center;background:#000;border:1px solid #3D434E;border-radius:3px 0 0 3px;padding:2px 4px}
#dbox.focus{border-color:#FF8C00}
.dseg{background:transparent;color:#FF8C00;border:none;outline:none;font-family:inherit;
      font-size:14px;font-weight:bold;text-align:center;padding:3px 1px;letter-spacing:.5px}
#dY{width:44px}#dM{width:26px}#dD{width:26px}
.dseg:focus{background:#2C5A8A;color:#FFF;border-radius:2px}
.dsep{color:#5A6270;font-size:14px;font-weight:bold;user-select:none}
.pnl{font-weight:bold;font-size:13px;margin-left:12px}.pnl.w{color:#FFF}.pnl.l{color:#FF4444}
.tc{color:#8B8F98;font-size:11px}
#zc{margin-left:auto;display:flex;gap:4px;align-items:center}
.ind-toggles{display:flex;gap:3px;margin-left:8px;align-items:center}
/* 原本 #666 on #2A2E35 只有 2.4:1，關閉狀態幾乎看不見；開/關也難分辨 */
/* 必須用 #tb 前綴：上面的 `#tb button` 是 ID 選擇器，特異性高於單純的
   class，否則 .on 的藍底會被它蓋掉，開/關看起來一模一樣。 */
#tb .ind-btn{background:#20242B;color:#9AA3B0;border:1px solid #3D434E;border-radius:3px;padding:4px 9px;font-size:12px;font-weight:bold;cursor:pointer;font-family:Consolas,monospace;line-height:1.2}
#tb .ind-btn:hover{background:#2E343D;color:#E6EAF0}
#tb .ind-btn.on{background:#2C5A8A;color:#FFF;border-color:#5A9BD8}
#tb .ind-btn.on:hover{background:#356BA3;color:#FFF}
#cvsel{margin-left:10px;background:#20242B;color:#D4D9E0;border:1px solid #3D434E;padding:5px 7px;font-size:12px;border-radius:3px;cursor:pointer;font-family:inherit}
#cvsel:hover{border-color:#5A6270}
#legend{display:inline-flex;gap:10px;margin-left:10px;font-size:12px;align-items:center}
#legend .sw{display:inline-block;width:12px;height:12px;margin-right:4px;vertical-align:middle;border:1px solid #555}
#zc button{font-size:14px;padding:5px 12px;min-width:34px}
#cc{position:relative;width:100%;flex:1 1 auto;min-height:0}
canvas{display:block;width:100%;height:100%}
#tb2{height:105px;flex:0 0 auto;background:#0A0C0F;border-top:1px solid #1A1D22;display:flex;align-items:stretch;overflow-x:auto;padding:6px 12px;gap:6px}
.tc2{flex:0 0 auto;min-width:130px;background:#0D0F13;border:1px solid #1A1D22;border-radius:4px;padding:6px 10px;font-size:11px;display:flex;flex-direction:column;justify-content:center;transition:border-color .15s,box-shadow .15s}
.tc2 .dr{color:#8B8F98;margin-bottom:3px}.tc2 .pv{font-weight:bold;font-size:14px}
.tc2 .pv.w{color:#FFF}.tc2 .pv.l{color:#FF4444}.tc2 .dt{color:#888;font-size:10px;margin-top:2px}
.tc2.hi{border-color:#FFD700!important;box-shadow:0 0 8px rgba(255,215,0,0.35)}
#nb{position:absolute;background:rgba(26,29,34,0.94);border:1px solid #3A3E48;border-radius:4px;padding:0;font-size:12px;color:#E0E0E0;line-height:1.7;min-width:180px;min-height:50px;cursor:move;z-index:20;white-space:pre-wrap;word-wrap:break-word;overflow:hidden;resize:both}
#nb .nb-hdr{display:flex;align-items:center;justify-content:space-between;padding:4px 10px 2px;border-bottom:1px solid #2A2E35;cursor:move;user-select:none}
#nb .nb-hdr span{color:#FF8C00;font-size:10px;font-weight:bold}
#nb .nb-min{background:none;border:1px solid #3A3E48;color:#8B8F98;font-size:14px;line-height:1;width:22px;height:18px;cursor:pointer;border-radius:2px;display:flex;align-items:center;justify-content:center;padding:0}
#nb .nb-min:hover{color:#FFF;border-color:#666}
#nb .nb-body{padding:8px 14px 10px;overflow:auto;max-height:calc(100% - 28px)}
#nb .eh{color:#555;font-size:9px;margin-top:6px}
#nb.mini{min-width:0;min-height:0;padding:0;cursor:move;resize:none;overflow:visible}
#nb.mini .nb-body{display:none}
#nb.mini .nb-hdr{border-bottom:none;padding:3px 8px}
#ned{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.7);z-index:100;justify-content:center;align-items:center}
#ned.show{display:flex}
#ned .pn{background:#1A1D22;border:1px solid #3A3E48;border-radius:8px;padding:20px;width:420px}
#ned .pn h3{color:#FF8C00;font-size:14px;margin-bottom:12px}
#ned textarea{width:100%;height:180px;background:#000;color:#E0E0E0;border:1px solid #3A3E48;border-radius:4px;padding:10px;font-family:inherit;font-size:12px;resize:vertical;line-height:1.6}
#ned button{background:#3D6FCC;color:#FFF;border:none;border-radius:4px;padding:6px 24px;cursor:pointer;font-family:inherit;font-size:12px;margin-top:12px}
#ld{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);color:#FF8C00;font-size:14px;z-index:50;display:none}
/* Was position:fixed bottom:6px in #333 — that sat ON TOP of the trade-card bar
   and at ~1.5:1 contrast was effectively invisible. Now it lives in the header
   strip, which has spare room, at a legible grey. */
/* 日期選擇器：年月日三格 + 迷你日曆 */
#dwrap{position:relative;display:flex;align-items:center}
#dtog{background:#20242B;color:#D4D9E0;border:1px solid #3D434E;border-left:none;
      border-radius:0 3px 3px 0;padding:5px 9px;font-size:13px;cursor:pointer;line-height:1.2}
#dtog:hover{background:#2E343D;color:#FFF}

#dcal{display:none;position:absolute;top:100%;left:0;margin-top:4px;z-index:60;
      background:#12151A;border:1px solid #3D434E;border-radius:5px;padding:8px;
      box-shadow:0 8px 24px rgba(0,0,0,.65);user-select:none}
#dcal.show{display:block}
#dcal .cal-hd{display:flex;align-items:center;justify-content:space-between;margin-bottom:6px;gap:6px}
#dcal .cal-hd .ttl{color:#E6EAF0;font-size:13px;font-weight:bold;min-width:96px;text-align:center}
#dcal .cal-nav{background:#20242B;color:#D4D9E0;border:1px solid #3D434E;border-radius:3px;
               width:26px;height:24px;font-size:13px;cursor:pointer;line-height:1;padding:0}
#dcal .cal-nav:hover{background:#2E343D;color:#FFF}
#dcal table{border-collapse:separate;border-spacing:2px}
#dcal th{color:#7A8290;font-size:11px;font-weight:normal;width:30px;padding:2px 0}
#dcal th.we{color:#5A6270}
#dcal td{width:30px;height:26px;text-align:center;font-size:12px;border-radius:3px;
         color:#3A4048;                    /* 預設＝非交易日，明顯壓暗 */
         cursor:default}
#dcal td.has{color:#D4D9E0;background:#20242B;cursor:pointer;font-weight:bold}  /* 有 K 線＝可點 */
#dcal td.has:hover{background:#2E343D;color:#FFF}
#dcal td.nok{color:#C08A3E;background:#241E14;cursor:pointer}   /* 有交易但無 K 線 */
#dcal td.nok:hover{background:#33291B;color:#E0A860}
#dcal td.cur{background:#FF8C00;color:#000}
#dcal td.today{outline:1px solid #5A6270}
#dcal .cal-ft{margin-top:6px;color:#7A8290;font-size:10px;text-align:center;line-height:1.5}
#dcal .cal-ft b{color:#D4D9E0}
#dlist .di-empty{padding:10px 12px;font-size:12px;color:#7A8290}
.hint{margin-left:20px;color:#6B7280;font-size:10px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
#syncbadge{margin-left:auto;margin-right:14px;font-size:11px;font-weight:bold;white-space:nowrap;cursor:default}
</style></head><body>
<div id="hdr"><span class="tk">SPY US Equity</span><span class="hint">拖曳平移（跨日無縫）｜ 滾輪/+- 縮放 ｜ ← → 切日期 ｜ 雙擊文字框編輯 ｜ 底部/右側邊緣拖曳可縮放軸</span><span id="syncbadge" style="display:none"></span><span class="lbl">Intraday Candle Chart</span></div>
<div id="tb">
<button id="bp" title="前一個交易日（← 鍵）">&#8592; Prev</button>
<span id="dwrap" title="可直接輸入數字；↑↓ 切換前後交易日；點日曆圖示選日期"><span id="dbox"><input class="dseg" id="dY" maxlength="4" inputmode="numeric" autocomplete="off" spellcheck="false"><span class="dsep">-</span><input class="dseg" id="dM" maxlength="2" inputmode="numeric" autocomplete="off" spellcheck="false"><span class="dsep">-</span><input class="dseg" id="dD" maxlength="2" inputmode="numeric" autocomplete="off" spellcheck="false"></span><button id="dtog" title="開啟日曆">&#128197;</button><div id="dcal"></div></span>
<button id="bn" title="後一個交易日（→ 鍵）">Next &#8594;</button>
<button id="bt" title="跳到最新美股交易日">Today</button>
<span class="pnl" id="dp">---</span><span class="tc" id="dtc"></span>
<select id="cvsel" title="色版"></select>
<span id="legend"></span>
<span class="ind-toggles"><button id="tVWAP" class="ind-btn" title="VWAP">V</button><button id="tBB" class="ind-btn" title="Bollinger Bands">BB</button><button id="tMA" class="ind-btn" title="20 MA">MA</button><button id="tPC" class="ind-btn" title="前日收盤價">C</button></span>
<div id="zc"><button id="zo" title="Zoom Out (-)">&#8722;</button><button id="zi" title="Zoom In (+)">&#43;</button><button id="zr" title="Reset View">R</button><button id="be" title="輸出 PNG">⤓ Export</button></div>
</div>
<div id="cc"><canvas id="cv"></canvas><div id="nb"></div><div id="ld">載入中 ...</div></div>
<div id="tb2"></div>
<div id="ned"><div class="pn"><h3 id="ned-title">市場概述</h3><textarea id="nt"></textarea><button id="ns">儲存</button></div></div>
<script>
const cv=document.getElementById("cv"),ctx=cv.getContext("2d"),cc=document.getElementById("cc");

// ═══════════════════════════════════════════════════
// VIRTUAL TIMELINE: each day stored independently
// ═══════════════════════════════════════════════════
const dayCache=new Map(); // date→{candles,trades,notes,noTrading}
let dayList=[];           // sorted asc: ["2026-03-30","2026-03-31","2026-04-01"]
let focusIdx=0;           // index into dayList for the "current" day
const BPD=391, GAP=0;     // bars per day, no gap between days
let panX=0,panY=0,zoom=1;
let userYZoom=1; // user's manual Y zoom multiplier (1 = default 7-grid view)
const ZS=.2,ZMIN=.25,ZMAX=6;
let mouse=null,hovT=-1;
// Notes box persistence: per-date position + size, anchored to that day's K-line X range.
// Storage format: { "2026-04-14": {x, y, w, h}, ..., _mini: bool }
//   x: fraction (0..1) within that day's X range in grid space (used to compute K-line-anchored X)
//   y: fraction (0..1) of canvas height
//   w,h: px dimensions (optional)
function _loadNotesState(){
  try{
    const s=localStorage.getItem("notesStateV2");
    if(!s)return null;
    return JSON.parse(s);
  }catch(e){return null;}
}
function _saveNotesState(){
  try{
    localStorage.setItem("notesStateV2",JSON.stringify(notesStateByDate));
  }catch(e){}
}
let notesStateByDate=_loadNotesState()||{};
// Market wrap (市場概述) starts COLLAPSED on every load. The stored _mini flag is
// still written by the toggle so it works within a session, but it is deliberately
// NOT read here — otherwise an existing localStorage value of false would keep the
// box expanded forever and the default would never apply.
let notesMini=true;
let nDrag=false,nOff={x:0,y:0};

// Helper: get or create state for a given date
function _getNoteState(date){
  if(!notesStateByDate[date]){
    notesStateByDate[date]={x:0.5,y:0.1,w:null,h:null};
  }
  return notesStateByDate[date];
}
let cDrag=false,cDS={x:0,y:0},cDP={x:0,y:0};
let yAxisDrag=false,yAxisDS=0,yAxisZoomStart=1;
let loading=false; // mutex for edge-loading

const C={bg:"#000",grid:"#454A55",gridM:"#5A6070",txt:"#8B8F98",uB:"#16A34A",dB:"#DC2626",uW:"#16A34A",dW:"#DC2626",
lo:"#000000",sh:"#FFFFFF",pW:"#FFF",pL:"#FF4444",lp:"#FFD700",bd:"#333640",cross:"#6A7080",sep:"#FF8C00"};
const fmt=n=>n.toFixed(2);
// Day-of-week abbreviation: "2026-04-16" → "Wed."
function dowStr(dateStr){
  try{const d=new Date(dateStr+"T12:00:00");const days=["Sun.","Mon.","Tue.","Wed.","Thu.","Fri.","Sat."];return days[d.getDay()]||"";}
  catch(e){return "";}
}

// ═══════════════════════════════════════════════════
// Indicator toggles: VWAP, Bollinger Bands, 20 MA
// ═══════════════════════════════════════════════════
let indState={vwap:false,bb:false,ma:false,pc:false};
try{const s=localStorage.getItem("indToggles");if(s){const o=JSON.parse(s);indState=Object.assign(indState,o);}}catch(e){}
function saveIndState(){try{localStorage.setItem("indToggles",JSON.stringify(indState));}catch(e){}}
function initIndButtons(){
  const bV=document.getElementById("tVWAP"),bBB=document.getElementById("tBB"),bMA=document.getElementById("tMA"),bPC=document.getElementById("tPC");
  function upd(){bV.className="ind-btn"+(indState.vwap?" on":"");bBB.className="ind-btn"+(indState.bb?" on":"");bMA.className="ind-btn"+(indState.ma?" on":"");bPC.className="ind-btn"+(indState.pc?" on":"");}
  upd();
  bV.addEventListener("click",()=>{indState.vwap=!indState.vwap;saveIndState();upd();draw();});
  bBB.addEventListener("click",()=>{indState.bb=!indState.bb;saveIndState();upd();draw();});
  bMA.addEventListener("click",()=>{indState.ma=!indState.ma;saveIndState();upd();draw();});
  bPC.addEventListener("click",()=>{indState.pc=!indState.pc;saveIndState();upd();draw();});
}

// Calculate VWAP for a day's candles: cumulative (typical_price * volume) / cumulative volume
function calcVWAP(candles){
  const out=[];let cumPV=0,cumV=0;
  for(const c of candles){
    const tp=(c.h+c.l+c.c)/3;
    cumPV+=tp*(c.v||0);cumV+=(c.v||0);
    out.push(cumV>0?cumPV/cumV:c.c);
  }
  return out;
}

// Calculate Simple Moving Average (period bars)
function calcSMA(candles,period){
  const out=[];
  for(let i=0;i<candles.length;i++){
    if(i<period-1){out.push(null);continue;}
    let sum=0;for(let j=i-period+1;j<=i;j++)sum+=candles[j].c;
    out.push(sum/period);
  }
  return out;
}

// Calculate Bollinger Bands (period, multiplier)
function calcBB(candles,period,mult){
  const sma=calcSMA(candles,period);
  const upper=[],lower=[];
  for(let i=0;i<candles.length;i++){
    if(sma[i]===null){upper.push(null);lower.push(null);continue;}
    let sq=0;for(let j=i-period+1;j<=i;j++){const d=candles[j].c-sma[i];sq+=d*d;}
    const sd=Math.sqrt(sq/period);
    upper.push(sma[i]+mult*sd);lower.push(sma[i]-mult*sd);
  }
  return{upper,middle:sma,lower};
}

// Draw indicator lines for a given day
function drawIndicators(ctx,di,candles,pn,px,exportMode){
  if(exportMode)return; // never draw indicators in export
  if(!indState.vwap&&!indState.bb&&!indState.ma&&!indState.pc)return;

  ctx.save();
  // Clip to chart area
  ctx.beginPath();ctx.rect(M.l,M.t,cW,cH);ctx.clip();

  // Previous day close line
  if(indState.pc){
    // Find previous day's last candle close
    let prevClose=null;
    for(let p=di-1;p>=0;p--){
      const pd=dayCache.get(dayList[p]);
      if(pd&&pd.candles&&pd.candles.length){
        prevClose=pd.candles[pd.candles.length-1].c;break;
      }
    }
    if(prevClose!==null){
      const y=yOf(prevClose,pn,px);
      ctx.strokeStyle="rgba(180,180,180,0.4)";ctx.lineWidth=1;ctx.setLineDash([6,4]);
      ctx.beginPath();ctx.moveTo(M.l,y);ctx.lineTo(W-M.r,y);ctx.stroke();
      ctx.setLineDash([]);
    }
  }

  // VWAP
  if(indState.vwap){
    const vwap=calcVWAP(candles);
    ctx.strokeStyle="rgba(255,152,0,0.7)";ctx.lineWidth=1.5;ctx.setLineDash([]);
    ctx.beginPath();
    let started=false;
    for(let i=0;i<candles.length;i++){
      const x=xOfGi(gi(di,i)),y=yOf(vwap[i],pn,px);
      if(!started){ctx.moveTo(x,y);started=true;}else ctx.lineTo(x,y);
    }
    ctx.stroke();
  }

  // 20 MA
  if(indState.ma||indState.bb){
    const sma=calcSMA(candles,20);
    if(indState.ma){
      ctx.strokeStyle="rgba(33,150,243,0.6)";ctx.lineWidth=1;ctx.setLineDash([]);
      ctx.beginPath();let started=false;
      for(let i=0;i<candles.length;i++){
        if(sma[i]===null)continue;
        const x=xOfGi(gi(di,i)),y=yOf(sma[i],pn,px);
        if(!started){ctx.moveTo(x,y);started=true;}else ctx.lineTo(x,y);
      }
      ctx.stroke();
    }
  }

  // Bollinger Bands
  if(indState.bb){
    const bb=calcBB(candles,20,2);
    ctx.setLineDash([3,3]);
    // Upper
    ctx.strokeStyle="rgba(158,158,158,0.7)";ctx.lineWidth=0.8;
    ctx.beginPath();let s1=false;
    for(let i=0;i<candles.length;i++){
      if(bb.upper[i]===null)continue;
      const x=xOfGi(gi(di,i)),y=yOf(bb.upper[i],pn,px);
      if(!s1){ctx.moveTo(x,y);s1=true;}else ctx.lineTo(x,y);
    }
    ctx.stroke();
    // Lower
    ctx.beginPath();let s2=false;
    for(let i=0;i<candles.length;i++){
      if(bb.lower[i]===null)continue;
      const x=xOfGi(gi(di,i)),y=yOf(bb.lower[i],pn,px);
      if(!s2){ctx.moveTo(x,y);s2=true;}else ctx.lineTo(x,y);
    }
    ctx.stroke();
    ctx.setLineDash([]);
  }

  ctx.restore();
}
function addDays(ds,n){const d=new Date(ds+"T12:00:00");d.setDate(d.getDate()+n);return d.toISOString().slice(0,10);}
function pStep(r){const raw=r/8,m=Math.pow(10,Math.floor(Math.log10(raw))),n=raw/m;return(n<=1?1:n<=2?2:n<=5?5:10)*m;}

// ── Coordinate system ──
let M={t:6,r:68,b:28,l:6},W,H,cW,cH;
let resizeFreeze=false;
function resize(){if(resizeFreeze)return;const dpr=devicePixelRatio||1,r=cc.getBoundingClientRect();W=r.width;H=r.height;
cv.width=W*dpr;cv.height=H*dpr;cv.style.width=W+"px";cv.style.height=H+"px";
ctx.setTransform(dpr,0,0,dpr,0,0);cW=W-M.l-M.r;cH=H-M.t-M.b;}

// Pixels per virtual unit (1 unit = 1 candle slot)
function ppb(){return cW*zoom/(BPD-1||1);}

// Global index for a candle: dayIndex * (BPD+GAP) + localIndex
function gi(dayI,localI){return dayI*(BPD+GAP)+localI;}
function xOfGi(g){return M.l+panX+g*ppb();}
function gi2x(x){return(x-M.l-panX)/ppb();}

// t2i within a specific day's candles
function t2iLocal(hm,candles){
  for(let i=0;i<candles.length;i++)if(candles[i].t===hm)return i;
  let b=0,bd=9e9;const[h,m]=hm.split(":").map(Number),tg=h*60+m;
  candles.forEach((c,i)=>{const d=Math.abs(c.t.split(":").map(Number).reduce((a,b2)=>a*60+b2)-tg);if(d<bd){bd=d;b=i;}});
  return b;
}

// ── Auto Y from focus day, forced to ~7 price grids ──
function autoY(){
  if(exportYOverride)return{pn:exportYOverride.pn,px:exportYOverride.px};
  let mn=1e9,mx=-1e9;

  // Y 範圍取自「畫面上實際看得到的 K 棒」，而不是焦點日。
  //
  // 原本是鎖定 dayList[focusIdx] 那一天的高低價當錨點。焦點日是離螢幕中心
  // 最近的那天 —— 拖曳跨日時它會在某個瞬間切換，Y 錨點跟著整個跳掉，
  // 畫面突然上下彈一下。焦點是離散的，所以跳動無法避免。
  // 改成連續量（可見範圍）之後，平移時 Y 只會平滑跟隨，不會有跳點。
  const gL=gi2x(M.l),gR=gi2x(W-M.r);
  dayList.forEach((date,di)=>{
    const d=dayCache.get(date);
    if(!d||!d.candles||!d.candles.length)return;
    const g0=gi(di,0);
    const i0=Math.max(0,Math.floor(gL-g0)),i1=Math.min(d.candles.length-1,Math.ceil(gR-g0));
    for(let i=i0;i<=i1;i++){
      const c=d.candles[i];
      if(c.l<mn)mn=c.l;
      if(c.h>mx)mx=c.h;
    }
    // 進出場點也要在範圍內，否則交易線會被切掉；留倉/跨日單價格可能離今日很遠，排除
    (d.trades||[]).forEach(t=>{
      if(t.isHold||t.crossDay)return;
      const ei=t.entryTime?t2iLocal(t.entryTime,d.candles):-1;
      if(ei>=i0&&ei<=i1&&t.entryPrice!=null){mn=Math.min(mn,t.entryPrice);mx=Math.max(mx,t.entryPrice);}
      const xi=t.exitTime?t2iLocal(t.exitTime,d.candles):-1;
      if(xi>=i0&&xi<=i1&&t.exitPrice!=null){mn=Math.min(mn,t.exitPrice);mx=Math.max(mx,t.exitPrice);}
    });
  });

  if(mn>mx){  // 畫面上沒有任何 K 棒（例如平移到空白區）→ 沿用上一次的範圍，不要亂跳
    if(_lastY)return{pn:_lastY.pn+panY-_lastY.panY,px:_lastY.px+panY-_lastY.panY};
    mn=650;mx=660;
  }

  const baseRange=7; // default visible range = $7 (7 grids of $1)
  const targetRange=baseRange*userYZoom;
  const ladder=[0.1,0.2,0.5,1,2,5,10,20,50,100,200,500];
  let step=1;
  for(const s of ladder){if(s*7>=targetRange*1.0){step=s;break;}}
  // 中心不再吸附到半格邊界 —— 那個吸附本身就是量化跳動的來源。
  // 格線仍然落在整數價位（draw() 用 Math.floor(pn/ps)*ps 起算），不受影響。
  const mid=(mn+mx)/2;
  const half=step*3.5;
  const r={pn:mid-half+panY,px:mid+half+panY};
  _lastY={pn:r.pn,px:r.px,panY:panY};
  return r;
}
let _lastY=null;

function yOf(p,pn,px){return M.t+(1-(p-pn)/(px-pn))*cH;}
function p2y(y,pn,px){return px-((y-M.t)/cH)*(px-pn);}
function sz(base){return Math.max(base*.4,Math.min(base*2.5,base*Math.pow(zoom,.35)));}

// ── Focus day = day closest to screen center ──
function updateFocus(){
  const centerG=gi2x(W/2);
  let bestDi=0,bestDist=1e9;
  dayList.forEach((_,di)=>{
    const mid=gi(di,Math.floor(BPD/2));
    const d=Math.abs(centerG-mid);
    if(d<bestDist){bestDist=d;bestDi=di;}
  });
  if(bestDi!==focusIdx){
    focusIdx=bestDi;
    updateToolbar();
  }
}

function updateToolbar(){
  const date=dayList[focusIdx]||"";
  const day=dayCache.get(date);
  setDateBoxes(date);
  const pe=document.getElementById("dp"),dtc=document.getElementById("dtc");
  if(!day||day.noTrading){pe.textContent="非交易日";pe.className="pnl";pe.style.color="#8B8F98";dtc.textContent="";}
  else if(!day.trades||!day.trades.length){pe.textContent="無當日交易記錄";pe.className="pnl";pe.style.color="#8B8F98";dtc.textContent="";}
  else{pe.style.color="";
    const intradayPnl=day.trades.filter(t=>!t.isHold&&!t.crossDay).reduce((s,t)=>s+t.pnl,0);
    const closePnl=day.trades.filter(t=>t.crossDay&&!t.isHold).reduce((s,t)=>s+t.pnl,0);
    const totalTrades=day.trades.filter(t=>!t.isHold).length;
    let pnlStr="";
    if(closePnl!==0){
      // Show both: intraday±X closeout±Y
      const iS=`${intradayPnl>=0?"+":""}${fmt(intradayPnl)}`;
      const cS=`${closePnl>=0?"+":""}${fmt(closePnl)}`;
      pnlStr=`Day P&L: $${iS} 平倉$${cS}`;
    } else {
      const total=intradayPnl;
      pnlStr=`Day P&L: ${total>=0?"+":""}$${fmt(Math.abs(total))}`;
    }
    const combined=intradayPnl+closePnl;
    pe.textContent=pnlStr;pe.className="pnl "+(combined>=0?"w":"l");
    dtc.textContent=`(${totalTrades} trades)`;}
  renderCards();
  uNotes();
}

// ── Hit test trades (across all visible days) ──
function hitTest(mx,my,pn,px){
  for(let di=0;di<dayList.length;di++){
    const day=dayCache.get(dayList[di]);if(!day||!day.trades||!day.candles||!day.candles.length)continue;
    for(let ti=0;ti<day.trades.length;ti++){
      const t=day.trades[ti];
      if(t.entryTime==null||t.entryPrice==null)continue;
      // 留倉 (isHold) rows carry exitTime/exitPrice = null — t2iLocal would throw on
      // null.split(), which killed the whole mousemove handler (crosshair + hover dead
      // on every day holding a position). Collapse them to a point at the entry.
      const ei=t2iLocal(t.entryTime,day.candles);
      const xi=t.exitTime!=null?t2iLocal(t.exitTime,day.candles):ei;
      const exitP=t.exitPrice!=null?t.exitPrice:t.entryPrice;
      const x1=xOfGi(gi(di,ei)),y1=yOf(t.entryPrice,pn,px);
      const x2=xOfGi(gi(di,xi)),y2=yOf(exitP,pn,px);
      const dx=x2-x1,dy=y2-y1,l2=dx*dx+dy*dy;let d2;
      if(l2===0)d2=Math.hypot(mx-x1,my-y1);
      else{const tt=Math.max(0,Math.min(1,((mx-x1)*dx+(my-y1)*dy)/l2));d2=Math.hypot(mx-(x1+tt*dx),my-(y1+tt*dy));}
      if(d2<14)return{di,ti};
    }
  }
  return null;
}

// ═══════════════════════════════════════════════════
// DRAW
// ═══════════════════════════════════════════════════
function draw(){
  resize();
  const{pn,px}=autoY();
  // In export, force step=$1 (otherwise pStep(9) returns 2 and breaks square cells)
  const ps=exportMode?1:pStep(px-pn);
  // Collected during pass 2; consumed by inline notes obstacle list
  const pnlLabelsForNotes=[];

  // Theme: light mode for printing, dark mode for screen
  const T=exportLightMode?{
    bg:"#FFFFFF",grid:"#666666",gridM:"#444444",txt:"#000000",
    uB:"#16A34A",dB:"#DC2626",uW:"#16A34A",dW:"#DC2626",
    pW:"#000000",pL:"#CC0000",sep:"#CC6600"
  }:{
    bg:C.bg,grid:C.grid,gridM:C.gridM,txt:C.txt,
    uB:C.uB,dB:C.dB,uW:C.uW,dW:C.dW,
    pW:C.pW,pL:C.pL,sep:"#FFB347"
  };

  ctx.clearRect(0,0,W,H);ctx.fillStyle=T.bg;ctx.fillRect(0,0,W,H);
  ctx.save();ctx.beginPath();ctx.rect(M.l,0,cW,H);ctx.clip();

  // H grid
  const gs=Math.floor(pn/ps)*ps;
  ctx.strokeStyle=T.grid;ctx.lineWidth=exportLightMode?0.7:0.5;ctx.setLineDash([1,3]);
  for(let p=gs;p<=px;p+=ps){const y=yOf(p,pn,px);if(y<-20||y>H+20)continue;
  ctx.beginPath();ctx.moveTo(M.l,y);ctx.lineTo(W-M.r,y);ctx.stroke();}
  ctx.setLineDash([]);

  // Global day separator obstacles (so each day's pass 2 can avoid all separators, not just its own)
  const sepObstacles=[];
  for(let di=0;di<dayList.length;di++){
    if(di===0)continue; // no separator before first day
    const sepX=xOfGi(gi(di,0)-0.5);
    if(sepX>=M.l-5&&sepX<=W-M.r+5){
      sepObstacles.push({x:sepX-3,y:M.t,w:6,h:H-M.b-M.t});
    }
  }

  // Per-day drawing
  dayList.forEach((date,di)=>{
    const day=dayCache.get(date);if(!day||!day.candles.length)return;
    const candles=day.candles;

    // Day separator line (thinner)
    if(di>0){
      const sepG=gi(di,0)-0.5;
      const sepX=xOfGi(sepG);
      if(sepX>=M.l-5&&sepX<=W-M.r+5){
        ctx.strokeStyle=T.sep;ctx.lineWidth=1;ctx.globalAlpha=.8;ctx.setLineDash([8,4]);
        ctx.beginPath();ctx.moveTo(sepX,M.t);ctx.lineTo(sepX,H-M.b);ctx.stroke();
        ctx.setLineDash([]);ctx.globalAlpha=.85;
        if(!exportMode){
          ctx.fillStyle=T.sep;ctx.font="bold 10px Consolas,monospace";ctx.textAlign="center";
          ctx.fillText(date,sepX,M.t+14);
        }
        ctx.globalAlpha=1;
      }
    }

    // V grid (Bloomberg-style fine dotted)
    candles.forEach((c,li)=>{
      const x=xOfGi(gi(di,li));if(x<M.l||x>W-M.r)return;
      const mm=+c.t.split(":")[1];
      if(mm===0){ctx.strokeStyle=T.gridM;ctx.lineWidth=exportLightMode?0.7:0.5;ctx.setLineDash([1,3]);ctx.beginPath();ctx.moveTo(x,M.t);ctx.lineTo(x,H-M.b);ctx.stroke();ctx.setLineDash([]);}
      else if(mm===30){ctx.strokeStyle=T.grid;ctx.lineWidth=exportLightMode?0.7:0.5;ctx.setLineDash([1,3]);ctx.beginPath();ctx.moveTo(x,M.t);ctx.lineTo(x,H-M.b);ctx.stroke();ctx.setLineDash([]);}
    });

    // Candles (1.5x thicker in export)
    const wickW=exportMode?Math.max(.8,sz(1.05)):Math.max(.5,sz(.7));
    const bw=Math.max(1,ppb()*(exportMode?.82:.55));
    candles.forEach((c,li)=>{
      const x=xOfGi(gi(di,li));if(x<M.l-bw*2||x>W-M.r+bw*2)return;
      const up=c.c>=c.o;
      ctx.strokeStyle=up?T.uW:T.dW;ctx.lineWidth=wickW;
      ctx.beginPath();ctx.moveTo(x,yOf(c.h,pn,px));ctx.lineTo(x,yOf(c.l,pn,px));ctx.stroke();
      const bt=yOf(Math.max(c.o,c.c),pn,px),bb=yOf(Math.min(c.o,c.c),pn,px);
      ctx.fillStyle=up?T.uB:T.dB;ctx.fillRect(x-bw/2,bt,bw,Math.max(bb-bt,.5));
    });

    // Draw indicators (VWAP, BB, MA) — after K-lines, before trade lines
    drawIndicators(ctx,di,candles,pn,px,exportMode);

    // Trades — pass 1: lines + dots
    // Color scheme:
    // Trade line colors: light green (long) / light red (short)
    //   App (dark bg): brighter pastels for visibility
    //   Export (white bg): mid-tone pastels distinguishable from K-line red/green
    // Colors from selected version (loaded from colors.xlsx)
    const cv=(window.colorVersions&&window.colorVersions[window.selectedColorIdx||0])||null;
    const longCol =cv?(exportLightMode?cv.Export_Long :cv.App_Long ):(exportLightMode?"#0040C0":"#5BA0FF");
    const shortCol=cv?(exportLightMode?cv.Export_Short:cv.App_Short):(exportLightMode?"#000000":"#B0B0B0");
    const holdCol =cv?(exportLightMode?cv.Export_Hold :cv.App_Hold ):"#9B59B6";
    // Border colors (optional — empty string = no border ring)
    const longBdr =cv?(exportLightMode?(cv.Export_Long_Border||""):(cv.App_Long_Border||"")):"";
    const shortBdr=cv?(exportLightMode?(cv.Export_Short_Border||""):(cv.App_Short_Border||"")):"";
    const holdBdr =cv?(exportLightMode?(cv.Export_Hold_Border||""):(cv.App_Hold_Border||"")):"";
    function getBorder(isHold,isCross,isLong){
      if(isHold||isCross)return holdBdr;
      return isLong?longBdr:shortBdr;
    }
    // Multi-page export filter helper: trade belongs on a page if its midpoint
    // price is within [pn, px). Use half-open interval to avoid duplicates on
    // page boundaries. The topmost page extends its upper bound to be inclusive.
    const tradeBelongsOnPage=(t)=>{
      // Trade LINES: draw on ALL pages (canvas auto-clips out-of-range portions)
      return true;
    };
    // Trade LABELS: only draw on pages where the trade's price is within visible Y range
    // (prevents labels from being clamped to top/bottom edge and piling up)
    const labelBelongsOnPage=(t)=>{
      if(!exportMode||!exportYOverride)return true;
      // Use the trade's anchor price (exit for normal/cross, entry for hold)
      const refPrice=t.isHold?t.entryPrice:(t.exitPrice!=null?t.exitPrice:t.entryPrice);
      if(refPrice==null)return true;
      // Allow some margin (1 grid step) so labels near page boundary still appear
      const margin=(px-pn)*0.15;
      return refPrice>=(pn-margin)&&refPrice<=(px+margin);
    };
    (day.trades||[]).forEach((t,ti)=>{
      // In export mode, only draw lines for the focus day (prev day shows 3 bars for context only)
      if(exportMode&&di!==focusIdx)return;
      if(!tradeBelongsOnPage(t))return;
      const isHold=!!t.isHold;
      const isCross=!!t.crossDay;
      const hov=hovHit&&hovHit.di===di&&hovHit.ti===ti;
      const isLong=t.dir==="多";
      // Hold / close → holdCol; normal → long/short
      const tradeCol=(isHold||isCross)?holdCol:(isLong?longCol:shortCol);
      // Compute entry/exit screen positions
      const ei=(t.entryTime)?t2iLocal(t.entryTime,candles):0;
      const xi=(t.exitTime)?t2iLocal(t.exitTime,candles):0;
      let x1=xOfGi(gi(di,ei)), y1=yOf(t.entryPrice,pn,px);
      let x2=null,y2=null;
      if(!isHold){
        if(isCross){
          // Cross-day close: entry is on t.entryDate (different day). Find its di in dayList.
          const entryDi=dayList.indexOf(t.entryDate);
          if(exportMode){
            // Export: entry day is off-screen. Use a proportional offset from left margin
            // so the line angle visually points toward the entry day's direction.
            // Offset = number of days between entry and exit * pixels_per_day equivalent
            const daysBetween=entryDi>=0?(di-entryDi):30; // fallback 30 days if not found
            const offsetPerDay=40; // px per day of separation (tuned for visual clarity)
            x1=M.l-Math.min(daysBetween*offsetPerDay,W*0.5); // cap at half page width off-screen
            y1=yOf(t.entryPrice,pn,px);
          } else if(entryDi>=0){
            x1=xOfGi(gi(entryDi,ei));y1=yOf(t.entryPrice,pn,px);
          } else {
            // Entry day not in loaded range — start line from off-screen left
            x1=M.l-20;y1=yOf(t.entryPrice,pn,px);
          }
          x2=xOfGi(gi(di,xi));y2=yOf(t.exitPrice,pn,px);
        } else {
          x2=xOfGi(gi(di,xi));y2=yOf(t.exitPrice,pn,px);
        }
      }
      // Skip if endpoints all off-screen
      if(!isHold){
        if((x1<M.l||x1>W-M.r)&&(x2<M.l||x2>W-M.r))return;
      } else {
        if(x1<M.l||x1>W-M.r)return;
      }

      // Hover glow
      if(hov&&!isHold){
        ctx.strokeStyle="#FFD700";
        ctx.lineWidth=sz(8);ctx.globalAlpha=.25;ctx.lineCap="round";
        ctx.beginPath();ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.stroke();
        ctx.globalAlpha=1;
      }

      // Trade line (skip for holds)
      if(!isHold){
        const alpha=isCross?0.375:1.0;
        if(exportMode){
          const bdrCol=getBorder(isHold,isCross,isLong);
          const hasBorder=bdrCol&&bdrCol!=="nan";
          const totalW=sz(3)*1.54+2;
          const haloW=totalW;
          const mainW=hasBorder?Math.max(1,totalW-2.5):Math.max(1,totalW-2.0);
          // Halo: border color if set, else default white
          const haloColor=hasBorder?bdrCol:"rgba(255,255,255,0.6)";
          ctx.globalAlpha=alpha;
          ctx.strokeStyle=haloColor;
          ctx.lineWidth=haloW;ctx.setLineDash([]);ctx.lineCap="round";
          ctx.beginPath();ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.stroke();
          ctx.strokeStyle=tradeCol;ctx.lineWidth=mainW;
          ctx.setLineDash([sz(2),sz(4.5)]);
          ctx.beginPath();ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.stroke();
          ctx.setLineDash([]);ctx.globalAlpha=1;
        } else {
          const bdrCol=getBorder(isHold,isCross,isLong);
          if(bdrCol&&bdrCol!=="nan"){
            ctx.globalAlpha=alpha;
            ctx.strokeStyle=bdrCol;ctx.lineWidth=sz(3)*1.1+1;
            ctx.setLineDash([]);ctx.lineCap="round";
            ctx.beginPath();ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.stroke();
          }
          ctx.globalAlpha=alpha;
          ctx.strokeStyle=tradeCol;ctx.lineWidth=sz(3)*1.1;
          ctx.setLineDash([sz(2),sz(4)]);ctx.lineCap="round";
          ctx.beginPath();ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.stroke();
          ctx.setLineDash([]);ctx.globalAlpha=1;
        }
      }

      // Hold: semi-transparent circle around entry point (no filled dot)
      if(isHold){
        const drBase=Math.max(1.5,Math.min(4.5,2.5*Math.pow(zoom,0.5)));
        const dr=exportMode?drBase*1.4:drBase;
        ctx.strokeStyle=holdCol;ctx.globalAlpha=0.6;ctx.lineWidth=2;
        ctx.beginPath();ctx.arc(x1,y1,dr+6,0,Math.PI*2);ctx.stroke();
        ctx.globalAlpha=1;
      }
    });

    // Trades — pass 2: PnL labels with collision avoidance
    // Build obstacle list: K-line wicks (high to low) for this day + ALL day separators
    const obstacles=[];
    candles.forEach((c,li)=>{
      const x=xOfGi(gi(di,li));
      if(x<M.l-2||x>W-M.r+2)return;
      const yh=yOf(c.h,pn,px),yl=yOf(c.l,pn,px);
      obstacles.push({x:x-Math.max(1,bw/2),y:Math.min(yh,yl),w:Math.max(2,bw),h:Math.abs(yl-yh)+1});
    });
    // Add all day separators (global) so labels never cross any of them
    for(const s of sepObstacles)obstacles.push(s);
    // Add crossDay trade lines as obstacles — use a chain of small rectangles
    // along the diagonal line (not one huge bounding box which blocks too much area)
    (day.trades||[]).forEach(t=>{
      if(!t.crossDay||t.isHold)return;
      const ei2=(t.entryTime)?t2iLocal(t.entryTime,candles):0;
      const xi2=(t.exitTime)?t2iLocal(t.exitTime,candles):0;
      const entryDi2=dayList.indexOf(t.entryDate);
      let lx1;
      if(exportMode){
        const daysBetween=entryDi2>=0?(di-entryDi2):30;
        lx1=M.l-Math.min(daysBetween*40,W*0.5);
      } else {
        lx1=entryDi2>=0?xOfGi(gi(entryDi2,ei2)):M.l-20;
      }
      let ly1=yOf(t.entryPrice,pn,px);
      let lx2=xOfGi(gi(di,xi2));
      let ly2=yOf(t.exitPrice,pn,px);
      // Break line into N segments, each becomes a small obstacle rectangle
      const segLen=40; // px per segment
      const dx=lx2-lx1,dy=ly2-ly1;
      const lineLen=Math.sqrt(dx*dx+dy*dy)||1;
      const nSegs=Math.max(1,Math.ceil(lineLen/segLen));
      const pad=6; // half-width of the obstacle strip
      for(let s=0;s<nSegs;s++){
        const t0=s/nSegs,t1=(s+1)/nSegs;
        const sx0=lx1+dx*t0,sy0=ly1+dy*t0;
        const sx1=lx1+dx*t1,sy1=ly1+dy*t1;
        obstacles.push({
          x:Math.min(sx0,sx1)-pad,
          y:Math.min(sy0,sy1)-pad,
          w:Math.abs(sx1-sx0)+pad*2,
          h:Math.abs(sy1-sy0)+pad*2,
        });
      }
    });
    // Normal (same-day) trade lines are obstacles too — a PnL label must never sit
    // ON a dotted trade line (its own or another trade's). Same segment-chain
    // technique as crossDay above; combined with the line-midpoint anchor below,
    // the dy-walk settles each label snugly just above/below its own line.
    (day.trades||[]).forEach(t=>{
      if(t.isHold||t.crossDay)return;
      if(t.entryPrice==null||t.exitPrice==null)return;
      const ei3=(t.entryTime)?t2iLocal(t.entryTime,candles):0;
      const xi3=(t.exitTime)?t2iLocal(t.exitTime,candles):0;
      const lx1=xOfGi(gi(di,ei3)),ly1=yOf(t.entryPrice,pn,px);
      const lx2=xOfGi(gi(di,xi3)),ly2=yOf(t.exitPrice,pn,px);
      const segLen=40,dx=lx2-lx1,dy=ly2-ly1,lineLen=Math.sqrt(dx*dx+dy*dy)||1;
      const nSegs=Math.max(1,Math.ceil(lineLen/segLen)),pad=5;
      for(let s=0;s<nSegs;s++){
        const t0=s/nSegs,t1=(s+1)/nSegs;
        const sx0=lx1+dx*t0,sy0=ly1+dy*t0,sx1=lx1+dx*t1,sy1=ly1+dy*t1;
        obstacles.push({x:Math.min(sx0,sx1)-pad,y:Math.min(sy0,sy1)-pad,w:Math.abs(sx1-sx0)+pad*2,h:Math.abs(sy1-sy0)+pad*2});
      }
    });
    const placedLabels=[];
    const labelFontSize=exportMode?16:Math.round(sz(11));
    ctx.font=`bold ${labelFontSize}px Consolas,monospace`;
    function rectsCollide(a,b){return a.x<b.x+b.w&&a.x+a.w>b.x&&a.y<b.y+b.h&&a.y+a.h>b.y;}
    function labelOverlap(rx,ry,rw,rh){
      let total=0;
      for(const o of obstacles){
        const ox=Math.max(0,Math.min(rx+rw,o.x+o.w)-Math.max(rx,o.x));
        const oy=Math.max(0,Math.min(ry+rh,o.y+o.h)-Math.max(ry,o.y));
        total+=ox*oy;
      }
      for(const o of placedLabels){
        const ox=Math.max(0,Math.min(rx+rw,o.x+o.w)-Math.max(rx,o.x));
        const oy=Math.max(0,Math.min(ry+rh,o.y+o.h)-Math.max(ry,o.y));
        total+=ox*oy*10; // heavy penalty for overlapping other labels
      }
      return total;
    }
    function labelHasCollision(rx,ry,rw,rh){
      for(const o of obstacles){if(rectsCollide({x:rx,y:ry,w:rw,h:rh},o))return true;}
      for(const o of placedLabels){if(rectsCollide({x:rx,y:ry,w:rw,h:rh},o))return true;}
      return false;
    }

    // In export mode, only draw PnL labels for the focus day. Previous day's trades
    // exist in dayCache (they may be used for review on other dates) but we only show
    // 3 bars of the previous day for visual context — drawing its trade labels would
    // force them to be clamped to leftBound and pile up at the left edge.
    if(exportMode&&di!==focusIdx){
      return; // skip this whole day's pass 2
    }

    // ── Multi-page export: compute PnL label positions ONCE in absolute price
    // coordinates, then replay the SAME positions on every page. This guarantees a
    // label that falls inside the overlap band of two adjacent pages is drawn at the
    // identical price-axis coordinate on both, so the printed tiles align seamlessly
    // (no "same PnL shown at two different spots" problem). Screen mode and single-page
    // export are untouched — they fall through to the original per-page search below.
    if(exportMode&&exportPageTops&&exportPageTops.length>1){
      if(exportLabelCache===null){
        const pageRangeL=px-pn;                 // each page spans exactly one page range
        const topEdge=exportPageTops[0];
        const bottomEdge=exportPageTops[exportPageTops.length-1]-pageRangeL;
        const pxPerDollar=cH/pageRangeL;        // identical scale on every page
        const yAbs=p=>M.t+(topEdge-p)*pxPerDollar;   // page-independent y
        const virtTop=M.t;
        const virtBot=M.t+(topEdge-bottomEdge)*pxPerDollar; // full virtual chart height
        // Absolute-coordinate obstacles (K-line wicks + separators + crossDay lines)
        const aObs=[];
        candles.forEach((c,li)=>{
          const x=xOfGi(gi(di,li));
          if(x<M.l-2||x>W-M.r+2)return;
          const yh=yAbs(c.h),yl=yAbs(c.l);
          aObs.push({x:x-Math.max(1,bw/2),y:Math.min(yh,yl),w:Math.max(2,bw),h:Math.abs(yl-yh)+1});
        });
        for(const s of sepObstacles)aObs.push({x:s.x,y:virtTop,w:s.w,h:virtBot-virtTop});
        (day.trades||[]).forEach(t=>{
          if(!t.crossDay||t.isHold)return;
          const xi2=(t.exitTime)?t2iLocal(t.exitTime,candles):0;
          const entryDi2=dayList.indexOf(t.entryDate);
          const daysBetween=entryDi2>=0?(di-entryDi2):30;
          const lx1=M.l-Math.min(daysBetween*40,W*0.5);
          const ly1=yAbs(t.entryPrice);
          const lx2=xOfGi(gi(di,xi2));
          const ly2=yAbs(t.exitPrice);
          const segLen=40,dx=lx2-lx1,dy=ly2-ly1,lineLen=Math.sqrt(dx*dx+dy*dy)||1;
          const nSegs=Math.max(1,Math.ceil(lineLen/segLen)),pad=6;
          for(let s=0;s<nSegs;s++){
            const t0=s/nSegs,t1=(s+1)/nSegs;
            const sx0=lx1+dx*t0,sy0=ly1+dy*t0,sx1=lx1+dx*t1,sy1=ly1+dy*t1;
            aObs.push({x:Math.min(sx0,sx1)-pad,y:Math.min(sy0,sy1)-pad,w:Math.abs(sx1-sx0)+pad*2,h:Math.abs(sy1-sy0)+pad*2});
          }
        });
        // Normal (same-day) trade lines are obstacles too — a label must never sit
        // ON a dotted trade line (its own or another trade's). Same segment-chain
        // technique as crossDay above. Combined with the line-midpoint anchor below,
        // the dy-walk settles each label snugly just above/below its own line.
        (day.trades||[]).forEach(t=>{
          if(t.isHold||t.crossDay)return;
          if(t.entryPrice==null||t.exitPrice==null)return;
          const ei2=(t.entryTime)?t2iLocal(t.entryTime,candles):0;
          const xi2=(t.exitTime)?t2iLocal(t.exitTime,candles):0;
          const lx1=xOfGi(gi(di,ei2)),ly1=yAbs(t.entryPrice);
          const lx2=xOfGi(gi(di,xi2)),ly2=yAbs(t.exitPrice);
          const segLen=40,dx=lx2-lx1,dy=ly2-ly1,lineLen=Math.sqrt(dx*dx+dy*dy)||1;
          const nSegs=Math.max(1,Math.ceil(lineLen/segLen)),pad=5;
          for(let s=0;s<nSegs;s++){
            const t0=s/nSegs,t1=(s+1)/nSegs;
            const sx0=lx1+dx*t0,sy0=ly1+dy*t0,sx1=lx1+dx*t1,sy1=ly1+dy*t1;
            aObs.push({x:Math.min(sx0,sx1)-pad,y:Math.min(sy0,sy1)-pad,w:Math.abs(sx1-sx0)+pad*2,h:Math.abs(sy1-sy0)+pad*2});
          }
        });
        const aPlaced=[];
        const aHasCol=(rx,ry,rw,rh)=>{
          for(const o of aObs){if(rectsCollide({x:rx,y:ry,w:rw,h:rh},o))return true;}
          for(const o of aPlaced){if(rectsCollide({x:rx,y:ry,w:rw,h:rh},o))return true;}
          return false;
        };
        const aOverlap=(rx,ry,rw,rh)=>{
          let total=0;
          for(const o of aObs){const ox=Math.max(0,Math.min(rx+rw,o.x+o.w)-Math.max(rx,o.x));const oy=Math.max(0,Math.min(ry+rh,o.y+o.h)-Math.max(ry,o.y));total+=ox*oy;}
          for(const o of aPlaced){const ox=Math.max(0,Math.min(rx+rw,o.x+o.w)-Math.max(rx,o.x));const oy=Math.max(0,Math.min(ry+rh,o.y+o.h)-Math.max(ry,o.y));total+=ox*oy*10;}
          return total;
        };
        // Sort by label anchor X (line midpoint), NOT entryTime. A trade held across
        // the whole day enters early but anchors far right (its midX); sorting it
        // early poisons the monotonic-X cursor (prevCX) so every later trade whose
        // candidate window lies left of that midX has ALL candidates rejected and
        // falls into the prevCX teleport fallback — that is exactly the "labels
        // pile up in one vertical column" bug. Anchor-X order keeps prevCX naturally
        // non-decreasing. Specials (hold/cross) still go last, in entry order.
        const anchorXOf=(t)=>{
          const aei=(t.entryTime)?t2iLocal(t.entryTime,candles):0;
          const axi=(t.exitTime)?t2iLocal(t.exitTime,candles):0;
          if(t.isHold)return xOfGi(gi(di,aei));
          if(t.crossDay)return xOfGi(gi(di,axi));
          return (xOfGi(gi(di,aei))+xOfGi(gi(di,axi)))/2;
        };
        const tIdx=(day.trades||[]).map((t,ti)=>({t,ti,ax:anchorXOf(t)}));
        tIdx.sort((a,b)=>{
          const aS=a.t.isHold||a.t.crossDay?1:0,bS=b.t.isHold||b.t.crossDay?1:0;
          if(aS!==bS)return aS-bS;
          if(aS===0&&a.ax!==b.ax)return a.ax-b.ax;
          const ka=a.t.entryTime||"",kb=b.t.entryTime||"";
          if(ka!==kb)return ka<kb?-1:1;
          return a.ti-b.ti;
        });
        const boundPad=12;
        const dayLeftX=xOfGi(gi(di,0)-0.5)+boundPad;
        const dayRightX=xOfGi(gi(di,BPD-1)+0.5)-boundPad;
        const leftBound=Math.max(M.l+boundPad,dayLeftX);
        const rightBound=Math.min(W-M.r-boundPad,dayRightX);
        const out=[];
        let prevCX=-Infinity;
        tIdx.forEach(({t})=>{
          const isHold=!!t.isHold,isCross=!!t.crossDay;
          let lines=[];
          if(isHold){
            lines=[`${t.entryTime} 留倉`,`${t.shares} @ ${fmt(t.entryPrice)}`];
          } else if(isCross){
            const pnlS=(t.pnl>=0?"+":"")+fmt(t.pnl);
            const divTag=(t.div_income&&Math.abs(t.div_income)>0.005)?"(計入配息)":"";
            lines=[`庫存賣出 @ ${fmt(t.exitPrice)} ${pnlS}${divTag}`,`(${t.entryDate} buy @ ${fmt(t.entryPrice)})`];
            if(t.dividends&&t.dividends.length){for(const d of t.dividends){lines.push(`(${d.date} 配息 ${d.ps}@0.7)`);}}
          } else {
            lines=[(t.pnl>=0?"+":"")+fmt(t.pnl)];
          }
          if(isHold||isCross){ctx.font=`${labelFontSize}px Consolas,monospace`;}
          else{ctx.font=`bold ${labelFontSize}px Consolas,monospace`;}
          let maxW=0;for(const L of lines){const w=ctx.measureText(L).width;if(w>maxW)maxW=w;}
          const lblW=maxW+4;
          const lblH=lines.length*(labelFontSize+2)+2;
          const ei=(t.entryTime)?t2iLocal(t.entryTime,candles):0;
          const xi=(t.exitTime)?t2iLocal(t.exitTime,candles):0;
          let x1=xOfGi(gi(di,ei)),y1=yAbs(t.entryPrice),x2=null,y2=null;
          if(isHold){x2=x1;y2=y1;}
          else if(isCross){
            const entryDi=dayList.indexOf(t.entryDate);
            const daysBetween=entryDi>=0?(di-entryDi):30;
            x1=M.l-Math.min(daysBetween*40,W*0.5);
            x2=xOfGi(gi(di,xi));y2=yAbs(t.exitPrice);
          } else {x2=xOfGi(gi(di,xi));y2=yAbs(t.exitPrice);}
          if(!isHold){if((x1<M.l||x1>W-M.r)&&(x2<M.l||x2>W-M.r))return;}
          else{if(x1<M.l||x1>W-M.r)return;}
          let midX,baseAbove,baseBelow;
          if(isHold){midX=x1;baseAbove=y1-lblH/2-6;baseBelow=y1+lblH/2+6;}
          else if(isCross){midX=x2;baseAbove=y2-lblH/2-6;baseBelow=y2+lblH/2+6;}
          else{
            // Anchor at the LINE's height at midX (= (y1+y2)/2 on a straight line),
            // not the bounding-box top/bottom: for a multi-hour diagonal min/max is
            // the entry/exit apex price, dollars away from the line at midX — the
            // label would float in empty space. The line itself is an obstacle, so
            // the dy-walk pushes the label off the line to hug it locally.
            midX=(x1+x2)/2;
            const yMid=(y1+y2)/2;
            baseAbove=yMid-lblH/2-6;baseBelow=yMid+lblH/2+6;
          }
          const candidates=[];
          const searchY=isCross?400:200,searchXL=isCross?60:120,searchXR=isCross?500:120;
          for(let dy=0;dy<=searchY;dy+=4){candidates.push({x:midX-lblW/2,y:baseAbove-dy});candidates.push({x:midX-lblW/2,y:baseBelow+dy});}
          for(let dx=6;dx<=searchXR;dx+=6){
            for(let dy=0;dy<=searchY/2;dy+=8){
              candidates.push({x:midX-lblW/2+dx,y:baseAbove-dy});
              candidates.push({x:midX-lblW/2+dx,y:baseBelow+dy});
              if(dx<=searchXL){candidates.push({x:midX-lblW/2-dx,y:baseAbove-dy});candidates.push({x:midX-lblW/2-dx,y:baseBelow+dy});}
            }
          }
          let chosen=null,bestFallback=null,bestOverlap=Infinity;
          for(const c of candidates){
            if(c.x<leftBound||c.x+lblW>rightBound||c.y<virtTop||c.y+lblH>virtBot)continue;
            if(!isHold&&!isCross){const centerX=c.x+lblW/2;if(centerX<prevCX)continue;}
            if(!aHasCol(c.x,c.y,lblW,lblH)){chosen=c;break;}
            const ov=aOverlap(c.x,c.y,lblW,lblH);
            if(ov<bestOverlap){bestOverlap=ov;bestFallback=c;}
          }
          if(!chosen)chosen=bestFallback;
          if(!chosen){
            // Last resort: stay AT the anchor (tolerate overlap) — never teleport
            // to prevCX. The old prevCX term is what stacked every rejected label
            // into one vertical column at the cursor's X.
            const fx=Math.max(leftBound,midX-lblW/2);
            const fxClamped=Math.min(fx,rightBound-lblW);
            chosen={x:fxClamped,y:baseAbove};
          }
          chosen.x=Math.max(leftBound,Math.min(chosen.x,rightBound-lblW));
          chosen.y=Math.max(virtTop,Math.min(chosen.y,virtBot-lblH));
          aPlaced.push({x:chosen.x,y:chosen.y,w:lblW,h:lblH});
          if(!isHold&&!isCross){prevCX=chosen.x+lblW/2;}
          out.push({x:chosen.x,priceForY:topEdge-(chosen.y-M.t)/pxPerDollar,
                    lines:lines.slice(),isHold,isCross,pnl:t.pnl,lblW,lblH});
        });
        exportLabelCache=out;
      }
      // Replay cached placements on THIS page (clip to plot area so a label sitting in
      // the overlap band is cleanly cut at the page edge instead of bleeding into margins;
      // the same label appears whole at the same coordinate on the neighbouring page).
      ctx.save();ctx.beginPath();ctx.rect(M.l,M.t,cW,cH);ctx.clip();
      for(const L of exportLabelCache){
        const yTop=yOf(L.priceForY,pn,px);
        if(yTop+L.lblH<M.t||yTop>H-M.b)continue; // not on this page at all
        // Register this page's visible labels so the notes block avoids them
        // (the original per-page path does this at draw time; the replay must too,
        // otherwise notes text can be laid out right on top of a PnL label).
        pnlLabelsForNotes.push({x:L.x,y:yTop,w:L.lblW,h:L.lblH});
        if(L.isHold||L.isCross){ctx.fillStyle=holdCol;ctx.font=`${labelFontSize}px Consolas,monospace`;}
        else{ctx.fillStyle=L.pnl>=0?T.pW:T.pL;ctx.font=`bold ${labelFontSize}px Consolas,monospace`;}
        ctx.textAlign="left";
        for(let li=0;li<L.lines.length;li++){ctx.fillText(L.lines[li],L.x+2,yTop+labelFontSize+li*(labelFontSize+2));}
      }
      ctx.restore();
      ctx.font=`bold ${labelFontSize}px Consolas,monospace`;
      return; // multi-page export labels done — skip original per-page search
    }

    // Sort trades by label ANCHOR X (line midpoint), not entryTime, so the
    // monotonic-X cursor is naturally non-decreasing. Sorting by entryTime breaks
    // when a trade entered early is held for hours: its anchor sits far right, the
    // cursor jumps there, and every later trade whose anchor is left of it has all
    // candidates rejected → teleport fallback → labels pile into one column.
    // Specials (hold/cross) still go last, in entry order (exempt from the cursor).
    const scrAnchorXOf=(t)=>{
      const aei=(t.entryTime)?t2iLocal(t.entryTime,candles):0;
      const axi=(t.exitTime)?t2iLocal(t.exitTime,candles):0;
      if(t.isHold)return xOfGi(gi(di,aei));
      if(t.crossDay)return xOfGi(gi(di,axi));
      return (xOfGi(gi(di,aei))+xOfGi(gi(di,axi)))/2;
    };
    const tradesIndexed=(day.trades||[]).map((t,ti)=>({t,ti,ax:scrAnchorXOf(t)}));
    tradesIndexed.sort((a,b)=>{
      const aSpecial=a.t.isHold||a.t.crossDay?1:0;
      const bSpecial=b.t.isHold||b.t.crossDay?1:0;
      if(aSpecial!==bSpecial)return aSpecial-bSpecial;
      if(aSpecial===0&&a.ax!==b.ax)return a.ax-b.ax;
      const ka=a.t.entryTime||"";const kb=b.t.entryTime||"";
      if(ka!==kb)return ka<kb?-1:1;
      return a.ti-b.ti;
    });
    let prevLabelCenterX=-Infinity; // monotonic constraint for normal trades

    tradesIndexed.forEach(({t,ti})=>{
      if(!labelBelongsOnPage(t))return;
      const isHold=!!t.isHold;
      const isCross=!!t.crossDay;
      // Build label text (possibly multi-line for holds/closes)
      let lines=[];
      if(isHold){
        // "[時間] 留倉" / "[股數] @ [價]"
        lines=[`${t.entryTime} 留倉`, `${t.shares} @ ${fmt(t.entryPrice)}`];
      } else if(isCross){
        const pnlS=(t.pnl>=0?"+":"")+fmt(t.pnl);
        const divTag=(t.div_income&&Math.abs(t.div_income)>0.005)?"(計入配息)":"";
        lines=[`庫存賣出 @ ${fmt(t.exitPrice)} ${pnlS}${divTag}`,
               `(${t.entryDate} buy @ ${fmt(t.entryPrice)})`];
        // Add dividend lines
        if(t.dividends&&t.dividends.length){
          for(const d of t.dividends){
            lines.push(`(${d.date} 配息 ${d.ps}@0.7)`);
          }
        }
      } else {
        lines=[(t.pnl>=0?"+":"")+fmt(t.pnl)];
      }
      // Measure max width + total height (use normal weight for hold/close, bold for normal)
      if(isHold||isCross){ctx.font=`${labelFontSize}px Consolas,monospace`;}
      else{ctx.font=`bold ${labelFontSize}px Consolas,monospace`;}
      let maxW=0;
      for(const L of lines){const w=ctx.measureText(L).width;if(w>maxW)maxW=w;}
      const lblW=maxW+4;
      const lblH=lines.length*(labelFontSize+2)+2;
      // Positions
      const ei=(t.entryTime)?t2iLocal(t.entryTime,candles):0;
      const xi=(t.exitTime)?t2iLocal(t.exitTime,candles):0;
      let x1=xOfGi(gi(di,ei)),y1=yOf(t.entryPrice,pn,px);
      let x2=null,y2=null;
      if(isHold){
        x2=x1;y2=y1;
      } else if(isCross){
        const entryDi=dayList.indexOf(t.entryDate);
        if(exportMode){
          const daysBetween=entryDi>=0?(di-entryDi):30;
          x1=M.l-Math.min(daysBetween*40,W*0.5);
        } else if(entryDi>=0){x1=xOfGi(gi(entryDi,ei));}
        else{x1=M.l-20;}
        x2=xOfGi(gi(di,xi));y2=yOf(t.exitPrice,pn,px);
      } else {
        x2=xOfGi(gi(di,xi));y2=yOf(t.exitPrice,pn,px);
      }
      // Skip if completely off-screen X
      if(!isHold){
        if((x1<M.l||x1>W-M.r)&&(x2<M.l||x2>W-M.r))return;
      } else {
        if(x1<M.l||x1>W-M.r)return;
      }
      const txt=lines.join("\n"); // legacy usage — not rendered directly
      // midX: normal trades = midpoint of entry/exit; crossDay = exit point; isHold = entry point
      // baseAbove/baseBelow: normal trades = above/below the trade line endpoints
      let midX, baseAbove, baseBelow;
      if(isHold){
        midX=x1;
        baseAbove=y1-lblH/2-6;
        baseBelow=y1+lblH/2+6;
      } else if(isCross){
        midX=x2; // anchor near exit point (entry may be off-screen)
        baseAbove=y2-lblH/2-6;
        baseBelow=y2+lblH/2+6;
      } else {
        // Anchor at the LINE's height at midX (=(y1+y2)/2 on a straight line), not
        // the bounding-box top/bottom: for a long diagonal min/max is the apex
        // price, dollars away from the line at midX — the label floats in space.
        // The line itself is an obstacle now, so the dy-walk hugs the label to it.
        midX=(x1+x2)/2;
        const yMid=(y1+y2)/2;
        baseAbove=yMid-lblH/2-6;
        baseBelow=yMid+lblH/2+6;
      }

      // Generate candidates: dense grid around the trade midpoint
      // For crossDay labels (multi-line, tall), bias search rightward (exit point is often
      // at left edge where K-lines are dense; right side has more open space)
      const candidates=[];
      const searchY=isCross?400:200;
      const searchXL=isCross?60:120;   // leftward search (limited for crossDay — left edge)
      const searchXR=isCross?500:120;  // rightward search (wide for crossDay)
      // First: directly above/below midX
      for(let dy=0;dy<=searchY;dy+=4){
        candidates.push({x:midX-lblW/2,y:baseAbove-dy});
        candidates.push({x:midX-lblW/2,y:baseBelow+dy});
      }
      // Then: spread horizontally (right-biased for crossDay)
      for(let dx=6;dx<=searchXR;dx+=6){
        for(let dy=0;dy<=searchY/2;dy+=8){
          // Right side
          candidates.push({x:midX-lblW/2+dx,y:baseAbove-dy});
          candidates.push({x:midX-lblW/2+dx,y:baseBelow+dy});
          // Left side (limited range)
          if(dx<=searchXL){
            candidates.push({x:midX-lblW/2-dx,y:baseAbove-dy});
            candidates.push({x:midX-lblW/2-dx,y:baseBelow+dy});
          }
        }
      }

      let chosen=null;
      let bestFallback=null,bestOverlap=Infinity;
      // Bound labels to this day's own K-line time range
      // In export mode add extra padding for bold 16px font metrics to ensure
      // the first/last characters (especially "-" and digits) are never clipped
      const boundPad=exportMode?12:2;
      const dayLeftX=xOfGi(gi(di,0)-0.5)+boundPad;
      const dayRightX=xOfGi(gi(di,BPD-1)+0.5)-boundPad;
      const leftBound=Math.max(M.l+boundPad,dayLeftX);
      const rightBound=Math.min(W-M.r-boundPad,dayRightX);
      for(const c of candidates){
        if(c.x<leftBound||c.x+lblW>rightBound||c.y<M.t||c.y+lblH>H-M.b)continue;
        // Monotonic-X constraint: only for normal intraday trades (crossDay/isHold are exempt)
        if(!isHold&&!isCross){
          const centerX=c.x+lblW/2;
          if(centerX<prevLabelCenterX)continue;
        }
        if(!labelHasCollision(c.x,c.y,lblW,lblH)){chosen=c;break;}
        const ov=labelOverlap(c.x,c.y,lblW,lblH);
        if(ov<bestOverlap){bestOverlap=ov;bestFallback=c;}
      }
      if(!chosen)chosen=bestFallback;
      // Last-resort fallback: stay AT the anchor (tolerate overlap) — never
      // teleport to prevLabelCenterX; that stacked rejected labels into a column.
      if(!chosen){
        const fx=Math.max(leftBound,midX-lblW/2);
        const fxClamped=Math.min(fx,rightBound-lblW);
        chosen={x:fxClamped,y:baseAbove};
      }
      chosen.x=Math.max(leftBound,Math.min(chosen.x,rightBound-lblW));
      chosen.y=Math.max(M.t,Math.min(chosen.y,H-M.b-lblH));
      const finalRect={x:chosen.x,y:chosen.y,w:lblW,h:lblH};
      placedLabels.push(finalRect);
      pnlLabelsForNotes.push(finalRect);
      // Only normal intraday trades advance the monotonic-X cursor
      if(!isHold&&!isCross){prevLabelCenterX=chosen.x+lblW/2;}
      // Color + font weight: hold/close → holdCol, normal weight; normal → bold
      if(isHold||isCross){
        ctx.fillStyle=holdCol;
        ctx.font=`${labelFontSize}px Consolas,monospace`; // normal weight (thinner)
      } else {
        ctx.fillStyle=t.pnl>=0?T.pW:T.pL;
        ctx.font=`bold ${labelFontSize}px Consolas,monospace`; // bold
      }
      ctx.textAlign="left";
      // Render each line
      for(let li=0;li<lines.length;li++){
        ctx.fillText(lines[li],chosen.x+2,chosen.y+labelFontSize+li*(labelFontSize+2));
      }
      // Restore bold font for next iteration's measureText
      ctx.font=`bold ${labelFontSize}px Consolas,monospace`;
    });
  });
  ctx.restore();

  // Y axis (always right for screen; both sides during dual-axis export)
  const yAxisFont=exportMode?"13px Consolas,monospace":"11px Consolas,monospace";
  const yAxisInt=ps>=1&&ps===Math.floor(ps); // integer step → no decimals
  const fmtY=v=>yAxisInt?String(Math.round(v)):fmt(v);
  // Bottom label adjustment: shift up if too close to chart bottom (export only)
  const yLabelBottomAdjust=exportMode?6:0;
  // Prev-close label (ISO-week-first-day only, left side only).
  // Pick the integer price closest to the prev-close value so we can skip it
  // when drawing regular labels — the prev-close label replaces it.
  let prevCloseSkipP=null;
  if(exportMode&&exportPrevClose&&exportPrevClose.price!=null){
    prevCloseSkipP=Math.round(exportPrevClose.price/ps)*ps;
  }
  const drawYLabels=(side)=>{
    if(side==="right"){
      ctx.fillStyle=T.bg;ctx.fillRect(W-M.r,0,M.r,H);
      for(let p=gs;p<=px;p+=ps){const y=yOf(p,pn,px);if(y<M.t-5||y>H-M.b+15)continue;
      const adj=(y>H-M.b-2)?-yLabelBottomAdjust:0;
      ctx.fillStyle=T.txt;ctx.font=yAxisFont;
      // Export: left-align so numbers sit just inside the right margin without overlapping K lines
      if(exportMode){ctx.textAlign="left";ctx.fillText(fmtY(p),W-M.r+2,y+4+adj);}
      else{ctx.textAlign="left";ctx.fillText(fmtY(p),W-M.r+6,y+4+adj);}
      }
    } else {
      ctx.fillStyle=T.bg;ctx.fillRect(0,0,exportMode?M.l:M.l+60,H);
      for(let p=gs;p<=px;p+=ps){const y=yOf(p,pn,px);if(y<M.t-5||y>H-M.b+15)continue;
      // Skip the integer label nearest to prev close — we'll overwrite it with the prev-close label
      if(prevCloseSkipP!=null&&Math.abs(p-prevCloseSkipP)<ps*0.01)continue;
      const adj=(y>H-M.b-2)?-yLabelBottomAdjust:0;
      ctx.fillStyle=T.txt;ctx.font=yAxisFont;ctx.textAlign="left";ctx.fillText(fmtY(p),exportMode?2:M.l+56,y+4+adj);}
      // Prev-close label: drawn as a yellow pill tag inside the chart area near
      // the left edge (yellow-on-white is unreadable, so use a filled background).
      // Only on the ISO week's first trading day, only on the left Y axis.
      if(exportMode&&exportPrevClose&&exportPrevClose.price!=null){
        const pcPrice=exportPrevClose.price;
        const yPc=yOf(pcPrice,pn,px);
        if(yPc>=M.t-5&&yPc<=H-M.b+15){
          const pcText=`Close: ${fmt(pcPrice)}`;
          ctx.font=`bold ${yAxisFont.split(" ")[0]} Consolas,monospace`;
          const tw=ctx.measureText(pcText).width;
          const padX=6,padY=3;
          const tagW=tw+padX*2;
          const tagH=parseInt(yAxisFont,10)+padY*2;
          // Position: just inside the left edge of the chart area
          const tagX=0;
          const tagY=yPc-tagH/2;
          // Yellow translucent background (lets K-lines/grid show through)
          ctx.fillStyle="rgba(245,215,110,0.7)";
          ctx.fillRect(tagX,tagY,tagW,tagH);
          ctx.strokeStyle="#B8860B"; // dark goldenrod border
          ctx.lineWidth=1;
          ctx.strokeRect(tagX+0.5,tagY+0.5,tagW-1,tagH-1);
          // Dark text on yellow — high contrast
          ctx.fillStyle="#000000";
          ctx.textAlign="left";
          ctx.fillText(pcText,tagX+padX,yPc+4);
        }
      }
    }
  };
  if(exportDualAxis){drawYLabels("right");drawYLabels("left");}
  else if(exportYAxisLeft){drawYLabels("left");}
  else{drawYLabels("right");}

  // Last price tag of focus day (skip in export)
  if(!exportMode){
  const fDay=dayCache.get(dayList[focusIdx]);
  if(fDay&&fDay.candles.length){const lp2=fDay.candles[fDay.candles.length-1].c,ly2=yOf(lp2,pn,px);
  if(ly2>M.t&&ly2<H-M.b){ctx.fillStyle=C.lp;ctx.fillRect(W-M.r+2,ly2-9,54,18);
  ctx.fillStyle="#000";ctx.font="bold 10px Consolas,monospace";ctx.textAlign="center";ctx.fillText(fmt(lp2),W-M.r+29,ly2+4);}}
  }

  // X axis
  ctx.fillStyle=T.bg;ctx.fillRect(0,H-M.b,W,M.b);
  ctx.fillStyle=T.txt;ctx.font=exportMode?"12px Consolas,monospace":"10px Consolas,monospace";ctx.textAlign="center";
  dayList.forEach((date,di)=>{
    const day=dayCache.get(date);if(!day||!day.candles.length)return;
    day.candles.forEach((c,li)=>{
      const mm=+c.t.split(":")[1];if(mm===0||mm===30){
        // In export, skip 09:30 (the very first label) entirely
        if(exportMode&&li===0)return;
        const x=xOfGi(gi(di,li));
        if(x>=M.l&&x<=W-M.r)ctx.fillText(c.t,x,H-M.b+14);
      }
    });
  });
  // Export: date label + stats line below time labels
  if(exportMode){
    const focusDate=dayList[focusIdx]||"";
    if(focusDate){
      const isMultiPage=exportPageTops&&exportPageTops.length>1;
      const isLastPage=!isMultiPage||(currentExportPageIdx===exportPageTops.length-1);
      const focusDay=dayCache.get(focusDate);
      const trades=(focusDay&&focusDay.trades)||[];
      const yLine=H-M.b+38;
      // Date in bold Consolas
      ctx.font="bold 16px Consolas,monospace";
      const dowText=" "+dowStr(focusDate);
      const wDate=ctx.measureText(focusDate).width;
      // Measure day-of-week in lighter font
      ctx.font="300 14px Consolas,monospace";
      const wDow=ctx.measureText(dowText).width;
      // Stats parts in Times New Roman bold
      ctx.font="bold 18px 'Times New Roman',Times,serif";
      const gap="   "; // 3-space gap
      let nStr="",pnlStr="",wN=0,wP=0,wGap=0;
      if(isLastPage){
        if(trades.length===0){
          nStr="No Trades";
          wN=ctx.measureText(nStr).width;
        } else {
          const activeTrades=trades.filter(t=>!t.isHold);
          const n=activeTrades.length;
          const intradayPnl=trades.filter(t=>!t.isHold&&!t.crossDay).reduce((s,t)=>s+(t.pnl||0),0);
          const closePnl=trades.filter(t=>t.crossDay&&!t.isHold).reduce((s,t)=>s+(t.pnl||0),0);
          nStr=`${n} trades`;
          if(closePnl!==0){
            const iS=`${intradayPnl>=0?"+":""}${intradayPnl.toFixed(2)}`;
            const cS=`${closePnl>=0?"+":""}${closePnl.toFixed(2)}`;
            pnlStr=` $${iS} 平倉$${cS}`;
          } else {
            const total=intradayPnl;
            const sign=total>=0?"+":"-";
            pnlStr=` ${sign}$${Math.abs(total).toFixed(2)}`;
          }
          wN=ctx.measureText(nStr).width;
          wP=ctx.measureText(pnlStr).width;
        }
        wGap=ctx.measureText(gap).width;
      }
      const totalW=wDate+wDow+(isLastPage?(wGap+wN+wP):0);
      let x=W/2-totalW/2;
      // Date
      ctx.font="bold 16px Consolas,monospace";
      ctx.fillStyle=T.txt;ctx.textAlign="left";
      ctx.fillText(focusDate,x,yLine);
      x+=wDate;
      // Day-of-week (thinner)
      ctx.font="300 14px Consolas,monospace";
      ctx.fillStyle="#888888";
      ctx.fillText(dowText,x,yLine);
      x+=wDow;
      if(isLastPage){
        x+=wGap;
        ctx.font="bold 18px 'Times New Roman',Times,serif";
        if(trades.length===0){
          ctx.fillStyle=T.txt;
          ctx.fillText(nStr,x,yLine);
        } else {
          // "N trades" in blue
          ctx.fillStyle="#0040C0";
          ctx.fillText(nStr,x,yLine);
          x+=wN;
          // PnL color based on combined total
          const combined=trades.filter(t=>!t.isHold).reduce((s,t)=>s+(t.pnl||0),0);
          ctx.fillStyle=combined>=0?"#16A34A":"#DC2626";
          ctx.fillText(pnlStr,x,yLine);
        }
      }
      ctx.textAlign="center";
    }
  }

  // Crosshair
  if(mouse&&!cDrag&&!exportMode&&mouse.x>M.l&&mouse.x<W-M.r&&mouse.y>M.t&&mouse.y<H-M.b){
    ctx.strokeStyle=C.cross;ctx.lineWidth=.5;ctx.setLineDash([3,3]);
    ctx.beginPath();ctx.moveTo(M.l,mouse.y);ctx.lineTo(W-M.r,mouse.y);ctx.stroke();
    ctx.beginPath();ctx.moveTo(mouse.x,M.t);ctx.lineTo(mouse.x,H-M.b);ctx.stroke();ctx.setLineDash([]);
    // Y label
    const cp=p2y(mouse.y,pn,px);
    ctx.fillStyle="#1A1D22";ctx.fillRect(W-M.r+2,mouse.y-9,54,18);
    ctx.strokeStyle="#4A5060";ctx.lineWidth=.5;ctx.strokeRect(W-M.r+2,mouse.y-9,54,18);
    ctx.fillStyle="#7FE0E0";ctx.font="bold 10px Consolas,monospace";ctx.textAlign="center";ctx.fillText(fmt(cp),W-M.r+29,mouse.y+4);
    // Find which candle mouse is over
    const gPos=gi2x(mouse.x);
    let hitCandle=null,hitDate="";
    for(let di=0;di<dayList.length;di++){
      const day=dayCache.get(dayList[di]);if(!day||!day.candles.length)continue;
      const g0=gi(di,0),gN=gi(di,day.candles.length-1);
      if(gPos>=g0-1&&gPos<=gN+1){
        const li=Math.round(gPos-g0);
        if(li>=0&&li<day.candles.length){hitCandle=day.candles[li];hitDate=dayList[di];}
        break;
      }
    }
    // X label
    if(hitCandle){
      ctx.fillStyle="#1A1D22";ctx.fillRect(mouse.x-26,H-M.b+2,52,16);
      ctx.strokeStyle="#4A5060";ctx.lineWidth=.5;ctx.strokeRect(mouse.x-26,H-M.b+2,52,16);
      ctx.fillStyle="#7FE0E0";ctx.font="bold 10px Consolas,monospace";ctx.textAlign="center";
      ctx.fillText(hitCandle.t,mouse.x,H-M.b+14);
    }
  }

  // Export mode: draw notes inline (Bloomberg-style: white text, no background)
  if(exportMode){
    const day=dayCache.get(dayList[focusIdx]);
    const notesText=day?(day.notes||"").trim():"";
    if(notesText){
      // Multi-page filter: only draw on the assigned page (notesPageIdx)
      // Falls back to range check if notesPageIdx wasn't set (e.g., single-page export)
      const anchorInRange=notesPageIdx!=null
        ? currentExportPageIdx===notesPageIdx
        : (notesAnchorPrice==null||(notesAnchorPrice>=pn&&notesAnchorPrice<=px));
      if(anchorInRange){
      // Build obstacles once (shared across width attempts)
      const obstacles=[];
      dayList.forEach((date,di)=>{
        const day2=dayCache.get(date);if(!day2||!day2.candles.length)return;
        day2.candles.forEach((c,li)=>{
          const x=xOfGi(gi(di,li));
          if(x<M.l-2||x>W-M.r+2)return;
          const yh=yOf(c.h,pn,px),yl=yOf(c.l,pn,px);
          obstacles.push({x:x-2,y:Math.min(yh,yl),w:4,h:Math.abs(yl-yh)+1});
        });
        if(di>0){
          const sepX=xOfGi(gi(di,0)-0.5);
          if(sepX>=M.l-5&&sepX<=W-M.r+5){
            obstacles.push({x:sepX-3,y:M.t,w:6,h:H-M.b-M.t});
          }
        }
      });
      // PnL labels (already placed in pass 2) — notes must avoid them
      for(const lbl of pnlLabelsForNotes)obstacles.push(lbl);

      function hasCollision(rx,ry,rw,rh){
        for(const o of obstacles){
          if(rx<o.x+o.w&&rx+rw>o.x&&ry<o.y+o.h&&ry+rh>o.y)return true;
        }
        return false;
      }
      function overlapArea(rx,ry,rw,rh){
        let total=0;
        for(const o of obstacles){
          const ox=Math.max(0,Math.min(rx+rw,o.x+o.w)-Math.max(rx,o.x));
          const oy=Math.max(0,Math.min(ry+rh,o.y+o.h)-Math.max(ry,o.y));
          total+=ox*oy;
        }
        return total;
      }

      // Wrap text at given character limit (with hard-break for long words)
      function wrapAt(maxChars){
        const rawLines=notesText.split("\n");
        const wrapped=[];
        function hardBreak(s){
          const out=[];
          for(let i=0;i<s.length;i+=maxChars)out.push(s.slice(i,i+maxChars));
          return out;
        }
        rawLines.forEach(line=>{
          if(line.length<=maxChars){wrapped.push(line);return;}
          let cur="";
          line.split(/(\s+)/).forEach(word=>{
            if(word.length>maxChars){
              if(cur.trim()){wrapped.push(cur.trimEnd());cur="";}
              const chunks=hardBreak(word);
              for(let i=0;i<chunks.length-1;i++)wrapped.push(chunks[i]);
              cur=chunks[chunks.length-1];
              return;
            }
            if((cur+word).length>maxChars&&cur){wrapped.push(cur.trimEnd());cur=word.trimStart();}
            else cur+=word;
          });
          if(cur.trim())wrapped.push(cur.trimEnd());
        });
        return wrapped;
      }

      // Try a candidate wrap width: returns {wrapped, blockW, blockH, pos, overlap, perfect}
      function tryWidth(maxChars){
        const wrapped=wrapAt(maxChars);
        ctx.font="21px 'Times New Roman',Times,serif";
        let blockW=0;
        wrapped.forEach(l=>{const w=ctx.measureText(l).width;if(w>blockW)blockW=w;});
        const lineH=26;
        const blockH=wrapped.length*lineH+8;
        blockW+=10;

        const _cd=dayList[focusIdx];
        const _npX=(_cd&&notesStateByDate[_cd])?notesStateByDate[_cd].x:0.5;
        const _npY=(_cd&&notesStateByDate[_cd])?notesStateByDate[_cd].y:0.1;
        const desiredX=W*_npX;
        const desiredY=notesAnchorPrice!=null?yOf(notesAnchorPrice,pn,px):H*_npY;
        const minX=M.l+5,maxX=W-M.r-blockW-5;
        const minY=M.t+5,maxY=H-M.b-blockH-5;
        if(maxX<minX||maxY<minY)return null; // doesn't even fit in chart area
        const clamp=(v,lo,hi)=>Math.max(lo,Math.min(hi,v));
        const dx=clamp(desiredX,minX,maxX),dy=clamp(desiredY,minY,maxY);
        if(!hasCollision(dx,dy,blockW,blockH))return{wrapped,blockW,blockH,lineH,pos:{x:dx,y:dy},overlap:0,perfect:true};
        let bestPos={x:dx,y:dy},bestOverlap=overlapArea(dx,dy,blockW,blockH);
        const stepPx=8;
        const maxR=Math.hypot(W,H);
        for(let r=stepPx;r<maxR;r+=stepPx){
          const cands=[
            {x:dx,y:dy-r},{x:dx,y:dy+r},
            {x:dx-r,y:dy},{x:dx+r,y:dy},
            {x:dx-r*0.7,y:dy-r*0.7},{x:dx+r*0.7,y:dy-r*0.7},
            {x:dx-r*0.7,y:dy+r*0.7},{x:dx+r*0.7,y:dy+r*0.7},
            {x:dx-r*0.4,y:dy-r},{x:dx+r*0.4,y:dy-r},
            {x:dx-r*0.4,y:dy+r},{x:dx+r*0.4,y:dy+r},
            {x:dx-r,y:dy-r*0.4},{x:dx+r,y:dy-r*0.4},
            {x:dx-r,y:dy+r*0.4},{x:dx+r,y:dy+r*0.4},
          ];
          for(const c of cands){
            const cx=clamp(c.x,minX,maxX),cy=clamp(c.y,minY,maxY);
            if(!hasCollision(cx,cy,blockW,blockH))return{wrapped,blockW,blockH,lineH,pos:{x:cx,y:cy},overlap:0,perfect:true};
            const ov=overlapArea(cx,cy,blockW,blockH);
            if(ov<bestOverlap){bestOverlap=ov;bestPos={x:cx,y:cy};}
          }
        }
        return{wrapped,blockW,blockH,lineH,pos:bestPos,overlap:bestOverlap,perfect:false};
      }

      // Try widths from wide to narrow; pick the widest that fits perfectly
      const widthLadder=[100,80,65,50,38,28];
      let chosen=null,fallback=null,fallbackOverlap=Infinity;
      for(const mc of widthLadder){
        const result=tryWidth(mc);
        if(!result)continue;
        if(result.perfect){chosen=result;break;}
        // Track narrowest fallback with smallest overlap
        if(result.overlap<fallbackOverlap){fallbackOverlap=result.overlap;fallback=result;}
      }
      if(!chosen)chosen=fallback;
      if(!chosen)chosen=tryWidth(50); // ultimate fallback

      const{wrapped,pos,lineH}=chosen;
      ctx.font="21px 'Times New Roman',Times,serif";
      if(exportLightMode){
        ctx.fillStyle="#000000";
        ctx.shadowBlur=0;
      } else {
        ctx.fillStyle="#FFFFFF";
        ctx.shadowColor="rgba(0,0,0,0.85)";ctx.shadowBlur=4;
      }
      ctx.textAlign="left";
      wrapped.forEach((line,i)=>{ctx.fillText(line,pos.x+5,pos.y+(i+1)*lineH);});
      ctx.shadowBlur=0;
      }
    }
  }

  // Export legend: draw color-coded legend above chart top (in top margin) — FIRST PAGE ONLY
  if(exportMode&&(!exportPageTops||!exportPageTops.length||currentExportPageIdx===0)){
    const cv=(window.colorVersions&&window.colorVersions[window.selectedColorIdx||0]);
    if(cv){
      const day=dayCache.get(dayList[focusIdx]);
      const trades=(day&&day.trades)||[];
      const hasHold=trades.some(t=>t.isHold);
      const hasClose=trades.some(t=>t.crossDay&&!t.isHold);
      const items=[
        {color:cv.Export_Long,label:"多"},
        {color:cv.Export_Short,label:"空"},
      ];
      if(hasHold&&hasClose)items.push({color:cv.Export_Hold,label:"留倉/平倉"});
      else if(hasHold)items.push({color:cv.Export_Hold,label:"留倉"});
      else if(hasClose)items.push({color:cv.Export_Hold,label:"平倉"});
      ctx.save();
      ctx.font="bold 14px 'Times New Roman',Times,serif";
      ctx.textBaseline="middle";
      const swSize=12,gapAfterSw=5,gapBetween=16;
      // Measure total width
      let totalW=0;
      for(const it of items){
        totalW+=swSize+gapAfterSw+ctx.measureText(it.label).width+gapBetween;
      }
      totalW-=gapBetween;
      // Position: centered horizontally, at y = M.t/2
      let x=W/2-totalW/2;
      const y=Math.max(10,M.t/2);
      for(const it of items){
        // Color swatch
        ctx.fillStyle=it.color;
        ctx.fillRect(x,y-swSize/2,swSize,swSize);
        ctx.strokeStyle="#555555";
        ctx.lineWidth=0.5;
        ctx.strokeRect(x+0.5,y-swSize/2+0.5,swSize-1,swSize-1);
        x+=swSize+gapAfterSw;
        // Label
        ctx.fillStyle="#000000";
        ctx.textAlign="left";
        ctx.fillText(it.label,x,y);
        x+=ctx.measureText(it.label).width+gapBetween;
      }
      ctx.restore();
    }
  }
  // Update notes box position (follows K-line on pan)
  if(!exportMode&&typeof pNotes==="function"&&document.getElementById("nb").style.display!=="none")pNotes();
}

// ═══════════════════════════════════════════════════
// EDGE DETECTION: load adjacent days when panning
// ═══════════════════════════════════════════════════
let hovHit=null;

// Load a trading day (skip weekends/holidays), returns date loaded or null
async function loadTradingDay(startDate,direction){
  for(let i=1;i<=7;i++){
    const ds=addDays(startDate,direction*i);
    if(dayCache.has(ds)){
      const d=dayCache.get(ds);
      if(d.candles&&d.candles.length)return ds; // already cached trading day
      continue; // cached but non-trading, skip
    }
    try{
      const res=await fetch(`/api/data?date=${ds}`,{cache:"no-store"});const data=await res.json();
      dayCache.set(ds,{candles:data.candles||[],trades:data.trades||[],notes:data.notes||"",noTrading:!!data.noTrading});
      if(data.candles&&data.candles.length)return ds; // found trading day
    }catch(e){dayCache.set(ds,{candles:[],trades:[],notes:"",noTrading:true});}
  }
  return null; // no trading day found within 7 days
}

// Shift the whole virtual timeline by dpx pixels.
// Prepending a day renumbers every global index by +(BPD+GAP), so panX must be
// pulled back by one day-width to keep the picture still. CRITICAL: an in-flight
// canvas drag recomputes panX from its own anchor every mousemove
// (panX = cDP.x + dx), which would silently DISCARD that compensation and
// teleport the view a full day — which then re-triggers the edge check and
// cascades into skipping many days per drag (root cause of the 跨日跳多天 bug).
// So the drag anchor has to move with the world.
function shiftWorldX(dpx){
  panX+=dpx;
  cDP.x+=dpx;   // keep an in-flight cDrag anchored to the shifted world
}

async function checkEdges(){
  if(loading||!dayList.length)return;
  loading=true;
  try{
    // One call may need several days when zoomed far out (ZMIN shows ~4 days).
    // Each prepend raises leftG by exactly (BPD+GAP) and each append raises
    // lastDayEnd by the same, so the loop is guaranteed to converge; the cap is
    // only a backstop against a pathological viewport.
    for(let n=0;n<6;n++){
      let didLoad=false;
      // Near left edge → load previous trading day
      if(gi2x(M.l)<gi(0,0)+60){
        const found=await loadTradingDay(dayList[0],-1);
        if(found){
          dayList.unshift(found);focusIdx++;
          // ppb() read AFTER the await, so a zoom change mid-fetch can't skew it
          shiftWorldX(-(BPD+GAP)*ppb());
          didLoad=true;
        }
      }
      // Near right edge → load next trading day
      if(gi2x(W-M.r)>gi(dayList.length-1,BPD-1)-60){
        const found=await loadTradingDay(dayList[dayList.length-1],1);
        if(found){dayList.push(found);didLoad=true;}
      }
      if(!didLoad)break;
    }
  } finally {
    loading=false;
  }
  draw();
}

// ═══════════════════════════════════════════════════
// Notes
// ═══════════════════════════════════════════════════
const nb=document.getElementById("nb");
function uNotes(){
  const day=dayCache.get(dayList[focusIdx]);
  const notes=day?day.notes||"":"";
  nb.style.display="block";
  const curDate=dayList[focusIdx]||"";
  const hdr='<div class="nb-hdr"><span>'+curDate+' '+dowStr(curDate)+'</span><button class="nb-min" id="nbtn" title="最小化/展開">'+(notesMini?'&#9633;':'&#8212;')+'</button></div>';
  if(notesMini){nb.className="mini";nb.innerHTML=hdr;}
  else{nb.className="";
    if(notes.trim())nb.innerHTML=hdr+'<div class="nb-body">'+notes.replace(/</g,"&lt;").replace(/\n/g,"<br>")+'<div class="eh">雙擊編輯 ｜ 拖曳移動</div></div>';
    else nb.innerHTML=hdr+'<div class="nb-body"><span style="color:#555;font-style:italic">雙擊新增市場概述</span></div>';
  }
  pNotes();
  document.getElementById("nbtn").addEventListener("click",e=>{e.stopPropagation();notesMini=!notesMini;notesStateByDate._mini=notesMini;_saveNotesState();uNotes();});
}
function pNotes(){
  const curDate=dayList[focusIdx];
  if(!curDate){nb.style.display="none";return;}
  nb.style.display="block";
  const st=_getNoteState(curDate);
  const r=cc.getBoundingClientRect();
  // X anchor: relative to focus day's K-line range (x fraction 0..1 maps to [gi(di,0), gi(di,BPD-1)])
  const di=focusIdx;
  const xStart=xOfGi(gi(di,0));
  const xEnd=xOfGi(gi(di,BPD-1));
  const leftX=xStart+(xEnd-xStart)*st.x;
  nb.style.left=leftX+"px";
  nb.style.top=(r.height*st.y)+"px";
  if(!notesMini&&st.w&&st.h){
    nb.style.width=st.w+"px";
    nb.style.height=st.h+"px";
  } else if(notesMini){
    nb.style.width="";nb.style.height="";
  } else {
    nb.style.width="";nb.style.height="";
  }
}
// Capture user-resized dimensions
let _nbResizeObs=null;
if(window.ResizeObserver){
  _nbResizeObs=new ResizeObserver(entries=>{
    if(notesMini||nDrag)return;
    const curDate=dayList[focusIdx];
    if(!curDate)return;
    for(const e of entries){
      const cr=e.contentRect;
      if(cr.width>=180&&cr.height>=50){
        const st=_getNoteState(curDate);
        st.w=Math.round(e.target.offsetWidth);
        st.h=Math.round(e.target.offsetHeight);
        _saveNotesState();
      }
    }
  });
  _nbResizeObs.observe(nb);
}
nb.addEventListener("mousedown",e=>{
  if(e.target.closest(".nb-min"))return;
  if(e.detail===2)return;
  // Check if in bottom-right resize handle zone (~16px corner)
  const r=nb.getBoundingClientRect();
  const inResizeZone=(e.clientX>r.right-16)&&(e.clientY>r.bottom-16);
  if(inResizeZone&&!notesMini)return; // let native resize take over
  nDrag=true;nOff={x:e.clientX-r.left,y:e.clientY-r.top};
  e.preventDefault();e.stopPropagation();
});
nb.addEventListener("dblclick",e=>{if(e.target.closest(".nb-min"))return;if(notesMini){notesMini=false;uNotes();return;}
  const day=dayCache.get(dayList[focusIdx]);
  document.getElementById("ned-title").textContent=(dayList[focusIdx]||"")+" 市場概述";
  document.getElementById("nt").value=day?day.notes||"":"";
  document.getElementById("ned").classList.add("show");});

// ═══════════════════════════════════════════════════
// Cards (show focus day's trades)
// ═══════════════════════════════════════════════════
function renderCards(){
  const bar=document.getElementById("tb2");bar.innerHTML="";
  const day=dayCache.get(dayList[focusIdx]);
  if(!day||!day.trades)return;
  day.trades.forEach((t,i)=>{
    const d=document.createElement("div");d.className="tc2";d.id="tc"+i;
    let rowHTML;
    if(t.isHold){
      rowHTML=`<div class="dr">⊙ ${t.entryTime} 留倉</div><div class="pv">${t.shares}股</div><div class="dt">@ ${fmt(t.entryPrice)}</div>`;
    } else if(t.crossDay){
      const pc=t.pnl>=0?"w":"l";
      rowHTML=`<div class="dr">📦 庫存賣出 ${t.exitTime}</div><div class="pv ${pc}">${t.pnl>=0?"+":""}${fmt(t.pnl)}</div><div class="dt">${fmt(t.entryPrice)} → ${fmt(t.exitPrice)}</div>`;
    } else {
      const ar=t.dir==="多"?"▲ 多":"▼ 空",pc=t.pnl>=0?"w":"l";
      rowHTML=`<div class="dr">${ar} ${t.entryTime} → ${t.exitTime}</div><div class="pv ${pc}">${t.pnl>=0?"+":""}${fmt(t.pnl)}</div><div class="dt">${fmt(t.entryPrice)} → ${fmt(t.exitPrice)}</div>`;
    }
    d.innerHTML=rowHTML;
    d.addEventListener("mouseenter",()=>{hovHit={di:focusIdx,ti:i};hlCard(i);draw();});
    d.addEventListener("mouseleave",()=>{hovHit=null;hlCard(-1);draw();});
    bar.appendChild(d);
  });
}
function hlCard(idx){document.querySelectorAll(".tc2").forEach((c,i)=>{c.classList.toggle("hi",i===idx);
if(i===idx)c.scrollIntoView({behavior:"smooth",block:"nearest",inline:"nearest"});});}

// ═══════════════════════════════════════════════════
// Mouse & Drag
// ═══════════════════════════════════════════════════
let xAxisDrag=false,xAxisDS=0,xAxisZoomStart=0;

cv.addEventListener("mousedown",e=>{if(e.button!==0)return;
const r=cv.getBoundingClientRect(),mx=e.clientX-r.left,my=e.clientY-r.top;
// X-axis zone: bottom 28px strip
if(my>=H-M.b&&mx>M.l&&mx<W-M.r){
  xAxisDrag=true;xAxisDS=e.clientX;xAxisZoomStart=zoom;
  cv.style.cursor="ew-resize";return;
}
// Y-axis zone: right 68px strip
if(mx>=W-M.r&&my>M.t&&my<H-M.b){
  yAxisDrag=true;yAxisDS=e.clientY;yAxisZoomStart=userYZoom;
  cv.style.cursor="ns-resize";return;
}
if(mx>M.l&&mx<W-M.r){cDrag=true;cDS={x:e.clientX,y:e.clientY};cDP={x:panX,y:panY};cv.style.cursor="grabbing";}});

document.addEventListener("mousemove",e=>{
  if(nDrag){
    const curDate=dayList[focusIdx];
    if(!curDate)return;
    const r=cc.getBoundingClientRect();
    const ccRect=cc.getBoundingClientRect();
    // Desired left position of nb in page coords, then convert to K-line X fraction
    const pageLeft=e.clientX-nOff.x;
    const di=focusIdx;
    const xStart=xOfGi(gi(di,0));
    const xEnd=xOfGi(gi(di,BPD-1));
    // xStart/xEnd are in canvas coords (cv), which equals cc coords for left offset
    const leftInCC=pageLeft-ccRect.left;
    const spanX=xEnd-xStart;
    const st=_getNoteState(curDate);
    if(spanX>0){
      st.x=(leftInCC-xStart)/spanX;
      // clamp to reasonable range (allow dragging slightly outside for flexibility)
      st.x=Math.max(-2,Math.min(2,st.x));
    }
    st.y=Math.max(0,Math.min(.9,(e.clientY-r.top-nOff.y)/r.height));
    pNotes();return;
  }
  if(xAxisDrag){
    const dx=e.clientX-xAxisDS;
    const sensitivity=0.004;
    const newZoom=Math.max(ZMIN,Math.min(ZMAX,xAxisZoomStart*Math.exp(-dx*sensitivity)));
    const centerG=gi2x(W/2);
    zoom=newZoom;
    panX=W/2-M.l-centerG*ppb();
    mouse=null;draw();checkEdges();return;
  }
  if(yAxisDrag){
    const dy=e.clientY-yAxisDS;
    // Drag down (positive dy) = narrower range (zoom in) = smaller userYZoom
    // Drag up (negative dy) = wider range (zoom out) = larger userYZoom
    const sensitivity=0.005;
    userYZoom=Math.max(0.1,Math.min(20,yAxisZoomStart*Math.exp(dy*sensitivity)));
    mouse=null;draw();return;
  }
  if(cDrag){
    panX=cDP.x+(e.clientX-cDS.x);
    const{pn,px}=autoY();const pRange=px-pn;
    panY=cDP.y+(e.clientY-cDS.y)*pRange/cH;
    mouse=null;draw();checkEdges();updateFocus();return;
  }
  const r=cv.getBoundingClientRect();mouse={x:e.clientX-r.left,y:e.clientY-r.top};
  // Axis strips are drag-to-zoom targets but looked identical to the chart body.
  // Mirror the cursor mousedown would set, so the affordance is visible on hover.
  if(mouse.x>=0&&mouse.x<=W&&mouse.y>=0&&mouse.y<=H){
    if(mouse.y>=H-M.b&&mouse.x>M.l&&mouse.x<W-M.r)cv.style.cursor="ew-resize";
    else if(mouse.x>=W-M.r&&mouse.y>M.t&&mouse.y<H-M.b)cv.style.cursor="ns-resize";
    else cv.style.cursor="crosshair";
  }
  const{pn,px}=autoY();
  const ht=hitTest(mouse.x,mouse.y,pn,px);
  if(JSON.stringify(ht)!==JSON.stringify(hovHit)){
    hovHit=ht;
    if(ht&&ht.di===focusIdx)hlCard(ht.ti);else hlCard(-1);
  }
  draw();
});

document.addEventListener("mouseup",()=>{
  if(cDrag){cDrag=false;cv.style.cursor="crosshair";}
  if(xAxisDrag){xAxisDrag=false;cv.style.cursor="crosshair";}
  if(yAxisDrag){yAxisDrag=false;cv.style.cursor="crosshair";}
  if(nDrag){nDrag=false;_saveNotesState();}
});
cv.addEventListener("mouseleave",()=>{if(!cDrag&&!xAxisDrag&&!yAxisDrag){mouse=null;hovHit=null;hlCard(-1);draw();}});

function doZoom(d){const o=zoom;zoom=Math.max(ZMIN,Math.min(ZMAX,zoom+d));
const ctr=cW/2;panX=panX*(zoom/o)+(ctr*(1-zoom/o));draw();checkEdges();}
document.getElementById("zi").addEventListener("click",()=>doZoom(ZS));
document.getElementById("zo").addEventListener("click",()=>doZoom(-ZS));
document.getElementById("zr").addEventListener("click",()=>{fitView();draw();});
cv.addEventListener("wheel",e=>{e.preventDefault();doZoom(e.deltaY<0?ZS:-ZS);},{passive:false});

// ═══════════════════════════════════════════════════
// EXPORT PNG (Bloomberg-style with inline notes)
// ═══════════════════════════════════════════════════
let exportMode=false;
let exportYOverride=null; // {pn, px} forces specific Y range during export
let exportYAxisLeft=false; // legacy single-side flag (kept for compat)
let exportLightMode=false; // white bg, black dotted grid for printing
let exportDualAxis=false; // show both left and right Y axis
let exportSquareCells=false; // force square grid cells (1 dollar = 30 min visually)
let notesAnchorPrice=null; // price the notes block is anchored to (multi-page filter)
let notesPageIdx=null; // exact page index notes should appear on (for multi-page export)
let currentExportPageIdx=null; // index of the page currently being rendered
let exportPageTops=null; // array of pageTop values for all pages (multi-page export)
let exportPrevClose=null; // {price, mmdd} for ISO-week-first-day label, or null
let exportLabelCache=null; // multi-page export: PnL label placements computed ONCE in
                           // absolute price coordinates, replayed identically on every
                           // page so overlap-band labels land at the same price axis spot

// ISO week number for a date string
function isoWeek(ds){
  const d=new Date(ds+"T12:00:00Z");
  const day=d.getUTCDay()||7;
  d.setUTCDate(d.getUTCDate()+4-day);
  const yearStart=new Date(Date.UTC(d.getUTCFullYear(),0,1));
  return Math.ceil((((d-yearStart)/86400000)+1)/7);
}

// Sync version: only checks dayCache (no fetch). If we don't have data for
// earlier days in the same week, we conservatively assume they exist as
// trading days, so this returns false unless we have actual evidence.
// Exception: if the date is a Monday (ISO weekday 1), it's always the first.
function isFirstTradingDayOfWeekSync(ds){
  const d=new Date(ds+"T12:00:00Z");
  const isoDay=d.getUTCDay()||7; // 1=Mon, 7=Sun
  if(isoDay===1)return true; // Monday is always first
  const wk=isoWeek(ds);
  const yr=d.getUTCFullYear();
  // For Tue-Fri, check if any earlier weekday in this ISO week has candles in cache
  for(let i=1;i<isoDay;i++){
    const prev=addDays(ds,-i);
    if(isoWeek(prev)!==wk||new Date(prev+"T12:00:00").getFullYear()!==yr)break;
    const dd=dayCache.get(prev);
    if(dd&&dd.candles&&dd.candles.length)return false;
  }
  // No earlier trading day found in cache for this week → treat as first
  return true;
}

// Is this date the first trading day of its ISO week within dayCache?
async function isFirstTradingDayOfWeek(ds){
  const wk=isoWeek(ds);
  const yr=new Date(ds+"T12:00:00").getFullYear();
  // Walk back day by day; if any earlier day in the same ISO week is a trading day, this is not first
  for(let i=1;i<=6;i++){
    const prev=addDays(ds,-i);
    if(isoWeek(prev)!==wk||new Date(prev+"T12:00:00").getFullYear()!==yr)break;
    if(!dayCache.has(prev)){
      try{const res=await fetch(`/api/data?date=${prev}`);const data=await res.json();
        dayCache.set(prev,{candles:data.candles||[],trades:data.trades||[],notes:data.notes||"",noTrading:!!data.noTrading});
      }catch(e){dayCache.set(prev,{candles:[],trades:[],notes:"",noTrading:true});}
    }
    const d=dayCache.get(prev);
    if(d&&d.candles&&d.candles.length)return false; // earlier trading day exists
  }
  return true;
}

let _exporting=false; // re-entrancy guard: a 2nd run would save the ALREADY-exported
                      // state as "saved", permanently corrupting the view on restore
async function exportPNG(){
  if(_exporting)return;
  const day=dayCache.get(dayList[focusIdx]);
  if(!day||!day.candles||!day.candles.length){alert("無 K 線資料無法輸出");return;}
  const date=dayList[focusIdx];
  _exporting=true;
  const _beBtn=document.getElementById("be"),_ldEl=document.getElementById("ld");
  _beBtn.disabled=true;_ldEl.textContent="輸出中 ...";_ldEl.style.display="block";

  // Compute full day price range (including trade prices)
  let mn=1e9,mx=-1e9;
  day.candles.forEach(c=>{mn=Math.min(mn,c.l);mx=Math.max(mx,c.h);});
  (day.trades||[]).forEach(t=>{
    if(t.isHold||t.crossDay)return;
    if(t.entryPrice!=null){mn=Math.min(mn,t.entryPrice);mx=Math.max(mx,t.entryPrice);}
    if(t.exitPrice!=null){mn=Math.min(mn,t.exitPrice);mx=Math.max(mx,t.exitPrice);}
  });

  // Export uses 9-grid page range (vs App's 7) to better fit A4 paper
  // and reduce both page count and notes-collision pressure
  const step=1; // $1 per grid
  const PAGE_GRIDS=9;
  const pageRange=step*PAGE_GRIDS;
  const TOTAL_BARS=BPD; // 391 — export shows ONLY the focus day; no prev-day tail

  // Previous trading day: still load it so we can stamp the prev-close label on
  // the left Y axis (ISO-week-first-day only), but we do NOT draw its candles.
  const prevDate=await loadTradingDay(date,-1);
  const isWeekFirst=await isFirstTradingDayOfWeek(date);
  let prevCloseInfo=null; // {price, mmdd} or null — set only if ISO-week-first
  if(isWeekFirst&&prevDate){
    const pd=dayCache.get(prevDate);
    if(pd&&pd.candles&&pd.candles.length){
      const lastBar=pd.candles[pd.candles.length-1];
      // Format date as M/D (no leading zero, no year)
      const [yy,mm,dd]=prevDate.split("-");
      const mmdd=`${parseInt(mm,10)}/${parseInt(dd,10)}`;
      prevCloseInfo={price:lastBar.c,mmdd:mmdd};
    }
  }
  const dayRange=mx-mn;

  // Export day list: focus day only, no prev day visible
  const exportDayList=[date];

  // Save current state (including canvas size)
  const savedDayList=dayList,savedFocusIdx=focusIdx,savedZoom=zoom,savedPanX=panX,savedPanY=panY,savedUserYZoom=userYZoom;
  const savedCvW=cv.width,savedCvH=cv.height,savedStyleW=cv.style.width,savedStyleH=cv.style.height;
  const savedM={...M};

  // Everything below mutates global view state (dayList, M, canvas size, export
  // flags). Wrapped in try/finally so a throw mid-render can't strand the app in
  // export mode — that state is unrecoverable without a page reload.
  try{
  dayList=exportDayList;
  focusIdx=0;

  // ── Set up high-resolution export canvas ──
  // Symmetric layout: ISO-week-first day shows the LEFT Y axis only (with the
  // prev-close label), other days show the RIGHT Y axis only. Margins and edge
  // paddings mirror each other accordingly.
  const EXPORT_W=1400;
  const EXPORT_DPR=3;
  M.t=44; // top margin hosts the color legend (at y=M.t/2=22)
  M.b=64;
  if(isWeekFirst){
    M.l=30; M.r=6;  // left Y axis hugs canvas left edge
  } else {
    M.l=6;  M.r=30; // right Y axis hugs canvas right edge
  }
  const EDGE_PAD_LEFT =8;
  const EDGE_PAD_RIGHT=8;
  const innerW=EXPORT_W-M.l-M.r;
  const kLineW=innerW-EDGE_PAD_LEFT-EDGE_PAD_RIGHT;
  const bwTarget=kLineW/TOTAL_BARS;
  const REF_TOTAL_BARS=BPD+60;
  const refBw=innerW/REF_TOTAL_BARS;
  const innerH=PAGE_GRIDS*30*refBw;
  const EXPORT_H=Math.round(innerH+M.t+M.b);
  cv.width=EXPORT_W*EXPORT_DPR;cv.height=EXPORT_H*EXPORT_DPR;
  cv.style.width=EXPORT_W+"px";cv.style.height=EXPORT_H+"px";
  ctx.setTransform(EXPORT_DPR,0,0,EXPORT_DPR,0,0);
  W=EXPORT_W;H=EXPORT_H;cW=W-M.l-M.r;cH=H-M.t-M.b;

  exportMode=true;
  exportLightMode=true;
  exportDualAxis=false;
  exportYAxisLeft=isWeekFirst;
  userYZoom=1;
  exportPrevClose=prevCloseInfo;
  exportLabelCache=null; // fresh computation for this export run

  // Compute notesAnchorPrice using App's default 7-grid range
  exportYOverride=null;
  const{pn:defPn,px:defPx}=autoY();
  const _curDate=dayList[focusIdx];
  const _noteY=(_curDate&&notesStateByDate[_curDate])?notesStateByDate[_curDate].y:0.1;
  notesAnchorPrice=defPx-_noteY*(defPx-defPn);
  const notesInUpperHalf=_noteY<0.5;
  const notesIsCentered=_noteY>=0.4&&_noteY<=0.6;

  // Compute zoom: ppb = cW/(BPD-1) * zoom = bwTarget → zoom = bwTarget*(BPD-1)/cW
  zoom=bwTarget*(BPD-1)/cW;

  // Position panX so the focus day's first candle sits EDGE_PAD_LEFT in from M.l
  panX=EDGE_PAD_LEFT;
  panY=0;

  resizeFreeze=true;

  async function capturePage(filename){
    draw();
    const dataURL=cv.toDataURL("image/png");
    const a=document.createElement("a");
    a.download=filename;a.href=dataURL;a.click();
    await new Promise(r=>setTimeout(r,250));
  }

  // Single-page case: dayRange fits in 9 grids
  if(dayRange<=pageRange-step*0.5){
    exportYOverride=null;
    notesPageIdx=null;
    currentExportPageIdx=0;
    // Build 9-grid range centered on day's mid, with asymmetric bias toward notes if needed
    const rawMid=(mn+mx)/2;
    const mid=Math.round(rawMid/step-0.5)*step+step/2; // snap mid to half-step
    let pn0=mid-pageRange/2;
    let px0=mid+pageRange/2;
    // If notes is offset (not centered), shift the extra 2 grids toward the notes side
    // by snapping pn0/px0 to integer dollar boundaries that favor the notes side
    if(!notesIsCentered){
      // Try to fit mn..mx within the range, but bias the extra space toward notes
      const gridTop=Math.ceil(mx);
      const gridBot=Math.floor(mn);
      const usedRange=gridTop-gridBot; // K-line range, integer
      const slack=pageRange-usedRange; // extra grids to distribute
      if(slack>0){
        let extraTop,extraBot;
        if(notesInUpperHalf){
          // Notes on top → put more space above K lines
          extraTop=Math.ceil(slack*0.75);
          extraBot=slack-extraTop;
        } else {
          // Notes on bottom → put more space below K lines
          extraBot=Math.ceil(slack*0.75);
          extraTop=slack-extraBot;
        }
        px0=gridTop+extraTop;
        pn0=gridBot-extraBot;
      }
    } else {
      // Symmetric: snap to integer boundaries with K lines centered
      const gridTop=Math.ceil(mx);
      const gridBot=Math.floor(mn);
      const usedRange=gridTop-gridBot;
      const slack=pageRange-usedRange;
      if(slack>0){
        const extraBot=Math.floor(slack/2);
        const extraTop=slack-extraBot;
        px0=gridTop+extraTop;
        pn0=gridBot-extraBot;
      }
    }
    exportYOverride={pn:pn0,px:px0};
    await capturePage(`SPY_${date.replace(/-/g,"")}_review.png`);
  } else {
    const topEdge=mx===Math.floor(mx)?mx+1:Math.ceil(mx);
    const bottomEdge=mn===Math.floor(mn)?mn-1:Math.floor(mn);
    const totalRange=topEdge-bottomEdge;
    const numPages=Math.ceil(totalRange/pageRange);

    // Compute each page's pageTop:
    //   page 0           → topEdge
    //   page numPages-1  → bottomEdge + pageRange
    //   middle pages     → linearly interpolated
    const pageTops=[];
    if(numPages===1){
      pageTops.push(topEdge);
    } else {
      const firstTop=topEdge;
      const lastTop=bottomEdge+pageRange;
      for(let page=0;page<numPages;page++){
        const t=page/(numPages-1);
        const pt=firstTop-t*(firstTop-lastTop);
        pageTops.push(Math.round(pt));
      }
    }
    exportPageTops=pageTops;

    // Pick the page whose center is closest to notesAnchorPrice → notes will only render there
    if(notesAnchorPrice!=null){
      let bestPage=0,bestDist=Infinity;
      for(let page=0;page<numPages;page++){
        const center=pageTops[page]-pageRange/2;
        const dist=Math.abs(center-notesAnchorPrice);
        if(dist<bestDist){bestDist=dist;bestPage=page;}
      }
      notesPageIdx=bestPage;
    } else {
      notesPageIdx=null;
    }

    for(let page=0;page<numPages;page++){
      currentExportPageIdx=page;
      const pageTop=pageTops[page];
      const pageBottom=pageTop-pageRange;
      exportYOverride={pn:pageBottom,px:pageTop};
      await capturePage(`SPY_${date.replace(/-/g,"")}_review_${page+1}of${numPages}.png`);
    }
  }

  } finally {
    // Restore resize function
    resizeFreeze=false;

    // Restore canvas state
    cv.width=savedCvW;cv.height=savedCvH;
    cv.style.width=savedStyleW;cv.style.height=savedStyleH;
    M=savedM;

    // Restore export flags & view state
    exportMode=false;exportLightMode=false;exportDualAxis=false;exportYAxisLeft=false;exportYOverride=null;notesAnchorPrice=null;notesPageIdx=null;currentExportPageIdx=null;exportPageTops=null;exportPrevClose=null;exportLabelCache=null;
    dayList=savedDayList;focusIdx=savedFocusIdx;zoom=savedZoom;panX=savedPanX;panY=savedPanY;userYZoom=savedUserYZoom;
    _exporting=false;
    _beBtn.disabled=false;_ldEl.style.display="none";_ldEl.textContent="載入中 ...";
    draw();
  }
}
document.getElementById("be").addEventListener("click",exportPNG);

// ═══════════════════════════════════════════════════
// Data loading & Navigation
// ═══════════════════════════════════════════════════
async function loadDay(ds){
  if(dayCache.has(ds))return;
  const res=await fetch(`/api/data?date=${ds}`,{cache:"no-store"});const data=await res.json();
  dayCache.set(ds,{candles:data.candles||[],trades:data.trades||[],notes:data.notes||"",noTrading:!!data.noTrading});
}

// Fit view: focus day fills ~95% of chart width with ~10 bars buffer each side
function fitView(){
  zoom=0.95;
  panY=0;
  userYZoom=1;
  // Center focus day in view
  const focusCenterG=focusIdx*(BPD+GAP)+(BPD-1)/2;
  panX=W/2-M.l-focusCenterG*ppb();
}

async function jumpTo(ds){
  document.getElementById("ld").style.display="block";
  try{
    await loadDay(ds);
    const day=dayCache.get(ds);
    if(!day||!day.candles||!day.candles.length){
      const asked=ds;
      const found=await loadTradingDay(ds,-1);
      if(found){
        ds=found;
        // 不要靜默改跳。使用者明確輸入/點選了某一天，卻看到別天的圖，
        // 沒有提示的話只會以為是自己看錯或程式壞了。
        if(typeof showToast==="function")showToast(`${asked} 無 K 線資料，已改顯示 ${found}`);
      }
      else{document.getElementById("ld").style.display="none";
        if(typeof showToast==="function")showToast(`${asked} 無 K 線資料`);return;}
    }
    // Reset and load neighbors first
    dayList=[ds];focusIdx=0;
    const prevDate=await loadTradingDay(ds,-1);
    const nextDate=await loadTradingDay(ds,1);
    const newList=[];
    if(prevDate)newList.push(prevDate);
    newList.push(ds);
    if(nextDate)newList.push(nextDate);
    dayList=newList;
    focusIdx=prevDate?1:0;
    // Resize first to get correct W/H, then fit
    resize();fitView();
    updateToolbar();draw();
    if(typeof updateLegend==="function")updateLegend();
  }catch(e){alert("載入失敗："+e.message);}
  finally{document.getElementById("ld").style.display="none";}
}

// Get latest US trading day (ET timezone aware)
function getLatestUSTradeDate(){
  // Get current date/time in US/Eastern
  const now=new Date();
  const parts=new Intl.DateTimeFormat("en-CA",{timeZone:"America/New_York",year:"numeric",month:"2-digit",day:"2-digit",hour:"numeric",minute:"numeric",hour12:false}).formatToParts(now);
  const get=t=>parts.find(p=>p.type===t).value;
  let y=+get("year"),m=+get("month"),day=+get("day"),h=+get("hour"),min=+get("minute");
  // Build a local date object for the ET date
  let d=new Date(y,m-1,day);
  // If market hasn't closed yet (before 16:00 ET) or before open, use previous trading day's completed session
  // After 16:00 ET = today's session is done, use today
  // Before 16:00 ET = today's session not complete, use previous completed session
  if(h<16) d.setDate(d.getDate()-1);
  // Skip weekends
  while(d.getDay()===0||d.getDay()===6) d.setDate(d.getDate()-1);
  // Format as YYYY-MM-DD
  const yy=d.getFullYear(),mm=String(d.getMonth()+1).padStart(2,"0"),dd=String(d.getDate()).padStart(2,"0");
  return `${yy}-${mm}-${dd}`;
}

async function loadColors(){
  try{
    const r=await fetch("/api/colors");
    window.colorVersions=await r.json();
    const sel=document.getElementById("cvsel");
    sel.innerHTML="";
    window.colorVersions.forEach((v,i)=>{
      const o=document.createElement("option");
      o.value=i;o.textContent="版本 "+(v.Version||(i+1));
      sel.appendChild(o);
    });
    window.selectedColorIdx=0;
    sel.addEventListener("change",()=>{
      window.selectedColorIdx=parseInt(sel.value,10);
      updateLegend();
      draw();
    });
    updateLegend();
  }catch(e){console.error("loadColors",e);}
}

function updateLegend(){
  const cv=(window.colorVersions&&window.colorVersions[window.selectedColorIdx||0]);
  if(!cv)return;
  const day=dayCache.get(dayList[focusIdx]);
  const trades=(day&&day.trades)||[];
  const hasHold=trades.some(t=>t.isHold);
  const hasClose=trades.some(t=>t.crossDay&&!t.isHold);
  const parts=[
    `<span><span class="sw" style="background:${cv.App_Long}"></span>多</span>`,
    `<span><span class="sw" style="background:${cv.App_Short}"></span>空</span>`,
  ];
  if(hasHold&&hasClose){
    parts.push(`<span><span class="sw" style="background:${cv.App_Hold}"></span>留倉/平倉</span>`);
  } else if(hasHold){
    parts.push(`<span><span class="sw" style="background:${cv.App_Hold}"></span>留倉</span>`);
  } else if(hasClose){
    parts.push(`<span><span class="sw" style="background:${cv.App_Hold}"></span>平倉</span>`);
  }
  document.getElementById("legend").innerHTML=parts.join(" ");
}

async function init(){
  await loadColors();
  const res=await fetch("/api/dates",{cache:"no-store"});const dj=await res.json();
  const dates=dj.dates||[];
  allDates=dates;                              // 日曆用來判斷哪些日子可點
  allDateSet=new Set(dates);
  noCandleDates=new Set(dj.no_candles||[]);    // 有交易但沒有 K 線，日曆上標成褐色
  if(dates.length){await jumpTo(dates[0]);}
  else{const today=getLatestUSTradeDate();await jumpTo(today);}
  seedDataVersion();
}

// ── Auto-reload when trades_all.xlsx is edited externally (no app restart needed) ──
function showToast(msg){
  let t=document.getElementById("_toast");
  if(!t){t=document.createElement("div");t.id="_toast";
    t.style.cssText="position:fixed;bottom:16px;left:50%;transform:translateX(-50%);background:rgba(20,20,20,.88);color:#fff;padding:8px 16px;border-radius:6px;font-size:14px;z-index:99999;transition:opacity .3s;pointer-events:none";
    document.body.appendChild(t);}
  t.textContent=msg;t.style.opacity="1";
  clearTimeout(t._h);t._h=setTimeout(()=>{t.style.opacity="0";},1700);
}
// Re-fetch the currently loaded days without resetting pan/zoom/focus.
// Fetch into a staging map FIRST, then swap in one synchronous step: the old code
// cleared dayCache up front, so every draw() during the refetch window (drag, edge
// load, the 4s poll itself) rendered an empty chart.
async function refreshData(){
  const dates=[...dayList];if(!dates.length)return;
  const fresh=new Map();
  for(const ds of dates){
    try{
      const res=await fetch(`/api/data?date=${ds}`,{cache:"no-store"});const d=await res.json();
      fresh.set(ds,{candles:d.candles||[],trades:d.trades||[],notes:d.notes||"",noTrading:!!d.noTrading});
    }catch(e){const old=dayCache.get(ds);if(old)fresh.set(ds,old);}
  }
  dayCache.clear();                       // drop stale off-screen days too
  for(const[k,v]of fresh)dayCache.set(k,v); // ...but never leave dayList uncached
  if(typeof updateToolbar==="function")updateToolbar();
  draw();
  if(typeof updateLegend==="function")updateLegend();
}
let _dataVer=null,_verBusy=false;
async function seedDataVersion(){
  try{const r=await fetch("/api/version",{cache:"no-store"});const j=await r.json();
    _dataVer=j.trades;renderSyncBadge(j.sync);}catch(e){}
}
async function checkDataVersion(){
  if(_verBusy||_exporting||exportMode)return;  // never yank data out from under an export
  _verBusy=true;
  try{
    const r=await fetch("/api/version",{cache:"no-store"});const j=await r.json();const v=j.trades;
    renderSyncBadge(j.sync);
    if(_dataVer===null){_dataVer=v;}
    else if(v!==_dataVer){_dataVer=v;await refreshData();showToast("資料已更新");}
  }catch(e){}finally{_verBusy=false;}
}

// 同步新鮮度：Boss PC 每次跑完寫 sync_state.json，隨 git 一起同步過來。
// 少了這個，push 失敗與「今天還沒開盤」在畫面上長得一模一樣。
// 落後幾個「美股交易日」——刻意不用「距上次更新幾小時」。
// Boss PC 週末關機，週五的帳要等週一較晚開機才補；用時數判斷的話
// 每個週末都會固定跨過 30 小時而變紅。每週都在誤報的告警等於沒有告警。
function tradingDaysBehind(lastDate){
  if(!lastDate)return 99;
  const latest=getLatestUSTradeDate();          // 已含週末回溯
  if(lastDate>=latest)return 0;
  let n=0,d=new Date(latest+"T12:00:00");
  const stop=new Date(lastDate+"T12:00:00");
  while(d>stop&&n<99){
    d.setDate(d.getDate()-1);
    if(d.getDay()!==0&&d.getDay()!==6)n++;      // 只數平日
  }
  return n;
}

function renderSyncBadge(s){
  const el=document.getElementById("syncbadge");
  if(!el)return;
  if(!s||!s.updated_utc){el.style.display="none";return;}
  el.style.display="";
  const behind=tradingDaysBehind(s.last_trade_date);
  let color="#6B7280",txt=`資料截至 ${s.last_trade_date||"?"}`;
  if(s.status==="blocked"){color="#E0A800";txt+=" ⚠需人工確認";}
  else if(s.status==="failed"){color="#FF4444";txt+=" ✕採集失敗";}
  // 落後 1 個交易日是正常待處理（例：週一早上，週五的帳還沒補）
  else if(behind>=2){color="#FF4444";txt+=` ✕落後 ${behind} 個交易日`;}
  else if(behind===1){color="#E0A800";txt+=" ⚠待補前一交易日";}
  el.style.color=color;
  el.textContent=txt;
  const ageH=Math.floor((Date.now()-Date.parse(s.updated_utc))/3600000);
  el.title=`status=${s.status||"?"}  host=${s.host||"?"}\n`
           +`最新美股交易日=${getLatestUSTradeDate()}  落後=${behind}\n`
           +`updated=${s.updated_utc}（${ageH}h 前）`
           +(s.note?`\n${s.note}`:"");
}
setInterval(checkDataVersion,4000);                 // poll every 4s
window.addEventListener("focus",checkDataVersion);  // and instantly when returning to the tab
document.addEventListener("visibilitychange",()=>{if(!document.hidden)checkDataVersion();});

document.getElementById("bp").addEventListener("click",async()=>{
  const cur=dayList[focusIdx]||dayList[0];
  const prev=await loadTradingDay(cur,-1);
  if(prev)jumpTo(prev);
});
document.getElementById("bn").addEventListener("click",async()=>{
  const cur=dayList[focusIdx]||dayList[0];
  const next=await loadTradingDay(cur,1);
  if(next)jumpTo(next);
});
document.getElementById("bt").addEventListener("click",()=>jumpTo(getLatestUSTradeDate()));

// ── 日期選擇器：年月日三格 + 迷你日曆 ────────────────────────────────
const dY=document.getElementById("dY"),dM=document.getElementById("dM"),dD=document.getElementById("dD"),
      dbox=document.getElementById("dbox"),dtog=document.getElementById("dtog"),dcal=document.getElementById("dcal");
const dsegs=[dY,dM,dD];
let allDates=[],allDateSet=new Set(),noCandleDates=new Set();
let calY=null,calM=null;   // 日曆目前顯示的年/月（1-12）

const pad=(n,w)=>String(n).padStart(w,"0");
function dateBoxFocused(){return dsegs.indexOf(document.activeElement)>=0;}
function setDateBoxes(ds){
  if(!ds||!/^\d{4}-\d{2}-\d{2}$/.test(ds)){dY.value="";dM.value="";dD.value="";return;}
  dY.value=ds.slice(0,4);dM.value=ds.slice(5,7);dD.value=ds.slice(8,10);
}
function readDateBoxes(){
  const y=dY.value.trim(),m=dM.value.trim(),d=dD.value.trim();
  if(!/^\d{4}$/.test(y)||!/^\d{1,2}$/.test(m)||!/^\d{1,2}$/.test(d))return null;
  const mm=+m,dd=+d;
  if(mm<1||mm>12||dd<1||dd>31)return null;
  return `${y}-${pad(mm,2)}-${pad(dd,2)}`;
}

// ── 迷你日曆 ──
function calRender(){
  const cur=dayList[focusIdx]||"";
  if(calY===null){const b=(cur||getLatestUSTradeDate());calY=+b.slice(0,4);calM=+b.slice(5,7);}
  const first=new Date(calY,calM-1,1), lead=first.getDay(),
        ndays=new Date(calY,calM,0).getDate(),
        todayStr=getLatestUSTradeDate();
  let cells="",dayN=1;
  for(let r=0;r<6;r++){
    let row="";
    for(let c=0;c<7;c++){
      if((r===0&&c<lead)||dayN>ndays){row+="<td></td>";continue;}
      const ds=`${calY}-${pad(calM,2)}-${pad(dayN,2)}`;
      const cls=[];
      if(noCandleDates.has(ds))cls.push("nok");
      else if(allDateSet.has(ds))cls.push("has");
      if(ds===cur)cls.push("cur");
      if(ds===todayStr)cls.push("today");
      row+=`<td class="${cls.join(" ")}" data-d="${ds}">${dayN}</td>`;
      dayN++;
    }
    cells+=`<tr>${row}</tr>`;
    if(dayN>ndays)break;
  }
  dcal.innerHTML=
    `<div class="cal-hd"><button class="cal-nav" data-mv="-1">&#8249;</button>`
   +`<span class="ttl">${calY} 年 ${calM} 月</span>`
   +`<button class="cal-nav" data-mv="1">&#8250;</button></div>`
   +`<table><tr><th class="we">日</th><th>一</th><th>二</th><th>三</th><th>四</th><th>五</th><th class="we">六</th></tr>`
   +cells+`</table>`
   +`<div class="cal-ft"><b>亮＝可檢視</b>　<span style="color:#FF8C00">橘底＝目前</span>`
   +`　<span style="color:#C08A3E">褐＝有交易無 K 線</span>　暗＝休市</div>`;
}
function calOpen(){
  const cur=dayList[focusIdx]||"";
  if(cur){calY=+cur.slice(0,4);calM=+cur.slice(5,7);}
  calRender();dcal.classList.add("show");
}
function calClose(){dcal.classList.remove("show");}

dtog.addEventListener("click",e=>{e.stopPropagation();
  if(dcal.classList.contains("show"))calClose();else calOpen();});
dcal.addEventListener("mousedown",e=>{     // mousedown 早於 blur，才點得到
  const nav=e.target.closest(".cal-nav");
  if(nav){e.preventDefault();calM+= +nav.dataset.mv;
    if(calM<1){calM=12;calY--;}else if(calM>12){calM=1;calY++;}
    calRender();return;}
  const td=e.target.closest("td[data-d]");
  if(!td||(!td.classList.contains("has")&&!td.classList.contains("nok")))return;  // 非交易日不可點
  e.preventDefault();calClose();jumpTo(td.dataset.d);
});
document.addEventListener("click",e=>{if(!e.target.closest("#dwrap"))calClose();});

// ── 三格輸入 ──
dsegs.forEach((el,idx)=>{
  el.addEventListener("focus",()=>{el.select();dbox.classList.add("focus");});
  el.addEventListener("blur",()=>setTimeout(()=>{if(!dateBoxFocused())dbox.classList.remove("focus");},0));
  el.addEventListener("input",()=>{
    el.value=el.value.replace(/\D/g,"");
    // 填滿就自動跳下一格，這是分段日期輸入的基本手感
    if(el.value.length>=el.maxLength&&idx<2)dsegs[idx+1].focus();
  });
  el.addEventListener("keydown",e=>{
    // ↑↓：前後交易日（跟 ←/→ 按鈕同語意，會自動跳過休市日）
    if(e.key==="ArrowUp"||e.key==="ArrowDown"){
      e.preventDefault();
      const cur=dayList[focusIdx]||dayList[0];
      if(!cur)return;
      loadTradingDay(cur,e.key==="ArrowUp"?1:-1).then(d=>{if(d)jumpTo(d);});
      return;
    }
    if(e.key==="Enter"){
      e.preventDefault();
      const v=readDateBoxes();
      if(v){calClose();el.blur();jumpTo(v);}
      else{showToast("日期格式不完整");setDateBoxes(dayList[focusIdx]||"");}
      return;
    }
    if(e.key==="Escape"){calClose();setDateBoxes(dayList[focusIdx]||"");el.blur();return;}
    // 左右鍵在格子邊界時換格
    if(e.key==="ArrowLeft"&&idx>0&&el.selectionStart===0){e.preventDefault();dsegs[idx-1].focus();}
    if(e.key==="ArrowRight"&&idx<2&&el.selectionStart===el.value.length){e.preventDefault();dsegs[idx+1].focus();}
    if(e.key==="Backspace"&&idx>0&&!el.value){e.preventDefault();dsegs[idx-1].focus();}
  });
});

document.getElementById("ns").addEventListener("click",()=>{
  const day=dayCache.get(dayList[focusIdx]);if(day)day.notes=document.getElementById("nt").value;
  document.getElementById("ned").classList.remove("show");uNotes();
  fetch("/api/notes",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({date:dayList[focusIdx],notes:document.getElementById("nt").value})});
});
document.getElementById("ned").addEventListener("click",e=>{if(e.target===document.getElementById("ned"))document.getElementById("ned").classList.remove("show");});

document.addEventListener("keydown",e=>{if(document.getElementById("ned").classList.contains("show")||dateBoxFocused())return;
if(e.key==="ArrowLeft")document.getElementById("bp").click();
if(e.key==="ArrowRight")document.getElementById("bn").click();
if(e.key==="+"||e.key==="=")doZoom(ZS);if(e.key==="-")doZoom(-ZS);});

window.addEventListener("resize",draw);
// URL hash 跳到指定日（jumpTo 已內建非交易日 fallback，不會報錯）
function gotoDateFromHash(){
  const h=(location.hash||"").replace(/^#/,"").trim();
  if(/^\d{4}-\d{2}-\d{2}$/.test(h)) jumpTo(h);
}
window.addEventListener("hashchange", gotoDateFromHash);
initIndButtons();
init().then(gotoDateFromHash);
</script></body></html>"""

if __name__=="__main__":
    print(f"\n  Trade Review Web App v3\n  Root:    {ROOT_FOLDER}\n  Trades:  {TRADES_FILE}\n  History: {HISTORY_FILE}\n  Notes:   {NOTES_FOLDER}\n  http://localhost:{PORT}\n")
    if not os.path.exists(ROOT_FOLDER):os.makedirs(ROOT_FOLDER,exist_ok=True)
    if not os.path.exists(NOTES_FOLDER):os.makedirs(NOTES_FOLDER,exist_ok=True)
    # 以下兩段只是「暖身快取」，讓第一個請求快一點。它們絕不能擋住伺服器啟動：
    # 這個 app 的核心價值是「隨時打得開、看得到過往所有 K 線」，而 K 線來自
    # history_minute.xlsx，跟帳本完全無關。就算 trades_all 壞掉、還沒寫入當日
    # 交易、或分析邏輯出例外，圖表都還是該看得到。
    # （pythonw 沒有主控台，這裡若拋例外只會靜默不啟動 —— 最難察覺的失敗。）
    if os.path.exists(HISTORY_FILE):
        print("  Loading history file (may take a moment for large files)...")
        try:
            hist = _load_history_by_date()
            print(f"  [OK] {len(hist)} trading days ready")
        except Exception as e:
            print(f"  ⚠ history 預載失敗（不影響啟動，改為每次請求時再讀）：{type(e).__name__}: {e}")
    print("  Analysing trades...")
    try:
        _analyse_all_trades()
        print("  [OK] Trade analysis complete")
    except Exception as e:
        print(f"  ⚠ 交易分析預載失敗（不影響啟動，K 線仍可正常瀏覽）：{type(e).__name__}: {e}")
    if not os.path.exists(TRADES_FILE):
        print(f"  ⚠ {TRADES_FILE} not found. Create it with columns:")
        print(f"     Date | Exec Time(EDT) | Symbol | Price | Type | 損益(AI辨識) | Shares")
    try:
        app.run(host="0.0.0.0",port=PORT,debug=False)
    except OSError as e:
        # 幾乎都是 port 被占用。pythonw 下這行會落在 _logs\trade_review_app.log，
        # 是排查「排程說在跑、網頁卻打不開」的唯一線索。
        print(f"  🔴 無法在 port {PORT} 啟動：{e}")
        print(f"     檢查誰占用了：Get-NetTCPConnection -LocalPort {PORT} -State Listen")
        raise
