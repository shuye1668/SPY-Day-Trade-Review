"""
Daily Fetch — yfinance edition
==============================
Pulls SPY 1-minute candles from Yahoo Finance and saves them as JSON
into D:\\fileserver_D\\TradeReview\\candles\\

Behavior:
  • Default mode: fetch today's session only (after market close)
  • --catchup: also fetch any weekday in the past 7 days that's missing locally
              (yfinance only allows 1m bars within last ~30 days; 7-day windows per call)
  • --days N: catchup window in days (default 7, max 30)
  • --force: overwrite even if file exists

Setup:
  pip install yfinance pandas

Schedule (Windows Task Scheduler):
  Run daily at 16:30 ET → maps to 04:30 (winter) / 05:30 (summer) Taipei time
  Action: python  D:\\fileserver_D\\TradeReview\\daily_fetch.py --catchup
  ALSO: trigger "At log on" with same command, so missed days are recovered when
        you turn the PC back on.
"""
import datetime as dt
import os, json, sys, argparse, time

# ── Config ──────────────────────────────────────────────────────────────────
_BASE = os.path.dirname(os.path.abspath(__file__))
CANDLES_FOLDER = os.path.join(_BASE, "candles")
TICKER = "SPY"
# ────────────────────────────────────────────────────────────────────────────


def is_weekday(d):
    return d.weekday() < 5


def get_today_et():
    """Today's date in US/Eastern timezone."""
    import pandas as pd
    now_utc = pd.Timestamp.utcnow().tz_convert("US/Eastern")
    return now_utc.date()


def fetch_yf_intraday(ticker, date_str):
    """Fetch 1-minute candles for a single date using yfinance.
    Returns list of {t, o, h, l, c, v} dicts (t = ET HH:MM), or [] on holiday/error.
    """
    import yfinance as yf
    import pandas as pd

    d = dt.datetime.strptime(date_str, "%Y-%m-%d").date()
    # yfinance: end is exclusive, so add 1 day
    start_str = d.strftime("%Y-%m-%d")
    end_str = (d + dt.timedelta(days=1)).strftime("%Y-%m-%d")
    try:
        df = yf.download(
            tickers=ticker,
            start=start_str,
            end=end_str,
            interval="1m",
            prepost=False,         # RTH only
            auto_adjust=False,
            progress=False,
            threads=False,
        )
    except Exception as e:
        print(f"    yfinance error: {e}")
        return []
    if df is None or len(df) == 0:
        return []
    # Some yfinance versions return MultiIndex columns when single ticker — flatten
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)
    # Ensure tz-aware in ET
    if df.index.tz is None:
        df.index = df.index.tz_localize("UTC").tz_convert("US/Eastern")
    else:
        df.index = df.index.tz_convert("US/Eastern")
    # Filter to RTH 09:30-16:00 ET (yfinance sometimes leaks pre/post even with prepost=False)
    df = df.between_time("09:30", "15:59")
    # Filter to the exact target date in ET
    df = df[df.index.date == d]
    if len(df) == 0:
        return []
    bars = []
    for ts, row in df.iterrows():
        try:
            o = float(row["Open"]); h = float(row["High"])
            lo = float(row["Low"]); c = float(row["Close"])
            v = int(row["Volume"]) if not pd.isna(row["Volume"]) else 0
        except Exception:
            continue
        # Skip rows with NaN (yfinance occasionally inserts gaps)
        if any(pd.isna(x) for x in (o, h, lo, c)):
            continue
        bars.append({
            "t": ts.strftime("%H:%M"),
            "o": round(o, 4),
            "h": round(h, 4),
            "l": round(lo, 4),
            "c": round(c, 4),
            "v": v,
        })
    return bars


def fetch_and_save(date_str, force=False):
    """Fetch one date and save to JSON. Returns ('saved'|'skipped'|'empty'|'failed', count)."""
    out_path = os.path.join(CANDLES_FOLDER, f"{TICKER}_{date_str}.json")
    if os.path.exists(out_path) and not force:
        return ("skipped", 0)
    try:
        bars = fetch_yf_intraday(TICKER, date_str)
    except Exception as e:
        print(f"  ✗ {date_str}  exception: {e}")
        return ("failed", 0)
    if not bars:
        return ("empty", 0)
    try:
        os.makedirs(CANDLES_FOLDER, exist_ok=True)
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(bars, f, separators=(",", ":"))
        return ("saved", len(bars))
    except Exception as e:
        print(f"  ✗ {date_str}  write failed: {e}")
        return ("failed", 0)


def main():
    parser = argparse.ArgumentParser(description="Daily SPY fetch via yfinance")
    parser.add_argument("--catchup", action="store_true",
                        help="Also fetch missing days in the past N days (see --days)")
    parser.add_argument("--days", type=int, default=7,
                        help="Catchup window in days (default 7, max 30)")
    parser.add_argument("--force", action="store_true",
                        help="Overwrite existing JSON files")
    parser.add_argument("--date", type=str, default=None,
                        help="Fetch a specific date YYYY-MM-DD instead of today")
    args = parser.parse_args()

    # Verify yfinance is installed
    try:
        import yfinance  # noqa: F401
    except ImportError:
        print("  ✗ yfinance not installed.  Run: pip install yfinance")
        sys.exit(1)

    today_et = get_today_et()
    print(f"\n  Daily Fetch — {dt.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}  (ET today: {today_et})")
    print(f"  Output: {CANDLES_FOLDER}")

    targets = []
    if args.date:
        targets = [args.date]
    else:
        # Always include today
        if is_weekday(today_et):
            targets.append(today_et.strftime("%Y-%m-%d"))
        # Catchup: walk back through past N days, add missing weekdays
        if args.catchup:
            window = max(1, min(30, args.days))
            for i in range(1, window + 1):
                d = today_et - dt.timedelta(days=i)
                if not is_weekday(d):
                    continue
                ds = d.strftime("%Y-%m-%d")
                out_path = os.path.join(CANDLES_FOLDER, f"{TICKER}_{ds}.json")
                if not os.path.exists(out_path):
                    targets.append(ds)
            # Order chronologically (oldest first) so logs read naturally
            targets = sorted(set(targets))

    if not targets:
        print("  Nothing to fetch.\n")
        return

    print(f"  Targets: {len(targets)} date(s)\n")

    fetched = skipped = empty = failed = 0
    for ds in targets:
        status, count = fetch_and_save(ds, force=args.force)
        if status == "saved":
            print(f"  ✓ {ds}  saved {count} bars")
            fetched += 1
        elif status == "skipped":
            print(f"  - {ds}  already exists")
            skipped += 1
        elif status == "empty":
            print(f"  · {ds}  no data (holiday or before market close)")
            empty += 1
        else:
            failed += 1
        # Be kind to Yahoo's rate limiter
        time.sleep(0.5)

    print(f"\n  Done. Fetched: {fetched}  Skipped: {skipped}  Empty: {empty}  Failed: {failed}\n")


if __name__ == "__main__":
    main()
