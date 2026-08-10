"""
Historical Backfill — Bloomberg edition
========================================
One-time use: pull SPY 1-minute candles from Bloomberg Terminal for a date range
and save them as JSON files into the local cache directory used by trade_review_app_free.py.

Run this on the Bloomberg Terminal computer ONCE to populate your historical archive.
After that, daily_fetch.py (yfinance-based) will keep adding new days from any computer.

Setup:
  pip install pandas openpyxl
  pip install --index-url=https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi

Usage:
  python historical_backfill.py                  # default: backfill past 365 days
  python historical_backfill.py --days 730       # backfill past 2 years
  python historical_backfill.py --start 2024-01-01 --end 2024-12-31
  python historical_backfill.py --force          # overwrite existing JSON files
"""
import datetime as dt
import os, json, sys, argparse, time

# ── Config ──────────────────────────────────────────────────────────────────
_BASE = os.path.dirname(os.path.abspath(__file__))
CANDLES_FOLDER = os.path.join(_BASE, "candles")
TICKER_BBG = "SPY US Equity"   # Bloomberg ticker
TICKER_FILE = "SPY"            # filename prefix (matches free app & daily_fetch)
# ────────────────────────────────────────────────────────────────────────────


def fetch_bbg_intraday(ticker, date_str):
    """Fetch 1-minute candles for a single date from Bloomberg Desktop API.
    Returns a list of {t, o, h, l, c, v} dicts (t = ET HH:MM)."""
    import blpapi
    so = blpapi.SessionOptions()
    so.setServerHost("localhost")
    so.setServerPort(8194)
    sess = blpapi.Session(so)
    if not sess.start():
        raise ConnectionError("Cannot start Bloomberg session — is the Terminal running?")
    if not sess.openService("//blp/refdata"):
        sess.stop()
        raise ConnectionError("Cannot open //blp/refdata service")
    svc = sess.getService("//blp/refdata")
    req = svc.createRequest("IntradayBarRequest")
    req.set("security", ticker)
    req.set("eventType", "TRADE")
    req.set("interval", 1)
    d = dt.datetime.strptime(date_str, "%Y-%m-%d")
    # ET 09:30-16:00 = UTC 13:30-20:00 (winter; summer is 13:30-20:00 too thanks to DST handled by BBG)
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
                        "t": str(b.getElementAsDatetime("time")),
                        "o": b.getElementAsFloat("open"),
                        "h": b.getElementAsFloat("high"),
                        "l": b.getElementAsFloat("low"),
                        "c": b.getElementAsFloat("close"),
                        "v": b.getElementAsInteger("volume"),
                    })
        if ev.eventType() == blpapi.Event.RESPONSE:
            break
    sess.stop()
    # Convert UTC timestamp string to ET HH:MM
    import pandas as pd
    for bar in bars:
        utc = pd.Timestamp(bar["t"])
        if utc.tzinfo is None:
            utc = utc.tz_localize("UTC")
        bar["t"] = utc.tz_convert("US/Eastern").strftime("%H:%M")
    return bars


def is_weekday(d):
    return d.weekday() < 5  # 0=Mon ... 4=Fri


def main():
    parser = argparse.ArgumentParser(description="One-time Bloomberg historical backfill")
    parser.add_argument("--days", type=int, default=365,
                        help="Number of days back from today (default 365). Ignored if --start is given.")
    parser.add_argument("--start", type=str, default=None, help="Start date YYYY-MM-DD (inclusive)")
    parser.add_argument("--end",   type=str, default=None, help="End date YYYY-MM-DD (inclusive, default today)")
    parser.add_argument("--force", action="store_true", help="Overwrite existing JSON files")
    args = parser.parse_args()

    # Determine date range
    today = dt.date.today()
    if args.start:
        start = dt.datetime.strptime(args.start, "%Y-%m-%d").date()
    else:
        start = today - dt.timedelta(days=args.days)
    if args.end:
        end = dt.datetime.strptime(args.end, "%Y-%m-%d").date()
    else:
        end = today

    if not os.path.exists(CANDLES_FOLDER):
        os.makedirs(CANDLES_FOLDER, exist_ok=True)

    print(f"\n  Bloomberg Historical Backfill")
    print(f"  Ticker:   {TICKER_BBG}")
    print(f"  Range:    {start} → {end}")
    print(f"  Output:   {CANDLES_FOLDER}")
    print(f"  Force:    {args.force}\n")

    # Verify Bloomberg connectivity once
    try:
        import blpapi  # noqa: F401
    except ImportError:
        print("  ✗ blpapi not installed.")
        print("    Run: pip install --index-url=https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi")
        sys.exit(1)

    # Iterate weekdays
    d = start
    fetched, skipped, failed, holidays = 0, 0, 0, 0
    while d <= end:
        if not is_weekday(d):
            d += dt.timedelta(days=1)
            continue
        ds = d.strftime("%Y-%m-%d")
        out_path = os.path.join(CANDLES_FOLDER, f"{TICKER_FILE}_{ds}.json")
        if os.path.exists(out_path) and not args.force:
            print(f"  [skip]  {ds}  (file exists)")
            skipped += 1
            d += dt.timedelta(days=1)
            continue
        try:
            print(f"  [fetch] {ds} ...", end=" ", flush=True)
            bars = fetch_bbg_intraday(TICKER_BBG, ds)
            if not bars:
                print("no data (holiday?)")
                holidays += 1
            else:
                with open(out_path, "w", encoding="utf-8") as f:
                    json.dump(bars, f, separators=(",", ":"))
                print(f"saved {len(bars)} bars")
                fetched += 1
        except Exception as e:
            print(f"FAILED: {e}")
            failed += 1
        # gentle pacing — Bloomberg has request limits
        time.sleep(0.2)
        d += dt.timedelta(days=1)

    print(f"\n  Done. Fetched: {fetched}  Skipped: {skipped}  Holidays: {holidays}  Failed: {failed}\n")


if __name__ == "__main__":
    main()
