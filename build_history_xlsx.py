"""One-time builder for C:\\TradeReview\\history_minute.xlsx

Merges:
  1. Existing history_minute.xlsx (user's own records, e.g. 2025/4/29 - 2026/3/18)
  2. Bloomberg API (fills any gaps: missing dates up to today)

After running this once (with Bloomberg Terminal running), history_minute.xlsx
becomes the single source of truth. The main app (trade_review_app.py) will then
read from this file exclusively and use yfinance for any future new days.

Usage (run once, Bloomberg Terminal must be open):
    python build_history_xlsx.py

Or to fill a specific date range:
    python build_history_xlsx.py --start 2026-03-19 --end 2026-04-16

Requirements:
    pip install pandas openpyxl
    pip install --index-url=https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi
"""
import os, sys, argparse, datetime as dt
import pandas as pd

HISTORY_FILE = r"C:\TradeReview\history_minute.xlsx"
TICKER = "SPY"


def fetch_bloomberg_day(ticker, date_str):
    """Fetch 1-minute bars for a single date from Bloomberg. Returns rows list (dicts)."""
    import blpapi
    so = blpapi.SessionOptions()
    so.setServerHost("localhost")
    so.setServerPort(8194)
    sess = blpapi.Session(so)
    if not sess.start():
        raise ConnectionError("Bloomberg Terminal 未連線")
    if not sess.openService("//blp/refdata"):
        sess.stop()
        raise ConnectionError("無法開啟 //blp/refdata")
    svc = sess.getService("//blp/refdata")
    req = svc.createRequest("IntradayBarRequest")
    req.set("security", f"{ticker} US Equity")
    req.set("eventType", "TRADE")
    req.set("interval", 1)
    d = dt.datetime.strptime(date_str, "%Y-%m-%d")
    req.set("startDateTime", dt.datetime(d.year, d.month, d.day, 13, 30))
    req.set("endDateTime", dt.datetime(d.year, d.month, d.day, 20, 0))
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
                        "_raw_t": str(b.getElementAsDatetime("time")),
                        "Open": float(b.getElementAsFloat("open")),
                        "High": float(b.getElementAsFloat("high")),
                        "Low": float(b.getElementAsFloat("low")),
                        "Close": float(b.getElementAsFloat("close")),
                        "Volume": int(b.getElementAsInteger("volume")),
                    })
        if ev.eventType() == blpapi.Event.RESPONSE:
            break
    sess.stop()
    # Convert UTC timestamps to US/Eastern datetime objects
    rows = []
    for bar in bars:
        utc = pd.Timestamp(bar["_raw_t"])
        if utc.tzinfo is None:
            utc = utc.tz_localize("UTC")
        et = utc.tz_convert("US/Eastern").tz_localize(None)  # strip tz for Excel
        rows.append({
            "Date": et,
            "Open": bar["Open"],
            "High": bar["High"],
            "Low": bar["Low"],
            "Close": bar["Close"],
            "Volume": bar["Volume"],
        })
    return rows


def us_trading_days(start_date, end_date):
    """Yield YYYY-MM-DD strings for US trading days (Mon-Fri, skip basic holidays).
    For a thorough holiday list use pandas_market_calendars; here we use a simple weekday filter
    plus a hardcoded list of common US market holidays."""
    holidays = {
        # 2025
        "2025-01-01", "2025-01-20", "2025-02-17", "2025-04-18", "2025-05-26",
        "2025-06-19", "2025-07-04", "2025-09-01", "2025-11-27", "2025-12-25",
        # 2026
        "2026-01-01", "2026-01-19", "2026-02-16", "2026-04-03", "2026-05-25",
        "2026-06-19", "2026-07-03", "2026-09-07", "2026-11-26", "2026-12-25",
    }
    d = start_date
    while d <= end_date:
        if d.weekday() < 5:  # Mon=0..Fri=4
            ds = d.strftime("%Y-%m-%d")
            if ds not in holidays:
                yield ds
        d += dt.timedelta(days=1)


def load_existing():
    """Load existing history_minute.xlsx; return (DataFrame, set-of-dates-present)."""
    if not os.path.exists(HISTORY_FILE):
        return pd.DataFrame(), set()
    df = pd.read_excel(HISTORY_FILE)
    df.columns = [str(c).strip() for c in df.columns]
    if "Date" in df.columns:
        df["Date"] = pd.to_datetime(df["Date"])
        dates = set(df["Date"].dt.strftime("%Y-%m-%d").unique())
    else:
        dates = set()
    return df, dates


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--start", help="Start date YYYY-MM-DD (default: today - 60 days)")
    parser.add_argument("--end", help="End date YYYY-MM-DD (default: today)")
    args = parser.parse_args()

    today = dt.datetime.now().date()
    end_date = dt.datetime.strptime(args.end, "%Y-%m-%d").date() if args.end else today
    start_date = dt.datetime.strptime(args.start, "%Y-%m-%d").date() if args.start else (end_date - dt.timedelta(days=60))

    print(f"Target range: {start_date} → {end_date}")
    print(f"Reading existing file: {HISTORY_FILE}")

    existing_df, existing_dates = load_existing()
    print(f"Existing file covers {len(existing_dates)} trading days")

    # Determine missing trading days in range
    missing = []
    for ds in us_trading_days(start_date, end_date):
        if ds not in existing_dates:
            missing.append(ds)

    if not missing:
        print("✓ No gaps in range. History file is up to date.")
        return

    print(f"Missing {len(missing)} trading days: {missing[:5]}{'...' if len(missing) > 5 else ''}")
    print("Fetching from Bloomberg...")

    all_new_rows = []
    for i, ds in enumerate(missing, 1):
        print(f"  [{i}/{len(missing)}] {ds}", end=" ... ", flush=True)
        try:
            rows = fetch_bloomberg_day(TICKER, ds)
            print(f"{len(rows)} bars")
            all_new_rows.extend(rows)
        except Exception as e:
            print(f"FAILED: {e}")
            continue

    if not all_new_rows:
        print("No new data fetched.")
        return

    new_df = pd.DataFrame(all_new_rows)
    combined = pd.concat([existing_df, new_df], ignore_index=True) if len(existing_df) else new_df
    # Sort by Date and dedupe (in case of overlap)
    combined = combined.sort_values("Date").drop_duplicates(subset=["Date"]).reset_index(drop=True)

    os.makedirs(os.path.dirname(HISTORY_FILE), exist_ok=True)
    combined.to_excel(HISTORY_FILE, index=False)
    print(f"\n✓ Wrote {len(combined)} rows ({combined['Date'].dt.strftime('%Y-%m-%d').nunique()} trading days) to {HISTORY_FILE}")


if __name__ == "__main__":
    main()
