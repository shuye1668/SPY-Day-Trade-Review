# Trade Review — FREE Version Setup Guide

A Bloomberg-independent version of the trade review app. SPY 1-minute candle data
is fetched from Yahoo Finance and stored locally in JSON files. Trades come from
a single consolidated Excel file.

---

## 1. Folder Structure

Set up this layout on the computer where you'll review trades:

```
C:\TradeReview\
├── trade_review_app_free.py    ← main web app
├── daily_fetch.py              ← yfinance daily fetch + catchup
├── historical_backfill.py      ← one-time Bloomberg backfill (only on BBG terminal)
├── start_app_free.bat          ← double-click to launch
├── trades_all.xlsx             ← YOUR consolidated trade log (see §3)
├── notes\                      ← per-day notes (auto-created on first save)
│   ├── 2026-04-08.txt
│   └── ...
└── candles\                    ← per-day OHLC cache (auto-populated)
    ├── SPY_2026-04-08.json
    └── ...
```

The first run of `trade_review_app_free.py` or `daily_fetch.py` will create
`notes\` and `candles\` automatically if they don't exist.

---

## 2. Install Dependencies

```cmd
pip install flask pandas numpy openpyxl yfinance
```

(For the one-time historical backfill on the Bloomberg Terminal computer, also
install `blpapi`:)

```cmd
pip install --index-url=https://blpapi.bloomberg.com/repository/releases/python/simple/ blpapi
```

---

## 3. Excel Trades File Format

Create `C:\TradeReview\trades_all.xlsx` with these columns (in this order):

| Date       | Exec Time(EDT) | Symbol | Price   | Type | 損益(AI辨識) | Shares |
|------------|----------------|--------|---------|------|--------------|--------|
| 2026-04-08 | 09:35          | SPY    | 654.30  |      |              | 100    |
| 2026-04-08 | 09:38          | SPY    | 654.74  | 多   | +43.20       | 100    |
| 2026-04-08 | 10:12          | SPY    | 655.10  |      |              | 100    |
| 2026-04-08 | 10:18          | SPY    | 654.50  | 空   | -60.30       | 100    |
| 2026-04-09 | 09:32          | SPY    | 656.20  |      |              | 100    |
| ...        | ...            | ...    | ...     | ...  | ...          | ...    |

**Pairing rules** (same as before):
- Each entry row has empty `Type` and `損益(AI辨識)` columns
- Each exit row has `多` (long) or `空` (short) in `Type`, and the realized PnL in `損益(AI辨識)`
- Entries and exits must alternate: entry → exit → entry → exit → ...
- A trade = the most recent entry row + the next exit row

**Date format**: ISO `YYYY-MM-DD` is preferred (e.g. `2026-04-08`). The app will
also accept `YYYY/MM/DD`, `MM/DD/YYYY`, or `YYYYMMDD` if that's what Excel gives you.

**Just append new rows** to the bottom each trading day. No need to keep them sorted.

---

## 4. One-Time Historical Backfill (run on Bloomberg Terminal computer)

This pulls all your past data from Bloomberg in one go and saves it as JSON files
in `C:\TradeReview\candles\`. Copy these files to your target computer afterwards
(or just run everything from the Bloomberg machine).

```cmd
cd C:\TradeReview
python historical_backfill.py --days 730        REM 2 years
```

Other options:
```cmd
python historical_backfill.py                   REM default: past 365 days
python historical_backfill.py --start 2024-01-01 --end 2024-12-31
python historical_backfill.py --force           REM overwrite existing files
```

**Time estimate**: about 0.5 sec per trading day, so 1 year ≈ 2 minutes.

---

## 5. Daily Fetch Setup (yfinance, runs forever after backfill)

### Manual test
```cmd
cd C:\TradeReview
python daily_fetch.py --catchup
```

You should see something like:
```
  ✓ 2026-04-08  saved 390 bars
  - 2026-04-07  already exists
```

### Schedule to run automatically (Windows Task Scheduler)

**Two triggers**, both running the same command:

1. **Daily at 16:30 ET**:
   - Convert to your local time: 04:30 (winter / EST) or 05:30 (summer / EDT) Taipei
   - For simplicity, use **05:30 daily** — if it's a winter month it'll just run an
     hour later than necessary, no harm done

2. **At log on (when you turn the computer on)**:
   - Catches up any missed days from when the PC was off

**Steps**:

1. Open **Task Scheduler** (Win+R → `taskschd.msc`)
2. Click **Create Task...** (NOT "Create Basic Task")
3. **General tab**:
   - Name: `SPY Daily Fetch`
   - Check ✅ "Run whether user is logged on or not"
   - Check ✅ "Run with highest privileges"
4. **Triggers tab → New...** (do this twice for two triggers)
   - **Trigger 1**: "On a schedule" → Daily → Start at `05:30:00` → OK
   - **Trigger 2**: "At log on" → Any user → OK
5. **Actions tab → New...**
   - Action: "Start a program"
   - Program/script: `python` (or full path like `C:\Python312\python.exe` if needed)
   - Add arguments: `daily_fetch.py --catchup`
   - Start in: `C:\TradeReview`
6. **Conditions tab**:
   - UNcheck ❌ "Start the task only if the computer is on AC power" (laptop users)
7. **Settings tab**:
   - Check ✅ "Run task as soon as possible after a scheduled start is missed"
   - Check ✅ "If the task fails, restart every: 5 minutes, up to 3 times"
8. Click OK, enter your Windows password if prompted

### Verify it works
```cmd
schtasks /run /tn "SPY Daily Fetch"
```

Then check `C:\TradeReview\candles\` to see the new JSON file.

---

## 6. Daily Workflow

After setup, your routine is:

1. **During trading day**: take notes, paste fills into `trades_all.xlsx`
2. **After 16:00 ET**: yfinance daily fetch runs automatically (or manually)
3. **Anytime**: double-click `start_app_free.bat` to open the review UI
4. The app opens in your browser at `http://localhost:5501`

If you turn on the PC after being away for a few days, the **At log on** trigger
will catch up missing days automatically (up to 7 days back by default).

---

## 7. Limitations

- **yfinance 1-minute history is only ~30 days deep**. If a day is missed and
  more than 30 days pass without backfilling it from Bloomberg, it's gone forever.
  → Mitigation: keep the catchup-on-logon trigger enabled, and back up the
  `candles\` folder periodically.
- **Data quality**: yfinance gives consolidated SPY 1-minute bars from Yahoo's
  feed. OHLC values are reliable; volume is approximate compared to Bloomberg.
- **Time alignment**: bars are timestamped at the start of each minute in ET.

---

## 8. Backup Recommendation

Once a week (or after big sessions), zip and back up:

```
C:\TradeReview\trades_all.xlsx
C:\TradeReview\notes\
C:\TradeReview\candles\
```

These three are your entire archive. Everything else is reproducible from this
guide and the `.py` files.
