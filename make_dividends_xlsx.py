"""Generate D:\\fileserver_D\\TradeReview\\dividends.xlsx (SPY ex-dividend dates).
Run once to create template; fill in actual ex-dividend dates and per-share amounts
from broker statements. The app re-reads this file at startup.

Columns:
  Date                  — ex-dividend date (YYYY-MM-DD)
  Dividend_Per_Share    — USD per share (gross; app applies *0.7 for Taiwan tax)
"""
import pandas as pd, os
OUT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(OUT_DIR, "dividends.xlsx")
os.makedirs(OUT_DIR, exist_ok=True)
df = pd.DataFrame([
    {"Date": "2026-03-21", "Dividend_Per_Share": 1.75},
    # Add more rows as SPY distributes dividends, e.g.:
    # {"Date": "2026-06-20", "Dividend_Per_Share": 1.82},
])
df.to_excel(OUT, index=False)
print(f"Wrote {OUT}")
