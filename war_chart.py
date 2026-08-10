"""
War Event Daily Chart — v1
===========================
Reads events.xlsx (in same folder), pulls SPY daily OHLC from yfinance,
renders a single-page daily K-line chart with annotated event labels.

Features:
  - Daily K lines (one bar = one trading day), red/green filled rectangles
  - Event labels: dashed thin border, color-coded by 性質 (利空/利多/中性)
  - Holiday events (no K line for that date) get arrows pointing to the
    midpoint between adjacent trading days
  - Hover tooltip in App shows 事件簡述 + 市場/SPY 反應
  - Export PNG shows only event title

Excel format (events.xlsx, in same folder as this script):
  Columns: # | 日期 | 星期 | 事件標題 | 事件簡述 | 性質 | 市場/SPY 反應
  Date format: 2026/02/26 (or any common format; auto-parsed)
  性質 values: 利空 | 利多 | 中性 | (or empty / "—")

Setup:
  pip install flask pandas numpy openpyxl yfinance
  python war_chart.py → http://localhost:5503
"""
import datetime as dt, json, os, warnings
warnings.filterwarnings("ignore")
import pandas as pd, numpy as np
from flask import Flask, jsonify, request

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
EVENTS_FILE = os.path.join(SCRIPT_DIR, "events.xlsx")
CACHE_DIR = os.path.join(SCRIPT_DIR, "daily_cache")
CACHE_FILE = os.path.join(CACHE_DIR, "SPY_daily.json")
TICKER = "SPY"
PORT = 5503


# ────────────────────── Date utilities ──────────────────────
def norm_date(v):
    """Normalize various date formats to YYYY-MM-DD string. Returns '' on failure."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%m/%d/%Y", "%Y%m%d"):
        try:
            return dt.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    try:
        return pd.to_datetime(s).strftime("%Y-%m-%d")
    except Exception:
        return ""


# ────────────────────── Events Excel reader ──────────────────────
def load_events():
    """Read events.xlsx and return a list of event dicts.
    Each event: {date, weekday, title, summary, nature, reaction}.
    Always re-reads the file (no caching) so user edits take effect immediately."""
    if not os.path.exists(EVENTS_FILE):
        print(f"[load_events] {EVENTS_FILE} not found")
        return []
    try:
        df = pd.read_excel(EVENTS_FILE, dtype=str)
    except Exception as e:
        print(f"[load_events] failed to read {EVENTS_FILE}: {e}")
        return []
    df.columns = [c.strip() for c in df.columns]

    def col(*names):
        """Find first column matching any of the given name fragments (case-insensitive)."""
        for c in df.columns:
            cl = c.lower().replace(" ", "")
            for n in names:
                if n.lower().replace(" ", "") in cl:
                    return c
        return None

    c_date = col("日期", "date")
    c_weekday = col("星期", "weekday")
    c_title = col("事件標題", "標題", "title")
    c_summary = col("事件簡述", "簡述", "summary", "description")
    c_nature = col("性質", "nature", "type")
    c_reaction = col("市場", "spy 反應", "reaction")

    if not c_date or not c_title:
        print(f"[load_events] required columns missing. Found: {list(df.columns)}")
        return []

    events = []
    for _, r in df.iterrows():
        date_str = norm_date(r.get(c_date, ""))
        if not date_str:
            continue
        title = str(r.get(c_title, "")).strip() if c_title else ""
        if not title or title in ("nan", "—", "-"):
            # No title → skip (allow placeholder rows with only date)
            continue
        nature = str(r.get(c_nature, "")).strip() if c_nature else ""
        if nature in ("nan", "—", "-"):
            nature = ""
        events.append({
            "date": date_str,
            "weekday": str(r.get(c_weekday, "")).strip() if c_weekday else "",
            "title": title,
            "summary": str(r.get(c_summary, "")).strip() if c_summary else "",
            "nature": nature,
            "reaction": str(r.get(c_reaction, "")).strip() if c_reaction else "",
        })
    return events


# ────────────────────── yfinance daily K line ──────────────────────
def fetch_daily(start_date, end_date):
    """Fetch daily OHLC for SPY between start_date and end_date (inclusive).
    Uses cache: if cache covers requested range, return cached; else re-fetch.
    Returns list of {date, o, h, l, c, v}."""
    cached = None
    if os.path.exists(CACHE_FILE):
        try:
            with open(CACHE_FILE, "r", encoding="utf-8") as f:
                cached = json.load(f)
        except Exception:
            cached = None

    needs_refetch = True
    if cached and isinstance(cached, list) and len(cached) > 0:
        cached_dates = [b["date"] for b in cached]
        cached_min = min(cached_dates)
        cached_max = max(cached_dates)
        if cached_min <= start_date and cached_max >= end_date:
            needs_refetch = False

    if not needs_refetch:
        # Filter cached to requested range
        return [b for b in cached if start_date <= b["date"] <= end_date]

    # Fetch fresh from yfinance
    try:
        import yfinance as yf
    except ImportError:
        print("✗ yfinance not installed. Run: pip install yfinance")
        return []

    # Pad both ends slightly so future re-runs with similar range hit cache
    fetch_start = (dt.datetime.strptime(start_date, "%Y-%m-%d") - dt.timedelta(days=14)).strftime("%Y-%m-%d")
    fetch_end = (dt.datetime.strptime(end_date, "%Y-%m-%d") + dt.timedelta(days=14)).strftime("%Y-%m-%d")

    print(f"[fetch_daily] yfinance: {fetch_start} → {fetch_end}")
    try:
        df = yf.download(
            tickers=TICKER,
            start=fetch_start,
            end=fetch_end,
            interval="1d",
            auto_adjust=False,
            progress=False,
            threads=False,
        )
    except Exception as e:
        print(f"[fetch_daily] yfinance error: {e}")
        return []

    if df is None or len(df) == 0:
        print("[fetch_daily] empty result from yfinance")
        return []
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = df.columns.get_level_values(0)

    bars = []
    for ts, row in df.iterrows():
        try:
            o = float(row["Open"]); h = float(row["High"])
            lo = float(row["Low"]); c = float(row["Close"])
            v = int(row["Volume"]) if not pd.isna(row["Volume"]) else 0
        except Exception:
            continue
        if any(pd.isna(x) for x in (o, h, lo, c)):
            continue
        d = ts.strftime("%Y-%m-%d") if hasattr(ts, "strftime") else str(ts)[:10]
        bars.append({
            "date": d,
            "o": round(o, 4), "h": round(h, 4), "l": round(lo, 4), "c": round(c, 4),
            "v": v,
        })

    # Save cache (pre-filter, full fetch range)
    try:
        os.makedirs(CACHE_DIR, exist_ok=True)
        with open(CACHE_FILE, "w", encoding="utf-8") as f:
            json.dump(bars, f, separators=(",", ":"))
        print(f"[fetch_daily] cached {len(bars)} bars to {CACHE_FILE}")
    except Exception as e:
        print(f"[fetch_daily] cache write failed: {e}")

    # Return only the requested range
    return [b for b in bars if start_date <= b["date"] <= end_date]


# ────────────────────── Flask routes ──────────────────────
app = Flask(__name__)


@app.route("/")
def index():
    return HTML


@app.route("/api/data")
def api_data():
    """Return events + candles for the date range derived from events.xlsx."""
    events = load_events()
    if not events:
        return jsonify({"events": [], "candles": [], "warning": "no events found in events.xlsx"})

    dates = [e["date"] for e in events]
    start_date = min(dates)
    end_date = max(dates)
    candles = fetch_daily(start_date, end_date)

    return jsonify({
        "events": events,
        "candles": candles,
        "start": start_date,
        "end": end_date,
    })


# ────────────────────── HTML / JS ──────────────────────
HTML = r"""<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="UTF-8">
<title>War Event Chart</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { background: #000; color: #C0C4CC; font-family: 'Consolas', 'Courier New', monospace;
       overflow: hidden; height: 100vh; }
#tb { display: flex; align-items: center; gap: 12px; padding: 8px 14px; background: #0f1115;
      border-bottom: 1px solid #2a2d34; height: 44px; }
#tb h1 { font-size: 14px; color: #C0C4CC; font-weight: normal; }
#tb .info { font-size: 12px; color: #8a9099; }
#tb .right { margin-left: auto; display: flex; gap: 4px; align-items: center; }
#tb button { font-size: 13px; padding: 3px 12px; background: #1f2229; color: #ccc;
             border: 1px solid #3a3d44; cursor: pointer; border-radius: 3px; }
#tb button:hover { background: #2a2d34; }
#cc { position: relative; width: 100%; height: calc(100vh - 44px); }
canvas { display: block; width: 100%; height: 100%; }
#tooltip { position: absolute; display: none; background: rgba(40,40,40,0.95); color: #fff;
           padding: 8px 12px; border-radius: 4px; font-family: 'Times New Roman', Times, serif;
           font-size: 14px; line-height: 1.5; max-width: 400px; pointer-events: none;
           z-index: 100; box-shadow: 0 2px 8px rgba(0,0,0,0.5); }
#tooltip .title { font-weight: bold; font-size: 15px; margin-bottom: 4px; }
#tooltip .row { margin-top: 4px; }
#tooltip .label { color: #aab; font-size: 12px; }
#ld { position: absolute; top: 50%; left: 50%; transform: translate(-50%, -50%);
      background: rgba(0,0,0,0.8); padding: 12px 24px; border-radius: 4px; color: #fff;
      font-size: 14px; }
</style></head>
<body>
<div id="tb">
  <h1>War Event Daily Chart</h1>
  <span class="info" id="info">Loading...</span>
  <div class="right">
    <button id="zo" title="Zoom Out">&#8722;</button>
    <button id="zi" title="Zoom In">&#43;</button>
    <button id="zr" title="Reset View">R</button>
    <button id="be" title="Export PNG">⤓ Export</button>
  </div>
</div>
<div id="cc">
  <canvas id="cv"></canvas>
  <div id="tooltip"></div>
  <div id="ld">載入中 ...</div>
</div>
<script>
"use strict";

// ────────── State ──────────
let candles = [];      // [{date, o, h, l, c, v}]
let events = [];       // [{date, weekday, title, summary, nature, reaction}]
let dateRange = {start: "", end: ""};
const cv = document.getElementById("cv");
const ctx = cv.getContext("2d");
let W = 0, H = 0;
const M = {t: 60, r: 70, b: 50, l: 6};  // top margin holds holiday annotation rows
let cW = 0, cH = 0;
let zoom = 1;
let panX = 0, panY = 0;
let userYZoom = 1;
let mouse = null;
let cDrag = null;
let exportMode = false;
let exportLightMode = false;
let resizeFreeze = false;
let hoveredLabel = null;       // currently hovered label (for visual highlight)
let placedLabelsGlobal = [];   // last drawn labels (for hit-test)

// ────────── Colors ──────────
const C = {
  bg: "#000",
  txt: "#C0C4CC",
  grid: "#2a2d34",
  gridM: "#3a3d44",
  uB: "#16A34A",  // up green
  dB: "#DC2626",  // down red
  cross: "#5BA0FF",
};

// Theme based on mode
function theme() {
  if (exportLightMode) return {
    bg: "#FFFFFF", txt: "#000000", grid: "#D8D8D8", gridM: "#E8E8E8",
    uB: "#16A34A", dB: "#DC2626",
  };
  return {bg: C.bg, txt: C.txt, grid: C.grid, gridM: C.gridM, uB: C.uB, dB: C.dB};
}

// Nature → label colors (light/pastel for low visual interference)
function natureColors(nature) {
  // bgApp: semi-transparent for App (so K lines/grid show through)
  // bg:    solid for Export (printable)
  if (nature === "利空") return {border: "#E08585", bg: "#FFE8E8", bgApp: "rgba(255,200,200,0.45)"};
  if (nature === "利多") return {border: "#7BB880", bg: "#E8F5EA", bgApp: "rgba(180,230,180,0.45)"};
  return {border: "#A0A0A0", bg: "#F0F0F0", bgApp: "rgba(180,180,180,0.40)"};  // 中性 / empty / —
}

// ────────── Layout helpers ──────────
function resize() {
  if (resizeFreeze) return;
  const cc = document.getElementById("cc");
  const r = cc.getBoundingClientRect();
  const dpr = window.devicePixelRatio || 1;
  cv.width = r.width * dpr;
  cv.height = r.height * dpr;
  cv.style.width = r.width + "px";
  cv.style.height = r.height + "px";
  ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
  W = r.width;
  H = r.height;
  cW = W - M.l - M.r;
  cH = H - M.t - M.b;
}

// X coordinate per K line index (after pan/zoom)
function xOfIdx(i) {
  const n = candles.length || 1;
  const bw = (cW * zoom) / n;
  return M.l + panX + (i + 0.5) * bw;
}

function bw() {
  const n = candles.length || 1;
  return (cW * zoom) / n;
}

// Y axis: auto-fit candle highs/lows + small padding
function yRange() {
  if (!candles.length) return {pn: 600, px: 700};
  let mn = Infinity, mx = -Infinity;
  for (const c of candles) {
    if (c.l < mn) mn = c.l;
    if (c.h > mx) mx = c.h;
  }
  // Add 8% padding top and bottom for event label space
  const pad = (mx - mn) * 0.12 * userYZoom;
  return {pn: mn - pad + panY, px: mx + pad + panY};
}

function yOf(p, pn, px) {
  return M.t + (1 - (p - pn) / (px - pn)) * cH;
}

function fmt(n) { return n.toFixed(2); }
function pStep(range) {
  const raw = range / 8;
  const mag = Math.pow(10, Math.floor(Math.log10(raw)));
  const n = raw / mag;
  if (n > 5) return 10 * mag;
  if (n > 2) return 5 * mag;
  if (n > 1) return 2 * mag;
  return mag;
}

// ────────── Drawing ──────────
function draw() {
  resize();
  const T = theme();
  ctx.fillStyle = T.bg;
  ctx.fillRect(0, 0, W, H);

  if (!candles.length) {
    ctx.fillStyle = T.txt;
    ctx.font = "14px sans-serif";
    ctx.textAlign = "center";
    ctx.fillText("No data", W / 2, H / 2);
    return;
  }

  const {pn, px} = yRange();
  const ps = pStep(px - pn);
  const gs = Math.ceil(pn / ps) * ps;

  // Grid lines — export uses light thin SOLID lines (to avoid clashing with
  // dashed connector lines of event labels); App uses dashed dark
  ctx.strokeStyle = T.grid;
  ctx.lineWidth = 0.5;
  if (exportMode) {
    ctx.setLineDash([]);
  } else {
    ctx.setLineDash([2, 4]);
  }
  // Horizontal grid (at price steps)
  for (let p = gs; p <= px; p += ps) {
    const y = yOf(p, pn, px);
    if (y < M.t || y > H - M.b) continue;
    ctx.beginPath();
    ctx.moveTo(M.l, y);
    ctx.lineTo(W - M.r, y);
    ctx.stroke();
  }
  // Vertical grid (every K line)
  ctx.strokeStyle = T.gridM;
  for (let i = 0; i < candles.length; i++) {
    const x = xOfIdx(i);
    if (x < M.l || x > W - M.r) continue;
    ctx.beginPath();
    ctx.moveTo(x, M.t);
    ctx.lineTo(x, H - M.b);
    ctx.stroke();
  }
  ctx.setLineDash([]);

  // K lines (filled rectangles + wicks)
  const candleW = Math.max(2, bw() * 0.6);
  for (let i = 0; i < candles.length; i++) {
    const c = candles[i];
    const x = xOfIdx(i);
    const yo = yOf(c.o, pn, px);
    const yc = yOf(c.c, pn, px);
    const yh = yOf(c.h, pn, px);
    const yl = yOf(c.l, pn, px);
    const up = c.c >= c.o;
    ctx.fillStyle = up ? T.uB : T.dB;
    ctx.strokeStyle = up ? T.uB : T.dB;
    ctx.lineWidth = 1.5;
    // Wick
    ctx.beginPath();
    ctx.moveTo(x, yh);
    ctx.lineTo(x, yl);
    ctx.stroke();
    // Body (filled rectangle)
    const bodyTop = Math.min(yo, yc);
    const bodyH = Math.max(1, Math.abs(yc - yo));
    ctx.fillRect(x - candleW / 2, bodyTop, candleW, bodyH);
  }

  // Y axis labels (right side)
  ctx.fillStyle = T.bg;
  ctx.fillRect(W - M.r, 0, M.r, H);
  ctx.fillStyle = T.txt;
  ctx.font = exportMode ? "13px Consolas,monospace" : "11px Consolas,monospace";
  ctx.textAlign = "left";
  for (let p = gs; p <= px; p += ps) {
    const y = yOf(p, pn, px);
    if (y < M.t - 5 || y > H - M.b + 15) continue;
    ctx.fillText(fmt(p), W - M.r + 4, y + 4);
  }

  // X axis labels (date MM/DD) — every trading day
  ctx.fillStyle = T.bg;
  ctx.fillRect(0, H - M.b, W, M.b);
  ctx.fillStyle = T.txt;
  ctx.font = exportMode ? "12px Consolas,monospace" : "10px Consolas,monospace";
  ctx.textAlign = "center";
  for (let i = 0; i < candles.length; i++) {
    const x = xOfIdx(i);
    if (x < M.l || x > W - M.r) continue;
    const [y, m, d] = candles[i].date.split("-");
    ctx.fillText(`${parseInt(m, 10)}/${parseInt(d, 10)}`, x, H - M.b + 14);
  }

  // Crosshair (App only)
  if (mouse && !exportMode && mouse.x > M.l && mouse.x < W - M.r &&
      mouse.y > M.t && mouse.y < H - M.b) {
    ctx.strokeStyle = C.cross;
    ctx.lineWidth = 0.5;
    ctx.setLineDash([3, 3]);
    ctx.beginPath();
    ctx.moveTo(M.l, mouse.y);
    ctx.lineTo(W - M.r, mouse.y);
    ctx.stroke();
    ctx.beginPath();
    ctx.moveTo(mouse.x, M.t);
    ctx.lineTo(mouse.x, H - M.b);
    ctx.stroke();
    ctx.setLineDash([]);
  }

  // ────────── Event labels ──────────
  // Group events by date
  const eventsByDate = {};
  for (const ev of events) {
    if (!eventsByDate[ev.date]) eventsByDate[ev.date] = [];
    eventsByDate[ev.date].push(ev);
  }

  // Index of candles by date (for fast lookup)
  const candleIdxByDate = {};
  for (let i = 0; i < candles.length; i++) candleIdxByDate[candles[i].date] = i;

  // Compute targets for each event-date
  // For trading days: target = (xOfIdx, yHigh or yLow depending on chart half)
  // For holidays: target = (midpoint X between prev/next trading day, edge Y)
  const targets = []; // {date, evList, tx, ty, side, isHoliday, color}
  const sortedEventDates = Object.keys(eventsByDate).sort();

  for (const date of sortedEventDates) {
    const evList = eventsByDate[date];
    // Color = first event's nature (mixed natures use first)
    const color = natureColors(evList[0].nature);

    if (candleIdxByDate[date] !== undefined) {
      // Trading day with K line
      const i = candleIdxByDate[date];
      const c = candles[i];
      const tx = xOfIdx(i);
      if (tx < M.l - bw()/2 || tx > W - M.r + bw()/2) continue;
      const yh = yOf(c.h, pn, px);
      const yl = yOf(c.l, pn, px);
      // Both sides are viable — algorithm picks later
      targets.push({date, evList, tx, tyAbove: yh - 8, tyBelow: yl + 8, isHoliday: false, color});
    } else {
      // Holiday — find adjacent trading days
      let prevIdx = -1, nextIdx = -1;
      for (let i = 0; i < candles.length; i++) {
        if (candles[i].date < date) prevIdx = i;
        else if (candles[i].date > date && nextIdx === -1) { nextIdx = i; break; }
      }
      let tx;
      if (prevIdx >= 0 && nextIdx >= 0) tx = (xOfIdx(prevIdx) + xOfIdx(nextIdx)) / 2;
      else if (prevIdx >= 0) tx = xOfIdx(prevIdx) + bw();
      else if (nextIdx >= 0) tx = xOfIdx(nextIdx) - bw();
      else continue;
      if (tx < M.l - bw() || tx > W - M.r + bw()) continue;
      let envHi = -Infinity, envLo = Infinity;
      if (prevIdx >= 0) { envHi = Math.max(envHi, candles[prevIdx].h); envLo = Math.min(envLo, candles[prevIdx].l); }
      if (nextIdx >= 0) { envHi = Math.max(envHi, candles[nextIdx].h); envLo = Math.min(envLo, candles[nextIdx].l); }
      const yEnvHi = isFinite(envHi) ? yOf(envHi, pn, px) : (M.t + H - M.b) / 2;
      const yEnvLo = isFinite(envLo) ? yOf(envLo, pn, px) : (M.t + H - M.b) / 2;
      // Arrow tips just outside envelope on each side
      targets.push({date, evList, tx, tyAbove: yEnvHi - 8, tyBelow: yEnvLo + 8, isHoliday: true, color});
    }
  }

  // Build obstacle list: K-line bounding boxes
  const obstacles = [];
  for (let i = 0; i < candles.length; i++) {
    const c = candles[i];
    const x = xOfIdx(i);
    const w = Math.max(2, bw() * 0.6);
    obstacles.push({
      x: x - w / 2,
      y: yOf(c.h, pn, px),
      w: w,
      h: yOf(c.l, pn, px) - yOf(c.h, pn, px),
    });
  }

  // Pre-measure label sizes with wrap
  const fontSize = exportMode ? 16 : 14;
  ctx.font = `bold ${fontSize}px 'Times New Roman',Times,serif`;
  const lineH = fontSize + 4;
  const padX = 8, padY = 6;

  // Wrap a string into up to 2 lines, splitting at the middle character count
  // but avoiding breaking inside a contiguous run of letters/digits.
  function wrapTitle(s) {
    const len = [...s].length; // Unicode-safe length
    if (len <= 6) return [s];
    const chars = [...s];
    // Ideal mid = floor(len/2) — first line gets floor, so shorter/equal first line
    let mid = Math.floor(len / 2);
    // Slide break toward the first whitespace or CJK-boundary so we don't split
    // a contiguous ASCII word (letters+digits+period+percent)
    const isAlnum = (ch) => /[A-Za-z0-9._%+\-]/.test(ch);
    // Try shrinking mid leftward first (prefer shorter first line)
    let best = mid;
    for (let delta = 0; delta <= 3; delta++) {
      for (const m of [mid - delta, mid + delta]) {
        if (m <= 0 || m >= len) continue;
        const prev = chars[m - 1], cur = chars[m];
        if (!(isAlnum(prev) && isAlnum(cur))) { best = m; delta = 99; break; }
      }
    }
    // Also skip splitting right after/before a space
    while (best > 0 && chars[best] === " ") best++;
    while (best > 0 && chars[best - 1] === " ") best--;
    if (best <= 0 || best >= len) return [s];
    return [chars.slice(0, best).join(""), chars.slice(best).join("")];
  }

  // Partition targets: holidays get rendered as plain text at top; trading days use boxes
  const holidayTargets = [];
  const tradingTargets = [];
  for (const tgt of targets) {
    if (tgt.isHoliday) holidayTargets.push(tgt);
    else tradingTargets.push(tgt);
  }

  // Pre-measure trading-day boxes with wrapped titles
  for (const tgt of tradingTargets) {
    const rawLines = tgt.evList.map(e => e.title);
    const wrapped = [];
    for (const ln of rawLines) {
      const w = wrapTitle(ln);
      for (const piece of w) wrapped.push(piece);
    }
    let maxW = 0;
    for (const ln of wrapped) {
      const w = ctx.measureText(ln).width;
      if (w > maxW) maxW = w;
    }
    tgt.lblW = maxW + padX * 2;
    tgt.lblH = wrapped.length * lineH + padY * 2;
    tgt.lines = wrapped;
  }

  // Placement: greedy by date, picking side with lower total cost
  const placed = []; // {x, y, w, h, target, evList, side, tx, ty, sx, sy}
  function rectsCollide(a, b) {
    return a.x < b.x + b.w && a.x + a.w > b.x && a.y < b.y + b.h && a.y + a.h > b.y;
  }
  function overlapArea(a, b) {
    const ox = Math.max(0, Math.min(a.x+a.w, b.x+b.w) - Math.max(a.x, b.x));
    const oy = Math.max(0, Math.min(a.y+a.h, b.y+b.h) - Math.max(a.y, b.y));
    return ox * oy;
  }
  // Line-rect intersection (segment from (x1,y1) to (x2,y2) crosses rect?)
  function segRectHits(x1, y1, x2, y2, rect) {
    // Quick reject
    if (Math.max(x1, x2) < rect.x || Math.min(x1, x2) > rect.x + rect.w) return false;
    if (Math.max(y1, y2) < rect.y || Math.min(y1, y2) > rect.y + rect.h) return false;
    // Parametric: test intersection with each of 4 edges
    function segSeg(a, b, c, d) {
      const d1 = (d.x-c.x)*(a.y-c.y) - (d.y-c.y)*(a.x-c.x);
      const d2 = (d.x-c.x)*(b.y-c.y) - (d.y-c.y)*(b.x-c.x);
      const d3 = (b.x-a.x)*(c.y-a.y) - (b.y-a.y)*(c.x-a.x);
      const d4 = (b.x-a.x)*(d.y-a.y) - (b.y-a.y)*(d.x-a.x);
      return ((d1>0&&d2<0)||(d1<0&&d2>0)) && ((d3>0&&d4<0)||(d3<0&&d4>0));
    }
    const a={x:x1,y:y1}, b={x:x2,y:y2};
    const tl={x:rect.x,y:rect.y}, tr={x:rect.x+rect.w,y:rect.y};
    const bl={x:rect.x,y:rect.y+rect.h}, br={x:rect.x+rect.w,y:rect.y+rect.h};
    return segSeg(a,b,tl,tr)||segSeg(a,b,tr,br)||segSeg(a,b,br,bl)||segSeg(a,b,bl,tl);
  }

  // Counts for balance
  let countAbove = 0, countBelow = 0;

  function findBestY(chosenX, lblW, lblH, tx, ty, side, isSearchInPreferredSide) {
    // Search outward from arrow tip in chosen side direction; return best {y, cost}
    const startY = side === "above" ? ty - lblH - 8 : ty + 8;
    const dir = side === "above" ? -1 : 1;
    let bestY = null, bestCost = Infinity;
    for (let dy = 0; dy < cH; dy += 4) {
      const y = startY + dir * dy;
      if (y < M.t + 2 || y + lblH > H - M.b - 2) break;
      const rect = {x: chosenX, y: y, w: lblW, h: lblH};
      // Compute cost for this candidate position
      let cost = 0;
      // K-line collision = hard (very high cost) — inflate K-line by 8px for spacing
      for (const o of obstacles) {
        const inflated = {x: o.x - 8, y: o.y - 8, w: o.w + 16, h: o.h + 16};
        cost += overlapArea(rect, inflated) * 100;
      }
      // Placed-label overlap — use inflated rect so labels have ≥8px visual spacing
      for (const p of placed) {
        const inflated = {x: p.x - 8, y: p.y - 8, w: p.w + 16, h: p.h + 16};
        cost += overlapArea(rect, inflated) * 100;
      }
      // Arrow line from (tx,ty) to rect edge — check if crosses any placed label
      const labelCx = chosenX + lblW / 2;
      const labelCy = y + lblH / 2;
      // My connector vs placed labels (inflated 8px so lines don't graze labels)
      for (const p of placed) {
        const infl = {x: p.x - 8, y: p.y - 8, w: p.w + 16, h: p.h + 16};
        if (segRectHits(labelCx, labelCy, tx, ty, infl)) cost += 200;
      }
      // My connector vs K lines (8px buffer) — skip the K line at my own tx
      for (const o of obstacles) {
        // Skip if this obstacle is the owner's K line (tx falls within it)
        if (tx >= o.x && tx <= o.x + o.w) continue;
        const infl = {x: o.x - 8, y: o.y - 8, w: o.w + 16, h: o.h + 16};
        if (segRectHits(labelCx, labelCy, tx, ty, infl)) cost += 200;
      }
      // Already-placed connectors vs THIS candidate rect (inflated 8px)
      const rectInfl = {x: rect.x - 8, y: rect.y - 8, w: rect.w + 16, h: rect.h + 16};
      for (const p of placed) {
        if (segRectHits(p.sx, p.sy, p.tx, p.ty, rectInfl)) cost += 200;
      }
      // Small penalty for distance from target (prefer closer)
      cost += dy * 0.5;
      if (cost < bestCost) { bestCost = cost; bestY = y; }
      // Perfect fit — return immediately
      if (cost < dy * 0.5 + 0.1) return {y: bestY, cost: bestCost};
    }
    return {y: bestY, cost: bestCost};
  }

  // Sort by date (left-to-right on time axis)
  const sortedTargets = [...tradingTargets].sort((a, b) => a.date.localeCompare(b.date));

  for (const tgt of sortedTargets) {
    const {tx, tyAbove, tyBelow, lblW, lblH, isHoliday} = tgt;
    let chosenX = tx - lblW / 2;
    chosenX = Math.max(M.l + 2, Math.min(W - M.r - lblW - 2, chosenX));
    // Try both sides, pick lower cost
    const above = findBestY(chosenX, lblW, lblH, tx, tyAbove, "above", true);
    const below = findBestY(chosenX, lblW, lblH, tx, tyBelow, "below", true);
    // Balance nudge: if above already has more, prefer below and vice versa
    const balanceBonus = 10;
    const aboveCost = above.cost + (countAbove > countBelow ? balanceBonus : 0);
    const belowCost = below.cost + (countBelow > countAbove ? balanceBonus : 0);
    let chosenY, chosenSide, ty;
    if (above.y !== null && (below.y === null || aboveCost <= belowCost)) {
      chosenY = above.y; chosenSide = "above"; ty = tyAbove;
      countAbove++;
    } else if (below.y !== null) {
      chosenY = below.y; chosenSide = "below"; ty = tyBelow;
      countBelow++;
    } else {
      // Both null — use last-resort edge
      chosenY = M.t + 4; chosenSide = "above"; ty = tyAbove;
      countAbove++;
    }
    // Compute connector start (edge of label toward target)
    const labelCx = chosenX + lblW / 2;
    const labelCy = chosenY + lblH / 2;
    const dxL = tx - labelCx, dyL = ty - labelCy;
    let sx = labelCx, sy = labelCy;
    if (dxL !== 0 || dyL !== 0) {
      const halfW = lblW / 2, halfH = lblH / 2;
      const tX = dxL === 0 ? Infinity : halfW / Math.abs(dxL);
      const tY = dyL === 0 ? Infinity : halfH / Math.abs(dyL);
      const t = Math.min(tX, tY);
      sx = labelCx + t * dxL;
      sy = labelCy + t * dyL;
    }
    placed.push({
      x: chosenX, y: chosenY, w: lblW, h: lblH,
      target: tgt, evList: tgt.evList,
      side: chosenSide, tx, ty, sx, sy,
    });
  }

  // Draw connector lines (dashed) + labels
  ctx.lineWidth = 1;
  ctx.setLineDash([3, 3]);
  for (const p of placed) {
    const tgt = p.target;
    // sx, sy, tx, ty were precomputed during placement
    ctx.strokeStyle = tgt.color.border;
    ctx.beginPath();
    ctx.moveTo(p.sx, p.sy);
    ctx.lineTo(p.tx, p.ty);
    ctx.stroke();
    // Arrow head for holiday targets
    if (tgt.isHoliday) {
      ctx.setLineDash([]);
      const dx = p.tx - p.sx, dy = p.ty - p.sy;
      const ah = 7;
      const len = Math.sqrt(dx * dx + dy * dy) || 1;
      const ux = dx / len, uy = dy / len;
      const px2 = -uy, py = ux;
      ctx.fillStyle = tgt.color.border;
      ctx.beginPath();
      ctx.moveTo(p.tx, p.ty);
      ctx.lineTo(p.tx - ux * ah + px2 * ah * 0.5, p.ty - uy * ah + py * ah * 0.5);
      ctx.lineTo(p.tx - ux * ah - px2 * ah * 0.5, p.ty - uy * ah - py * ah * 0.5);
      ctx.closePath();
      ctx.fill();
      ctx.setLineDash([3, 3]);
    }
  }

  // Draw label boxes + text (trading day events)
  ctx.font = `bold ${fontSize}px 'Times New Roman',Times,serif`;
  for (const p of placed) {
    const tgt = p.target;
    const isHover = hoveredLabel === p;
    ctx.fillStyle = exportMode ? tgt.color.bg : tgt.color.bgApp;
    ctx.fillRect(p.x, p.y, p.w, p.h);
    ctx.strokeStyle = tgt.color.border;
    ctx.lineWidth = isHover ? 2 : 1;
    ctx.setLineDash([3, 3]);
    ctx.strokeRect(p.x + 0.5, p.y + 0.5, p.w - 1, p.h - 1);
    ctx.setLineDash([]);
    ctx.fillStyle = exportMode ? "#202020" : "#F0F0F0";
    ctx.textAlign = "left";
    ctx.textBaseline = "top";
    // Use wrapped lines from tgt.lines (one string per rendered line)
    const lines = tgt.lines || tgt.evList.map(e => e.title);
    for (let li = 0; li < lines.length; li++) {
      ctx.fillText(lines[li], p.x + padX, p.y + padY + li * lineH + 2);
    }
    ctx.textBaseline = "alphabetic";
  }
  ctx.lineWidth = 1;

  // Draw holiday events as two-line plain text at the chart top:
  //   Line 1: event title(s)         (13px App / 14px Export)
  //   Line 2: (非交易日) — smaller   (11px App / 12px Export)
  // Gray color, no background. Positioned high (near canvas top).
  const holidayFont = exportMode ? 14 : 13;
  const holidaySubFont = holidayFont - 2;
  const holidayGray = exportMode ? "#666666" : "#999999";
  ctx.textAlign = "center";
  ctx.textBaseline = "top";
  const holidayRows = []; // each row: list of {x1, x2}
  const holidayBlockH = holidayFont + 2 + holidaySubFont + 4; // total height of one holiday's 2-line block
  const holidayHitRects = [];
  // Center the holiday block vertically in the top margin area
  const holidayTopY = Math.max(2, Math.round(M.t / 2 - holidayBlockH / 2));
  for (const ht of holidayTargets) {
    // Combine multiple titles for same date with " / "
    const titleText = ht.evList.map(e => e.title).join(" / ");
    const subText = "(非交易日)";
    // Measure using the title font (larger of the two) for width
    ctx.font = `${holidayFont}px 'Times New Roman',Times,serif`;
    const twTitle = ctx.measureText(titleText).width;
    ctx.font = `${holidaySubFont}px 'Times New Roman',Times,serif`;
    const twSub = ctx.measureText(subText).width;
    const tw = Math.max(twTitle, twSub);
    const x = Math.max(M.l + 2 + tw / 2, Math.min(W - M.r - 2 - tw / 2, ht.tx));
    const x1 = x - tw / 2, x2 = x + tw / 2;
    // Find first row (from top) that doesn't horizontally overlap
    let rowIdx = 0;
    for (; rowIdx < holidayRows.length; rowIdx++) {
      let collide = false;
      for (const seg of holidayRows[rowIdx]) {
        if (x1 < seg.x2 + 8 && x2 > seg.x1 - 8) { collide = true; break; }
      }
      if (!collide) break;
    }
    if (rowIdx === holidayRows.length) holidayRows.push([]);
    holidayRows[rowIdx].push({x1, x2});
    const y0 = holidayTopY + rowIdx * holidayBlockH;
    // Draw title (larger)
    ctx.fillStyle = holidayGray;
    ctx.font = `${holidayFont}px 'Times New Roman',Times,serif`;
    ctx.fillText(titleText, x, y0);
    // Draw (非交易日) below, smaller
    ctx.font = `${holidaySubFont}px 'Times New Roman',Times,serif`;
    ctx.fillText(subText, x, y0 + holidayFont + 2);
    // Hover hit rect covers both lines
    holidayHitRects.push({
      x: x1, y: y0, w: tw, h: holidayFont + 2 + holidaySubFont + 2,
      target: ht, evList: ht.evList, _isHoliday: true,
    });
  }
  ctx.textBaseline = "alphabetic";

  placedLabelsGlobal = placed.concat(holidayHitRects);
}

// ────────── Interaction ──────────
cv.addEventListener("mousedown", (e) => {
  const r = cv.getBoundingClientRect();
  cDrag = {x: e.clientX - r.left, y: e.clientY - r.top, panX, panY};
});
window.addEventListener("mouseup", () => { cDrag = null; });
window.addEventListener("mousemove", (e) => {
  const r = cv.getBoundingClientRect();
  mouse = {x: e.clientX - r.left, y: e.clientY - r.top};
  if (cDrag) {
    panX = cDrag.panX + (mouse.x - cDrag.x);
    hoveredLabel = null;
    document.getElementById("tooltip").style.display = "none";
    draw();
    return;
  }
  // Hit-test against placed labels
  let hit = null;
  for (const p of placedLabelsGlobal) {
    if (mouse.x >= p.x && mouse.x <= p.x + p.w &&
        mouse.y >= p.y && mouse.y <= p.y + p.h) {
      hit = p;
      break;
    }
  }
  const changed = hit !== hoveredLabel;
  hoveredLabel = hit;
  // Update tooltip
  const tip = document.getElementById("tooltip");
  if (hit) {
    let html = "";
    for (let i = 0; i < hit.evList.length; i++) {
      const ev = hit.evList[i];
      if (i > 0) html += '<hr style="border:none;border-top:1px solid #555;margin:8px 0;">';
      html += `<div class="title">${escapeHtml(ev.title)}</div>`;
      html += `<div class="row"><span class="label">日期:</span> ${escapeHtml(ev.date)} ${escapeHtml(ev.weekday || "")}</div>`;
      if (ev.summary) html += `<div class="row"><span class="label">事件簡述:</span> ${escapeHtml(ev.summary)}</div>`;
      if (ev.reaction) html += `<div class="row"><span class="label">市場/SPY 反應:</span> ${escapeHtml(ev.reaction)}</div>`;
    }
    tip.innerHTML = html;
    tip.style.display = "block";
    // Position tooltip near mouse, but keep on screen
    const tipR = tip.getBoundingClientRect();
    let tx = e.clientX + 16;
    let ty = e.clientY + 12;
    if (tx + tipR.width > window.innerWidth - 8) tx = e.clientX - tipR.width - 12;
    if (ty + tipR.height > window.innerHeight - 8) ty = e.clientY - tipR.height - 8;
    tip.style.left = tx + "px";
    tip.style.top = ty + "px";
  } else {
    tip.style.display = "none";
  }
  if (changed) draw(); else draw();
});

function escapeHtml(s) {
  return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}
cv.addEventListener("wheel", (e) => {
  e.preventDefault();
  const factor = e.deltaY < 0 ? 1.1 : 1 / 1.1;
  zoom = Math.max(0.3, Math.min(8, zoom * factor));
  draw();
}, {passive: false});

document.getElementById("zi").onclick = () => { zoom = Math.min(8, zoom * 1.2); draw(); };
document.getElementById("zo").onclick = () => { zoom = Math.max(0.3, zoom / 1.2); draw(); };
document.getElementById("zr").onclick = () => { zoom = 1; panX = 0; panY = 0; userYZoom = 1; draw(); };
document.getElementById("be").onclick = async () => {
  // Export at high resolution with light theme
  const savedW = cv.width, savedH = cv.height;
  const savedSW = cv.style.width, savedSH = cv.style.height;
  const savedZoom = zoom, savedPanX = panX, savedPanY = panY;
  const savedYZoom = userYZoom;
  resizeFreeze = true;
  exportMode = true;
  exportLightMode = true;
  hoveredLabel = null;
  // Reset view so all candles fit
  zoom = 1; panX = 0; panY = 0; userYZoom = 1;
  // High-resolution canvas: 1800 CSS px wide, 3x DPR
  const EW = 1800, EDPR = 3;
  // Aspect: keep similar to current view but fit all candles
  const EH = Math.round(EW * 0.5);
  cv.width = EW * EDPR; cv.height = EH * EDPR;
  cv.style.width = EW + "px"; cv.style.height = EH + "px";
  ctx.setTransform(EDPR, 0, 0, EDPR, 0, 0);
  W = EW; H = EH; cW = W - M.l - M.r; cH = H - M.t - M.b;
  draw();
  const dataURL = cv.toDataURL("image/png");
  const a = document.createElement("a");
  a.download = `war_chart_${dateRange.start}_to_${dateRange.end}.png`;
  a.href = dataURL;
  a.click();
  await new Promise(r => setTimeout(r, 200));
  // Restore
  exportMode = false; exportLightMode = false;
  cv.width = savedW; cv.height = savedH;
  cv.style.width = savedSW; cv.style.height = savedSH;
  zoom = savedZoom; panX = savedPanX; panY = savedPanY; userYZoom = savedYZoom;
  resizeFreeze = false;
  resize(); draw();
};

window.addEventListener("resize", draw);

// ────────── Init: fetch data ──────────
async function init() {
  document.getElementById("ld").style.display = "block";
  try {
    const r = await fetch("/api/data");
    const data = await r.json();
    candles = data.candles || [];
    events = data.events || [];
    dateRange = {start: data.start || "", end: data.end || ""};
    const info = `${dateRange.start} → ${dateRange.end}  |  ${candles.length} bars  |  ${events.length} events`;
    document.getElementById("info").textContent = info;
    if (data.warning) console.warn(data.warning);
  } catch (e) {
    document.getElementById("info").textContent = "Failed to load: " + e.message;
  }
  document.getElementById("ld").style.display = "none";
  resize();
  draw();
}

init();
</script></body></html>"""


if __name__ == "__main__":
    print(f"\n  War Event Daily Chart")
    print(f"  Script dir:   {SCRIPT_DIR}")
    print(f"  Events file:  {EVENTS_FILE}  {'(found)' if os.path.exists(EVENTS_FILE) else '(MISSING)'}")
    print(f"  Cache:        {CACHE_FILE}")
    print(f"  http://localhost:{PORT}\n")
    if not os.path.exists(CACHE_DIR):
        os.makedirs(CACHE_DIR, exist_ok=True)
    if not os.path.exists(EVENTS_FILE):
        print(f"  ⚠ {EVENTS_FILE} not found.")
        print(f"     Required columns: 日期 | 星期 | 事件標題 | 事件簡述 | 性質 | 市場/SPY 反應")
        print(f"     Date format: 2026/02/26 (or any common format)\n")
    try:
        import yfinance  # noqa
    except ImportError:
        print(f"  ⚠ yfinance not installed.  Run: pip install yfinance\n")
    app.run(host="0.0.0.0", port=PORT, debug=False)
