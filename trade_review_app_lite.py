"""
Trade Review — Web App (LITE / Portable version)
=================================================
A self-contained, portable trade review tool. Just drop this .py file and a
trades_all.xlsx into any folder on any computer and run — no Bloomberg, no
daily fetch script, no Windows Task Scheduler, no historical backfill needed.

K-line data is fetched on demand from Yahoo Finance (yfinance) and cached
locally so re-opening the same date is instant. yfinance only allows 1-minute
bars within the past ~30 days; older dates will silently show no data.

Folder layout (same folder as this script):
  trade_review_app_lite.py
  trades_all.xlsx           ← all trades, with Date column (YYYY-MM-DD)
  notes_2026-04-08.txt      ← per-day notes (auto-saved)
  .cache/
    SPY_2026-04-08.json     ← per-day OHLC cache (auto-fetched on demand)

Setup:
  pip install flask pandas numpy openpyxl yfinance
  python trade_review_app_lite.py → http://localhost:5500
"""
import datetime as dt,json,os,glob,warnings,sys
warnings.filterwarnings("ignore")
import pandas as pd,numpy as np
from flask import Flask,jsonify,request

# All paths are relative to this script's directory, so the whole thing is portable
SCRIPT_DIR=os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
TRADES_FILE=os.path.join(SCRIPT_DIR,"trades_all.xlsx")
NOTES_DIR=SCRIPT_DIR  # notes_YYYY-MM-DD.txt sit alongside the script
CACHE_DIR=os.path.join(SCRIPT_DIR,".cache")
TICKER="SPY"
PORT=5502

# In-memory cache for parsed candles
_cache={}

def fetch_intraday(ticker,date_str):
    """Return 1-minute candles for a date.
    Lookup order: in-memory cache → on-disk JSON cache → yfinance (then save to disk).
    Returns [] if yfinance has no data for that day (holiday or > 30 days old)."""
    k=f"{ticker}|{date_str}"
    if k in _cache:return _cache[k]
    # Disk cache
    cache_path=os.path.join(CACHE_DIR,f"{ticker}_{date_str}.json")
    if os.path.exists(cache_path):
        try:
            with open(cache_path,"r",encoding="utf-8") as f:
                bars=json.load(f)
            _cache[k]=bars
            return bars
        except Exception as e:
            print(f"[fetch_intraday] cache read failed for {date_str}: {e}")
    # Cache miss → fetch from yfinance
    bars=_fetch_from_yfinance(ticker,date_str)
    _cache[k]=bars
    if bars:
        # Save to disk for next time
        try:
            os.makedirs(CACHE_DIR,exist_ok=True)
            with open(cache_path,"w",encoding="utf-8") as f:
                json.dump(bars,f,separators=(",",":"))
        except Exception as e:
            print(f"[fetch_intraday] cache write failed for {date_str}: {e}")
    return bars

def _fetch_from_yfinance(ticker,date_str):
    """Pull 1-minute bars from yfinance for a single date (RTH only)."""
    try:
        import yfinance as yf
    except ImportError:
        print("  ✗ yfinance not installed. Run: pip install yfinance")
        return []
    try:
        d=dt.datetime.strptime(date_str,"%Y-%m-%d").date()
    except ValueError:
        return []
    start_str=d.strftime("%Y-%m-%d")
    end_str=(d+dt.timedelta(days=1)).strftime("%Y-%m-%d")
    try:
        df=yf.download(
            tickers=ticker,
            start=start_str,
            end=end_str,
            interval="1m",
            prepost=False,
            auto_adjust=False,
            progress=False,
            threads=False,
        )
    except Exception as e:
        print(f"[yfinance] {date_str} error: {e}")
        return []
    if df is None or len(df)==0:
        return []
    if isinstance(df.columns,pd.MultiIndex):
        df.columns=df.columns.get_level_values(0)
    if df.index.tz is None:
        df.index=df.index.tz_localize("UTC").tz_convert("US/Eastern")
    else:
        df.index=df.index.tz_convert("US/Eastern")
    df=df.between_time("09:30","15:59")
    df=df[df.index.date==d]
    if len(df)==0:
        return []
    bars=[]
    for ts,row in df.iterrows():
        try:
            o=float(row["Open"]);h=float(row["High"])
            lo=float(row["Low"]);c=float(row["Close"])
            v=int(row["Volume"]) if not pd.isna(row["Volume"]) else 0
        except Exception:
            continue
        if any(pd.isna(x) for x in (o,h,lo,c)):
            continue
        bars.append({
            "t":ts.strftime("%H:%M"),
            "o":round(o,4),"h":round(h,4),"l":round(lo,4),"c":round(c,4),
            "v":v,
        })
    print(f"[yfinance] fetched {len(bars)} bars for {date_str}")
    return bars

# Cache the loaded trades DataFrame; reload if file mtime changes
_trades_cache={"mtime":None,"df":None}

def load_trades_df():
    """Load trades_all.xlsx, with cache invalidation on file change."""
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
    if v is None or (isinstance(v,float) and pd.isna(v)):return ""
    s=str(v).strip()
    for fmt in ("%Y-%m-%d","%Y/%m/%d","%m/%d/%Y","%Y%m%d"):
        try:return dt.datetime.strptime(s,fmt).strftime("%Y-%m-%d")
        except:pass
    try:return pd.to_datetime(s).strftime("%Y-%m-%d")
    except:return s

def read_and_pair_for_date(date_str):
    df=load_trades_df()
    if df is None or len(df)==0:return []
    if "Date" not in df.columns:
        print("[read_and_pair_for_date] trades_all.xlsx is missing Date column")
        return []
    sub=df.copy()
    sub["_d"]=sub["Date"].apply(_norm_date)
    sub=sub[sub["_d"]==date_str]
    if len(sub)==0:return []
    sub["Price"]=sub["Price"].astype(float)
    sub["Shares"]=sub["Shares"].astype(int)
    sub["Type"]=sub["Type"].fillna("").str.strip() if "Type" in sub.columns else ""
    pc=[c for c in sub.columns if "損益" in c or "PnL" in c.lower()]
    sub["PnL_raw"]=sub[pc[0]].fillna("").str.strip() if pc else ""
    trades,entry=[],None
    for _,r in sub.iterrows():
        ie=r["Type"] in ("多","空") and r["PnL_raw"]!=""
        if not ie:entry=r
        elif entry is not None:
            ps2=r["PnL_raw"].replace("+","").replace("\uff0b","").replace(",","")
            try:pnl=float(ps2)
            except:pnl=0.0
            trades.append({
                "entryTime":entry["Exec Time(EDT)"].strip()[:5],
                "entryPrice":entry["Price"],
                "exitTime":r["Exec Time(EDT)"].strip()[:5],
                "exitPrice":r["Price"],
                "dir":r["Type"],
                "pnl":round(pnl,2),
                "shares":r["Shares"],
            })
    return trades

app=Flask(__name__)

@app.route("/")
def index():return HTML

@app.route("/api/dates")
def api_dates():
    """Return all dates that appear in trades_all.xlsx Date column,
    plus any dates already cached on disk."""
    dates=set()
    df=load_trades_df()
    if df is not None and "Date" in df.columns:
        for v in df["Date"].dropna().unique():
            nd=_norm_date(v)
            if nd:dates.add(nd)
    if os.path.isdir(CACHE_DIR):
        for f in glob.glob(os.path.join(CACHE_DIR,f"{TICKER}_*.json")):
            n=os.path.basename(f)
            try:
                d=n.replace(f"{TICKER}_","").replace(".json","")
                dt.datetime.strptime(d,"%Y-%m-%d")
                dates.add(d)
            except:pass
    return jsonify(sorted(dates,reverse=True))

@app.route("/api/data")
def api_data():
    ds=request.args.get("date","")
    if not ds:return jsonify({"error":"missing date"}),400
    notes_path=os.path.join(NOTES_DIR,f"notes_{ds}.txt")
    notes=""
    if os.path.exists(notes_path):
        with open(notes_path,"r",encoding="utf-8") as f:notes=f.read()
    c=fetch_intraday(TICKER,ds)
    if not c:
        return jsonify({"date":ds,"candles":[],"trades":[],"notes":notes,"noTrading":True})
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
    with open(os.path.join(NOTES_DIR,f"notes_{ds}.txt"),"w",encoding="utf-8") as f:
        f.write(d.get("notes",""))
    return jsonify({"ok":True})

HTML=r"""<!DOCTYPE html><html lang="zh-Hant"><head><meta charset="UTF-8">
<title>Trade Review</title>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{background:#000;color:#C0C4CC;font-family:'Consolas','Courier New',monospace;overflow:hidden;height:100vh}
#hdr{height:26px;display:flex;align-items:center;padding:0 12px;background:#0D0F13;border-bottom:1px solid #1A1D22}
#hdr .tk{color:#FFF;font-weight:bold;font-size:13px}
#hdr .lbl{margin-left:auto;background:#FF8C00;color:#000;font-weight:bold;font-size:11px;padding:2px 10px;border-radius:2px}
#tb{height:30px;display:flex;align-items:center;gap:8px;padding:0 12px;background:#0A0C0F;border-bottom:1px solid #1A1D22;font-size:12px}
#tb button{background:#1A1D22;color:#8B8F98;border:1px solid #333640;padding:2px 12px;cursor:pointer;font-family:inherit;font-size:11px;border-radius:2px}
#tb button:hover{background:#252830;color:#FFF}
#di{background:#000;color:#FF8C00;border:1px solid #333640;font-family:inherit;font-size:13px;font-weight:bold;width:110px;text-align:center;padding:2px 4px;border-radius:2px}
#di:focus{border-color:#FF8C00;outline:none}
.pnl{font-weight:bold;font-size:13px;margin-left:12px}.pnl.w{color:#FFF}.pnl.l{color:#FF4444}
.tc{color:#8B8F98;font-size:11px}
#zc{margin-left:auto;display:flex;gap:4px;align-items:center}
#zc button{font-size:13px;padding:2px 10px}
#ip{font-size:12px;color:#9aa0aa;display:flex;align-items:center;gap:3px;cursor:pointer;padding:0 6px;user-select:none}
#ip input{margin:0;cursor:pointer}
#cc{position:relative;width:100%;height:calc(100vh - 56px - 105px)}
canvas{display:block;width:100%;height:100%}
#tb2{height:105px;background:#0A0C0F;border-top:1px solid #1A1D22;display:flex;align-items:stretch;overflow-x:auto;padding:6px 12px;gap:6px}
.tc2{flex:0 0 auto;min-width:130px;background:#0D0F13;border:1px solid #1A1D22;border-radius:4px;padding:6px 10px;font-size:11px;display:flex;flex-direction:column;justify-content:center;transition:border-color .15s,box-shadow .15s}
.tc2 .dr{color:#8B8F98;margin-bottom:3px}.tc2 .pv{font-weight:bold;font-size:14px}
.tc2 .pv.w{color:#FFF}.tc2 .pv.l{color:#FF4444}.tc2 .dt{color:#888;font-size:10px;margin-top:2px}
.tc2.hi{border-color:#FFD700!important;box-shadow:0 0 8px rgba(255,215,0,0.35)}
#nb{position:absolute;background:rgba(26,29,34,0.94);border:1px solid #3A3E48;border-radius:4px;padding:0;font-size:12px;color:#E0E0E0;line-height:1.7;max-width:320px;cursor:move;user-select:none;z-index:20;white-space:pre-wrap;word-wrap:break-word}
#nb .nb-hdr{display:flex;align-items:center;justify-content:space-between;padding:4px 10px 2px;border-bottom:1px solid #2A2E35;cursor:move}
#nb .nb-hdr span{color:#FF8C00;font-size:10px;font-weight:bold}
#nb .nb-min{background:none;border:1px solid #3A3E48;color:#8B8F98;font-size:14px;line-height:1;width:22px;height:18px;cursor:pointer;border-radius:2px;display:flex;align-items:center;justify-content:center;padding:0}
#nb .nb-min:hover{color:#FFF;border-color:#666}
#nb .nb-body{padding:8px 14px 10px}
#nb .eh{color:#555;font-size:9px;margin-top:6px}
#nb.mini{max-width:none;padding:0;cursor:move}
#nb.mini .nb-body{display:none}
#nb.mini .nb-hdr{border-bottom:none;padding:3px 8px}
#ned{display:none;position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.7);z-index:100;justify-content:center;align-items:center}
#ned.show{display:flex}
#ned .pn{background:#1A1D22;border:1px solid #3A3E48;border-radius:8px;padding:20px;width:420px}
#ned .pn h3{color:#FF8C00;font-size:14px;margin-bottom:12px}
#ned textarea{width:100%;height:180px;background:#000;color:#E0E0E0;border:1px solid #3A3E48;border-radius:4px;padding:10px;font-family:inherit;font-size:12px;resize:vertical;line-height:1.6}
#ned button{background:#3D6FCC;color:#FFF;border:none;border-radius:4px;padding:6px 24px;cursor:pointer;font-family:inherit;font-size:12px;margin-top:12px}
#ld{position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);color:#FF8C00;font-size:14px;z-index:50;display:none}
.hint{position:fixed;bottom:6px;right:12px;color:#333;font-size:10px;z-index:10}
</style></head><body>
<div id="hdr"><span class="tk">SPY US Equity</span><span class="lbl">Intraday Candle Chart</span></div>
<div id="tb">
<button id="bp">&#8592; Prev</button>
<input type="text" id="di" value="---" spellcheck="false">
<button id="bn">Next &#8594;</button>
<button id="bt" title="跳到最新美股交易日">Today</button>
<span class="pnl" id="dp">---</span><span class="tc" id="dtc"></span>
<div id="zc"><button id="zo" title="Zoom Out (-)">&#8722;</button><button id="zi" title="Zoom In (+)">&#43;</button><button id="zr" title="Reset View">R</button><label id="ip" title="輸出時包含前一日尾盤"><input type="checkbox" id="ipc">含前日</label><button id="be" title="輸出 PNG">⤓ Export</button></div>
</div>
<div id="cc"><canvas id="cv"></canvas><div id="nb"></div><div id="ld">載入中 ...</div></div>
<div id="tb2"></div>
<div id="ned"><div class="pn"><h3>市場概述</h3><textarea id="nt"></textarea><button id="ns">儲存</button></div></div>
<div class="hint">拖曳平移（跨日無縫）｜ 滾輪/+- 縮放 ｜ ← → 切日期 ｜ 雙擊文字框編輯</div>
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
let notePos={x:.72,y:.08},nDrag=false,nOff={x:0,y:0},notesMini=false;
let cDrag=false,cDS={x:0,y:0},cDP={x:0,y:0};
let yAxisDrag=false,yAxisDS=0,yAxisZoomStart=1;
let loading=false; // mutex for edge-loading

const C={bg:"#000",grid:"#454A55",gridM:"#5A6070",txt:"#8B8F98",uB:"#16A34A",dB:"#DC2626",uW:"#16A34A",dW:"#DC2626",
lo:"#000000",sh:"#FFFFFF",pW:"#FFF",pL:"#FF4444",lp:"#FFD700",bd:"#333640",cross:"#6A7080",sep:"#FF8C00"};
const fmt=n=>n.toFixed(2);
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
  const focusDay=dayCache.get(dayList[focusIdx]);
  if(focusDay&&focusDay.candles&&focusDay.candles.length){
    focusDay.candles.forEach(c=>{mn=Math.min(mn,c.l);mx=Math.max(mx,c.h);});
    (focusDay.trades||[]).forEach(t=>{
      mn=Math.min(mn,t.entryPrice,t.exitPrice);
      mx=Math.max(mx,t.entryPrice,t.exitPrice);
    });
  }
  if(mn>mx){mn=650;mx=660;}
  // Default: step=1 always (matches export). userYZoom > 1 widens the visible range,
  // userYZoom < 1 narrows it. Step is chosen so step*7 covers the requested range.
  const baseRange=7; // default visible range = $7 (7 grids of $1)
  const targetRange=baseRange*userYZoom;
  const ladder=[0.1,0.2,0.5,1,2,5,10,20,50,100,200,500];
  let step=1;
  for(const s of ladder){if(s*7>=targetRange*1.0){step=s;break;}}
  // Snap mid to half-step boundary so pn lands exactly on a line position
  const rawMid=(mn+mx)/2;
  const mid=Math.round(rawMid/step-0.5)*step+step/2;
  const half=step*3.5;
  return{pn:mid-half+panY,px:mid+half+panY};
}

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
  document.getElementById("di").value=date;
  const pe=document.getElementById("dp"),dtc=document.getElementById("dtc");
  if(!day||day.noTrading){pe.textContent="非交易日";pe.className="pnl";pe.style.color="#8B8F98";dtc.textContent="";}
  else if(!day.trades||!day.trades.length){pe.textContent="無當日交易記錄";pe.className="pnl";pe.style.color="#8B8F98";dtc.textContent="";}
  else{pe.style.color="";const total=day.trades.reduce((s,t)=>s+t.pnl,0);
    pe.textContent=`Day P&L: ${total>=0?"+":""}$${fmt(Math.abs(total))}`;pe.className="pnl "+(total>=0?"w":"l");
    dtc.textContent=`(${day.trades.length} trades)`;}
  renderCards();
  uNotes();
}

// ── Hit test trades (across all visible days) ──
function hitTest(mx,my,pn,px){
  for(let di=0;di<dayList.length;di++){
    const day=dayCache.get(dayList[di]);if(!day||!day.trades)continue;
    for(let ti=0;ti<day.trades.length;ti++){
      const t=day.trades[ti];
      const ei=t2iLocal(t.entryTime,day.candles),xi=t2iLocal(t.exitTime,day.candles);
      const x1=xOfGi(gi(di,ei)),y1=yOf(t.entryPrice,pn,px);
      const x2=xOfGi(gi(di,xi)),y2=yOf(t.exitPrice,pn,px);
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

    // Trades — pass 1: lines + dots
    // Color scheme:
    //   Long  (多) - App: bright blue (#5BA0FF) / Export: deep blue (#0040C0)
    //   Short (空) - App: light gray  (#B0B0B0) / Export: black     (#000000)
    const longCol=exportLightMode?"#0040C0":"#5BA0FF";
    const shortCol=exportLightMode?"#000000":"#B0B0B0";
    // Multi-page export filter helper: trade belongs on a page if its midpoint
    // price is within [pn, px). Use half-open interval to avoid duplicates on
    // page boundaries. The topmost page extends its upper bound to be inclusive.
    const tradeBelongsOnPage=(t)=>{
      if(!exportMode||!exportYOverride)return true;
      const midPrice=(t.entryPrice+t.exitPrice)/2;
      // Multi-page: find page whose center is closest to trade's midPrice
      if(exportPageTops&&exportPageTops.length>1&&currentExportPageIdx!=null){
        const pageRangeLocal=9; // step*PAGE_GRIDS (export uses 9 grids)
        let bestPage=0,bestDist=Infinity;
        for(let p=0;p<exportPageTops.length;p++){
          const center=exportPageTops[p]-pageRangeLocal/2;
          const dist=Math.abs(center-midPrice);
          if(dist<bestDist){bestDist=dist;bestPage=p;}
        }
        return bestPage===currentExportPageIdx;
      }
      // Single page: use simple range check
      return midPrice>=pn&&midPrice<px+0.0001;
    };
    (day.trades||[]).forEach((t,ti)=>{
      if(!tradeBelongsOnPage(t))return;
      const ei=t2iLocal(t.entryTime,candles),xi=t2iLocal(t.exitTime,candles);
      const x1=xOfGi(gi(di,ei)),y1=yOf(t.entryPrice,pn,px);
      const x2=xOfGi(gi(di,xi)),y2=yOf(t.exitPrice,pn,px);
      const hov=hovHit&&hovHit.di===di&&hovHit.ti===ti;
      const isLong=t.dir==="多";
      const tradeCol=isLong?longCol:shortCol;

      // Glow on hover (gold halo)
      if(hov){
        ctx.strokeStyle="#FFD700";
        ctx.lineWidth=sz(8);ctx.globalAlpha=.25;ctx.lineCap="round";
        ctx.beginPath();ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.stroke();
        ctx.globalAlpha=1;
      }

      // Dashed line — thicker in export for printability
      const lineW=exportMode?sz(3)*1.4:sz(3);
      ctx.strokeStyle=tradeCol;ctx.lineWidth=lineW;
      // Export uses a slightly wider gap so thicker dashes don't blur together
      ctx.setLineDash(exportMode?[sz(2),sz(4.5)]:[sz(2),sz(4)]);
      ctx.lineCap="round";
      ctx.beginPath();ctx.moveTo(x1,y1);ctx.lineTo(x2,y2);ctx.stroke();
      ctx.setLineDash([]);

      // Solid filled dots — slightly larger in export
      const drBase=Math.max(1.5,Math.min(4.5,2.5*Math.pow(zoom,0.5)));
      const dr=exportMode?drBase*1.4:drBase;
      [[x1,y1],[x2,y2]].forEach(([px2,py])=>{
        ctx.fillStyle=tradeCol;
        ctx.beginPath();ctx.arc(px2,py,dr,0,Math.PI*2);ctx.fill();
        if(hov){ctx.strokeStyle="#FFD700";ctx.lineWidth=1.5;ctx.beginPath();ctx.arc(px2,py,dr+2,0,Math.PI*2);ctx.stroke();}
      });
    });

    // Trades — pass 2: PnL labels with collision avoidance
    // Build obstacle list: K-line wicks (high-low rectangles) for this day + ALL day separators
    const obstacles=[];
    candles.forEach((c,li)=>{
      const x=xOfGi(gi(di,li));
      if(x<M.l-2||x>W-M.r+2)return;
      const yh=yOf(c.h,pn,px),yl=yOf(c.l,pn,px);
      obstacles.push({x:x-Math.max(1,bw/2),y:Math.min(yh,yl),w:Math.max(2,bw),h:Math.abs(yl-yh)+1});
    });
    // Add all day separators (global) so labels never cross any of them
    for(const s of sepObstacles)obstacles.push(s);
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

    // Sort trades by entry time so labels are placed in chronological order.
    // We track the previous label's X-center to enforce monotonic left-to-right ordering.
    const tradesIndexed=(day.trades||[]).map((t,ti)=>({t,ti}));
    tradesIndexed.sort((a,b)=>{
      const ka=a.t.entryTime||"";const kb=b.t.entryTime||"";
      if(ka!==kb)return ka<kb?-1:1;
      return a.ti-b.ti;
    });
    let prevLabelCenterX=-Infinity; // monotonic constraint: current label center >= prev

    tradesIndexed.forEach(({t,ti})=>{
      // Multi-page export: skip trades not belonging on this page
      if(!tradeBelongsOnPage(t))return;
      const ei=t2iLocal(t.entryTime,candles),xi=t2iLocal(t.exitTime,candles);
      const x1=xOfGi(gi(di,ei)),y1=yOf(t.entryPrice,pn,px);
      const x2=xOfGi(gi(di,xi)),y2=yOf(t.exitPrice,pn,px);
      const txt=(t.pnl>=0?"+":"")+fmt(t.pnl);
      const lblW=ctx.measureText(txt).width+4;
      const lblH=labelFontSize+2;
      const midX=(x1+x2)/2;
      const baseAbove=Math.min(y1,y2)-lblH/2-6;
      const baseBelow=Math.max(y1,y2)+lblH/2+6;

      // Generate candidates: dense grid around the trade midpoint
      const candidates=[];
      for(let dy=0;dy<=200;dy+=4){
        candidates.push({x:midX-lblW/2,y:baseAbove-dy});
        candidates.push({x:midX-lblW/2,y:baseBelow+dy});
      }
      for(let dx=6;dx<=120;dx+=6){
        for(let dy=0;dy<=100;dy+=8){
          candidates.push({x:midX-lblW/2+dx,y:baseAbove-dy});
          candidates.push({x:midX-lblW/2-dx,y:baseAbove-dy});
          candidates.push({x:midX-lblW/2+dx,y:baseBelow+dy});
          candidates.push({x:midX-lblW/2-dx,y:baseBelow+dy});
        }
      }

      let chosen=null;
      let bestFallback=null,bestOverlap=Infinity;
      // Bound labels to this day's own K-line time range
      const dayLeftX=xOfGi(gi(di,0)-0.5)+2;
      const dayRightX=xOfGi(gi(di,BPD-1)+0.5)-2;
      const leftBound=Math.max(M.l,dayLeftX);
      const rightBound=Math.min(W-M.r,dayRightX);
      for(const c of candidates){
        if(c.x<leftBound||c.x+lblW>rightBound||c.y<M.t||c.y+lblH>H-M.b)continue;
        // Monotonic-X constraint: label center must be >= prev label's center
        const centerX=c.x+lblW/2;
        if(centerX<prevLabelCenterX)continue;
        if(!labelHasCollision(c.x,c.y,lblW,lblH)){chosen=c;break;}
        const ov=labelOverlap(c.x,c.y,lblW,lblH);
        if(ov<bestOverlap){bestOverlap=ov;bestFallback=c;}
      }
      if(!chosen)chosen=bestFallback;
      // Last-resort fallback: place at midX above (still respecting monotonic constraint)
      if(!chosen){
        const fx=Math.max(midX-lblW/2,prevLabelCenterX-lblW/2);
        chosen={x:fx,y:baseAbove};
      }
      const finalRect={x:chosen.x,y:chosen.y,w:lblW,h:lblH};
      placedLabels.push(finalRect);
      pnlLabelsForNotes.push(finalRect); // shared with notes obstacle list
      prevLabelCenterX=chosen.x+lblW/2;
      ctx.fillStyle=t.pnl>=0?T.pW:T.pL;
      ctx.textAlign="left";
      ctx.fillText(txt,chosen.x+2,chosen.y+labelFontSize);
    });
  });
  ctx.restore();

  // Y axis (always right for screen; both sides during dual-axis export)
  const yAxisFont=exportMode?"13px Consolas,monospace":"11px Consolas,monospace";
  const yAxisInt=ps>=1&&ps===Math.floor(ps); // integer step → no decimals
  const fmtY=v=>yAxisInt?String(Math.round(v)):fmt(v);
  // Bottom label adjustment: shift up if too close to chart bottom (export only)
  const yLabelBottomAdjust=exportMode?6:0;
  const drawYLabels=(side)=>{
    if(side==="right"){
      ctx.fillStyle=T.bg;ctx.fillRect(W-M.r,0,M.r,H);
      for(let p=gs;p<=px;p+=ps){const y=yOf(p,pn,px);if(y<M.t-5||y>H-M.b+15)continue;
      const adj=(y>H-M.b-2)?-yLabelBottomAdjust:0;
      ctx.fillStyle=T.txt;ctx.font=yAxisFont;ctx.textAlign="left";ctx.fillText(fmtY(p),W-M.r+6,y+4+adj);}
    } else {
      ctx.fillStyle=T.bg;ctx.fillRect(0,0,M.l+60,H);
      for(let p=gs;p<=px;p+=ps){const y=yOf(p,pn,px);if(y<M.t-5||y>H-M.b+15)continue;
      const adj=(y>H-M.b-2)?-yLabelBottomAdjust:0;
      ctx.fillStyle=T.txt;ctx.font=yAxisFont;ctx.textAlign="right";ctx.fillText(fmtY(p),M.l+56,y+4+adj);}
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
        const x=xOfGi(gi(di,li));if(x>=M.l&&x<=W-M.r)ctx.fillText(c.t,x,H-M.b+14);
      }
    });
  });
  // Export: date label centered below time labels
  if(exportMode){
    const focusDate=dayList[focusIdx]||"";
    if(focusDate){
      ctx.fillStyle=T.txt;ctx.font="bold 14px Consolas,monospace";ctx.textAlign="center";
      ctx.fillText(focusDate,W/2,H-M.b+34);
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
      // OHLCV box only when on a candle
      const bw2=Math.max(3,ppb()*.55);
      const nearestX=xOfGi(Math.round(gi2x(mouse.x)));
      if(Math.abs(mouse.x-nearestX)<bw2*1.8){
        const hdr=`SPY US Equity ${hitDate.replace(/-/g,"/")} ${hitCandle.t}`;
        const rows=[{l:"Open",v:fmt(hitCandle.o)},{l:"High",v:fmt(hitCandle.h)},{l:"Low",v:fmt(hitCandle.l)},{l:"Close",v:fmt(hitCandle.c)}];
        const pad=5,rh=13,hdrH=16;
        // Measure actual content width
        ctx.font="bold 10px Consolas,monospace";
        let contentW=ctx.measureText(hdr).width;
        ctx.font="10px Consolas,monospace";
        rows.forEach(r=>{const w=ctx.measureText(r.l+"  "+r.v).width;if(w>contentW)contentW=w;});
        const tw=Math.ceil(contentW)+pad*2;
        const th=hdrH+rows.length*rh+pad;
        let tx=mouse.x+14,ty=mouse.y-th-6;
        if(tx+tw>W-M.r)tx=mouse.x-tw-14;if(ty<M.t)ty=mouse.y+14;
        ctx.fillStyle="rgba(180,210,220,0.88)";ctx.fillRect(tx,ty,tw,th);
        ctx.strokeStyle="rgba(100,140,160,0.6)";ctx.lineWidth=.5;ctx.strokeRect(tx,ty,tw,th);
        ctx.fillStyle="#1A3A4A";ctx.font="bold 10px Consolas,monospace";ctx.textAlign="left";ctx.fillText(hdr,tx+pad,ty+12);
        ctx.font="10px Consolas,monospace";
        rows.forEach((r,ri)=>{const ry=ty+hdrH+ri*rh+10;
          ctx.fillStyle="#2A4A5A";ctx.textAlign="left";ctx.fillText(r.l,tx+pad,ry);
          ctx.fillStyle="#1A3A4A";ctx.textAlign="right";ctx.fillText(r.v,tx+tw-pad,ry);});
      }
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

        const desiredX=W*notePos.x;
        const desiredY=notesAnchorPrice!=null?yOf(notesAnchorPrice,pn,px):H*notePos.y;
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
      const res=await fetch(`/api/data?date=${ds}`);const data=await res.json();
      dayCache.set(ds,{candles:data.candles||[],trades:data.trades||[],notes:data.notes||"",noTrading:!!data.noTrading});
      if(data.candles&&data.candles.length)return ds; // found trading day
    }catch(e){dayCache.set(ds,{candles:[],trades:[],notes:"",noTrading:true});}
  }
  return null; // no trading day found within 7 days
}

async function checkEdges(){
  if(loading||!dayList.length)return;
  const leftG=gi2x(M.l),rightG=gi2x(W-M.r);
  const firstDayStart=gi(0,0);
  const lastDayEnd=gi(dayList.length-1,BPD-1);
  // Near left edge → load previous trading day
  if(leftG<firstDayStart+60){
    loading=true;
    const found=await loadTradingDay(dayList[0],-1);
    if(found){
      dayList.unshift(found);focusIdx++;
      panX-=(BPD+GAP)*ppb();
    }
    loading=false;draw();
  }
  // Near right edge → load next trading day
  if(rightG>lastDayEnd-60){
    loading=true;
    const found=await loadTradingDay(dayList[dayList.length-1],1);
    if(found){dayList.push(found);}
    loading=false;draw();
  }
}

// ═══════════════════════════════════════════════════
// Notes
// ═══════════════════════════════════════════════════
const nb=document.getElementById("nb");
function uNotes(){
  const day=dayCache.get(dayList[focusIdx]);
  const notes=day?day.notes||"":"";
  nb.style.display="block";
  const hdr='<div class="nb-hdr"><span>Notes</span><button class="nb-min" id="nbtn" title="最小化/展開">'+(notesMini?'&#9633;':'&#8212;')+'</button></div>';
  if(notesMini){nb.className="mini";nb.innerHTML=hdr;}
  else{nb.className="";
    if(notes.trim())nb.innerHTML=hdr+'<div class="nb-body">'+notes.replace(/</g,"&lt;").replace(/\n/g,"<br>")+'<div class="eh">雙擊編輯 ｜ 拖曳移動</div></div>';
    else nb.innerHTML=hdr+'<div class="nb-body"><span style="color:#555;font-style:italic">雙擊新增市場概述</span></div>';
  }
  pNotes();
  document.getElementById("nbtn").addEventListener("click",e=>{e.stopPropagation();notesMini=!notesMini;uNotes();});
}
function pNotes(){const r=cc.getBoundingClientRect();nb.style.left=(r.width*notePos.x)+"px";nb.style.top=(r.height*notePos.y)+"px";}
nb.addEventListener("mousedown",e=>{if(e.target.closest(".nb-min"))return;if(e.detail===2)return;nDrag=true;const r=nb.getBoundingClientRect();nOff={x:e.clientX-r.left,y:e.clientY-r.top};e.preventDefault();e.stopPropagation();});
nb.addEventListener("dblclick",e=>{if(e.target.closest(".nb-min"))return;if(notesMini){notesMini=false;uNotes();return;}
  const day=dayCache.get(dayList[focusIdx]);
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
    const ar=t.dir==="多"?"▲ 多":"▼ 空",pc=t.pnl>=0?"w":"l";
    d.innerHTML=`<div class="dr">${ar} ${t.entryTime} → ${t.exitTime}</div><div class="pv ${pc}">${t.pnl>=0?"+":""}${fmt(t.pnl)}</div><div class="dt">${fmt(t.entryPrice)} → ${fmt(t.exitPrice)}</div>`;
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
  if(nDrag){const r=cc.getBoundingClientRect();notePos.x=Math.max(0,Math.min(.95,(e.clientX-r.left-nOff.x)/r.width));
  notePos.y=Math.max(0,Math.min(.9,(e.clientY-r.top-nOff.y)/r.height));pNotes();return;}
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
  nDrag=false;
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

async function exportPNG(){
  const day=dayCache.get(dayList[focusIdx]);
  if(!day||!day.candles||!day.candles.length){alert("無 K 線資料無法輸出");return;}
  const date=dayList[focusIdx];

  // Compute full day price range (including trade prices)
  let mn=1e9,mx=-1e9;
  day.candles.forEach(c=>{mn=Math.min(mn,c.l);mx=Math.max(mx,c.h);});
  (day.trades||[]).forEach(t=>{mn=Math.min(mn,t.entryPrice,t.exitPrice);mx=Math.max(mx,t.entryPrice,t.exitPrice);});

  // Export uses 9-grid page range (vs App's 7) to better fit A4 paper
  // and reduce both page count and notes-collision pressure
  const step=1; // $1 per grid
  const PAGE_GRIDS=9;
  const pageRange=step*PAGE_GRIDS;
  const PREV_BARS=60; // show last 60 bars (1 hour) of previous day
  const TOTAL_BARS=BPD+PREV_BARS; // 391 + 60 = 451

  // Make sure previous trading day is loaded
  const prevDate=await loadTradingDay(date,-1);

  // Include previous day's tail price range if either:
  //   (1) this is the ISO week's first trading day (always)
  //   (2) the user has checked the "含前日" checkbox
  const includePrevByCheckbox=document.getElementById("ipc").checked;
  const isWeekFirst=await isFirstTradingDayOfWeek(date);
  if((isWeekFirst||includePrevByCheckbox)&&prevDate){
    const pd=dayCache.get(prevDate);
    if(pd&&pd.candles&&pd.candles.length){
      const tail=pd.candles.slice(-PREV_BARS);
      tail.forEach(c=>{mn=Math.min(mn,c.l);mx=Math.max(mx,c.h);});
    }
  }
  const dayRange=mx-mn;

  const exportDayList=[];
  if(prevDate)exportDayList.push(prevDate);
  exportDayList.push(date);

  // Save current state (including canvas size)
  const savedDayList=dayList,savedFocusIdx=focusIdx,savedZoom=zoom,savedPanX=panX,savedPanY=panY,savedUserYZoom=userYZoom;
  const savedCvW=cv.width,savedCvH=cv.height,savedStyleW=cv.style.width,savedStyleH=cv.style.height;
  const savedM={...M};

  dayList=exportDayList;
  focusIdx=prevDate?1:0;

  // ── Set up high-resolution export canvas ──
  // Square cells: pxPerCandle = pxPerDollar (1 dollar high = 30 minutes wide visually)
  // cW_export = TOTAL_BARS * bw = 451 * bw
  // cH_export = PAGE_GRIDS * 30 * bw = 270 * bw (was 210 with 7 grids)
  // Aspect ratio: 270/451 ≈ 0.5987 (vs 0.4656 before) — closer to A4 landscape 1.41:1
  const EXPORT_W=1400; // CSS pixels (close to typical screen so font ratios match)
  const EXPORT_DPR=3;  // 3x for crisp printing → 4200px actual
  // Increase bottom margin in export to fit date label below time labels
  M.b=48;
  const innerW=EXPORT_W-M.l-68; // 68 = right Y axis margin
  const bwTarget=innerW/TOTAL_BARS;
  const innerH=PAGE_GRIDS*30*bwTarget;
  const EXPORT_H=Math.round(innerH+M.t+M.b);
  // Resize canvas
  cv.width=EXPORT_W*EXPORT_DPR;
  cv.height=EXPORT_H*EXPORT_DPR;
  cv.style.width=EXPORT_W+"px";
  cv.style.height=EXPORT_H+"px";
  ctx.setTransform(EXPORT_DPR,0,0,EXPORT_DPR,0,0);
  W=EXPORT_W;H=EXPORT_H;cW=W-M.l-M.r;cH=H-M.t-M.b;

  // Setup export flags
  exportMode=true;
  exportLightMode=true;
  exportDualAxis=true;
  exportYAxisLeft=false;
  userYZoom=1; // ignore user's Y zoom during export

  // Compute the price the notes are anchored to: map notePos.y through the
  // focus day's default autoY range (not multi-page override)
  // Compute notesAnchorPrice using App's default 7-grid range
  // (notePos.y is relative to App's chart, not export's larger 9-grid view)
  exportYOverride=null;
  const{pn:defPn,px:defPx}=autoY();
  notesAnchorPrice=defPx-notePos.y*(defPx-defPn);
  // Note position in chart: top half if notePos.y < 0.5, bottom half otherwise
  // (used to bias asymmetric expansion of the 9-grid range)
  const notesInUpperHalf=notePos.y<0.5;
  // Symmetric is preferred when notes is roughly centered
  const notesIsCentered=notePos.y>=0.4&&notePos.y<=0.6;

  // Compute zoom: ppb = cW/(BPD-1) * zoom = bwTarget → zoom = bwTarget*(BPD-1)/cW
  zoom=bwTarget*(BPD-1)/cW;

  // Position panX so prev day's last 60 bars + focus day full = visible window
  const firstFocusG=gi(focusIdx,0);
  panX=(PREV_BARS-firstFocusG)*bwTarget;
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

  // Restore resize function
  resizeFreeze=false;

  // Restore canvas state
  cv.width=savedCvW;cv.height=savedCvH;
  cv.style.width=savedStyleW;cv.style.height=savedStyleH;
  M=savedM;

  // Restore export flags & view state
  exportMode=false;exportLightMode=false;exportDualAxis=false;exportYAxisLeft=false;exportYOverride=null;notesAnchorPrice=null;notesPageIdx=null;currentExportPageIdx=null;exportPageTops=null;
  dayList=savedDayList;focusIdx=savedFocusIdx;zoom=savedZoom;panX=savedPanX;panY=savedPanY;userYZoom=savedUserYZoom;
  draw();
}
document.getElementById("be").addEventListener("click",exportPNG);

// ═══════════════════════════════════════════════════
// Data loading & Navigation
// ═══════════════════════════════════════════════════
async function loadDay(ds){
  if(dayCache.has(ds))return;
  const res=await fetch(`/api/data?date=${ds}`);const data=await res.json();
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
      const found=await loadTradingDay(ds,-1);
      if(found){ds=found;}
      else{document.getElementById("ld").style.display="none";return;}
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

async function init(){
  const res=await fetch("/api/dates");const dates=await res.json();
  if(dates.length){await jumpTo(dates[0]);}
  else{const today=getLatestUSTradeDate();await jumpTo(today);}
}

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

const di=document.getElementById("di");
di.addEventListener("keydown",e=>{if(e.key==="Enter"){e.preventDefault();di.blur();
const v=di.value.trim();if(/^\d{4}-\d{2}-\d{2}$/.test(v))jumpTo(v);}});
di.addEventListener("focus",()=>di.select());

document.getElementById("ns").addEventListener("click",()=>{
  const day=dayCache.get(dayList[focusIdx]);if(day)day.notes=document.getElementById("nt").value;
  document.getElementById("ned").classList.remove("show");uNotes();
  fetch("/api/notes",{method:"POST",headers:{"Content-Type":"application/json"},body:JSON.stringify({date:dayList[focusIdx],notes:document.getElementById("nt").value})});
});
document.getElementById("ned").addEventListener("click",e=>{if(e.target===document.getElementById("ned"))document.getElementById("ned").classList.remove("show");});

document.addEventListener("keydown",e=>{if(document.getElementById("ned").classList.contains("show")||document.activeElement===di)return;
if(e.key==="ArrowLeft")document.getElementById("bp").click();
if(e.key==="ArrowRight")document.getElementById("bn").click();
if(e.key==="+"||e.key==="=")doZoom(ZS);if(e.key==="-")doZoom(-ZS);});

window.addEventListener("resize",draw);
init();
</script></body></html>"""

if __name__=="__main__":
    print(f"\n  Trade Review Web App — LITE / Portable version")
    print(f"  Script dir:  {SCRIPT_DIR}")
    print(f"  Trades file: {TRADES_FILE}  {'(found)' if os.path.exists(TRADES_FILE) else '(MISSING)'}")
    print(f"  Cache dir:   {CACHE_DIR}")
    print(f"  http://localhost:{PORT}\n")
    if not os.path.exists(CACHE_DIR):os.makedirs(CACHE_DIR,exist_ok=True)
    if not os.path.exists(TRADES_FILE):
        print(f"  ⚠ {TRADES_FILE} not found.")
        print(f"     Create an Excel file with these columns (in this order):")
        print(f"     Date | Exec Time(EDT) | Symbol | Price | Type | 損益(AI辨識) | Shares")
        print(f"     Date format: YYYY-MM-DD")
        print(f"     The app will still start so you can verify candles are fetched.\n")
    try:
        import yfinance  # noqa: F401
    except ImportError:
        print(f"  ⚠ yfinance is not installed. K-line fetching will fail.")
        print(f"     Install with:  pip install yfinance\n")
    app.run(host="0.0.0.0",port=PORT,debug=False)
