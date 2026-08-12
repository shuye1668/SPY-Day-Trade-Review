#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""從 trade_review_app.py 抽出前端 HTML，轉成公開靜態版 index.html。

設計原則：前端主體共用 app 的（app 改了就 rebuild，不分岔）。build 時只做四件事：
  1. fetch 目標改讀靜態 JSON（/api/data?date=X → data/X.json 等）
  2. 停用只有後端才有的功能（notes 寫入、version 輪詢）
  3. 靜態 data 一律不含 trades → app 的交易渲染自然全空（不必動那 256 處）
  4. 注入 CSS 隱藏交易相關 UI（P&L、卡片列、匯出、筆記、圖例、同步徽章）

用法：python build_site.py [--app <path>] [--out <dir>]
"""
import argparse
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
DEFAULT_APP = r"D:\fileserver_D\TradeReview\trade_review_app.py"

# 注入到 </head> 前：隱藏所有交易相關 UI，只留純 K 線
HIDE_CSS = """
<style id="public-mode">
/* 公開版：只顯示 K 線，移除所有交易紀錄相關 UI */
#dp, #dtc,            /* Day P&L 與交易筆數 */
#legend,              /* 多/空/留倉 圖例 */
#be,                  /* 匯出 PNG 按鈕 */
#tb2,                 /* 底部交易卡片列 */
#nb, #ned,            /* 市場概述筆記框與編輯彈窗 */
#syncbadge            /* 內部同步新鮮度徽章 */
{ display: none !important; }
/* 交易卡片列拿掉後，圖表區佔滿剩餘高度 */
#cc { flex: 1 1 auto !important; }
</style>
"""

# 頁尾免責（取代原本的操作提示 .hint 內容意義不變，但補上公開聲明）
DISCLAIMER_JS = """
<script>
/* 公開靜態版：資料每分鐘由雲端 GitHub Actions 以 yfinance 更新，與任何本機主機無關。 */
window.__PUBLIC_STATIC__ = true;
</script>
"""


def _replace_call(s, start_token, replacement):
    """把從 start_token 開始的一個完整函式呼叫（含平衡括號、跳過字串內容）
    整段換成 replacement。正則處理不了巢狀大括號與字串，這裡手動掃描。"""
    i = s.find(start_token)
    if i < 0:
        return s
    # 從 start_token 的第一個 '(' 開始配對括號
    p = s.find("(", i)
    depth, j, in_str, esc = 0, p, None, False
    while j < len(s):
        c = s[j]
        if in_str:
            if esc:
                esc = False
            elif c == "\\":
                esc = True
            elif c == in_str:
                in_str = None
        else:
            if c in "\"'`":
                in_str = c
            elif c == "(":
                depth += 1
            elif c == ")":
                depth -= 1
                if depth == 0:
                    j += 1
                    break
        j += 1
    return s[:i] + replacement + s[j:]


def extract_html(app_path):
    src = open(app_path, encoding="utf-8").read()
    m = re.search(r'HTML\s*=\s*r"""(.*?)"""', src, re.S)
    if not m:
        raise SystemExit("找不到 app 的 HTML=r\"\"\"...\"\"\" 區塊")
    return m.group(1)


def staticize(html):
    # ── 1. 資料抓取改讀靜態 JSON ──────────────────────────────
    # /api/data?date=${ds}  →  data/${ds}.json
    html = re.sub(r"`/api/data\?date=\$\{([^}]+)\}`", r"`data/${\1}.json`", html)
    html = re.sub(r'"/api/data\?date="\s*\+\s*(\w+)', r'"data/"+\1+".json"', html)
    # /api/dates → dates.json，但回傳格式從 {dates,no_candles} 變成純陣列，
    # 前端 init 用 dj.dates —— 包一層讓兩者相容
    html = html.replace('await fetch("/api/dates",{cache:"no-store"});const dj=await res.json();',
                        'await fetch("dates.json",{cache:"no-store"});'
                        'const _a=await res.json();const dj={dates:_a,no_candles:[]};')
    # /api/colors → colors.json
    html = html.replace('fetch("/api/colors")', 'fetch("colors.json")')

    # ── 2. 停用後端專屬功能 ───────────────────────────────────
    # 版本輪詢：靜態版改讀 meta.json 的 updated_utc 當版本；沒有就跳過
    html = html.replace('fetch("/api/version",{cache:"no-store"})',
                        'fetch("meta.json",{cache:"no-store"})')
    # meta.json 沒有 .trades 欄，輪詢比較會恆等（不觸發 reload），可接受
    # notes 寫入（POST）：靜態版無後端。用「平衡括號掃描」把整個 fetch(...) 呼叫
    # 換成 (void 0)，避免正則在巢狀大括號/字串裡誤斷（第一版就是這樣壞的）。
    html = _replace_call(html, 'fetch("/api/notes"', "(void 0)")

    # ── 3. 標題與品牌 ────────────────────────────────────────
    html = html.replace("<title>Trade Review</title>", "<title>SPY 1分K 走勢圖</title>")
    html = html.replace('<span class="lbl">Intraday Candle Chart</span>',
                        '<span class="lbl">SPY Intraday · 公開版</span>')

    # ── 4. 注入隱藏 CSS 與公開旗標 ───────────────────────────
    html = html.replace("</head>", HIDE_CSS + "</head>", 1)
    html = html.replace("</body>", DISCLAIMER_JS + "</body>", 1)

    return html


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--app", default=DEFAULT_APP)
    ap.add_argument("--out", default=os.path.join(HERE, "site"))
    a = ap.parse_args()
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

    html = staticize(extract_html(a.app))
    os.makedirs(a.out, exist_ok=True)
    out = os.path.join(a.out, "index.html")
    open(out, "w", encoding="utf-8").write(html)

    # 把色版一起匯出成 colors.json（前端 loadColors 會讀）。
    # 直接用 app 的載入函式，確保與內網版同一份色票。
    try:
        import importlib.util
        spec = importlib.util.spec_from_file_location("tra", a.app)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        import json
        colors = mod._load_colors()
        with open(os.path.join(a.out, "colors.json"), "w", encoding="utf-8") as f:
            json.dump(colors, f, ensure_ascii=False)
        print(f"colors.json: {len(colors)} 個色版")
    except Exception as e:
        # 色版失敗不致命，前端有內建預設；但要讓人看到
        print(f"⚠ colors.json 產生失敗（前端會用預設色）：{type(e).__name__}: {e}")

    # 快速自檢：確認沒有殘留的後端 API 呼叫
    leaks = re.findall(r"/api/\w+", html)
    print(f"寫出 {out}  ({len(html):,} bytes)")
    print(f"殘留 /api/ 呼叫: {sorted(set(leaks)) if leaks else '無 ✅'}")


if __name__ == "__main__":
    main()
