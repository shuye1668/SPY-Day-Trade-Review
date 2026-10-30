#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
julia_mail.py — 明細信內容產生器（2026-10-20，取代 4d 的 PNG 附件）

使用者要求：信裡直接呈現彙總表格本身（Excel 區塊原樣），不要附件縮圖。
做法：用 julia_block_spec.py 的同一份引擎數字，產生 HTML 表格（新細明體、負數紅字括號、
千分位、當日損益列上細線+下雙底線、% 兩位小數），不再傳 base64 圖片。

輸出 JSON：{subject, body(純文字), html}
用法：python julia_mail.py <csv> --date YYYY-MM-DD
"""
import argparse
import json
import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))


def money(v):
    s = f"{abs(v):,.2f}"
    return f'<span style="color:#ff0000">({s})</span>' if v < 0 else s


def pct(v):
    s = f"{abs(v) * 100:.2f}%"
    return f'<span style="color:#ff0000">-{s}</span>' if v < 0 else s


TD = "border:1px solid #d9d9d9;padding:2px 8px;white-space:nowrap;"


def table(rows):
    """rows: [(W, X, Y_html, style)]，style in {'', 'pnl', 'pct'}"""
    out = ['<table style="border-collapse:collapse;font-family:PMingLiU,\'新細明體\',serif;'
           'font-size:16px;margin:8px 0">']
    for w, x, y, st in rows:
        ys = TD + "text-align:right;min-width:130px;"
        if st == "pnl":
            ys += "border-top:1px solid #000;border-bottom:3px double #000;"
        out.append(f'<tr><td style="{TD}min-width:120px">{w}</td>'
                   f'<td style="{TD}min-width:90px">{x}</td><td style="{ys}">{y}</td></tr>')
    out.append("</table>")
    return "".join(out)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    ap = argparse.ArgumentParser()
    ap.add_argument("csv_path")
    ap.add_argument("--date", required=True)
    a = ap.parse_args()
    spec = json.loads(subprocess.run(
        [sys.executable, os.path.join(HERE, "julia_block_spec.py"), a.csv_path, "--date", a.date, "--flag"],
        capture_output=True, text=True, encoding="utf-8", check=True).stdout)
    if not spec["ok"]:
        print(json.dumps({"error": spec["reasons"]}, ensure_ascii=False)); sys.exit(3)

    legs, dot = spec["legs"], spec["dot"]
    lines = spec["summary_lines"]
    html = [f'<div style="font-family:PMingLiU,\'新細明體\',serif;font-size:16px">']
    html.append("".join(f"<div>{l}</div>" for l in lines))
    text = list(lines)
    flags = spec.get("flags") or []
    if flags:   # 2026-10-30：不擋，照寄但標明待修正
        warn = "⚠️待修正：" + "；".join(flags)
        html.append(f'<div style="color:#ff0000;font-weight:bold;margin:6px 0">{warn}</div>')
        text.append(warn)

    exits = [l for l in legs if l["kind"] == "day" and l["pnl"] is not None]
    if exits:
        day = [l for l in legs if l["kind"] == "day"]
        buy = round(sum(l["net"] for l in day if l["net"] < 0), 2)
        sell = round(sum(l["net"] for l in day if l["net"] > 0), 2)
        p = round(buy + sell, 2)
        html.append(table([(f"{dot}當沖", "買", money(buy), ""),
                           (f"SPY {len(exits) * 100}股", "賣", money(sell), ""),
                           ("", "當日損益", money(p), "pnl"),
                           ("", "當日損益 %", pct(p / abs(buy)), "")]))
    closes = [l for l in legs if l["kind"] == "close"]
    if closes:
        buy = round(sum(l["entry"]["net"] for l in closes), 2)
        sell = round(sum(l["net"] for l in closes), 2)
        p = round(buy + sell, 2)
        html.append(table([(f"{dot} 平倉", "買", money(buy), ""),
                           (f"SPY {len(closes) * 100}股", "賣", money(sell), ""),
                           ("", "當日損益", money(p), "pnl"),
                           ("", "當日損益 %", pct(p / abs(buy)), "")]))
    holds = spec["holds_after"]
    if holds:
        rows = [(f"累積留倉 {sum(h['qty'] for h in holds)}股", "", "", "")]
        rows += [(f"{h['date'].replace('-', '.')} 留倉", f"買 {h['qty']}股", money(h["net"]), "") for h in holds]
        html.append(table(rows))

    # 逐筆明細（台灣時間＝TOS 對帳畫面時間）
    det = ["<table style=\"border-collapse:collapse;font-size:14px;margin-top:10px\">",
           f'<tr><th style="{TD}">台灣時間</th><th style="{TD}">買賣</th><th style="{TD}">成交價</th>'
           f'<th style="{TD}">淨額</th><th style="{TD}">損益</th><th style="{TD}">多空</th></tr>']
    lab = {"day": lambda l: "空單" if l["type_tag"] == "空" else "多單", "hold": lambda l: "留倉(多)",
           "close": lambda l: "平倉(多)"}
    for l in legs:
        det.append(f'<tr><td style="{TD}">{l["time_tw"]}</td><td style="{TD}">{"買" if l["side"] == "BOT" else "賣"}</td>'
                   f'<td style="{TD}text-align:right">{l["price"]:.4f}'.rstrip("0").rstrip(".") + "</td>"
                   f'<td style="{TD}text-align:right">{money(l["net"])}</td>'
                   f'<td style="{TD}text-align:right">{"" if l["pnl"] is None else money(l["pnl"])}</td>'
                   f'<td style="{TD}">{lab[l["kind"]](l)}</td></tr>')
        text.append(f'{l["time_tw"]} {"買" if l["side"] == "BOT" else "賣"} @{l["price"]} 淨額 {l["net"]:,.2f}'
                    + ("" if l["pnl"] is None else f' 損益 {l["pnl"]:+.2f}') + f' {lab[l["kind"]](l)}')
    det.append("</table>")
    html.append("".join(det))
    html.append("</div>")
    print(json.dumps({"subject": ("【待修正】" if flags else "") + (lines[0] if lines else f"{dot} SPY"), "body": "\n".join(text),
                      "html": "".join(html)}, ensure_ascii=False))


if __name__ == "__main__":
    main()
