#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
blank_grid.py - 產生與 trade_review_app.py export PNG「完全同尺寸、同比例」的空白價格格線補頁。

SPY 同一週內隔日跳太多時，相鄰兩天落在不重疊價格帶，並排成牆共用價格軸時跳空處會斷層。
本工具印出只有格線(水平 $1 線 + 垂直 :00/:30 時間線)、價格刻度、時間刻度，但沒有任何 K 線/
交易線的補頁，插進去就能把每欄補到相同價格範圍、讓整週價格軸連續對齊，也能用格子數看出跳空多大。

幾何完全比照 exportPNG()：頁寬 1400(CSS)×DPR3=4200px、每頁 9 格每格 $1、每格高 817/9≈90.78
CSS px(與哪天無關)、ISO 週首日(週一)左軸其餘右軸 → 確保補頁能與真實輸出實體拼接。
完全不 import 也不修改 trade_review_app.py，對原 app 零影響。

輸出：預設存到「這支腳本旁邊」的 blank_out\ (即 D:\\fileserver_D\\TradeReview\\blank_out\\)，不管從哪裡啟動；
       執行完會印出完整路徑。也可用第 8 題 / --out 改。

用法:
  python blank_grid.py                              # 互動式(推薦)：一題一題問，免記 flag
  python blank_grid.py --low 597 --high 606 --date 2026-06-15        # flag 模式
  python blank_grid.py --low 588 --high 612 --date 2026-06-16        # 大跳空自動分頁
  python blank_grid.py --low 593 --high 602 --date 2026-06-15 --prev-close 595.32 --prev-close-date 6/12
  python blank_grid.py --low 597 --high 606 --no-labels --no-caption # 純墊片
"""

import argparse
import math
import os
import datetime

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    import sys
    print("缺少 Pillow 套件。請在命令列執行：  pip install pillow")
    try:
        input("按 Enter 關閉。")
    except EOFError:
        pass
    sys.exit(1)

# === 與 trade_review_app.py export 一致的幾何常數 (CSS px，繪圖再乘 DPR) ===
# 哪天改了 app 的 export 尺寸/格數，只要同步改這區塊，補頁就會繼續對齊。
EXPORT_W       = 1400      # exportPNG: EXPORT_W
DPR            = 3         # exportPNG: EXPORT_DPR
BPD            = 391       # bars per day 09:30-16:00
N_CANDLES      = 390       # 典型一天實際 K 棒(09:30-15:59)；決定 V 線/時間標籤畫到哪
M_T, M_B       = 44, 64    # 上 / 下 margin
EDGE_PAD       = 8         # EDGE_PAD_LEFT / RIGHT
PAGE_GRIDS     = 9         # 每頁 9 格
STEP           = 1         # $1 / 格
REF_TOTAL_BARS = BPD + 60  # 451
PAGE_RANGE     = PAGE_GRIDS * STEP

# 色彩 (exportLightMode 的 T 主題)
C_BG, C_GRID, C_GRIDM, C_TXT = "#FFFFFF", "#666666", "#444444", "#000000"
C_PILL_FILL, C_PILL_BORDER = (245, 215, 110), "#B8860B"

# 預設輸出資料夾：永遠在這支腳本旁邊的 blank_out\，不管從哪裡啟動都存同一處
try:
    _SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
except NameError:
    _SCRIPT_DIR = os.getcwd()
OUT_DEFAULT = os.path.join(_SCRIPT_DIR, "blank_out")


def _derive(week_first):
    """由常數推導頁面尺寸與座標參數 (與 app 相同算法)。"""
    M_l, M_r = (30, 6) if week_first else (6, 30)
    innerW = EXPORT_W - M_l - M_r           # 1364
    refBw  = innerW / REF_TOTAL_BARS        # ≈3.02439
    innerH = PAGE_GRIDS * 30 * refBw        # ≈816.585
    H      = round(innerH + M_T + M_B)      # 925
    cH     = H - M_T - M_B                  # 817
    bwTarget = (innerW - 2 * EDGE_PAD) / BPD  # ppb = 1348/391
    return dict(M_l=M_l, M_r=M_r, W=EXPORT_W, H=H, cW=innerW, cH=cH, ppb=bwTarget)


def y_of(p, pn, px, g):                     # yOf
    return (M_T + (1 - (p - pn) / (px - pn)) * g["cH"]) * DPR


def x_of(li, g):                            # xOfGi(gi(0,li)), panX=EDGE_PAD
    return (g["M_l"] + EDGE_PAD + li * g["ppb"]) * DPR


def _clock(li):
    """li 分鐘後的時鐘 (從 09:30) -> (hh, mm)。"""
    total = 9 * 60 + 30 + li
    return total // 60, total % 60


_FONTS = ["/usr/share/fonts/truetype/dejavu/DejaVuSansMono.ttf",
          "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
          "C:/Windows/Fonts/consola.ttf", "C:/Windows/Fonts/cour.ttf"]
_FONTS_B = ["/usr/share/fonts/truetype/dejavu/DejaVuSansMono-Bold.ttf",
            "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
            "C:/Windows/Fonts/consolab.ttf", "C:/Windows/Fonts/courbd.ttf"]


def _font(size_css, bold=False):
    for p in (_FONTS_B if bold else _FONTS):
        if os.path.exists(p):
            return ImageFont.truetype(p, int(round(size_css * DPR)))
    return ImageFont.load_default()


# 虛線 (PIL 無原生 dash)，比照 canvas setLineDash([1,3]) 乘 DPR 後 on=3 off=9
def _dashed_h(d, x0, x1, y, fill, w, on=DPR, off=3 * DPR):
    x = x0
    while x < x1:
        d.line([(x, y), (min(x + on, x1), y)], fill=fill, width=w)
        x += on + off


def _dashed_v(d, x, y0, y1, fill, w, on=DPR, off=3 * DPR):
    y = y0
    while y < y1:
        d.line([(x, y), (x, min(y + on, y1))], fill=fill, width=w)
        y += on + off


def render_page(pn, px, week_first, date=None, prev_close=None, prev_close_date=None,
                draw_labels=True, draw_caption=True, n_candles=N_CANDLES):
    """回傳一張 PIL Image (device px)。pn/px = 該頁價格下/上界 (整數美元，相差 PAGE_RANGE)。
    week_first=True 左軸；False 右軸。"""
    g = _derive(week_first)
    W, H = g["W"] * DPR, g["H"] * DPR
    img = Image.new("RGB", (W, H), C_BG)
    d = ImageDraw.Draw(img)
    lw = max(1, round(0.7 * DPR))
    left_px, right_px = g["M_l"] * DPR, (g["W"] - g["M_r"]) * DPR
    top_px, bot_px = M_T * DPR, (g["H"] - M_B) * DPR

    # 水平 $1 格線
    gs = math.floor(pn)
    p = gs
    while p <= px + 1e-9:
        y = y_of(p, pn, px, g)
        if -20 * DPR <= y <= H + 20 * DPR:
            _dashed_h(d, left_px, right_px, y, C_GRID, lw)
        p += STEP

    # 垂直時間線 (mm==0 深、mm==30 淺；09:30 有線無標籤)
    for li in range(n_candles):
        hh, mm = _clock(li)
        if mm not in (0, 30):
            continue
        x = x_of(li, g)
        if left_px <= x <= right_px:
            _dashed_v(d, x, top_px, bot_px, C_GRIDM if mm == 0 else C_GRID, lw)

    # Y 軸價格刻度 (整數、13px)
    if draw_labels:
        f_y = _font(13)
        p = gs
        while p <= px + 1e-9:
            y = y_of(p, pn, px, g)
            if y < (M_T - 5) * DPR or y > (g["H"] - M_B + 15) * DPR:
                p += STEP
                continue
            if prev_close is not None and week_first and abs(p - round(prev_close)) < 0.01:
                p += STEP
                continue
            adj = (-6 * DPR) if y > (g["H"] - M_B - 2) * DPR else 0
            txt = str(int(round(p)))
            xL = 2 * DPR if week_first else right_px + 2 * DPR
            d.text((xL, y + 4 * DPR + adj), txt, font=f_y, fill=C_TXT, anchor="lm")
            p += STEP
        # 前一交易日收盤黃標 (僅週首日左軸)
        if prev_close is not None and week_first:
            y_pc = y_of(prev_close, pn, px, g)
            if (M_T - 5) * DPR <= y_pc <= (g["H"] - M_B + 15) * DPR:
                label = ("%s Close: %.2f" % (prev_close_date, prev_close)) if prev_close_date \
                    else ("Close: %.2f" % prev_close)
                f_pc = _font(13, bold=True)
                bb = d.textbbox((0, 0), label, font=f_pc)
                padX, padY = 6 * DPR, 3 * DPR
                tagW, tagH = (bb[2] - bb[0]) + padX * 2, int(round(13 * DPR)) + padY * 2
                tagY = y_pc - tagH / 2
                d.rectangle([0, tagY, tagW, tagY + tagH], fill=C_PILL_FILL,
                            outline=C_PILL_BORDER, width=DPR)
                d.text((padX, y_pc), label, font=f_pc, fill="#000000", anchor="lm")

    # X 軸時間刻度 (12px，mm 0/30，跳過 09:30)
    if draw_labels:
        f_x = _font(12)
        for li in range(n_candles):
            hh, mm = _clock(li)
            if mm not in (0, 30) or li == 0:
                continue
            x = x_of(li, g)
            if left_px <= x <= right_px:
                d.text((x, (g["H"] - M_B + 14) * DPR), "%02d:%02d" % (hh, mm),
                       font=f_x, fill=C_TXT, anchor="mm")

    # 底部說明列：日期 + 價格帶 (標明這是空白補頁)
    if draw_caption:
        parts = []
        if date:
            try:
                dow = ["Mon.", "Tue.", "Wed.", "Thu.", "Fri.", "Sat.", "Sun."][
                    datetime.date.fromisoformat(date).weekday()]
                parts.append("%s %s" % (date, dow))
            except ValueError:
                parts.append(str(date))
        parts.append("blank $%d-$%d" % (int(round(pn)), int(round(px))))
        d.text((g["W"] * DPR / 2, (g["H"] - M_B + 38) * DPR), "   ".join(parts),
               font=_font(16, bold=True), fill="#888888", anchor="mm")
    return img


def paginate(low, high):
    """把 [low, high] 切成連續、不重疊的 9 格頁 (由上往下)。"""
    top, bot = math.ceil(high - 1e-9), math.floor(low + 1e-9)
    if top - bot <= PAGE_RANGE:
        return [(bot, bot + PAGE_RANGE)]
    n = math.ceil((top - bot) / PAGE_RANGE)
    return [(top - PAGE_RANGE * k - PAGE_RANGE, top - PAGE_RANGE * k) for k in range(n)]


def resolve_axis(axis, date):
    """auto：週一->左(週首日)；其餘->右。可被 left/right 覆寫。"""
    if axis == "left":
        return True
    if axis == "right":
        return False
    if date:
        try:
            return datetime.date.fromisoformat(date).weekday() == 0
        except ValueError:
            pass
    return False


def build(low, high, date, axis="auto", out_dir=None, prev_close=None,
          prev_close_date=None, draw_labels=True, draw_caption=True):
    out_dir = out_dir or OUT_DEFAULT
    week_first = resolve_axis(axis, date)
    pages = paginate(low, high)
    os.makedirs(out_dir, exist_ok=True)
    n = len(pages)
    saved = []
    tag = date.replace("-", "") if date else ("%g-%g" % (low, high))
    for i, (pn, px) in enumerate(pages):
        pc = prev_close if (prev_close is not None and pn <= prev_close <= px) else None
        img = render_page(pn, px, week_first, date=date, prev_close=pc,
                          prev_close_date=prev_close_date,
                          draw_labels=draw_labels, draw_caption=draw_caption)
        fn = ("BLANK_%s_%d-%d.png" % (tag, int(pn), int(px))) if n == 1 \
            else ("BLANK_%s_%dof%d_%d-%d.png" % (tag, i + 1, n, int(pn), int(px)))
        path = os.path.abspath(os.path.join(out_dir, fn))
        img.save(path)
        saved.append(path)
    return saved, week_first


def _report(saved, week_first):
    g = _derive(week_first)
    print("軸側：%s    頁面：%d×%d px    每格高 ≈ %.2f CSS px" % (
        "左(週首日)" if week_first else "右", g["W"] * DPR, g["H"] * DPR, g["cH"] / PAGE_GRIDS))
    if saved:
        print("輸出資料夾：" + os.path.dirname(saved[0]))
    print("共 %d 頁：" % len(saved))
    for p in saved:
        print("  " + p)


# === 互動式輸入：直接 python blank_grid.py，一題一題問，免記 flag、免怕打錯 ===
def _ask_float(prompt, required=True, default=None, check=None):
    while True:
        try:
            raw = input(prompt).strip()
        except EOFError:
            return default
        if raw == "":
            if required:
                print("  ! 必填，請輸入數字。")
                continue
            return default
        try:
            v = float(raw)
        except ValueError:
            print("  ! 請輸入數字 (例 597 或 597.5)。")
            continue
        if check:
            err = check(v)
            if err:
                print("  ! " + err)
                continue
        return v


def _ask_str(prompt, default=None, check=None):
    while True:
        try:
            raw = input(prompt).strip()
        except EOFError:
            return default
        if raw == "":
            return default
        if check:
            err = check(raw)
            if err:
                print("  ! " + err)
                continue
        return raw


def _check_date(s):
    try:
        datetime.date.fromisoformat(s)
        return None
    except ValueError:
        return "日期格式須為 YYYY-MM-DD (例 2026-06-15)。"


def interactive():
    print("=" * 52)
    print("  空白價格格線補頁產生器")
    print("=" * 52)
    print("一題一題輸入即可；標示「可留空」的直接按 Enter 跳過。\n")
    low = _ask_float("1) 價格下界 low (例 597): ")
    high = _ask_float("2) 價格上界 high (例 606): ",
                      check=lambda v: ("必須大於 low %.2f" % low) if v <= low else None)
    date = _ask_str("3) 日期 YYYY-MM-DD (例 2026-06-15，可留空): ", check=_check_date)
    axis = _ask_str("4) 軸側 [Enter=自動 / L=左 / R=右] (可留空): ", default="auto",
                    check=lambda s: None if s.lower() in ("l", "left", "r", "right", "auto")
                    else "請輸入 L 或 R，或直接 Enter。")
    axis = {"l": "left", "left": "left", "r": "right", "right": "right",
            "auto": "auto"}[str(axis).lower()]
    prev_close = _ask_float("5) 前一交易日收盤價 (畫黃標，可留空): ", required=False)
    prev_close_date = _ask_str("6) 前收日期 M/D (例 6/12，可留空): ") if prev_close is not None else None
    no_text = str(_ask_str("7) 純墊片、不要任何文字？[y / Enter=否]: ", default="")).lower() in ("y", "yes")
    out = _ask_str("8) 輸出資料夾 (可留空＝%s): " % OUT_DEFAULT, default=OUT_DEFAULT)
    print("\n產生中…")
    saved, week_first = build(low, high, date, axis=axis, out_dir=out, prev_close=prev_close,
                              prev_close_date=prev_close_date,
                              draw_labels=not no_text, draw_caption=not no_text)
    _report(saved, week_first)
    try:
        input("\n完成，按 Enter 關閉。")
    except EOFError:
        pass


def main():
    ap = argparse.ArgumentParser(description="產生與 trade_review_app export 同尺寸的空白價格格線補頁")
    ap.add_argument("--low", type=float, required=True, help="價格帶下界")
    ap.add_argument("--high", type=float, required=True, help="價格帶上界")
    ap.add_argument("--date", type=str, default=None, help="日期 YYYY-MM-DD (決定軸側 auto + 底部標籤)")
    ap.add_argument("--axis", choices=["auto", "left", "right"], default="auto", help="軸側 (預設 auto)")
    ap.add_argument("--prev-close", type=float, default=None, help="前一交易日收盤 (畫黃標，僅左軸頁)")
    ap.add_argument("--prev-close-date", type=str, default=None, help="前收日期 M/D")
    ap.add_argument("--out", type=str, default=OUT_DEFAULT, help="輸出資料夾 (預設：腳本旁的 blank_out)")
    ap.add_argument("--no-labels", action="store_true", help="不畫價格/時間刻度文字")
    ap.add_argument("--no-caption", action="store_true", help="不畫底部日期/價格帶說明")
    args = ap.parse_args()
    if args.high <= args.low:
        ap.error("--high 必須大於 --low")
    saved, week_first = build(args.low, args.high, args.date, axis=args.axis, out_dir=args.out,
                              prev_close=args.prev_close, prev_close_date=args.prev_close_date,
                              draw_labels=not args.no_labels, draw_caption=not args.no_caption)
    _report(saved, week_first)


if __name__ == "__main__":
    import sys
    try:
        if len(sys.argv) > 1:
            main()          # 有給 flag -> flag 模式
        else:
            interactive()   # 直接執行 -> 一題一題問
    except SystemExit:
        raise
    except Exception:
        import traceback
        traceback.print_exc()
        try:
            input("\n發生錯誤，按 Enter 關閉。")
        except EOFError:
            pass
