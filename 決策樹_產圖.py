# -*- coding: utf-8 -*-
"""產生《人工SOP》第 2 節一頁決策樹的 PNG（可列印貼牆）。"""
import os
from PIL import Image, ImageDraw, ImageFont

OUT = os.path.join(os.path.dirname(os.path.abspath(__file__)), "決策樹.png")
W, H = 1760, 2180
BG = (255, 255, 255)

REG = r"C:\Windows\Fonts\msjh.ttc"
BLD = r"C:\Windows\Fonts\msjhbd.ttc"

f_title = ImageFont.truetype(BLD, 46)
f_sub = ImageFont.truetype(REG, 24)
f_node = ImageFont.truetype(BLD, 27)
f_body = ImageFont.truetype(REG, 24)
f_small = ImageFont.truetype(REG, 21)
f_edge = ImageFont.truetype(BLD, 22)
f_tag = ImageFont.truetype(BLD, 23)

INK = (28, 32, 38)
MUTE = (105, 115, 125)
LINE = (120, 130, 140)

GREEN = (34, 120, 62)
GREEN_BG = (233, 246, 236)
AMBER = (188, 106, 8)
AMBER_BG = (253, 243, 226)
RED = (188, 46, 46)
RED_BG = (253, 235, 235)
BLUE = (36, 96, 156)
BLUE_BG = (233, 241, 250)
GREY = (96, 106, 116)
GREY_BG = (243, 245, 247)

img = Image.new("RGB", (W, H), BG)
d = ImageDraw.Draw(img)


def wtext(txt, font):
    return d.textbbox((0, 0), txt, font=font)[2]


def box(cx, cy, w, h, lines, border, fill, tag=None, tagcol=None, radius=16):
    """lines = [(text, font, colour), ...] 垂直置中。回傳 (top, bottom)。"""
    x0, y0, x1, y1 = cx - w // 2, cy - h // 2, cx + w // 2, cy + h // 2
    d.rounded_rectangle([x0, y0, x1, y1], radius=radius, fill=fill,
                        outline=border, width=3)
    total = sum(f.size + 9 for _, f, _ in lines) - 9
    y = cy - total // 2
    for t, f, c in lines:
        d.text((cx - wtext(t, f) // 2, y), t, font=f, fill=c)
        y += f.size + 9
    if tag:
        tw = wtext(tag, f_tag)
        d.rounded_rectangle([x0 + 14, y0 - 17, x0 + 14 + tw + 22, y0 + 17],
                            radius=15, fill=tagcol or border)
        d.text((x0 + 25, y0 - 13), tag, font=f_tag, fill=(255, 255, 255))
    return y0, y1


def diamond(cx, cy, w, h, lines, border, fill):
    d.polygon([(cx, cy - h // 2), (cx + w // 2, cy), (cx, cy + h // 2),
               (cx - w // 2, cy)], fill=fill, outline=border)
    for k in range(3):                      # 加粗邊框
        d.polygon([(cx, cy - h // 2 + k), (cx + w // 2 - k, cy),
                   (cx, cy + h // 2 - k), (cx - w // 2 + k, cy)],
                  outline=border)
    total = sum(f.size + 8 for _, f, _ in lines) - 8
    y = cy - total // 2
    for t, f, c in lines:
        d.text((cx - wtext(t, f) // 2, y), t, font=f, fill=c)
        y += f.size + 8
    return cy - h // 2, cy + h // 2


def head(x, y, dx, dy, col, size=11):
    """在 (x,y) 畫指向 (dx,dy) 方向的箭頭"""
    if dy:                                   # 垂直
        s = 1 if dy > 0 else -1
        d.polygon([(x, y), (x - size, y - s * size * 1.5),
                   (x + size, y - s * size * 1.5)], fill=col)
    else:                                    # 水平
        s = 1 if dx > 0 else -1
        d.polygon([(x, y), (x - s * size * 1.5, y - size),
                   (x - s * size * 1.5, y + size)], fill=col)


def vline(x, y0, y1, col=LINE, w=3, arrow=True):
    d.line([(x, y0), (x, y1)], fill=col, width=w)
    if arrow:
        head(x, y1, 0, 1 if y1 > y0 else -1, col)


def hline(y, x0, x1, col=LINE, w=3, arrow=True):
    d.line([(x0, y), (x1, y)], fill=col, width=w)
    if arrow:
        head(x1, y, 1 if x1 > x0 else -1, 0, col)


def elbow(x0, y0, x1, y1, col=LINE, w=3, first="v"):
    if first == "v":
        d.line([(x0, y0), (x0, y1)], fill=col, width=w)
        d.line([(x0, y1), (x1, y1)], fill=col, width=w)
        head(x1, y1, 1 if x1 > x0 else -1, 0, col)
    else:
        d.line([(x0, y0), (x1, y0)], fill=col, width=w)
        d.line([(x1, y0), (x1, y1)], fill=col, width=w)
        head(x1, y1, 0, 1 if y1 > y0 else -1, col)


def label(x, y, txt, col=INK, anchor="c"):
    tw = wtext(txt, f_edge)
    x0 = x - tw // 2 if anchor == "c" else (x if anchor == "l" else x - tw)
    d.rectangle([x0 - 9, y - 5, x0 + tw + 9, y + f_edge.size + 5], fill=BG)
    d.text((x0, y), txt, font=f_edge, fill=col)


# ── 標題 ────────────────────────────────────────────────────────────────────
d.rectangle([0, 0, W, 118], fill=(38, 48, 60))
d.text((56, 24), "SPY 當沖每日複盤 — 每日決策樹", font=f_title, fill=(255, 255, 255))
d.text((58, 78), "一頁看完：今天的資料從哪來、遇到狀況怎麼走。細節見《人工SOP_每日取料與判斷.md》",
       font=f_sub, fill=(178, 190, 202))

L, M, R = 350, 880, 1420          # 三個欄位中心

# ── 前提：今天到底要不要做（第 1 節） ──────────────────────────────────────
box(240, 330, 400, 228, [
    ("先確認今天要不要做", f_node, INK),
    ("", f_small, INK),
    ("台北 週二～週六 早上 → 要做", f_body, GREEN),
    ("台北 週日、週一 → 不用做", f_body, RED),
    ("", f_small, INK),
    ("老闆當天沒下單 → 也不用做", f_small, MUTE),
], GREY, GREY_BG)

# ── 開始 ────────────────────────────────────────────────────────────────────
box(M, 190, 330, 66, [("開始（台北早上）", f_node, INK)], GREY, GREY_BG, radius=33)
vline(M, 223, 258)

# ① Gmail
box(M, 340, 830, 128, [
    ("① 打開 Gmail：shuye1668@gmail.com →「草稿」匣", f_node, INK),
    ("找主旨  [SPY-DayTrade-Autosend] YYYY-MM-DD", f_body, INK),
    ("日期通常＝今天的台北日期；不確定就取最新那封", f_small, MUTE),
], BLUE, BLUE_BG, tag="①", tagcol=BLUE)
vline(M, 404, 442)

# ◆ 有草稿嗎
diamond(M, 510, 420, 128, [("有找到草稿嗎？", f_node, INK)], BLUE, BLUE_BG)

# 右：沒有 → ④
d.line([(M + 210, 510), (R, 510)], fill=LINE, width=3)
d.line([(R, 510), (R, 604)], fill=LINE, width=3)
head(R, 604, 0, 1, LINE)
label(1170, 476, "沒有／主旨含 FAILED", RED)

# 左：有 → ②
d.line([(M - 210, 510), (600, 510)], fill=LINE, width=3)
d.line([(600, 510), (600, 604)], fill=LINE, width=3)
head(600, 604, 0, 1, LINE)
label(700, 476, "有草稿", GREEN)

# ◆ ② 有 CSV 區段
diamond(600, 676, 470, 144, [
    ("② 內文有這兩行嗎？", f_node, INK),
    ("===== CSV CONTENT START / END =====", f_small, INK),
], BLUE, BLUE_BG)

# ④ 老闆電腦
box(R, 676, 480, 128, [
    ("④ 去老闆電腦重跑匯出", f_node, INK),
    ("（第 6 節・實際步驟待填）", f_body, AMBER),
    ("老闆正在交易就不要動", f_small, MUTE),
], AMBER, AMBER_BG, tag="④", tagcol=AMBER)

# ④ 兩個出口
d.line([(R, 740), (R, 790)], fill=LINE, width=3)
d.line([(1640, 790), (R, 790)], fill=LINE, width=3)
d.line([(1640, 790), (1640, 340)], fill=LINE, width=3)
d.line([(1640, 340), (1298, 340)], fill=LINE, width=3)
head(1298, 340, -1, 0, LINE)
label(1520, 800, "跑起來了 → 回到 ①", GREEN, anchor="c")
d.line([(R, 790), (R, 880)], fill=LINE, width=3)
d.line([(R, 880), (M + 150, 880)], fill=LINE, width=3)
d.line([(M + 150, 880), (M + 150, 966)], fill=LINE, width=3)
head(M + 150, 966, 0, 1, LINE)
label(1240, 892, "跑不動／老闆在用", AMBER)

# ② 三個出口
d.line([(365, 676), (L, 676)], fill=LINE, width=3)
d.line([(L, 676), (L, 838)], fill=LINE, width=3)
head(L, 838, 0, 1, LINE)
label(292, 716, "有", GREEN, anchor="c")

d.line([(600, 748), (600, 908)], fill=LINE, width=3)
d.line([(600, 908), (M - 150, 908)], fill=LINE, width=3)
d.line([(M - 150, 908), (M - 150, 966)], fill=LINE, width=3)
head(M - 150, 966, 0, 1, LINE)
label(612, 800, "只有截圖／OCR", AMBER, anchor="l")

d.line([(835, 676), (1090, 676)], fill=LINE, width=3)
d.line([(1090, 676), (1090, 652)], fill=LINE, width=3)
d.line([(1090, 652), (1180, 652)], fill=LINE, width=3)
head(1180, 652, 1, 0, LINE)
label(960, 588, "什麼都沒有", RED)

# ③ 存 CSV（左線）
box(L, 902, 500, 128, [
    ("③ 複製兩行「中間」整段", f_node, INK),
    ("存成 _inbox\\YYYY-MM-DD.csv", f_body, INK),
    ("UTF-8・存檔類型選「所有檔案」", f_small, MUTE),
], GREEN, GREEN_BG, tag="③ 最常見 9 成", tagcol=GREEN)

# ⑤ 截圖
box(M, 1022, 520, 112, [
    ("⑤ 截圖對帳單", f_node, INK),
    ("BAL 行／全部 TRD 列／Misc Fees", f_body, INK),
    ("／BALANCE／午夜後那幾列都要拍到", f_small, MUTE),
], AMBER, AMBER_BG, tag="⑤", tagcol=AMBER)
vline(M, 1078, 1128)

# ⑥ 貼 prompt
box(M, 1250, 820, 232, [
    ("⑥ 在 502 這台（放 D:\\fileserver_D\\TradeReview\\ 的那台）", f_node, INK),
    ("開 Claude 對話，貼 PROMPT_轉trades_all.md", f_node, INK),
    ("整段 ＋ 截圖  （第 7 節）", f_body, INK),
    ("→ AI 更新 trades_all.xlsx", f_body, GREEN),
    ("→ 並回報「期初餘額／收盤餘額」兩個數字", f_body, GREEN),
    ("不在這台 → 改用 PROMPT_截圖轉CSV_雲端版.md（7.5）", f_small, AMBER),
], AMBER, AMBER_BG, tag="⑥ 標準做法", tagcol=AMBER)

# ⑥ → ⑧ 手抄
d.line([(1290, 1250), (1360, 1250)], fill=LINE, width=3)
d.line([(1360, 1250), (1360, 1387)], fill=LINE, width=3)
head(1360, 1387, 0, 1, LINE)
label(1300, 1196, "AI 也不能用", RED, anchor="l")

box(1420, 1452, 470, 130, [
    ("⑦ 逐列手抄成 CSV", f_node, INK),
    ("（第 8 節・最後手段）", f_body, RED),
    ("抄錯不用怕：引擎逐列驗餘額會擋下", f_small, MUTE),
], RED, RED_BG, tag="⑦", tagcol=RED)

# 匯流到 ★
MERGE_Y = 1560
d.line([(L, 966), (L, MERGE_Y)], fill=LINE, width=3)
d.line([(L, MERGE_Y), (M - 340, MERGE_Y)], fill=LINE, width=3)
d.line([(M, 1345), (M, MERGE_Y)], fill=LINE, width=3)
d.line([(1420, 1517), (1420, MERGE_Y)], fill=LINE, width=3)
d.line([(1420, MERGE_Y), (M + 340, MERGE_Y)], fill=LINE, width=3)
d.line([(M - 340, MERGE_Y), (M + 340, MERGE_Y)], fill=LINE, width=3)
vline(M, MERGE_Y, 1622)

# ★ 一鍵檔
box(M, 1728, 940, 176, [
    ("★  雙擊  D:\\fileserver_D\\TradeReview\\每日一鍵複盤.bat", f_node, INK),
    ("", f_small, INK),
    ("手上有 CSV（③ 或 ⑦）→ 它自己挑檔、自己跑完", f_body, INK),
    ("只有 trades_all 更新了（⑥）→ 選單選 2，輸入 交易日／期初／收盤", f_body, INK),
], GREEN, GREEN_BG, tag="★", tagcol=GREEN)
vline(M, 1816, 1854)

# ◆ 有紅旗嗎
diamond(M, 1926, 430, 128, [("過程中有紅旗嗎？", f_node, INK)], BLUE, BLUE_BG)

d.line([(M - 215, 1926), (470, 1926)], fill=LINE, width=3)
d.line([(470, 1926), (470, 2004)], fill=LINE, width=3)
head(470, 2004, 0, 1, LINE)
label(600, 1892, "沒有", GREEN)

d.line([(M + 215, 1926), (1300, 1926)], fill=LINE, width=3)
d.line([(1300, 1926), (1300, 2004)], fill=LINE, width=3)
head(1300, 2004, 0, 1, LINE)
label(1170, 1892, "有", RED)

box(470, 2076, 620, 128, [
    ("開 http://localhost:5500 看複盤", f_node, INK),
    ("貼總結行 → 完成", f_body, GREEN),
    ("看完把 TradeReview App 視窗關掉", f_small, MUTE),
], GREEN, GREEN_BG)
d.ellipse([196, 2052, 224, 2080], fill=GREEN)
d.line([(202, 2066), (208, 2073), (218, 2059)], fill=(255, 255, 255), width=4)

box(1300, 2076, 700, 128, [
    ("停手。兩本帳沒有被改到，不要自己修", f_node, INK),
    ("截圖 ＋ _logs\\engine_last.txt、writer_last.txt", f_body, INK),
    ("交負責人（第 9 節）", f_body, RED),
], RED, RED_BG)
d.ellipse([938, 2052, 966, 2080], fill=RED)

img.save(OUT)
print("OK", OUT, os.path.getsize(OUT), "bytes", img.size)
