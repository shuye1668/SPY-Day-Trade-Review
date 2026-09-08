#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""美股（NYSE/NASDAQ）交易日曆 —— 判斷某天是否休市、以及休市原因。

為什麼需要（2026-09-08 加）：
原本 app 的主動補 K 線只排除週末，遇到平日休市（例如 2026-09-07 勞動節）會照樣
去問 yfinance，抓到空值後靜默重試三次然後放棄。結果是使用者看到「最新只到 9/4」
卻無從分辨那是「正確，因為 9/7 休市」還是「壞了，沒補進來」。
把休市日變成系統「知道」的事實，安靜就有了意義。

只處理「全天休市」。半日盤（感恩節隔日、聖誕夜等 13:00 收盤）仍是交易日，
資料完整，不需特別處理 —— 補 K 線一律等 16:05 ET 之後才收，半日盤早就結束了。
"""
import datetime as dt

__all__ = ["market_closed_reason", "us_market_holidays", "is_trading_day",
           "latest_completed_session", "next_trading_day"]


def _easter(year):
    """復活節（Anonymous Gregorian algorithm）。耶穌受難日 = 復活節前兩天。"""
    a = year % 19
    b, c = year // 100, year % 100
    d, e = b // 4, b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i, k = c // 4, c % 4
    L = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * L) // 451
    month = (h + L - 7 * m + 114) // 31
    day = ((h + L - 7 * m + 114) % 31) + 1
    return dt.date(year, month, day)


def _nth_weekday(year, month, weekday, n):
    """該月第 n 個 <weekday>（weekday: 週一=0）。"""
    first = dt.date(year, month, 1)
    offset = (weekday - first.weekday()) % 7
    return first + dt.timedelta(days=offset + 7 * (n - 1))


def _last_weekday(year, month, weekday):
    """該月最後一個 <weekday>。"""
    nxt = dt.date(year + 1, 1, 1) if month == 12 else dt.date(year, month + 1, 1)
    last = nxt - dt.timedelta(days=1)
    return last - dt.timedelta(days=(last.weekday() - weekday) % 7)


def _observed(d):
    """NYSE 慣例：落在週六提前到週五，落在週日順延到週一。"""
    if d.weekday() == 5:
        return d - dt.timedelta(days=1)
    if d.weekday() == 6:
        return d + dt.timedelta(days=1)
    return d


_cache = {}


def us_market_holidays(year):
    """{date: 名稱} —— 該年度全天休市日。"""
    if year in _cache:
        return _cache[year]
    h = {
        _observed(dt.date(year, 1, 1)): "元旦",
        _nth_weekday(year, 1, 0, 3): "馬丁路德金恩日",
        _nth_weekday(year, 2, 0, 3): "總統日",
        _easter(year) - dt.timedelta(days=2): "耶穌受難日",
        _last_weekday(year, 5, 0): "陣亡將士紀念日",
        _observed(dt.date(year, 6, 19)): "六月節",
        _observed(dt.date(year, 7, 4)): "美國國慶",
        _nth_weekday(year, 9, 0, 1): "勞動節",
        _nth_weekday(year, 11, 3, 4): "感恩節",
        _observed(dt.date(year, 12, 25)): "聖誕節",
    }
    # 六月節自 2021 年才成為聯邦假日，之前美股照常交易
    if year < 2021:
        h.pop(_observed(dt.date(year, 6, 19)), None)
    _cache[year] = h
    return h


def market_closed_reason(d):
    """休市原因字串；正常交易日回 None。d 可為 date 或 'YYYY-MM-DD'。"""
    if isinstance(d, str):
        d = dt.datetime.strptime(d, "%Y-%m-%d").date()
    elif isinstance(d, dt.datetime):
        d = d.date()
    if d.weekday() >= 5:
        return "週末"
    return us_market_holidays(d.year).get(d)


def is_trading_day(d):
    return market_closed_reason(d) is None


def next_trading_day(d):
    """d 之後的下一個交易日。"""
    if isinstance(d, str):
        d = dt.datetime.strptime(d, "%Y-%m-%d").date()
    d += dt.timedelta(days=1)
    while not is_trading_day(d):
        d += dt.timedelta(days=1)
    return d


def latest_completed_session(now_et, close_hour=16, close_min=5):
    """給美東現在時間，回傳「最近一個已經收盤的交易日」。

    收盤緩衝 16:05：盤中抓只會拿到半天 bar，而寫入端見日期已存在就不再更新，
    那筆殘缺資料會被永久凍結。
    """
    d = now_et.date()
    if now_et < dt.datetime.combine(d, dt.time(close_hour, close_min)):
        d -= dt.timedelta(days=1)
    while not is_trading_day(d):
        d -= dt.timedelta(days=1)
    return d
