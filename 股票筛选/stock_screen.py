#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
简单 A 股技术筛选工具。

用途：
1. 输入股票池、股票代码列表或方向主题；
2. 拉取公开日 K 行情；
3. 按趋势、量价、启动信号、位置、风险输出观察报告。

说明：本脚本只用于观察和复盘，不构成任何买卖建议。
"""

from __future__ import annotations

import argparse
import csv
import datetime as dt
import importlib
import json
import math
import sys
from pathlib import Path
from typing import Any

import numpy as np
import pandas as pd
import requests


BASE_DIR = Path(__file__).resolve().parent
PROJECT_DIR = BASE_DIR.parent
DEFAULT_OUTPUT_DIR = BASE_DIR / "output"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0 Safari/537.36"
    )
}

LOCAL_THEME_POOLS: dict[str, list[dict[str, str]]] = {
    "半导体材料": [
        {"code": "603688", "name": "石英股份", "theme": "半导体材料/光伏材料"},
        {"code": "300666", "name": "江丰电子", "theme": "半导体靶材"},
        {"code": "300054", "name": "鼎龙股份", "theme": "CMP材料/半导体材料"},
        {"code": "688019", "name": "安集科技", "theme": "CMP抛光液"},
        {"code": "688268", "name": "华特气体", "theme": "电子特气"},
        {"code": "688126", "name": "沪硅产业", "theme": "半导体硅片"},
        {"code": "002409", "name": "雅克科技", "theme": "电子材料/前驱体"},
    ],
    "新能源材料": [
        {"code": "300750", "name": "宁德时代", "theme": "动力电池"},
        {"code": "002812", "name": "恩捷股份", "theme": "隔膜"},
        {"code": "002709", "name": "天赐材料", "theme": "电解液"},
        {"code": "300037", "name": "新宙邦", "theme": "电解液/氟化工"},
        {"code": "300073", "name": "当升科技", "theme": "正极材料"},
    ],
    "锂电材料": [
        {"code": "002466", "name": "天齐锂业", "theme": "锂资源"},
        {"code": "002460", "name": "赣锋锂业", "theme": "锂资源"},
        {"code": "002709", "name": "天赐材料", "theme": "电解液"},
        {"code": "300073", "name": "当升科技", "theme": "正极材料"},
        {"code": "300568", "name": "星源材质", "theme": "隔膜"},
    ],
    "AI金属": [
        {"code": "000657", "name": "中钨高新", "theme": "AI金属/钨材料"},
        {"code": "600549", "name": "厦门钨业", "theme": "钨/稀土"},
        {"code": "000960", "name": "锡业股份", "theme": "锡材料"},
        {"code": "600497", "name": "驰宏锌锗", "theme": "锗/有色金属"},
        {"code": "000831", "name": "中国稀土", "theme": "稀土"},
    ],
    "AIDC硬件": [
        {"code": "601138", "name": "工业富联", "theme": "AI服务器"},
        {"code": "300476", "name": "胜宏科技", "theme": "AI服务器PCB"},
        {"code": "002463", "name": "沪电股份", "theme": "服务器高速板"},
        {"code": "300308", "name": "中际旭创", "theme": "光模块/CPO"},
        {"code": "300502", "name": "新易盛", "theme": "光模块/CPO"},
        {"code": "000063", "name": "中兴通讯", "theme": "通信设备/算力网络"},
    ],
    "稀土": [
        {"code": "600111", "name": "北方稀土", "theme": "轻稀土"},
        {"code": "000831", "name": "中国稀土", "theme": "中重稀土"},
        {"code": "600549", "name": "厦门钨业", "theme": "钨/稀土"},
        {"code": "000795", "name": "英洛华", "theme": "稀土永磁"},
        {"code": "300748", "name": "金力永磁", "theme": "稀土永磁"},
    ],
}


def get_akshare() -> Any | None:
    """懒加载 AkShare，未安装或导入失败时返回 None。"""
    try:
        return importlib.import_module("akshare")
    except Exception:
        return None


def normalize_code(code: Any) -> str:
    """统一股票代码格式，保留 6 位 A 股代码。"""
    text = str(code).strip().upper()
    text = text.replace(".SH", "").replace(".SZ", "").replace(".BJ", "")
    digits = "".join(ch for ch in text if ch.isdigit())
    return digits.zfill(6)[-6:] if digits else text


def market_prefix(code: str) -> str:
    """返回腾讯接口需要的市场前缀。"""
    code = normalize_code(code)
    if code.startswith(("6", "9")):
        return "sh"
    if code.startswith(("8", "4")):
        return "bj"
    return "sz"


def eastmoney_secid(code: str) -> str:
    """返回东方财富接口 secid。"""
    code = normalize_code(code)
    if code.startswith(("6", "9")):
        market = "1"
    elif code.startswith(("8", "4")):
        market = "0"
    else:
        market = "0"
    return f"{market}.{code}"


def read_name_cache() -> dict[str, dict[str, str]]:
    """读取仓库已有股票简称缓存，用于补全 --stocks 输入的名称。"""
    cache_path = PROJECT_DIR / "stock_name_cache.json"
    if not cache_path.exists():
        return {}
    try:
        with cache_path.open("r", encoding="utf-8") as f:
            raw = json.load(f)
        result: dict[str, dict[str, str]] = {}
        for item in raw.values():
            code = normalize_code(item.get("code", ""))
            if code:
                result[code] = {
                    "code": code,
                    "name": item.get("name", code),
                    "theme": item.get("direction", ""),
                }
        return result
    except Exception:
        return {}


def read_input_pool(path: Path) -> list[dict[str, str]]:
    """读取 CSV 股票池，要求至少包含 code，可选 name、theme。"""
    rows: list[dict[str, str]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        if not reader.fieldnames or "code" not in reader.fieldnames:
            raise ValueError("输入 CSV 至少需要包含 code 列，可选 name、theme 列。")
        for row in reader:
            code = normalize_code(row.get("code", ""))
            if not code:
                continue
            rows.append(
                {
                    "code": code,
                    "name": (row.get("name") or code).strip(),
                    "theme": (row.get("theme") or "").strip(),
                }
            )
    return dedupe_pool(rows)


def parse_stocks_arg(stocks: str) -> list[dict[str, str]]:
    """解析 --stocks 代码列表，并尽量从本地缓存补全名称。"""
    cache = read_name_cache()
    rows: list[dict[str, str]] = []
    for part in stocks.replace("，", ",").split(","):
        code = normalize_code(part)
        if not code:
            continue
        cached = cache.get(code, {})
        rows.append(
            {
                "code": code,
                "name": cached.get("name", code),
                "theme": cached.get("theme", "手动输入"),
            }
        )
    return dedupe_pool(rows)


def dedupe_pool(pool: list[dict[str, str]]) -> list[dict[str, str]]:
    """按股票代码去重，保留首次出现的主题说明。"""
    seen: set[str] = set()
    result: list[dict[str, str]] = []
    for item in pool:
        code = normalize_code(item.get("code", ""))
        if not code or code in seen:
            continue
        seen.add(code)
        result.append(
            {
                "code": code,
                "name": item.get("name") or code,
                "theme": item.get("theme") or "",
            }
        )
    return result


def resolve_theme_pool(theme: str, limit: int = 80) -> list[dict[str, str]]:
    """根据方向查找股票池：先尝试 AkShare 东方财富板块，再使用本地预设。"""
    theme = theme.strip()
    remote_pool = fetch_theme_pool_from_akshare(theme, limit=limit)
    if remote_pool:
        return remote_pool

    for key, pool in LOCAL_THEME_POOLS.items():
        if theme in key or key in theme:
            return dedupe_pool(pool[:limit])

    raise ValueError(
        f"未能自动找到方向“{theme}”对应股票池，请改用 --stocks 或 --input 提供股票列表。"
    )


def fetch_theme_pool_from_akshare(theme: str, limit: int = 80) -> list[dict[str, str]]:
    """通过 AkShare 查询东方财富概念/行业板块成分股。"""
    ak = get_akshare()
    if ak is None:
        return []

    board_jobs = [
        ("概念", "stock_board_concept_name_em", "stock_board_concept_cons_em"),
        ("行业", "stock_board_industry_name_em", "stock_board_industry_cons_em"),
    ]
    for board_type, name_func, cons_func in board_jobs:
        try:
            name_df = getattr(ak, name_func)()
            if name_df.empty:
                continue
            name_col = "板块名称" if "板块名称" in name_df.columns else name_df.columns[0]
            matched = name_df[name_df[name_col].astype(str).str.contains(theme, na=False)]
            if matched.empty:
                matched = name_df[name_df[name_col].astype(str).apply(lambda x: x in theme)]
            if matched.empty:
                continue
            board_name = str(matched.iloc[0][name_col])
            cons_df = getattr(ak, cons_func)(symbol=board_name)
            if cons_df.empty:
                continue
            code_col = "代码" if "代码" in cons_df.columns else "证券代码"
            name_col2 = "名称" if "名称" in cons_df.columns else "证券简称"
            rows = []
            for _, row in cons_df.head(limit).iterrows():
                rows.append(
                    {
                        "code": normalize_code(row.get(code_col, "")),
                        "name": str(row.get(name_col2, "")),
                        "theme": f"{board_type}:{board_name}",
                    }
                )
            return dedupe_pool(rows)
        except Exception:
            continue
    return []


def fetch_kline(code: str, days: int = 160) -> tuple[pd.DataFrame | None, str, str | None]:
    """按 AkShare、东方财富、腾讯财经顺序获取日 K 数据。"""
    errors: list[str] = []
    fetchers = [
        ("AkShare", fetch_kline_akshare),
        ("东方财富", fetch_kline_eastmoney),
        ("腾讯财经", fetch_kline_tencent),
    ]
    for source, fetcher in fetchers:
        try:
            df = fetcher(code, days)
            df = normalize_kline_df(df)
            if df is not None and not df.empty:
                return df.tail(days).reset_index(drop=True), source, None
            errors.append(f"{source}: 空数据")
        except Exception as exc:
            errors.append(f"{source}: {exc}")
    return None, "", "；".join(errors)


def fetch_kline_akshare(code: str, days: int) -> pd.DataFrame:
    """使用 AkShare 获取东方财富日 K。"""
    ak = get_akshare()
    if ak is None:
        raise RuntimeError("未安装 akshare")
    start = (dt.date.today() - dt.timedelta(days=max(260, days * 3))).strftime("%Y%m%d")
    df = ak.stock_zh_a_hist(
        symbol=normalize_code(code),
        period="daily",
        start_date=start,
        end_date="20500101",
        adjust="qfq",
    )
    rename_map = {
        "日期": "date",
        "开盘": "open",
        "收盘": "close",
        "最高": "high",
        "最低": "low",
        "成交量": "volume",
        "成交额": "amount",
        "涨跌幅": "pct_chg",
        "换手率": "turnover",
    }
    return df.rename(columns=rename_map)


def fetch_kline_eastmoney(code: str, days: int) -> pd.DataFrame:
    """使用东方财富公开 K 线接口获取日 K。"""
    url = "https://push2his.eastmoney.com/api/qt/stock/kline/get"
    params = {
        "secid": eastmoney_secid(code),
        "fields1": "f1,f2,f3,f4,f5,f6",
        "fields2": "f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61",
        "klt": "101",
        "fqt": "1",
        "beg": "19900101",
        "end": "20500101",
        "lmt": str(max(days, 160)),
    }
    resp = requests.get(url, params=params, headers=HEADERS, timeout=12)
    resp.raise_for_status()
    data = resp.json().get("data") or {}
    klines = data.get("klines") or []
    rows = []
    for line in klines:
        parts = line.split(",")
        if len(parts) < 11:
            continue
        rows.append(
            {
                "date": parts[0],
                "open": parts[1],
                "close": parts[2],
                "high": parts[3],
                "low": parts[4],
                "volume": parts[5],
                "amount": parts[6],
                "pct_chg": parts[8],
                "turnover": parts[10],
            }
        )
    return pd.DataFrame(rows)


def fetch_kline_tencent(code: str, days: int) -> pd.DataFrame:
    """使用腾讯财经公开 K 线接口作为备用数据源。"""
    symbol = f"{market_prefix(code)}{normalize_code(code)}"
    url = "https://web.ifzq.gtimg.cn/appstock/app/fqkline/get"
    params = {"param": f"{symbol},day,,,{max(days, 160)},qfq"}
    resp = requests.get(url, params=params, headers=HEADERS, timeout=12)
    resp.raise_for_status()
    data = resp.json().get("data", {}).get(symbol, {})
    lines = data.get("qfqday") or data.get("day") or []
    rows = []
    for parts in lines:
        if len(parts) < 6:
            continue
        rows.append(
            {
                "date": parts[0],
                "open": parts[1],
                "close": parts[2],
                "high": parts[3],
                "low": parts[4],
                "volume": parts[5],
                "amount": parts[6] if len(parts) > 6 else np.nan,
            }
        )
    return pd.DataFrame(rows)


def normalize_kline_df(df: pd.DataFrame | None) -> pd.DataFrame | None:
    """清洗行情字段，转为统一的 date/open/high/low/close/volume/amount。"""
    if df is None or df.empty:
        return None
    required = ["date", "open", "high", "low", "close", "volume"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(f"行情缺少字段: {missing}")

    result = df.copy()
    result["date"] = pd.to_datetime(result["date"], errors="coerce")
    for col in ["open", "high", "low", "close", "volume", "amount", "pct_chg", "turnover"]:
        if col in result.columns:
            result[col] = pd.to_numeric(result[col], errors="coerce")
    if "amount" not in result.columns:
        result["amount"] = np.nan
    if "pct_chg" not in result.columns or result["pct_chg"].isna().all():
        result["pct_chg"] = result["close"].pct_change() * 100
    if "turnover" not in result.columns:
        result["turnover"] = np.nan

    result = result.dropna(subset=["date", "open", "high", "low", "close"])
    result = result.sort_values("date").drop_duplicates("date")
    return result.reset_index(drop=True)


def add_indicators(df: pd.DataFrame) -> pd.DataFrame:
    """计算 MA、MACD、ATR、涨幅、成交均值等核心指标。"""
    data = df.copy()
    data["ma5"] = data["close"].rolling(5).mean()
    data["ma10"] = data["close"].rolling(10).mean()
    data["ma20"] = data["close"].rolling(20).mean()

    ema12 = data["close"].ewm(span=12, adjust=False).mean()
    ema26 = data["close"].ewm(span=26, adjust=False).mean()
    data["dif"] = ema12 - ema26
    data["dea"] = data["dif"].ewm(span=9, adjust=False).mean()
    data["macd"] = 2 * (data["dif"] - data["dea"])

    prev_close = data["close"].shift(1)
    tr1 = data["high"] - data["low"]
    tr2 = (data["high"] - prev_close).abs()
    tr3 = (data["low"] - prev_close).abs()
    data["tr"] = pd.concat([tr1, tr2, tr3], axis=1).max(axis=1)
    data["atr"] = data["tr"].rolling(14).mean()

    data["gain_20"] = (data["close"] / data["close"].shift(20) - 1) * 100
    data["gain_60"] = (data["close"] / data["close"].shift(60) - 1) * 100
    data["avg_vol_5"] = data["volume"].rolling(5).mean()
    data["avg_amount_20"] = data["amount"].rolling(20).mean()
    return data


def as_float(value: Any, default: float = 0.0) -> float:
    """安全转 float，避免 NaN 进入评分。"""
    try:
        number = float(value)
        if math.isnan(number) or math.isinf(number):
            return default
        return number
    except Exception:
        return default


def pct(value: Any, digits: int = 2) -> float:
    """保留百分比小数。"""
    return round(as_float(value), digits)


def analyze_trend(df: pd.DataFrame) -> dict[str, Any]:
    """判断趋势结构：MA20、均线排列、MACD 和趋势破坏。"""
    last = df.iloc[-1]
    prev = df.iloc[-2] if len(df) > 1 else last
    ma20_ref = df["ma20"].iloc[-6] if len(df) >= 26 else np.nan
    avg_vol5 = as_float(last.get("avg_vol_5"))

    ma20_up = as_float(last["ma20"]) > as_float(ma20_ref)
    close_above_ma20 = as_float(last["close"]) >= as_float(last["ma20"])
    bull_order = as_float(last["ma5"]) > as_float(last["ma10"]) > as_float(last["ma20"])
    macd_cross = as_float(prev["dif"]) <= as_float(prev["dea"]) and as_float(last["dif"]) > as_float(last["dea"])
    macd_near_zero = as_float(last["dif"]) >= -as_float(last["close"]) * 0.01
    gain20_positive = as_float(last.get("gain_20")) > 0
    markup = ma20_up and close_above_ma20 and (bull_order or gain20_positive)
    trend_broken = as_float(last["close"]) < as_float(last["ma20"]) and as_float(last["pct_chg"]) < -1
    volume_break_ma20 = (
        as_float(last["close"]) < as_float(last["ma20"])
        and as_float(last["pct_chg"]) <= -2
        and avg_vol5 > 0
        and as_float(last["volume"]) > avg_vol5 * 1.3
    )

    score = 40
    reasons: list[str] = []
    if ma20_up:
        score += 15
        reasons.append("MA20 向上")
    else:
        reasons.append("MA20 尚未向上")
    if close_above_ma20:
        score += 15
        reasons.append("收盘站上 MA20")
    else:
        score -= 15
        reasons.append("收盘跌破 MA20")
    if bull_order:
        score += 15
        reasons.append("MA5/MA10/MA20 多头排列")
    if macd_cross:
        score += 10
        reasons.append("MACD 金叉")
    elif as_float(last["dif"]) > as_float(last["dea"]):
        score += 6
        reasons.append("MACD 维持多头")
    if macd_near_zero:
        score += 5
        reasons.append("MACD 位于零轴附近或上方")
    if gain20_positive:
        score += 8
        reasons.append("近 20 日涨幅为正")
    if trend_broken:
        score -= 25
        reasons.append("趋势有破坏迹象")
    if volume_break_ma20:
        score = min(score, 35)
        reasons.append("放量跌破 MA20")

    score = int(max(0, min(100, score)))
    status = "趋势良好" if score >= 80 else "趋势可观察" if score >= 65 else "趋势一般" if score >= 50 else "趋势偏弱"
    return {
        "trend_score": score,
        "trend_status": status,
        "ma20_up": ma20_up,
        "close_above_ma20": close_above_ma20,
        "bull_order": bull_order,
        "macd_cross": macd_cross,
        "macd_near_zero": macd_near_zero,
        "markup": markup,
        "trend_broken": trend_broken,
        "volume_break_ma20": volume_break_ma20,
        "comment": "，".join(reasons),
    }


def analyze_volume_price(df: pd.DataFrame) -> dict[str, Any]:
    """判断量价质量：缩量回调、长下影、放量阴线和供应进场。"""
    last = df.iloc[-1]
    recent = df.tail(3)
    prior_avg5 = df["volume"].shift(1).rolling(5).mean()

    shrink_pullback = False
    for idx in recent.index:
        row = df.loc[idx]
        prev_avg = as_float(prior_avg5.loc[idx])
        if (
            as_float(row["pct_chg"]) < 0
            and prev_avg > 0
            and as_float(row["volume"]) < prev_avg * 0.8
            and as_float(row["close"]) >= as_float(row["ma20"])
        ):
            shrink_pullback = True
            break

    day_range = as_float(last["high"]) - as_float(last["low"])
    lower_shadow = min(as_float(last["open"]), as_float(last["close"])) - as_float(last["low"])
    long_lower_shadow = day_range > 0 and lower_shadow / day_range > 0.4

    up_volume = (
        as_float(last["pct_chg"]) > 1
        and as_float(last["avg_vol_5"]) > 0
        and as_float(last["volume"]) > as_float(last["avg_vol_5"]) * 1.2
    )
    pullback_shrink = shrink_pullback
    close_near_low = day_range > 0 and (as_float(last["close"]) - as_float(last["low"])) / day_range < 0.25
    big_volume_bear = (
        as_float(last["pct_chg"]) <= -4
        and as_float(last["avg_vol_5"]) > 0
        and as_float(last["volume"]) > as_float(last["avg_vol_5"]) * 1.5
    )
    supply_today = big_volume_bear and close_near_low

    bear_rows = df.tail(4).copy()
    bear_rows["prev_avg_vol"] = prior_avg5.reindex(bear_rows.index)
    bear_rows["volume_bear"] = (
        (bear_rows["close"] < bear_rows["open"])
        & (bear_rows["volume"] > bear_rows["prev_avg_vol"] * 1.25)
    )
    continuous_volume_bear = bool(bear_rows["volume_bear"].tail(3).sum() >= 2)
    supply_in = bool(supply_today or continuous_volume_bear)

    score = 60
    reasons: list[str] = []
    if shrink_pullback:
        score += 15
        reasons.append("近 1-3 日出现缩量回调")
    if long_lower_shadow:
        score += 10
        reasons.append("出现长下影止跌特征")
    if up_volume:
        score += 8
        reasons.append("上涨时成交量放大")
    if pullback_shrink:
        score += 7
        reasons.append("回调量能收缩")
    if big_volume_bear:
        score -= 25
        reasons.append("出现放量大阴线")
    if supply_in:
        score = min(score, 45)
        reasons.append("存在明显供应进场")
    if not reasons:
        reasons.append("量价信号不突出，暂按中性观察")

    score = int(max(0, min(100, score)))
    status = "量价健康" if score >= 80 else "量价尚可" if score >= 65 else "量价一般" if score >= 50 else "量价偏弱"
    return {
        "volume_score": score,
        "volume_status": status,
        "shrink_pullback": shrink_pullback,
        "long_lower_shadow": long_lower_shadow,
        "up_volume": up_volume,
        "pullback_shrink": pullback_shrink,
        "big_volume_bear": big_volume_bear,
        "supply_in": supply_in,
        "comment": "，".join(reasons),
    }


def detect_start_signal(df: pd.DataFrame) -> dict[str, Any]:
    """检测近 30 个交易日是否出现放量突破、涨停、MACD 金叉等启动信号。"""
    candidates: list[dict[str, Any]] = []
    start_idx = max(20, len(df) - 30)
    for idx in range(start_idx, len(df)):
        row = df.iloc[idx]
        prev = df.iloc[idx - 1]
        prior20 = df.iloc[max(0, idx - 20) : idx]
        prior5 = df.iloc[max(0, idx - 5) : idx]
        if prior20.empty or prior5.empty:
            continue

        prior20_high = as_float(prior20["high"].max())
        prior20_low = as_float(prior20["low"].min())
        prior5_vol = as_float(prior5["volume"].mean())
        prior20_amount = as_float(prior20["amount"].mean())
        volume_expand = prior5_vol > 0 and as_float(row["volume"]) > prior5_vol * 1.3
        amount_expand = prior20_amount > 0 and as_float(row["amount"]) > prior20_amount * 1.5
        break_high = as_float(row["close"]) > prior20_high and volume_expand
        platform = prior20_high > 0 and (prior20_high - prior20_low) / prior20_high < 0.18
        platform_break = platform and break_high
        near_limit = as_float(row["pct_chg"]) >= 9
        macd_cross = as_float(prev["dif"]) <= as_float(prev["dea"]) and as_float(row["dif"]) > as_float(row["dea"])
        bull_order = as_float(row["ma5"]) > as_float(row["ma10"]) > as_float(row["ma20"])
        prev_bull_order = as_float(prev["ma5"]) > as_float(prev["ma10"]) > as_float(prev["ma20"])
        ma_turn_bull = bull_order and not prev_bull_order

        types: list[str] = []
        score = 35
        if break_high:
            score += 22
            types.append("放量突破近 20 日新高")
        if platform_break:
            score += 15
            types.append("放量突破平台")
        if near_limit:
            score += 18
            types.append("涨停或接近涨停")
        if macd_cross:
            score += 12
            types.append("MACD金叉")
        if ma_turn_bull:
            score += 10
            types.append("均线开始多头排列")
        if amount_expand:
            score += 10
            types.append("成交额明显放大")

        if types:
            candidates.append(
                {
                    "date": row["date"],
                    "score": min(100, score),
                    "types": types,
                    "close": as_float(row["close"]),
                    "index": idx,
                }
            )

    if not candidates:
        return {
            "has_start_signal": False,
            "signal_date": "",
            "gain_since_signal": 0.0,
            "signal_type": "无明显启动信号",
            "signal_score": 35,
            "comment": "近 30 个交易日未检测到明确启动信号",
        }

    best = sorted(candidates, key=lambda x: (x["score"], x["date"]), reverse=True)[0]
    latest_close = as_float(df.iloc[-1]["close"])
    gain_since = (latest_close / best["close"] - 1) * 100 if best["close"] > 0 else 0.0
    signal_type = " + ".join(best["types"][:3])
    comment = "近期已出现启动信号"
    if gain_since > 25:
        comment += "，启动后涨幅较大，后续更适合等待确认"
    elif gain_since >= 0:
        comment += "，启动后趋势仍保持"
    else:
        comment += "，启动后出现回撤，需重新确认"

    return {
        "has_start_signal": True,
        "signal_date": pd.Timestamp(best["date"]).strftime("%Y-%m-%d"),
        "gain_since_signal": pct(gain_since),
        "signal_type": signal_type,
        "signal_score": int(best["score"]),
        "comment": comment,
    }


def analyze_position(df: pd.DataFrame) -> dict[str, Any]:
    """计算当前价格在近 120 日区间的位置分位。"""
    window = df.tail(min(120, len(df)))
    last_close = as_float(df.iloc[-1]["close"])
    low120 = as_float(window["low"].min())
    high120 = as_float(window["high"].max())
    if high120 <= low120:
        percentile = 50.0
    else:
        percentile = (last_close - low120) / (high120 - low120) * 100
    percentile = max(0.0, min(100.0, percentile))

    if percentile < 30:
        status, score = "低位", 90
    elif percentile < 60:
        status, score = "中位", 82
    elif percentile < 80:
        status, score = "偏高但可接受", 70
    elif percentile < 90:
        status, score = "高位，谨慎追高", 55
    else:
        status, score = "明显高位，不适合追高", 35

    return {
        "position_percentile": pct(percentile),
        "position_status": status,
        "position_score": score,
        "comment": f"当前处于近 {len(window)} 日区间约 {pct(percentile, 1)}% 位置，{status}",
    }


def analyze_risk(df: pd.DataFrame) -> dict[str, Any]:
    """计算 ATR/收盘价，判断波动风险。"""
    last = df.iloc[-1]
    close = as_float(last["close"])
    atr_percent = as_float(last.get("atr")) / close * 100 if close > 0 else 0.0
    if atr_percent < 3:
        status, score = "低波动", 90
    elif atr_percent < 6:
        status, score = "中等波动", 78
    elif atr_percent < 9:
        status, score = "高波动", 58
    else:
        status, score = "极高波动", 38
    return {
        "atr_percent": pct(atr_percent),
        "risk_status": status,
        "risk_score": score,
        "comment": f"ATR/收盘价约 {pct(atr_percent)}%，{status}，按纪律控制观察节奏",
    }


def calculate_score(result: dict[str, Any]) -> dict[str, Any]:
    """按权重计算总分，并应用强制降级规则。"""
    if result.get("data_status") != "正常":
        result["total_score"] = 0
        result["rating"] = result.get("data_status", "数据缺失")
        return result

    total = (
        result["trend"]["trend_score"] * 0.30
        + result["volume"]["volume_score"] * 0.25
        + result["signal"]["signal_score"] * 0.15
        + result["position"]["position_score"] * 0.15
        + result["risk"]["risk_score"] * 0.15
    )
    total_score = int(round(max(0, min(100, total))))
    rating = rating_by_score(total_score)
    forced_reasons: list[str] = []

    if result["trend"].get("volume_break_ma20"):
        rating = "暂时剔除"
        forced_reasons.append("放量跌破 MA20")
    if result["volume"].get("supply_in"):
        rating = cap_rating(rating, "仅跟踪")
        forced_reasons.append("明显供应进场")
    if result["position"]["position_percentile"] > 90 and result["risk"]["atr_percent"] > 9:
        rating = cap_rating(rating, "等待买点")
        forced_reasons.append("位置超过 90% 且 ATR 超过 9%")
    if result["history_days"] < 120:
        total_score = min(total_score, 74)
        rating = cap_rating(rating_by_score(total_score), "仅跟踪")
        forced_reasons.append("历史日 K 少于 120 日")
    avg_amount20 = as_float(result.get("avg_amount_20"))
    if 0 < avg_amount20 < 50_000_000:
        rating = downgrade_rating(rating)
        forced_reasons.append("近 20 日平均成交额偏低")

    result["total_score"] = total_score
    result["rating"] = rating
    result["forced_reasons"] = forced_reasons
    return result


def rating_by_score(score: int) -> str:
    """按分数得到观察评级。"""
    if score >= 85:
        return "重点观察"
    if score >= 75:
        return "等待买点"
    if score >= 65:
        return "仅跟踪"
    return "暂时剔除"


def cap_rating(current: str, cap: str) -> str:
    """把评级限制在指定上限，不让风险项拿到过高评级。"""
    order = ["暂时剔除", "仅跟踪", "等待买点", "重点观察"]
    return current if order.index(current) <= order.index(cap) else cap


def downgrade_rating(current: str) -> str:
    """评级下调一档。"""
    order = ["重点观察", "等待买点", "仅跟踪", "暂时剔除"]
    if current not in order:
        return current
    return order[min(order.index(current) + 1, len(order) - 1)]


def generate_watch_plan(result: dict[str, Any]) -> dict[str, Any]:
    """生成观察计划，只写观察条件，不写买入、卖出或推荐。"""
    if result.get("data_status") != "正常":
        return {
            "watch_type": "数据补齐型",
            "key_price": "",
            "watch_conditions": ["补齐最近日 K 数据后再复盘", "确认股票代码和数据源是否可用"],
            "invalid_conditions": ["连续多个数据源均无法获取行情"],
            "comment": "当前数据不足，暂不参与高分观察。",
        }

    last = result["latest_row"]
    key_price = as_float(last["low"])
    trend = result["trend"]
    volume = result["volume"]
    position = result["position"]
    risk = result["risk"]

    if volume.get("shrink_pullback") or volume.get("long_lower_shadow"):
        watch_type = "缩量回踩确认型"
        watch_conditions = [
            "次日继续缩量或温和放量修复",
            f"不破最近低点 {key_price:.2f}",
            "收盘重新站稳 MA5 或 MA10",
            "没有放量大阴线",
            "板块方向没有明显退潮",
        ]
    elif trend.get("markup") and result["signal"].get("has_start_signal"):
        watch_type = "趋势延续确认型"
        watch_conditions = [
            "MA20 继续向上",
            "回踩不有效跌破 MA10/MA20",
            "成交额维持在近 20 日均额附近",
            "盘中回落时没有明显供应进场",
        ]
    else:
        watch_type = "等待结构转强型"
        watch_conditions = [
            "重新站上 MA20",
            "MA5/MA10/MA20 形成多头排列",
            "出现放量突破或 MACD 金叉",
            "回调时成交量收缩",
        ]

    invalid_conditions = [
        f"放量跌破最近低点 {key_price:.2f}",
        "跌破 MA20 且无法快速收回",
        "出现高位巨量阴线",
        "明显供应进场",
    ]
    if position["position_percentile"] >= 80:
        invalid_conditions.append("高位继续放量冲高但收盘转弱")
    if risk["atr_percent"] >= 9:
        invalid_conditions.append("ATR 继续维持极高波动")

    return {
        "watch_type": watch_type,
        "key_price": round(key_price, 2),
        "watch_conditions": watch_conditions,
        "invalid_conditions": invalid_conditions,
        "comment": "按观察条件等待确认，不输出买卖建议。",
    }


def analyze_stock(stock: dict[str, str], days: int) -> dict[str, Any]:
    """拉取并分析单只股票。"""
    code = normalize_code(stock["code"])
    base = {"code": code, "name": stock.get("name", code), "theme": stock.get("theme", "")}
    df, source, error = fetch_kline(code, days=days)
    if df is None or df.empty:
        result = {
            **base,
            "data_status": "数据缺失",
            "data_source": source,
            "error": error or "无法获取行情",
            "history_days": 0,
        }
        result["watch_plan"] = generate_watch_plan(result)
        result["total_score"] = 0
        result["rating"] = "数据缺失"
        return result

    df = add_indicators(df)
    history_days = len(df)
    if history_days < 80:
        last = df.iloc[-1]
        result = {
            **base,
            "data_status": "数据不足",
            "data_source": source,
            "history_days": history_days,
            "latest_row": last,
            "latest_price": as_float(last["close"]),
            "latest_pct_chg": as_float(last["pct_chg"]),
            "error": f"仅获取到 {history_days} 个交易日数据",
        }
        result["watch_plan"] = generate_watch_plan(result)
        result["total_score"] = 0
        result["rating"] = "数据不足"
        return result

    trend = analyze_trend(df)
    volume = analyze_volume_price(df)
    signal = detect_start_signal(df)
    position = analyze_position(df)
    risk = analyze_risk(df)
    last = df.iloc[-1]

    result = {
        **base,
        "data_status": "正常",
        "data_source": source,
        "history_days": history_days,
        "latest_row": last,
        "latest_price": as_float(last["close"]),
        "latest_pct_chg": as_float(last["pct_chg"]),
        "avg_amount_20": as_float(last.get("avg_amount_20")),
        "trend": trend,
        "volume": volume,
        "signal": signal,
        "position": position,
        "risk": risk,
    }
    result = calculate_score(result)
    result["watch_plan"] = generate_watch_plan(result)
    result["summary"] = build_summary(result)
    return result


def build_summary(result: dict[str, Any]) -> str:
    """生成一句话摘要，保持观察/等待/跟踪/剔除语气。"""
    if result.get("data_status") != "正常":
        return f"{result.get('data_status')}：{result.get('error', '')}"
    parts = [
        result["trend"]["trend_status"],
        result["volume"]["volume_status"],
        result["position"]["position_status"],
        result["risk"]["risk_status"],
    ]
    forced = result.get("forced_reasons") or []
    suffix = f"；纪律限制：{'、'.join(forced)}" if forced else ""
    return f"{'，'.join(parts)}；评级为{result['rating']}{suffix}"


def result_to_csv_row(result: dict[str, Any]) -> dict[str, Any]:
    """把分析结果压平成 CSV 行。"""
    if result.get("data_status") != "正常":
        return {
            "code": result["code"],
            "name": result["name"],
            "theme": result.get("theme", ""),
            "latest_price": result.get("latest_price", ""),
            "latest_pct_chg": result.get("latest_pct_chg", ""),
            "total_score": result.get("total_score", 0),
            "rating": result.get("rating", result.get("data_status", "")),
            "trend_score": "",
            "volume_score": "",
            "signal_score": "",
            "position_score": "",
            "risk_score": "",
            "position_percentile": "",
            "atr_percent": "",
            "watch_type": result["watch_plan"]["watch_type"],
            "key_price": result["watch_plan"]["key_price"],
            "summary": result.get("error", "数据缺失"),
            "invalid_conditions": "；".join(result["watch_plan"]["invalid_conditions"]),
        }

    plan = result["watch_plan"]
    return {
        "code": result["code"],
        "name": result["name"],
        "theme": result.get("theme", ""),
        "latest_price": round(result["latest_price"], 2),
        "latest_pct_chg": round(result["latest_pct_chg"], 2),
        "total_score": result["total_score"],
        "rating": result["rating"],
        "trend_score": result["trend"]["trend_score"],
        "volume_score": result["volume"]["volume_score"],
        "signal_score": result["signal"]["signal_score"],
        "position_score": result["position"]["position_score"],
        "risk_score": result["risk"]["risk_score"],
        "position_percentile": result["position"]["position_percentile"],
        "atr_percent": result["risk"]["atr_percent"],
        "watch_type": plan["watch_type"],
        "key_price": plan["key_price"],
        "summary": result.get("summary", ""),
        "invalid_conditions": "；".join(plan["invalid_conditions"]),
    }


def sort_results(results: list[dict[str, Any]]) -> list[dict[str, Any]]:
    """按评分排序，数据缺失或不足放在后面。"""
    return sorted(
        results,
        key=lambda x: (
            0 if x.get("data_status") == "正常" else -1,
            x.get("total_score", 0),
        ),
        reverse=True,
    )


def write_csv(results: list[dict[str, Any]], output_path: Path) -> None:
    """写出 screening_result.csv。"""
    fieldnames = [
        "code",
        "name",
        "theme",
        "latest_price",
        "latest_pct_chg",
        "total_score",
        "rating",
        "trend_score",
        "volume_score",
        "signal_score",
        "position_score",
        "risk_score",
        "position_percentile",
        "atr_percent",
        "watch_type",
        "key_price",
        "summary",
        "invalid_conditions",
    ]
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for result in results:
            writer.writerow(result_to_csv_row(result))


def markdown_escape(text: Any) -> str:
    """转义 Markdown 表格中的竖线和换行。"""
    return str(text).replace("|", "/").replace("\n", " ")


def md_table(rows: list[list[Any]], headers: list[str]) -> str:
    """生成 Markdown 表格。"""
    lines = [
        "| " + " | ".join(headers) + " |",
        "| " + " | ".join("---:" if h in {"分数"} else "---" for h in headers) + " |",
    ]
    if not rows:
        lines.append("| " + " | ".join(["-"] * len(headers)) + " |")
    else:
        for row in rows:
            lines.append("| " + " | ".join(markdown_escape(cell) for cell in row) + " |")
    return "\n".join(lines)


def write_report(results: list[dict[str, Any]], output_path: Path, source_desc: str) -> None:
    """写出 screening_report.md。"""
    today = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    sections: dict[str, list[dict[str, Any]]] = {
        "重点观察": [],
        "等待买点": [],
        "仅跟踪": [],
        "暂时剔除": [],
    }
    for result in results:
        rating = result.get("rating", "")
        if rating in sections:
            sections[rating].append(result)
        else:
            sections["暂时剔除"].append(result)

    lines = [
        "# A股股票池筛选结果",
        "",
        "## 一、筛选说明",
        "",
        f"- 生成时间：{today}",
        f"- 输入来源：{source_desc}",
        "- 本次筛选基于指定股票池或方向，优先使用 AkShare，并以东方财富 / 腾讯财经公开行情数据兜底。",
        "- 本工具只用于观察和复盘，不构成买卖建议。",
        "",
        "## 二、重点观察",
        "",
        md_table(
            [
                [
                    r["name"],
                    r["code"],
                    r.get("total_score", 0),
                    r.get("rating", ""),
                    r.get("summary", ""),
                    "；".join(r["watch_plan"]["watch_conditions"][:3]),
                ]
                for r in sections["重点观察"]
            ],
            ["股票", "代码", "分数", "评级", "核心理由", "观察条件"],
        ),
        "",
        "## 三、等待买点",
        "",
        md_table(
            [
                [
                    r["name"],
                    r["code"],
                    r.get("total_score", 0),
                    r.get("rating", ""),
                    r.get("summary", ""),
                    "；".join(r["watch_plan"]["watch_conditions"][:3]),
                ]
                for r in sections["等待买点"]
            ],
            ["股票", "代码", "分数", "评级", "核心理由", "等待什么"],
        ),
        "",
        "## 四、仅跟踪",
        "",
        md_table(
            [
                [
                    r["name"],
                    r["code"],
                    r.get("total_score", 0),
                    r.get("rating", ""),
                    r.get("summary", r.get("error", "")),
                ]
                for r in sections["仅跟踪"]
            ],
            ["股票", "代码", "分数", "评级", "主要问题"],
        ),
        "",
        "## 五、暂时剔除",
        "",
        md_table(
            [
                [
                    r["name"],
                    r["code"],
                    r.get("total_score", 0),
                    r.get("rating", ""),
                    r.get("summary", r.get("error", "")),
                ]
                for r in sections["暂时剔除"]
            ],
            ["股票", "代码", "分数", "评级", "剔除原因"],
        ),
        "",
        "## 六、个股简评",
        "",
    ]

    for r in results:
        lines.extend(stock_comment_lines(r))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def stock_comment_lines(result: dict[str, Any]) -> list[str]:
    """生成单只股票的 Markdown 简评。"""
    title = f"### {result['name']} {result['code']}"
    if result.get("data_status") != "正常":
        return [
            title,
            "",
            f"- 数据：{result.get('data_status')}，{result.get('error', '')}",
            "- 一句话结论：数据不足，暂不参与观察评级。",
            "",
        ]

    plan = result["watch_plan"]
    return [
        title,
        "",
        f"- 趋势：{result['trend']['comment']}。",
        f"- 量价：{result['volume']['comment']}。",
        f"- 启动信号：{result['signal']['comment']}；{result['signal']['signal_type']}。",
        f"- 位置：{result['position']['comment']}。",
        f"- 风险：{result['risk']['comment']}。",
        f"- 观察计划：{plan['watch_type']}；{'；'.join(plan['watch_conditions'])}。",
        f"- 失效条件：{'；'.join(plan['invalid_conditions'])}。",
        f"- 一句话结论：{result.get('summary', '')}。",
        "",
    ]


def build_stock_pool(args: argparse.Namespace) -> tuple[list[dict[str, str]], str]:
    """根据命令行参数生成股票池。"""
    provided = [bool(args.theme), bool(args.stocks), bool(args.input)]
    if sum(provided) != 1:
        raise ValueError("请在 --theme、--stocks、--input 中选择且只选择一种输入方式。")
    if args.theme:
        return resolve_theme_pool(args.theme, limit=args.limit), f"方向：{args.theme}"
    if args.stocks:
        return parse_stocks_arg(args.stocks), f"代码列表：{args.stocks}"
    input_path = Path(args.input).resolve()
    return read_input_pool(input_path), f"CSV：{input_path}"


def parse_args(argv: list[str]) -> argparse.Namespace:
    """解析命令行参数。"""
    parser = argparse.ArgumentParser(description="A 股股票池技术筛选工具")
    parser.add_argument("--theme", help="方向主题，例如：半导体材料、AIDC硬件、稀土")
    parser.add_argument("--stocks", help="股票代码列表，例如：603688,300666,300054")
    parser.add_argument("--input", help="股票池 CSV，至少包含 code 列，可选 name、theme")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="输出目录")
    parser.add_argument("--days", type=int, default=160, help="拉取日 K 数量，建议不少于 160")
    parser.add_argument("--limit", type=int, default=80, help="方向主题最多纳入多少只股票")
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    """主流程：构建股票池、逐只筛选、写出 CSV 和 Markdown 报告。"""
    args = parse_args(argv or sys.argv[1:])
    try:
        pool, source_desc = build_stock_pool(args)
    except Exception as exc:
        print(f"[错误] {exc}", file=sys.stderr)
        return 2

    if not pool:
        print("[错误] 股票池为空，请检查输入。", file=sys.stderr)
        return 2

    print(f"本次股票池 {len(pool)} 只，开始获取行情并筛选...")
    results: list[dict[str, Any]] = []
    for idx, stock in enumerate(pool, start=1):
        print(f"[{idx}/{len(pool)}] {stock.get('name', stock['code'])} {stock['code']}")
        try:
            results.append(analyze_stock(stock, days=max(args.days, 120)))
        except Exception as exc:
            code = normalize_code(stock.get("code", ""))
            results.append(
                {
                    "code": code,
                    "name": stock.get("name", code),
                    "theme": stock.get("theme", ""),
                    "data_status": "数据缺失",
                    "error": str(exc),
                    "history_days": 0,
                    "total_score": 0,
                    "rating": "数据缺失",
                    "watch_plan": {
                        "watch_type": "数据补齐型",
                        "key_price": "",
                        "watch_conditions": ["检查代码或稍后重试数据源"],
                        "invalid_conditions": ["数据源持续失败"],
                    },
                }
            )

    results = sort_results(results)
    output_dir = Path(args.output_dir).resolve()
    csv_path = output_dir / "screening_result.csv"
    report_path = output_dir / "screening_report.md"
    write_csv(results, csv_path)
    write_report(results, report_path, source_desc)

    print(f"已生成：{csv_path}")
    print(f"已生成：{report_path}")
    print("提示：报告只用于观察和复盘，不构成买卖建议。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
