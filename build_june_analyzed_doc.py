#!/usr/bin/env python3
r"""
Build the June analyzed document for the NGA thread "科学技术打头阵".

Output design:
- keep all June posts in chronological order;
- analyze all A/B class posts with the same framework;
- highlight wolf's posts and make their analysis fuller;
- remove repetitive "小白学习笔记";
- resolve stock codes and common abbreviations to full names and directions.
"""

from __future__ import annotations

import argparse
import calendar
import html
import json
import re
import time
import urllib.parse
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

import requests
from docx import Document
from docx.enum.section import WD_SECTION_START
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor


THREAD_ID = "45974302"
THREAD_TITLE = "科学技术打头阵"
WOLF_UID = "150058"

INDEXES = {
    "1.000001": "上证指数",
    "0.399001": "深证成指",
    "0.399006": "创业板指",
    "1.000688": "科创50",
    "1.000852": "中证1000",
}

TENCENT_INDEXES = {
    "sh000001": "上证指数",
    "sz399001": "深证成指",
    "sz399006": "创业板指",
    "sh000688": "科创50",
    "sh000852": "中证1000",
}

TRADING_KEYWORDS = [
    "大盘", "指数", "上证", "深成", "创业板", "科创", "中证", "黄线", "白线",
    "板块", "主线", "方向", "细分", "科技", "光", "算力", "液冷", "AI",
    "半导体", "芯片", "封装", "PCB", "有色", "金属", "白酒", "券商",
    "银行", "医药", "消费", "机器人", "军工", "港股", "北向",
    "个股", "龙头", "中军", "补涨", "杂毛", "涨停", "跌停",
    "量能", "成交", "放量", "缩量", "换手", "资金", "承接", "兑现",
    "低吸", "追高", "高位", "低位", "仓", "减仓", "加仓", "止损",
    "止盈", "T", "做T", "回补", "风险", "分歧", "退潮", "轮动",
    "预期", "超预期", "不及预期", "利好", "利空", "落地",
    "K", "红K", "黑K", "揉搓线", "上影", "下影", "缺口", "回踩",
    "突破", "破位", "均线", "支撑", "压力", "V", "岛形",
]

SECTOR_KEYWORDS = {
    "核心科技/国产算力": ["科技", "核心科技", "国产算力", "华为", "昇腾", "算力", "服务器", "PCB"],
    "光通信/光模块": ["光模块", "光通信", "CPO", "光上游", "光芯片", "光器件", "光迅", "中际", "新易盛"],
    "半导体设备/材料/封测": ["半导体", "芯片", "封装", "封测", "设备", "材料", "长电", "通富", "北方华创", "中微"],
    "AI应用/软件": ["AI应用", "AI 软", "软件", "大模型", "传媒", "游戏"],
    "液冷/散热": ["液冷", "散热", "英维克", "强瑞"],
    "有色金属": ["有色", "金属", "铜", "铝", "小金属", "贵金属"],
    "消费/白酒": ["消费", "白酒", "食品", "旅游"],
    "金融权重": ["券商", "银行", "保险", "金融", "权重"],
    "医药": ["医药", "药", "创新药", "医疗"],
    "机器人": ["机器人", "减速器", "电机"],
}

ACTION_KEYWORDS = {
    "观察": ["观察", "看", "等", "确认", "不急"],
    "低吸": ["低吸", "低挂", "回落", "接"],
    "持有": ["持有", "锁仓", "拿", "坚定"],
    "减仓": ["减仓", "降仓", "跑", "卖", "出清", "止盈"],
    "止损": ["止损", "破位"],
    "做T": ["做T", "T", "回补"],
    "不操作": ["不做", "别追", "不追", "不要追", "严禁"],
}

TERM_DEFS = {
    "主线": "持续获得资金选择、容量足够、逻辑可延展的方向，不等同于当天涨得多。",
    "黄线": "分时图里更接近中小盘/平均股表现的线，可辅助判断题材股强弱。",
    "白线": "按权重计算的指数线，可辅助判断大票和权重护盘强弱。",
    "量能": "成交量和成交额，代表资金参与程度；没有量能的形态可靠性较弱。",
    "低吸": "在回落、支撑或分歧处买入，而不是上涨后情绪化追高。",
    "追高": "价格已经明显上涨后再买，常见风险是买在短线兑现点。",
    "分歧": "资金意见不一致，常表现为冲高回落、放量震荡、板块内部强弱分化。",
    "预期差": "实际消息或盘面相对市场原先预期的偏离，关键在位置和资金是否提前埋伏。",
    "揉搓线": "上下影线明显、实体偏小的 K 线结构，表示多空拉扯，必须结合趋势和量能。",
    "止损": "触发预设风险条件后退出，避免亏损扩大。",
    "止盈": "达到目标或跌破保护位后锁定利润。",
    "回踩": "突破后价格回落测试支撑或关键位置。",
    "突破": "价格或指数上穿关键压力位，需要量能和板块核心跟随确认。",
}

STOCK_ALIAS_DIRECTION = {
    "英子": ("英维克", "002837", "液冷/数据中心温控"),
    "英维克": ("英维克", "002837", "液冷/数据中心温控"),
    "寒武": ("寒武纪", "688256", "国产 AI 芯片/算力"),
    "寒武纪": ("寒武纪", "688256", "国产 AI 芯片/算力"),
    "海光": ("海光信息", "688041", "国产 CPU/GPU/算力芯片"),
    "中芯": ("中芯国际", "688981", "晶圆制造/半导体制造"),
    "长电": ("长电科技", "600584", "先进封装/封测"),
    "通富": ("通富微电", "002156", "先进封装/封测"),
    "华天": ("华天科技", "002185", "封装测试"),
    "甬矽": ("甬矽电子", "688362", "封装测试"),
    "北方": ("北方华创", "002371", "半导体设备"),
    "北方华创": ("北方华创", "002371", "半导体设备"),
    "中微": ("中微公司", "688012", "半导体设备"),
    "拓荆": ("拓荆科技", "688072", "半导体设备"),
    "华海": ("华海清科", "688120", "半导体设备/CMP"),
    "鼎龙": ("鼎龙股份", "300054", "半导体材料/CMP 抛光垫"),
    "华丰": ("华丰科技", "688629", "高速连接器/算力硬件"),
    "高新": ("高新发展", "000628", "华为昇腾/算力集成"),
    "拓维": ("拓维信息", "002261", "华为昇腾/软件与算力"),
    "神码": ("神州数码", "000034", "华为鲲鹏/昇腾生态"),
    "神州数码": ("神州数码", "000034", "华为鲲鹏/昇腾生态"),
    "华大": ("华大九天", "301269", "EDA/芯片设计软件"),
    "概伦": ("概伦电子", "688206", "EDA/芯片设计软件"),
    "华工": ("华工科技", "000988", "光通信/激光设备"),
    "光迅": ("光迅科技", "002281", "光通信/光模块"),
    "中际": ("中际旭创", "300308", "光模块/CPO"),
    "新易盛": ("新易盛", "300502", "光模块/CPO"),
    "龙蟠": ("龙蟠科技", "603906", "材料/电池材料"),
    "强瑞": ("强瑞技术", "301128", "液冷/测试设备/算力硬件"),
    "工富": ("工业富联", "601138", "AI 服务器/算力硬件"),
    "工业富联": ("工业富联", "601138", "AI 服务器/算力硬件"),
    "沪电": ("沪电股份", "002463", "PCB/服务器高速板"),
    "兴森": ("兴森科技", "002436", "PCB/IC载板"),
    "沃尔": ("沃尔核材", "002130", "线缆/高速连接/材料"),
    "胜宏": ("胜宏科技", "300476", "PCB/AI服务器板"),
    "英伟达": ("英伟达", "NVDA", "美股 AI GPU/算力龙头"),
}


@dataclass
class MarketDay:
    date: str
    indexes: dict[str, dict[str, float | str]]


class StockResolver:
    def __init__(self, cache_path: Path) -> None:
        self.cache_path = cache_path
        self.session = requests.Session()
        self.session.trust_env = False
        if cache_path.exists():
            self.cache: dict[str, Any] = json.loads(cache_path.read_text(encoding="utf-8-sig"))
        else:
            self.cache = {}

    def save(self) -> None:
        self.cache_path.write_text(json.dumps(self.cache, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    def direction_for_name(self, name: str) -> str:
        for _, (full, _, direction) in STOCK_ALIAS_DIRECTION.items():
            if name == full:
                return direction
        if any(k in name for k in ["光", "旭创", "易盛"]):
            return "光通信/光模块"
        if any(k in name for k in ["封", "电", "微"]):
            return "半导体/电子"
        if any(k in name for k in ["算", "联", "信息"]):
            return "算力/软件信息"
        return "待人工确认所属方向"

    def search(self, query: str) -> dict[str, str] | None:
        if query in self.cache:
            return self.cache[query]
        if query in STOCK_ALIAS_DIRECTION:
            full, code, direction = STOCK_ALIAS_DIRECTION[query]
            payload = {"alias": query, "name": full, "code": code, "direction": direction, "source": "内置映射"}
            self.cache[query] = payload
            return payload

        url = (
            "http://searchapi.eastmoney.com/api/suggest/get?"
            f"input={urllib.parse.quote(query)}&type=14&token=1"
        )
        try:
            response = self.session.get(
                url,
                timeout=10,
                headers={"User-Agent": "Mozilla/5.0", "Referer": "http://quote.eastmoney.com/"},
            )
            response.raise_for_status()
            data = response.json().get("QuotationCodeTable", {}).get("Data") or []
        except Exception:
            data = []
        result = None
        for item in data:
            if item.get("SecurityTypeName") in {"沪A", "深A", "科创板", "创业板", "美股"}:
                name = item.get("Name", "")
                code = item.get("Code", "")
                result = {
                    "alias": query,
                    "name": name,
                    "code": code,
                    "direction": self.direction_for_name(name),
                    "source": "东方财富搜索",
                }
                break
        self.cache[query] = result
        return result

    def detect(self, body: str) -> list[dict[str, str]]:
        candidates: set[str] = set()
        for alias in STOCK_ALIAS_DIRECTION:
            if alias in body:
                candidates.add(alias)
        for code in re.findall(r"(?<!\d)((?:60|68|00|30)\d{4})(?!\d)", body):
            candidates.add(code)
        found = []
        for query in sorted(candidates, key=len, reverse=True):
            item = self.search(query)
            if item:
                found.append(item)
        deduped = []
        seen = set()
        for item in found:
            key = (item["code"], item["name"])
            if key not in seen:
                seen.add(key)
                deduped.append(item)
        return deduped[:8]


def clean_text(value: str) -> str:
    text = html.unescape(value or "")
    text = re.sub(r"\[/?(?:quote|collapse|b|i|u|url|img|color|size|pid|uid)[^\]]*\]", "", text, flags=re.I)
    text = re.sub(r"Reply(?: to)?(?: Post by)?", "Reply", text)
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def parse_time(value: str) -> datetime:
    return datetime.strptime(value[:16], "%Y-%m-%d %H:%M")


def time_stage(dt: datetime) -> str:
    hhmm = dt.hour * 100 + dt.minute
    if hhmm < 930:
        return "盘前/开盘前"
    if hhmm <= 1000:
        return "开盘确认段"
    if hhmm < 1130:
        return "早盘"
    if hhmm < 1300:
        return "午间"
    if hhmm < 1430:
        return "午后"
    if hhmm <= 1500:
        return "尾盘"
    return "收盘后"


def load_posts(path: Path, month: str) -> list[dict[str, Any]]:
    posts = json.loads(path.read_text(encoding="utf-8-sig"))
    result = []
    for post in posts:
        if post.get("posted_at", "").startswith(month):
            item = dict(post)
            item["body"] = clean_text(item.get("body", ""))
            item["_dt"] = parse_time(item["time"])
            result.append(item)
    result.sort(key=lambda p: (p["_dt"], int(p["floor"]) if str(p.get("floor", "")).isdigit() else 0, p["uid"]))
    return result


def _date_dash(value: str) -> str:
    return f"{value[:4]}-{value[4:6]}-{value[6:8]}"


def fetch_tencent_klines(begin: str, end: str) -> dict[str, MarketDay]:
    begin_dash = _date_dash(begin)
    end_dash = _date_dash(end)
    session = requests.Session()
    session.trust_env = False
    by_day: dict[str, MarketDay] = {}

    for code, name in TENCENT_INDEXES.items():
        url = (
            "https://web.ifzq.gtimg.cn/appstock/app/fqkline/get?"
            f"param={code},day,{begin_dash},{end_dash},320,qfq"
        )
        response = None
        for attempt in range(1, 5):
            try:
                response = session.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0"})
                response.raise_for_status()
                break
            except Exception as exc:
                if attempt == 4:
                    print(f"Warning: failed fetching Tencent {name}: {exc}")
                else:
                    time.sleep(1.2 * attempt)
        if response is None:
            continue

        payload = response.json().get("data", {}).get(code, {})
        previous_close: float | None = None
        for line in payload.get("day", []):
            if len(line) < 6:
                continue
            date, open_, close, high, low, volume = line[:6]
            open_value = float(open_)
            close_value = float(close)
            high_value = float(high)
            low_value = float(low)
            volume_value = float(volume)
            if previous_close:
                change = close_value - previous_close
                pct = change / previous_close * 100
                amplitude = (high_value - low_value) / previous_close * 100
            else:
                change = 0.0
                pct = 0.0
                amplitude = 0.0
            by_day.setdefault(date, MarketDay(date=date, indexes={}))
            by_day[date].indexes[name] = {
                "open": open_value,
                "close": close_value,
                "high": high_value,
                "low": low_value,
                "volume": volume_value,
                "amount": 0.0,
                "amplitude": amplitude,
                "pct_change": pct,
                "change": change,
                "turnover": 0.0,
                "source": "腾讯财经日线",
            }
            previous_close = close_value

        quote = payload.get("qt", {}).get(code)
        if quote and len(quote) > 37:
            raw_date = quote[30][:8]
            date = _date_dash(raw_date)
            if begin_dash <= date <= end_dash:
                close_value = float(quote[3])
                previous_close = float(quote[4])
                open_value = float(quote[5])
                high_value = float(quote[33])
                low_value = float(quote[34])
                volume_value = float(quote[36])
                amount_value = float(quote[37]) * 10000
                by_day.setdefault(date, MarketDay(date=date, indexes={}))
                by_day[date].indexes[name] = {
                    "open": open_value,
                    "close": close_value,
                    "high": high_value,
                    "low": low_value,
                    "volume": volume_value,
                    "amount": amount_value,
                    "amplitude": (high_value - low_value) / previous_close * 100 if previous_close else 0.0,
                    "pct_change": float(quote[32]),
                    "change": float(quote[31]),
                    "turnover": 0.0,
                    "source": "腾讯财经实时行情",
                }

    return by_day


def fetch_index_klines(begin: str, end: str, cache_path: Path) -> dict[str, MarketDay]:
    if cache_path.exists():
        cached = json.loads(cache_path.read_text(encoding="utf-8-sig"))
        return {day: MarketDay(date=day, indexes=payload["indexes"]) for day, payload in cached.items()}

    session = requests.Session()
    session.trust_env = False
    by_day: dict[str, MarketDay] = {}
    for secid, name in INDEXES.items():
        url = (
            "http://push2his.eastmoney.com/api/qt/stock/kline/get"
            f"?secid={secid}&fields1=f1,f2,f3,f4,f5,f6"
            "&fields2=f51,f52,f53,f54,f55,f56,f57,f58,f59,f60,f61"
            f"&klt=101&fqt=1&beg={begin}&end={end}"
        )
        response = None
        for attempt in range(1, 5):
            try:
                response = session.get(
                    url,
                    timeout=20,
                    headers={"User-Agent": "Mozilla/5.0", "Referer": "http://quote.eastmoney.com/"},
                )
                response.raise_for_status()
                break
            except Exception as exc:
                if attempt == 4:
                    print(f"Warning: failed fetching {name}: {exc}")
                else:
                    time.sleep(1.2 * attempt)
        if response is None:
            continue
        data = response.json().get("data") or {}
        for line in data.get("klines", []):
            parts = line.split(",")
            if len(parts) < 11:
                continue
            date, open_, close, high, low, volume, amount, amplitude, pct, change, turnover = parts[:11]
            by_day.setdefault(date, MarketDay(date=date, indexes={}))
            by_day[date].indexes[name] = {
                "open": float(open_),
                "close": float(close),
                "high": float(high),
                "low": float(low),
                "volume": float(volume),
                "amount": float(amount),
                "amplitude": float(amplitude),
                "pct_change": float(pct),
                "change": float(change),
                "turnover": float(turnover),
            }
    if not by_day:
        print("Warning: Eastmoney index data unavailable; falling back to Tencent Finance.")
        by_day = fetch_tencent_klines(begin, end)

    cache_path.write_text(
        json.dumps({day: {"date": md.date, "indexes": md.indexes} for day, md in by_day.items()}, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    return by_day


def market_summary(day: str, market_days: dict[str, MarketDay]) -> str:
    md = market_days.get(day)
    if not md:
        return "未取得当日指数数据。"
    parts = []
    for name in ["上证指数", "深证成指", "创业板指", "科创50", "中证1000"]:
        item = md.indexes.get(name)
        if item:
            parts.append(f"{name}{item['pct_change']:+.2f}%")
    up_count = sum(1 for item in md.indexes.values() if float(item["pct_change"]) > 0)
    amount = md.indexes.get("上证指数", {}).get("amount")
    volume_note = f"；上证成交额约{float(amount)/1e8:.0f}亿元" if amount else ""
    style_note = "指数整体偏强" if up_count >= 3 else "指数分化或偏弱"
    return "，".join(parts) + volume_note + f"；{style_note}。"


def parse_referenced_dates(body: str, post_dt: datetime) -> list[str]:
    dates: set[str] = set()
    for match in re.finditer(r"(20\d{2})[-/年](\d{1,2})[-/月](\d{1,2})日?", body):
        year, month, day = map(int, match.groups())
        try:
            dates.add(datetime(year, month, day).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    for match in re.finditer(r"(?<!\d)(\d{1,2})月(\d{1,2})日", body):
        month, day = map(int, match.groups())
        try:
            dates.add(datetime(post_dt.year, month, day).strftime("%Y-%m-%d"))
        except ValueError:
            pass
    current_day = post_dt.strftime("%Y-%m-%d")
    return sorted(day for day in dates if day != current_day)


def market_days_for_month(month: str, current_month: str, current_market_days: dict[str, MarketDay]) -> dict[str, MarketDay]:
    if month == current_month:
        return current_market_days
    cache_path = Path(f"market_cache_{month}.json")
    year, month_num = map(int, month.split("-"))
    last_day = calendar.monthrange(year, month_num)[1]
    begin = f"{year}{month_num:02d}01"
    end = f"{year}{month_num:02d}{last_day:02d}"
    return fetch_index_klines(begin, end, cache_path)


def referenced_market_notes(body: str, post_dt: datetime, current_market_days: dict[str, MarketDay]) -> list[str]:
    notes: list[str] = []
    current_month = post_dt.strftime("%Y-%m")
    for day in parse_referenced_dates(body, post_dt):
        month = day[:7]
        month_days = market_days_for_month(month, current_month, current_market_days)
        notes.append(f"{day}：{market_summary(day, month_days)}")
    return notes


def classify_post(post: dict[str, Any]) -> tuple[str, str]:
    body = post["body"]
    if not body:
        return "C", "空正文或无有效信息。"
    hits = [kw for kw in TRADING_KEYWORDS if kw.lower() in body.lower()]
    if post["uid"] == WOLF_UID and len(hits) >= 2:
        return "A", "狼大发言且包含多个交易/盘面关键词，适合重点详析。"
    if post["uid"] == WOLF_UID:
        return "B", "狼大发言，交易信息较弱或偏互动，但仍保留结构化分析。"
    if len(hits) >= 2:
        return "A", "包含多个交易/盘面关键词，适合完整分析。"
    if hits:
        return "B", "包含少量交易关键词，适合作为上下文分析。"
    return "C", "偏互动、闲聊或情绪表达，不展开。"


def detect_terms(body: str) -> list[tuple[str, str]]:
    return [(term, desc) for term, desc in TERM_DEFS.items() if term in body][:8]


def detect_sectors(body: str) -> list[str]:
    return [sector for sector, keywords in SECTOR_KEYWORDS.items() if any(keyword in body for keyword in keywords)]


def detect_action(body: str) -> str:
    for action, keywords in ACTION_KEYWORDS.items():
        if any(keyword in body for keyword in keywords):
            return action
    return "观察/未明确"


def technical_notes(body: str) -> list[str]:
    notes = []
    if "揉搓线" in body or "上影" in body or "下影" in body:
        notes.append("涉及揉搓线/影线结构：必须先判断当前是上涨趋势还是下跌趋势，再结合红黑 K、上下影线长短和量能。")
    if "放量" in body or "缩量" in body:
        notes.append("涉及量能：放量说明资金参与提高，缩量说明分歧或参与不足，要看放量后能否延续。")
    if "缺口" in body:
        notes.append("涉及缺口：需要观察是否回补，以及是否形成跳空岛形等结构风险。")
    if "突破" in body or "回踩" in body:
        notes.append("涉及突破/回踩：关键是量能、核心标的和板块跟随是否确认。")
    if "黄线" in body or "白线" in body:
        notes.append("涉及黄白线：黄线强偏题材和中小票，白线强偏权重和大票，不能混用。")
    return notes


def plain_language(post: dict[str, Any], cls: str, stocks: list[dict[str, str]]) -> str:
    if cls == "C":
        return "这条不作为交易知识重点，保留原文即可。"
    sectors = detect_sectors(post["body"])
    action = detect_action(post["body"])
    actor = "狼大" if post["uid"] == WOLF_UID else post["username"]
    text = f"{actor}这条发言需要先还原成条件判断，而不是直接当作买卖指令。"
    if sectors:
        text += f" 讨论方向涉及：{'、'.join(sectors)}。"
    if stocks:
        text += " 涉及个股已补全为：" + "；".join(f"{s['alias']}→{s['name']}({s['code']})/{s['direction']}" for s in stocks) + "。"
    text += f" 动作倾向更接近：{action}。"
    return text


def make_analysis(post: dict[str, Any], market_days: dict[str, MarketDay], resolver: StockResolver) -> dict[str, Any]:
    day = post["posted_at"][:10]
    cls, reason = classify_post(post)
    stocks = resolver.detect(post["body"]) if cls != "C" else []
    market = market_summary(day, market_days) if cls != "C" else ""
    referenced_markets = referenced_market_notes(post["body"], post["_dt"], market_days) if cls != "C" else []
    sectors = detect_sectors(post["body"])
    terms = detect_terms(post["body"])
    tech = technical_notes(post["body"])
    action = detect_action(post["body"]) if cls != "C" else "不展开"
    chain = []
    if cls != "C":
        chain = [
            f"时间位置：{time_stage(post['_dt'])}，盘前、盘中和收盘后观点不能混用。",
            f"层级判断：{'、'.join(sectors) if sectors else '未明显点名板块，侧重交易纪律/盘面结构'}。",
            f"个股/简称：{'；'.join(f'{s['alias']}→{s['name']}({s['code']})/{s['direction']}' for s in stocks) if stocks else '未识别到明确个股或简称'}。",
            f"动作条件：更接近“{action}”，但只有原文条件和当时盘面同时满足才成立。",
        ]
    return {
        "class": cls,
        "reason": reason,
        "translation": plain_language(post, cls, stocks),
        "market": market,
        "referenced_markets": referenced_markets,
        "sectors": sectors,
        "stocks": stocks,
        "terms": terms,
        "tech": tech,
        "action": action,
        "chain": chain,
    }


def set_cell_shading(cell: Any, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_font(run: Any, size: float, bold: bool = False, color: str | None = None) -> None:
    run.font.name = "Microsoft YaHei"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor.from_string(color)


def style_document(doc: Document) -> None:
    section = doc.sections[0]
    section.top_margin = Pt(48)
    section.bottom_margin = Pt(48)
    section.left_margin = Pt(54)
    section.right_margin = Pt(54)
    styles = doc.styles
    for name in ["Normal", "Title", "Heading 1", "Heading 2", "Heading 3"]:
        style = styles[name]
        style.font.name = "Microsoft YaHei"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Microsoft YaHei")
    styles["Normal"].font.size = Pt(10)
    styles["Title"].font.size = Pt(20)
    styles["Heading 1"].font.size = Pt(15)
    styles["Heading 2"].font.size = Pt(12)


def add_cell_line(cell: Any, head: str, value: str, head_color: str = "111827") -> None:
    if not value:
        return
    paragraph = cell.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(2)
    label = paragraph.add_run(f"{head}：")
    set_font(label, 9.5, bold=True, color=head_color)
    body = paragraph.add_run(value)
    set_font(body, 9.5, color="374151")


def add_post_card(doc: Document, post: dict[str, Any], analysis: dict[str, Any]) -> None:
    is_wolf = post["uid"] == WOLF_UID
    has_analysis = analysis["class"] != "C"
    table = doc.add_table(rows=3 if has_analysis else 2, cols=1)
    table.style = "Table Grid"
    header = table.rows[0].cells[0]
    set_cell_shading(header, "DFF3E4" if is_wolf else "F3F4F6")
    hp = header.paragraphs[0]
    hp.paragraph_format.space_after = Pt(0)
    run = hp.add_run(f"{post['username']}  |  {post['time']}  |  {post['floor']}楼  |  UID {post['uid']}  |  {analysis['class']}类")
    set_font(run, 11.5 if is_wolf else 9.5, bold=True, color="166534" if is_wolf else "374151")

    body_cell = table.rows[1].cells[0]
    for idx, paragraph_text in enumerate(post["body"].split("\n\n")):
        para = body_cell.paragraphs[0] if idx == 0 else body_cell.add_paragraph()
        para.paragraph_format.line_spacing = 1.15
        para.paragraph_format.space_after = Pt(4)
        for line_idx, line in enumerate(paragraph_text.split("\n")):
            if line_idx:
                para.add_run().add_break()
            rr = para.add_run(line.strip())
            set_font(rr, 11 if is_wolf else 9.5)

    if has_analysis:
        analysis_cell = table.rows[2].cells[0]
        set_cell_shading(analysis_cell, "F7FCF8" if is_wolf else "FAFAFA")
        title = "狼大发言详析" if is_wolf else "发言分析"
        p = analysis_cell.paragraphs[0]
        p.paragraph_format.space_after = Pt(2)
        r = p.add_run(title)
        set_font(r, 10.5 if is_wolf else 9.8, bold=True, color="166534" if is_wolf else "111827")
        add_cell_line(analysis_cell, "一句话还原", analysis["translation"])
        if analysis["sectors"]:
            add_cell_line(analysis_cell, "板块/细分方向", "、".join(analysis["sectors"]))
        if analysis["stocks"]:
            add_cell_line(
                analysis_cell,
                "个股简称补全",
                "；".join(f"{s['alias']} → {s['name']}（{s['code']}，{s['direction']}，{s['source']}）" for s in analysis["stocks"]),
            )
        if analysis["referenced_markets"]:
            add_cell_line(analysis_cell, "跨日市场环境", "；".join(analysis["referenced_markets"]))
        if analysis["terms"]:
            add_cell_line(analysis_cell, "术语解释", "；".join(f"{term}：{desc}" for term, desc in analysis["terms"]))
        if analysis["tech"]:
            add_cell_line(analysis_cell, "技术/量能结构", "；".join(analysis["tech"]))
        add_cell_line(
            analysis_cell,
            "动作和风控",
            f"更接近“{analysis['action']}”。先确认触发条件、失效条件和退出条件，不能脱离原文条件照抄。",
        )
        add_cell_line(analysis_cell, "判断链", " / ".join(analysis["chain"]))
    doc.add_paragraph()


def add_day_toc(doc: Document, day_posts: list[dict[str, Any]]) -> None:
    doc.add_heading("当日发言目录", level=2)
    table = doc.add_table(rows=1, cols=5)
    table.style = "Table Grid"
    headers = ["用户", "时间", "楼层", "类型", "方向提示"]
    for index, head in enumerate(headers):
        cell = table.rows[0].cells[index]
        set_cell_shading(cell, "E5E7EB")
        run = cell.paragraphs[0].add_run(head)
        set_font(run, 9, bold=True)
    for post in day_posts:
        cls, _ = classify_post(post)
        sectors = detect_sectors(post["body"])
        row = table.add_row().cells
        values = [
            post["username"],
            post["time"][11:16],
            f"{post['floor']}楼",
            f"{cls}类",
            "、".join(sectors[:3]) if sectors else "原文保留",
        ]
        for cell, value in zip(row, values):
            cell.paragraphs[0].paragraph_format.space_after = Pt(0)
            run = cell.paragraphs[0].add_run(value)
            set_font(run, 8.5, color="374151")
    doc.add_paragraph()


def build_doc(posts: list[dict[str, Any]], market_days: dict[str, MarketDay], resolver: StockResolver, out_path: Path) -> None:
    doc = Document()
    style_document(doc)
    title = doc.add_paragraph(style="Title")
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.add_run(f"{THREAD_TITLE}：2026年6月发言逐条分析")

    wolf_count = sum(1 for p in posts if p["uid"] == WOLF_UID)
    summary = doc.add_paragraph()
    summary.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = summary.add_run(f"共 {len(posts)} 条；狼大 {wolf_count} 条重点详析；其他 A/B 类发言按同框架分析；按发言时间顺序排列")
    set_font(r, 10, color="6B7280")

    doc.add_heading("用户发言目录", level=1)
    counts = Counter((p["uid"], p["username"]) for p in posts)
    overview = doc.add_table(rows=1, cols=3)
    overview.style = "Table Grid"
    for i, head in enumerate(["用户", "UID", "6月发言数"]):
        cell = overview.rows[0].cells[i]
        set_cell_shading(cell, "E5E7EB")
        rr = cell.paragraphs[0].add_run(head)
        set_font(rr, 9.5, bold=True)
    for (uid, name), count in counts.most_common():
        row = overview.add_row().cells
        row[0].text = name
        row[1].text = uid
        row[2].text = str(count)
    doc.add_paragraph()

    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for post in posts:
        grouped[post["posted_at"][:10]].append(post)

    first_day = True
    for day in sorted(grouped):
        if not first_day:
            doc.add_section(WD_SECTION_START.NEW_PAGE)
        first_day = False
        day_posts = grouped[day]
        doc.add_heading(f"{day}（{len(day_posts)}条）", level=1)
        p = doc.add_paragraph()
        rr = p.add_run("当日指数环境：")
        set_font(rr, 10, bold=True)
        vv = p.add_run(market_summary(day, market_days))
        set_font(vv, 10)
        add_day_toc(doc, day_posts)

        for post in day_posts:
            analysis = make_analysis(post, market_days, resolver)
            add_post_card(doc, post, analysis)

    doc.save(out_path)


def write_markdown(posts: list[dict[str, Any]], market_days: dict[str, MarketDay], resolver: StockResolver, out_path: Path) -> None:
    lines = [f"# {THREAD_TITLE}：2026年6月发言逐条分析", ""]
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for post in posts:
        grouped[post["posted_at"][:10]].append(post)
    for day in sorted(grouped):
        day_posts = grouped[day]
        lines.append(f"## {day}（{len(day_posts)}条）")
        lines.append("")
        lines.append(f"当日指数环境：{market_summary(day, market_days)}")
        lines.append("")
        lines.append("### 当日发言目录")
        lines.append("")
        lines.append("| 用户 | 时间 | 楼层 | 类型 | 方向提示 |")
        lines.append("| --- | --- | --- | --- | --- |")
        for post in day_posts:
            cls, _ = classify_post(post)
            sectors = detect_sectors(post["body"])
            hint = "、".join(sectors[:3]) if sectors else "原文保留"
            lines.append(f"| {post['username']} | {post['time'][11:16]} | {post['floor']}楼 | {cls}类 | {hint} |")
        lines.append("")
        for post in day_posts:
            analysis = make_analysis(post, market_days, resolver)
            lines.append(f"### {post['username']} | {post['time']} | {post['floor']}楼 | UID {post['uid']} | {analysis['class']}类")
            lines.append("")
            lines.append(post["body"])
            lines.append("")
            if analysis["class"] != "C":
                lines.append(f"**{'狼大发言详析' if post['uid'] == WOLF_UID else '发言分析'}**")
                lines.append("")
                lines.append(f"- 一句话还原：{analysis['translation']}")
                if analysis["sectors"]:
                    lines.append(f"- 板块/细分方向：{'、'.join(analysis['sectors'])}")
                if analysis["stocks"]:
                    lines.append("- 个股简称补全：" + "；".join(f"{s['alias']} → {s['name']}（{s['code']}，{s['direction']}，{s['source']}）" for s in analysis["stocks"]))
                if analysis["referenced_markets"]:
                    lines.append("- 跨日市场环境：" + "；".join(analysis["referenced_markets"]))
                if analysis["terms"]:
                    lines.append("- 术语解释：" + "；".join(f"{term}：{desc}" for term, desc in analysis["terms"]))
                if analysis["tech"]:
                    lines.append("- 技术/量能结构：" + "；".join(analysis["tech"]))
                lines.append(f"- 动作和风控：更接近“{analysis['action']}”。需要先确认触发条件、失效条件和退出条件。")
                lines.append("- 判断链还原：" + " / ".join(analysis["chain"]))
            lines.append("")
    out_path.write_text("\n".join(lines), encoding="utf-8")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Build analyzed June document.")
    parser.add_argument("--input", default="selected_users_posts.json")
    parser.add_argument("--month", default="2026-06")
    parser.add_argument("--out", default="monthly_docs/科学技术打头阵_发言逐条分析_2026-06.docx")
    parser.add_argument("--markdown", default="monthly_docs/科学技术打头阵_发言逐条分析_2026-06.md")
    parser.add_argument("--market-cache", default="market_cache_2026-06.json")
    parser.add_argument("--stock-cache", default="stock_name_cache.json")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    posts = load_posts(Path(args.input), args.month)
    begin = args.month.replace("-", "") + "01"
    end = args.month.replace("-", "") + "30"
    market_days = fetch_index_klines(begin, end, Path(args.market_cache))
    resolver = StockResolver(Path(args.stock_cache))
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    build_doc(posts, market_days, resolver, out)
    write_markdown(posts, market_days, resolver, Path(args.markdown))
    resolver.save()
    print(f"Generated {out} with {len(posts)} posts.")
    print(f"Generated {args.markdown}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
