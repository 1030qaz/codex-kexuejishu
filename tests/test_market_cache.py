import json
import tempfile
import unittest
from datetime import datetime
from pathlib import Path
from unittest.mock import patch

import build_june_analyzed_doc as analyzed


class MarketCacheTests(unittest.TestCase):
    def test_current_day_cache_prevents_future_month_fetch(self) -> None:
        today = datetime.now()
        month = today.strftime("%Y%m")
        today_dash = today.strftime("%Y-%m-%d")
        cache_payload = {
            today_dash: {
                "date": today_dash,
                "indexes": {
                    "上证指数": {
                        "open": 1.0,
                        "close": 1.0,
                        "high": 1.0,
                        "low": 1.0,
                        "volume": 1.0,
                        "amount": 1.0,
                        "amplitude": 0.0,
                        "pct_change": 0.0,
                        "change": 0.0,
                        "turnover": 0.0,
                        "source": "test",
                    }
                },
            }
        }
        with tempfile.TemporaryDirectory() as tmp:
            cache_path = Path(tmp) / "market_cache.json"
            cache_path.write_text(json.dumps(cache_payload, ensure_ascii=False), encoding="utf-8")
            with patch.object(analyzed, "fetch_eastmoney_klines", side_effect=AssertionError("unexpected fetch")):
                market_days = analyzed.fetch_index_klines(f"{month}01", f"{month}31", cache_path)

        self.assertEqual([today_dash], list(market_days))
        self.assertEqual("test", market_days[today_dash].indexes["上证指数"]["source"])


if __name__ == "__main__":
    unittest.main()
