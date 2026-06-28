#!/usr/bin/env python3
"""
Update all tracked NGA users and rebuild only the current month's documents.

Historical monthly documents are intentionally left untouched.
"""

from __future__ import annotations

import argparse
import json
import os
import subprocess
import sys
from dataclasses import asdict
from datetime import datetime
from pathlib import Path

from build_multiuser_monthly_docs import UserPost, build_month_doc, sorted_posts, write_markdown


THREAD_TITLE = "科学技术打头阵"


def default_month() -> str:
    return datetime.now().strftime("%Y-%m")


def load_posts(path: Path) -> list[UserPost]:
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    return sorted_posts([UserPost(**item) for item in data])


def run_step(command: list[str]) -> None:
    print("Running: " + " ".join(command), flush=True)
    subprocess.run(command, check=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Update all users and rebuild only one month.")
    parser.add_argument("--month", default=default_month(), help="Month to rebuild, e.g. 2026-06.")
    parser.add_argument("--combined-json", default="selected_users_posts.json")
    parser.add_argument("--output-dir", default="monthly_docs")
    parser.add_argument("--delay", type=float, default=2.2)
    parser.add_argument("--timeout", type=int, default=35)
    parser.add_argument("--overlap", type=int, default=2)
    parser.add_argument("--page-limit", type=int, default=50)
    parser.add_argument("--skip-scrape", action="store_true", help="Only rebuild month docs from existing JSON.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    combined_json = Path(args.combined_json)
    output_dir = Path(args.output_dir)

    if not args.skip_scrape:
        if not os.environ.get("NGA_COOKIE", "").strip():
            print("NGA_COOKIE is required unless --skip-scrape is used.")
            return 2
        run_step(
            [
                sys.executable,
                "update_latest_posts.py",
                "--combined-json",
                str(combined_json),
                "--delay",
                str(args.delay),
                "--timeout",
                str(args.timeout),
                "--overlap",
                str(args.overlap),
                "--page-limit",
                str(args.page_limit),
            ]
        )

    posts = load_posts(combined_json)
    month_posts = [post for post in posts if post.posted_at.startswith(args.month)]
    if not month_posts:
        print(f"No posts found for {args.month}.")
        return 1

    output_dir.mkdir(parents=True, exist_ok=True)
    month_path = build_month_doc(args.month, month_posts, output_dir)
    write_markdown(month_posts, output_dir)
    print(f"Generated current-month整理文档: {month_path}")

    market_cache = Path(f"market_cache_{args.month}.json")
    analysis_doc = output_dir / f"{THREAD_TITLE}_发言逐条分析_{args.month}.docx"
    analysis_md = output_dir / f"{THREAD_TITLE}_发言逐条分析_{args.month}.md"
    run_step(
        [
            sys.executable,
            "build_june_analyzed_doc.py",
            "--input",
            str(combined_json),
            "--month",
            args.month,
            "--out",
            str(analysis_doc),
            "--markdown",
            str(analysis_md),
            "--market-cache",
            str(market_cache),
        ]
    )

    latest = max(post.posted_at for post in posts)
    latest_month = max(post.posted_at for post in month_posts)
    print(
        json.dumps(
            {
                "updated_at": datetime.now().isoformat(timespec="minutes"),
                "total_posts": len(posts),
                "month": args.month,
                "month_posts": len(month_posts),
                "latest_post": latest,
                "latest_month_post": latest_month,
                "month_doc": str(month_path),
                "analysis_doc": str(analysis_doc),
                "analysis_markdown": str(analysis_md),
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
