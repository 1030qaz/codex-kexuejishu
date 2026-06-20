#!/usr/bin/env python3
"""
Incrementally update selected NGA user posts for thread 45974302.

This script starts near each user's last captured page, follows pagination to
the current end, de-duplicates posts, and rewrites selected_users_posts.json.
"""

from __future__ import annotations

import argparse
import json
import os
import random
import time
from dataclasses import asdict
from pathlib import Path
from urllib.error import HTTPError, URLError

from build_multiuser_monthly_docs import (
    TARGET_USERS,
    THREAD_URL,
    UserPost,
    clean_body,
    parse_time,
    sorted_posts,
)
from scrape_nga_user_thread import build_page_url, fetch, has_next_page, parse_posts


def load_existing(path: Path) -> list[UserPost]:
    if not path.exists():
        return []
    data = json.loads(path.read_text(encoding="utf-8-sig"))
    return [UserPost(**item) for item in data]


def post_key(post: UserPost) -> tuple[str, str, str, str]:
    return (post.uid, post.floor, post.posted_at, post.body)


def page_window(posts: list[UserPost], uid: str, overlap: int) -> int:
    pages = [post.page for post in posts if post.uid == uid and isinstance(post.page, int)]
    if not pages:
        return 1
    return max(1, max(pages) - overlap)


def fetch_user_updates(
    uid: str,
    username: str,
    existing: list[UserPost],
    seen: set[tuple[str, str, str, str]],
    cookie: str,
    delay: float,
    timeout: int,
    overlap: int,
    page_limit: int,
) -> list[UserPost]:
    start_page = page_window(existing, uid, overlap)
    start_url = f"{THREAD_URL}&authorid={uid}"
    new_posts: list[UserPost] = []

    page = start_page
    last_page = start_page + page_limit
    while page <= last_page:
        page_url = build_page_url(start_url, page)
        print(f"Fetching {username} ({uid}) page {page}", flush=True)
        try:
            page_html = fetch(page_url, cookie, timeout)
        except (HTTPError, URLError, TimeoutError, RuntimeError) as exc:
            raise RuntimeError(f"Failed fetching {username} ({uid}) page {page}: {exc}") from exc

        parsed_posts = parse_posts(page_html, page, page_url)
        for parsed in parsed_posts:
            parsed_time = parse_time(parsed.time)
            if parsed_time is None:
                continue
            post = UserPost(
                uid=uid,
                username=username,
                floor=parsed.floor,
                time=parsed.time,
                posted_at=parsed_time.isoformat(timespec="minutes"),
                body=clean_body(parsed.body),
                page=page,
                url=page_url,
            )
            key = post_key(post)
            if key not in seen:
                seen.add(key)
                new_posts.append(post)

        if not has_next_page(page_html, page):
            break
        page += 1
        time.sleep(delay + random.uniform(0, delay * 0.4))

    return new_posts


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Incrementally update selected NGA user posts.")
    parser.add_argument("--combined-json", default="selected_users_posts.json")
    parser.add_argument("--delay", type=float, default=0.6)
    parser.add_argument("--timeout", type=int, default=25)
    parser.add_argument("--overlap", type=int, default=2, help="Pages before last captured page to rescan.")
    parser.add_argument("--page-limit", type=int, default=50, help="Safety cap beyond the start page per user.")
    parser.add_argument(
        "--allow-failures",
        action="store_true",
        help="Save successful users and return 0 even if one or more users fail.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    cookie = os.environ.get("NGA_COOKIE", "").strip()
    if not cookie:
        print("NGA_COOKIE is required.")
        return 2

    output_json = Path(args.combined_json)
    existing = load_existing(output_json)
    seen = {post_key(post) for post in existing}
    all_posts = list(existing)

    total_added = 0
    failures: list[str] = []
    for index, (uid, username) in enumerate(TARGET_USERS, start=1):
        before = len(all_posts)
        print(f"[{index}/{len(TARGET_USERS)}] Updating {username} ({uid})", flush=True)
        try:
            updates = fetch_user_updates(
                uid=uid,
                username=username,
                existing=all_posts,
                seen=seen,
                cookie=cookie,
                delay=args.delay,
                timeout=args.timeout,
                overlap=args.overlap,
                page_limit=args.page_limit,
            )
        except RuntimeError as exc:
            print(f"  warning: {exc}; continuing with next user", flush=True)
            failures.append(f"{username} ({uid})")
            updates = []
        all_posts.extend(updates)
        all_posts = sorted_posts(all_posts)
        total_added += len(all_posts) - before
        output_json.write_text(
            json.dumps([asdict(post) for post in all_posts], ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        latest = max((p.posted_at for p in all_posts if p.uid == uid), default="n/a")
        print(f"  added {len(updates)}; latest {latest}", flush=True)
        time.sleep(args.delay)

    latest_all = max((post.posted_at for post in all_posts), default="n/a")
    print(f"Saved {len(all_posts)} posts to {output_json}; added {total_added}; latest {latest_all}")
    if failures:
        print("Failed users: " + ", ".join(failures))
        return 0 if args.allow_failures else 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
