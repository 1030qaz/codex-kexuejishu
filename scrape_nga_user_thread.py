#!/usr/bin/env python3
"""
Scrape all posts from one NGA "only this user" thread view.

Usage:
  $env:NGA_COOKIE = "your_cookie_here"
  python scrape_nga_user_thread.py "https://bbs.nga.cn/read.php?tid=45974302&authorid=42587317" --out wolf_posts

The start URL should be the "只看该用户" link. The script keeps all query
parameters from that URL and only changes the page number while crawling.
"""

from __future__ import annotations

import argparse
import html
import json
import os
import random
import re
import sys
import time
from dataclasses import asdict, dataclass
from html.parser import HTMLParser
from typing import Iterable
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qsl, urlencode, urljoin, urlparse, urlunparse
from urllib.request import Request, urlopen

try:
    import requests
except ImportError:  # pragma: no cover - urllib fallback remains available.
    requests = None


DEFAULT_BASE = "https://bbs.nga.cn/read.php?tid=45974302&authorid=42587317"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/125.0 Safari/537.36"
)


@dataclass
class Post:
    floor: str
    time: str
    body: str
    page: int
    url: str


class PostBlockParser(HTMLParser):
    """Small, dependency-free parser for NGA read.php post blocks."""

    def __init__(self) -> None:
        super().__init__(convert_charrefs=False)
        self.posts: list[dict[str, str]] = []
        self._stack: list[tuple[str, dict[str, str]]] = []
        self._in_post = False
        self._post_depth = 0
        self._current: dict[str, str] | None = None
        self._text_target: str | None = None
        self._skip_depth = 0
        self._content_depth = 0
        self._buffer: list[str] = []

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        attr = {k: v or "" for k, v in attrs}
        self._stack.append((tag, attr))
        classes = set(attr.get("class", "").split())
        element_id = attr.get("id", "")

        postcontent_match = re.match(r"^postcontent(\d+)", element_id)
        if postcontent_match or "postcontent" in classes:
            if self._current is None:
                self._current = {"floor": "", "time": "", "body": ""}
            if postcontent_match:
                self._current["floor"] = postcontent_match.group(1)
            self._in_post = True
            self._post_depth = len(self._stack)
            self._content_depth = len(self._stack)
            self._text_target = "body"
            self._buffer = []
            return

        if self._in_post and tag in {"script", "style", "textarea"}:
            self._skip_depth = len(self._stack)
            return

        if self._in_post and tag in {"br", "p", "div", "li", "tr"} and self._text_target == "body":
            self._buffer.append("\n")

    def handle_endtag(self, tag: str) -> None:
        if self._skip_depth and len(self._stack) == self._skip_depth:
            self._skip_depth = 0

        if self._in_post and self._text_target == "body":
            if tag in {"p", "div", "li", "tr"}:
                self._buffer.append("\n")
            if len(self._stack) == self._content_depth:
                body = clean_text("".join(self._buffer))
                if self._current is not None:
                    self._current["body"] = body
                    self.posts.append(self._current)
                self._current = None
                self._in_post = False
                self._post_depth = 0
                self._content_depth = 0
                self._text_target = None
                self._buffer = []

        if self._stack:
            self._stack.pop()

    def handle_data(self, data: str) -> None:
        if self._skip_depth:
            return
        if self._in_post and self._text_target == "body":
            self._buffer.append(data)

    def handle_entityref(self, name: str) -> None:
        self.handle_data(f"&{name};")

    def handle_charref(self, name: str) -> None:
        self.handle_data(f"&#{name};")


def clean_text(value: str) -> str:
    value = html.unescape(value)
    value = value.replace("\u00a0", " ")
    value = re.sub(r"\r\n?", "\n", value)
    value = re.sub(r"[ \t\f\v]+", " ", value)
    value = re.sub(r" *\n *", "\n", value)
    value = re.sub(r"\n{3,}", "\n\n", value)
    return value.strip()


def decode_response(raw: bytes, headers: object) -> str:
    content_type = ""
    try:
        content_type = headers.get("Content-Type", "")  # type: ignore[attr-defined]
    except Exception:
        pass
    match = re.search(r"charset=([\w.-]+)", content_type, re.I)
    encodings = []
    if match:
        encodings.append(match.group(1))
    encodings.extend(["utf-8", "gb18030", "gbk"])

    for encoding in encodings:
        try:
            return raw.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def build_page_url(start_url: str, page: int) -> str:
    parsed = urlparse(start_url)
    query = dict(parse_qsl(parsed.query, keep_blank_values=True))
    query["page"] = str(page)
    new_query = urlencode(query, doseq=True)
    return urlunparse(parsed._replace(query=new_query))


def fetch(url: str, cookie: str, timeout: int, retries: int = 3) -> str:
    if requests is not None:
        headers = {
            "User-Agent": USER_AGENT,
            "Cookie": cookie,
            "Referer": "https://bbs.nga.cn/",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        }
        session = requests.Session()
        session.trust_env = False
        last_error: Exception | None = None
        for attempt in range(1, retries + 1):
            try:
                response = session.get(url, headers=headers, timeout=timeout)
                if response.status_code in {401, 403, 404}:
                    response.raise_for_status()
                response.raise_for_status()
                if "login" in response.url.lower() and "read.php" not in response.url.lower():
                    raise RuntimeError(f"Looks redirected to login page: {response.url}")
                if not response.encoding or response.encoding.lower() in {"iso-8859-1", "ascii"}:
                    response.encoding = response.apparent_encoding or "gb18030"
                return response.text
            except Exception as exc:
                last_error = exc
                if attempt == retries:
                    raise
                time.sleep(1.5 * attempt)
        raise RuntimeError(f"Failed fetching {url}: {last_error}")

    request = Request(
        url,
        headers={
            "User-Agent": USER_AGENT,
            "Cookie": cookie,
            "Referer": "https://bbs.nga.cn/",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        },
    )
    last_error: Exception | None = None
    for attempt in range(1, retries + 1):
        try:
            with urlopen(request, timeout=timeout) as response:
                raw = response.read()
                final_url = response.geturl()
                text = decode_response(raw, response.headers)
            break
        except (HTTPError, URLError, TimeoutError) as exc:
            last_error = exc
            if isinstance(exc, HTTPError) and exc.code in {401, 403, 404}:
                raise
            if attempt == retries:
                raise
            time.sleep(1.5 * attempt)
    else:
        raise RuntimeError(f"Failed fetching {url}: {last_error}")
    if "login" in final_url.lower() and "read.php" not in final_url.lower():
        raise RuntimeError(f"Looks redirected to login page: {final_url}")
    return text


def extract_post_meta(page_html: str) -> list[tuple[str, str]]:
    """Return [(floor, time), ...] found near post content blocks."""
    floors = re.findall(r'id=["\']postnum\d+["\'][^>]*>\s*#?\s*([^<\s]+)', page_html, re.I)
    if not floors:
        floors = re.findall(r'class=["\'][^"\']*\bpostnum\b[^"\']*["\'][^>]*>\s*#?\s*([^<\s]+)', page_html, re.I)

    times = re.findall(r'id=["\']postdate\d+["\'][^>]*>\s*([^<]+)', page_html, re.I)
    if not times:
        times = re.findall(r'class=["\'][^"\']*\bpostdate\b[^"\']*["\'][^>]*>\s*([^<]+)', page_html, re.I)

    count = max(len(floors), len(times))
    result: list[tuple[str, str]] = []
    for index in range(count):
        floor = clean_text(floors[index]) if index < len(floors) else ""
        posted_at = clean_text(times[index]) if index < len(times) else ""
        result.append((floor, posted_at))
    return result


def parse_posts(page_html: str, page: int, page_url: str) -> list[Post]:
    parser = PostBlockParser()
    parser.feed(page_html)
    meta = extract_post_meta(page_html)

    posts: list[Post] = []
    for index, parsed in enumerate(parser.posts):
        floor, posted_at = meta[index] if index < len(meta) else ("", "")
        floor = parsed.get("floor") or floor or f"page-{page}-post-{index + 1}"
        posted_at = parsed.get("time") or posted_at
        body = parsed.get("body", "")
        if body:
            posts.append(Post(floor=floor, time=posted_at, body=body, page=page, url=page_url))
    return posts


def has_next_page(page_html: str, current_page: int) -> bool:
    if re.search(rf'[?&]page={current_page + 1}(?:["\'&\s>]|&amp;)', page_html):
        return True
    page_numbers = [int(x) for x in re.findall(r'[?&]page=(\d+)', page_html) if x.isdigit()]
    return bool(page_numbers and max(page_numbers) > current_page)


def crawl(start_url: str, cookie: str, delay: float, timeout: int, max_pages: int) -> list[Post]:
    all_posts: list[Post] = []
    seen: set[tuple[str, str, str]] = set()

    for page in range(1, max_pages + 1):
        page_url = build_page_url(start_url, page)
        print(f"Fetching page {page}: {page_url}", file=sys.stderr)
        try:
            page_html = fetch(page_url, cookie, timeout)
        except (HTTPError, URLError, TimeoutError, RuntimeError) as exc:
            raise RuntimeError(f"Failed fetching page {page}: {exc}") from exc

        posts = parse_posts(page_html, page, page_url)
        new_posts = []
        for post in posts:
            key = (post.floor, post.time, post.body)
            if key not in seen:
                seen.add(key)
                new_posts.append(post)

        if not new_posts:
            print(f"No new posts found on page {page}; stopping.", file=sys.stderr)
            break

        all_posts.extend(new_posts)
        if not has_next_page(page_html, page):
            print(f"No next-page link after page {page}; stopping.", file=sys.stderr)
            break

        time.sleep(delay + random.uniform(0, delay * 0.5))

    return all_posts


def write_json(posts: Iterable[Post], path: str) -> None:
    with open(path, "w", encoding="utf-8") as file:
        json.dump([asdict(post) for post in posts], file, ensure_ascii=False, indent=2)
        file.write("\n")


def write_markdown(posts: Iterable[Post], path: str) -> None:
    with open(path, "w", encoding="utf-8") as file:
        file.write("# NGA user posts\n\n")
        for post in posts:
            title = f"## {post.floor}"
            if post.time:
                title += f" - {post.time}"
            file.write(title + "\n\n")
            file.write(f"- Page: {post.page}\n")
            file.write(f"- URL: {post.url}\n\n")
            file.write(post.body.strip() + "\n\n")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Scrape NGA only-author thread pages.")
    parser.add_argument(
        "url",
        nargs="?",
        default=DEFAULT_BASE,
        help="NGA 只看该用户 link. Defaults to tid=45974302&authorid=42587317.",
    )
    parser.add_argument("--out", default="nga_user_posts", help="Output filename prefix.")
    parser.add_argument("--delay", type=float, default=1.5, help="Delay between page requests.")
    parser.add_argument("--timeout", type=int, default=20, help="Request timeout in seconds.")
    parser.add_argument("--max-pages", type=int, default=200, help="Safety cap for pagination.")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    cookie = os.environ.get("NGA_COOKIE", "").strip()
    if not cookie:
        print("NGA_COOKIE is required. Set it to your logged-in NGA Cookie.", file=sys.stderr)
        return 2

    start_url = urljoin("https://bbs.nga.cn/", args.url)
    posts = crawl(start_url, cookie, args.delay, args.timeout, args.max_pages)
    json_path = f"{args.out}.json"
    markdown_path = f"{args.out}.md"
    write_json(posts, json_path)
    write_markdown(posts, markdown_path)
    print(f"Saved {len(posts)} posts to {json_path} and {markdown_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
