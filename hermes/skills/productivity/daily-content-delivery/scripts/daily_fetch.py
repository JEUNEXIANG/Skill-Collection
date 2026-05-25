#!/usr/bin/env python3
"""
Daily content fetcher for Hermes cron jobs.
Fetches from priority sources via RSS + HTTP scraping.
Outputs structured raw content for LLM processing into formatted delivery.
"""

import re
from datetime import datetime

import requests
import feedparser

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/125.0.0.0 Safari/537.36"
}
TIMEOUT = 20
TODAY = datetime.now().strftime("%Y-%m-%d")
HOUR = datetime.now().strftime("%H:%M")


def safe_get(url):
    try:
        resp = requests.get(url, headers=HEADERS, timeout=TIMEOUT)
        resp.raise_for_status()
        return resp.text
    except Exception:
        return None


def safe_feed(url):
    try:
        return feedparser.parse(url)
    except Exception:
        return None


def print_header(s):
    print(f"\n{'='*60}")
    print(f"  {s}  |  {TODAY} {HOUR}")
    print('=' * 60)


# ===== Part 1: Priority Tech Sources =====

def fetch_google_dev_feed():
    """Google Developers Blog via RSS — best source for AI Agent content."""
    print_header("【谷歌开发者博客 Google Developers Blog】")
    feed = safe_feed("https://developers.googleblog.com/feeds/posts/default")
    if not feed or not feed.entries:
        print("  ⚠️ RSS unavailable")
        return
    for entry in feed.entries[:10]:
        print(f"  📄 {entry.title}")
        print(f"     {entry.link}")
        if hasattr(entry, 'summary'):
            summary = re.sub(r'<[^>]+>', '', entry.summary).strip()[:300]
            print(f"     {summary}...")


def fetch_aliyun_blog():
    """阿里云开发者社区 — try atom feed first, then HTTP fallback."""
    print_header("【阿里云开发者社区】")
    feed = safe_feed("https://developer.aliyun.com/blog/atom.xml")
    if feed and feed.entries:
        for entry in feed.entries[:10]:
            print(f"  📄 {entry.title}")
            print(f"     {entry.link}")
        return
    html = safe_get("https://developer.aliyun.com/blog/")
    if html:
        titles = re.findall(r'data-title=["\']([^"\']+)["\']', html)
        links = re.findall(r'href=["\'](https?://developer\.aliyun\.com/article/\d+)["\']', html)
        seen = set()
        for title, link in zip(titles, links):
            if link not in seen:
                seen.add(link)
                print(f"  📄 {title}")
                print(f"     {link}")
        if not seen:
            print("  ⚠️ JS-heavy page; try browser in interactive mode")
            print(f"  ℹ️ Manual: https://developer.aliyun.com/blog/")
    else:
        print("  ⚠️ Unreachable")


def fetch_tencent_dev():
    """腾讯云开发者社区."""
    print_header("【腾讯云开发者社区】")
    html = safe_get("https://cloud.tencent.com/developer")
    if html:
        links = re.findall(
            r'href=["\'](https?://cloud\.tencent\.com/developer/article/\d+)["\']', html
        )
        seen = set()
        for link in links:
            if link not in seen:
                seen.add(link)
        for link in list(seen)[:10]:
            print(f"  📄 (article ID: {link.split('/')[-1]})")
            print(f"     {link}")
    else:
        print("  ⚠️ Unreachable")


def fetch_ifeng_tech():
    """凤凰网科技 — headline extraction with relevance filter."""
    print_header("【凤凰网科技 - 今日热点】")
    html = safe_get("https://tech.ifeng.com")
    if not html:
        print("  ⚠️ Unreachable")
        return

    # Extract article links
    links = re.findall(r'href=["\'](https?://tech\.ifeng\.com/c/[^"\']+)["\']', html)
    seen_links = set()
    for link in links:
        if link not in seen_links:
            seen_links.add(link)

    # Filter for relevant headlines using keyword matching
    KEYWORDS = ['AI', '大模型', 'DeepSeek', 'Agent', '智能', '芯片', '融资',
                '谷歌', '微软', '苹果', '华为', '字节', '腾讯', '阿里', '百度',
                '数据', '财报', '非农', '关税', '黄金', '美股', 'A股', 'IPO',
                '机器人', '自动驾驶', '云计算', '开源']
    seen_titles = set()
    pattern = r'>([^<]{10,80})</a>'
    for m in re.finditer(pattern, html):
        title = m.group(1).strip()
        if any(k in title for k in KEYWORDS):
            if title not in seen_titles and len(title) > 8:
                seen_titles.add(title)
                print(f"  🔥 {title}")
    for link in list(seen_links)[:15]:
        print(f"     {link}")


def fetch_cls_news():
    """财联社 — market data + headline news."""
    print_header("【财联社 - 要闻速览】")
    html = safe_get("https://www.cls.cn")
    if not html:
        print("  ⚠️ Unreachable")
        return

    # Extract article links with text
    links = re.findall(r'href=["\'](/detail/\d+)["\'][^>]*>([^<]+)', html)
    for url, text in links[:20]:
        text = text.strip()
        if text and len(text) > 10:
            print(f"  📰 {text}")
            print(f"     https://www.cls.cn{url}")

    # Extract market data (indices embedded in HTML)
    for pattern, name in [
        (r'上证指数[^<]*<[^>]*>([\d,.]+)', '上证指数'),
        (r'深证成指[^<]*<[^>]*>([\d,.]+)', '深证成指'),
        (r'创业板指[^<]*<[^>]*>([\d,.]+)', '创业板指'),
    ]:
        m = re.search(pattern, html)
        if m:
            print(f"  📊 {name}: {m.group(1)}")


# ===== Main =====

def main():
    print("=" * 60)
    print(f"  HERMES 每日内容抓取")
    print(f"  {TODAY} {HOUR}")
    print("=" * 60)

    print(f"\n{'#'*60}")
    print(f"# 第一部分：深度技术文章")
    print(f"{'#'*60}")
    fetch_google_dev_feed()
    fetch_aliyun_blog()
    fetch_tencent_dev()

    print(f"\n{'#'*60}")
    print(f"# 第二部分：新闻速报 + 数据")
    print(f"{'#'*60}")
    fetch_ifeng_tech()
    fetch_cls_news()

    print(f"\n{'='*60}")
    print(f"  ✅ 抓取完成 | 来源: Google Blog RSS / 凤凰网科技 / 财联社")
    print(f"{'='*60}")


if __name__ == "__main__":
    main()
