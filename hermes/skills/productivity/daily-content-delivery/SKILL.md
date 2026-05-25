---
name: daily-content-delivery
description: Set up automated daily content aggregation + LLM curation + scheduled delivery to WeChat (or other platforms). Combines RSS feeds, HTTP scraping, and browser fallback to gather articles/news from Chinese tech sources, then uses an LLM to filter, summarize, explain concepts, and format into structured push notifications via cron jobs.
tags: [cron, rss, scraping, wechat, content-curation, daily-briefing, agent-pipeline]
---

# Daily Content Delivery Pipeline

Set up an automated pipeline that fetches content daily from multiple sources, curates it with an LLM, and delivers it to WeChat (or any connected platform).

## Architecture

```
cron trigger
    │
    ├── script (Python) → fetches raw content via RSS + HTTP + browser
    │   └── stdout injected as context into LLM prompt
    │
    └── LLM prompt → processes raw content into structured format
        └── result delivered to platform (WeChat, Telegram, etc.)
```

## Key Sources and Their Availability

| Source | URL | Method | Status | Best For |
|--------|-----|--------|--------|----------|
| Google Developers Blog | developers.googleblog.com | RSS (`/feeds/posts/default`) | ✅ Best source | Tech deep dive |
| 凤凰网科技 | tech.ifeng.com | HTTP + browser | ✅ Good | Daily news |
| 财联社 | www.cls.cn | HTTP + browser | ✅ Good | News + market data |
| 科创板日报 | www.chinastarmarket.cn | HTTP | ⚠️ JS-blocked, limited | Tech news |
| AIGC开放社区 | www.aigcopen.com | HTTP | ⚠️ WordPress, limited parsing | AIGC articles |
| 阿里云开发者 | developer.aliyun.com | HTTP | ⚠️ Heavy JS, limited parsing | Tech articles |
| 腾讯云开发者 | cloud.tencent.com/developer | HTTP | ⚠️ Heavy JS, limited parsing | Tech articles |
| 36氪 | 36kr.com | HTTP | ⚠️ JS-rendered | Tech business |
| 哈佛商业评论 | hbrchina.org | HTTP | ✅ Available | Business depth |
| 路透 | cn.reuters.com | HTTP | ⚠️ 401 blocked | — |

### Why RSS is best
- Google Developers Blog RSS (`/feeds/posts/default`) returns clean, structured entries with title, link, and summary
- No JS rendering needed
- Reliable and fast

### When to use browser vs HTTP
- **HTTP/requests**: Fast, works for server-rendered pages (凤凰网, 财联社, AIGC开放社区)
- **Browser tool**: Needed for JS-heavy single-page apps (some sites). But **not available in cron scripts** — cron can only use terminal/HTTP tools
- For cron jobs: stick to HTTP + RSS. Browser scraping is only for interactive sessions

## Time-Prioritized Split Delivery Pattern

For best engagement, split content into separate cron jobs at different times:

| Push | Time | Content | Script Arg |
|------|------|---------|-----------|
| 📰 News brief | **08:00** daily | Categorized headlines + market data with brief analysis | `--section news` |
| 📖 Tech deep dive | **12:30** daily | 2-4 articles with full summaries + concept explanations | `--section tech` |
| 💼 Jobs | **10:00** every 3 days | Filtered job listings from LinkedIn/Boss直聘/etc | (no script) |

```bash
# News only (daily 08:00)
hermes cron create \
  --name "每日科技早报" \
  --schedule "0 8 * * *" \
  --script "daily_fetch.py --section news"

# Tech articles only (daily 12:30)
hermes cron create \
  --name "深度技术文章推送" \
  --schedule "30 12 * * *" \
  --script "daily_fetch.py --section tech"

# Jobs (every 3 days, 10:00)
hermes cron create \
  --name "职位推送" \
  --schedule "0 10 */3 * *"
```

### Job search cron (separate pattern)

Unlike news/tech, job search needs **interactive tools** (browser, web_search). Don't use `--script` for this — put all instructions in the `--prompt`:

```
--prompt "你是一个求职猎头。请用浏览器工具访问LinkedIn/Boss直聘...
      搜索AI Agent产品经理/商业分析岗位，筛选后推送给用户。"
```

The LLM in the cron session will use its tools to search live, rather than relying on pre-fetched data.

### 1. Create the fetch script with `--section` argument

Place at `~/.hermes/scripts/daily_fetch.py`. Use argparse to support separated deliveries:

```python
import argparse

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--section', choices=['news', 'tech', 'all'], default='all')
    args = parser.parse_args()
    
    if args.section in ('tech', 'all'):
        # fetch Google Dev Blog, 阿里云, 腾讯云
        pass
    if args.section in ('news', 'all'):
        # fetch 凤凰网科技, 财联社, 科创板日报
        pass
```

This allows splitting content into separate cron jobs without duplicating code.

```python
import requests
import feedparser

# RSS feed
feed = feedparser.parse("https://developers.googleblog.com/feeds/posts/default")
for entry in feed.entries[:10]:
    print(f"📄 {entry.title}")
    print(f"   {entry.link}")
    # summary is in entry.summary (HTML, strip tags)

# HTTP scraping (use regex for titles/links)
html = requests.get("https://tech.ifeng.com", headers={"User-Agent": "..."}).text
# extract headlines with regex patterns
```

Dependencies: `requests`, `feedparser`, `beautifulsoup4`, `lxml`
Install via: `python -m pip install feedparser beautifulsoup4 lxml`

### 2. Create cron job

```bash
# Daily delivery at 08:00
hermes cron create \
  --name "每日AI科技早报" \
  --schedule "0 8 * * *" \
  --script daily_fetch.py \
  --prompt "你是一个专业的科技内容编辑..."
```

The `--script` output is injected as context before the prompt.

### 3. Structured LLM prompt

The prompt should instruct the LLM to produce three parts:

**Part 1: Deep tech articles** — Pick 3-5 most relevant articles (prioritize AI Agent, LLM, infra), include:
- Title + source
- 200-300 word summary
- Key concept explanations (complex terms explained simply)
- Why it matters from a PM/BA perspective

**Part 2: News + data** — Categorized headlines with interpretation:
- Market data (indexes, commodities)
- Top 5 tech news with brief analysis
- Macro news summary

**Part 3: Concept explainer** — 2-3 deep concept dives:
- What it is
- Why it matters
- Application perspective

### 4. For job posting delivery (every 3 days)

Create a separate cron job without a script:

```bash
hermes cron create \
  --name "AI产品经理职位推送" \
  --schedule "0 10 */3 * *" \
  --prompt "你是一个求职猎头..."
```

The prompt itself should instruct the LLM to use browser/web search tools to find positions.

## Important: Cron job limitations

- Cron scripts (`--script`) run **before** the LLM session. Only stdout is injected as context
- Cron scripts CANNOT use Hermes tools (browser, web_search, etc.)
- Cron sessions CAN use Hermes tools (browser, web_search, etc.) — the LLM decides
- Scripts are good for data collection; prompts with tool instructions are good for interactive search

## Pitfalls

- **Some Chinese sites block non-browser requests** → use proper User-Agent headers
- **财联社** articles often require clicking through → extract from homepage list
- **JS-heavy sites** (阿里云, 腾讯云开发者) → HTTP only gets partial data; browser needed for full content. These are NOT ideal for cron scripts since cron can't use browser
- **Market data** from 财联社 is embedded in HTML text, not structured — regex extraction needed
- **Tavily API** may be down → fall back to browser tool for interactive sessions, or RSS for cron. Cron scripts rely on HTTP requests, not Tavily
- **Memory limit** (2200 chars) → keep cron job IDs and configs compact, remove old entries
- **Time zones** → cron schedules are in system timezone. Use `0 8 * * *` for 8 AM
- **Cron scripts can't approve dangerous commands** — script output is injected directly as context before LLM processes it. Keep scripts safe (read-only HTTP requests only)
- **Dangerous commands in interactive sessions** — always explain what the command does, what it affects, and the risk level before executing
- **Google Developers Blog RSS** is the most reliable high-quality source; prioritize it for tech deep dives

## Testing

Before deploying, test the fetch script:
```bash
cd ~/.hermes/hermes-agent && source venv/bin/activate
python ~/.hermes/scripts/daily_fetch.py
```

Then test the cron job immediately:
```bash
hermes cron run <job_id>
```

Check cron logs for errors and format issues.
