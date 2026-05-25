f---
name: wechat-daily-curation
description: Agentic content curation pipeline — 3 autonomous cron jobs that research, fetch, curate, and deliver WeChat-pushed content using AI-driven search, RSS parsing, web scraping, and browser automation.
---

# WeChat Daily Content Curation — Agentic Workflow

## Architecture (3 Layers)

```
┌─────────────────────────────────────────────────────┐
│  LAYER 1: Cron Triggers (3 jobs)                    │
│  ┌──────────────┐ ┌──────────────┐ ┌──────────────┐│
│  │ Daily 08:00  │ │ Daily 12:30  │ │ Every 3d 10:00││
│  │ News/Finance │ │ Tech/AI Deep │ │ Job Postings  ││
│  └──────┬───────┘ └──────┬───────┘ └──────┬───────┘│
├─────────┼────────────────┼────────────────┼─────────┤
│  LAYER 2: Agent Runtime (autonomous LLM-driven)     │
│                                                     │
│  Each cron → agent loads Source Registry +          │
│  autonomously decides:                              │
│  • Which sources to check (today's updates)         │
│  • Which tools to use (RSS / HTTP / Browser)        │
│  • What content to select                           │
│  • What concepts to explain                         │
│  • How to format output                             │
│                                                     │
│  Tool selection per source type:                    │
│  ┌──────────┬────────────┬──────────────────┐       │
│  │ RSS Feed │ feedparser │ Terminal: python3 │       │
│  │ HTTP     │ requests+BS4 │ Terminal: python3│       │
│  │ Browser  │ browser_*  │ Built-in tools    │       │
│  │ Search   │ tavily_sh  │ Terminal: bash     │       │
│  └──────────┴────────────┴──────────────────┘       │
├─────────────────────────────────────────────────────┤
│  LAYER 3: Source Registry (27 accounts mapped)      │
│  references/source_registry.md — structured table   │
│  of all accounts with URLs, RSS feeds, access       │
│  methods, and confidence ratings.                   │
└─────────────────────────────────────────────────────┘
```

## Source Registry (27 Accounts)

See `references/source_registry.md` for the full structured table.

### Summary by Task

| Task | Accounts | Web-Accessible | RSS | WeChat-Only |
|------|----------|---------------|-----|-------------|
| A: Tech/AI/Biz | 10 | 8 (80%) | 1 | 2 |
| B: Job Search | 7 | 4 (57%) | 0 | 3 |
| C: News/Finance | 10 | 10 (100%) | 5 | 0 |

### Key RSS Feeds (fastest, most reliable)

| Source | RSS XML URL | Notes |
|--------|------------|-------|
| **腾讯科技** | `https://feedmaker.kindle4rss.com/feeds/qqtech.weixin.xml` | WeChat mirror via kindle4rss, ~10 articles |
| **凤凰网科技** | `http://finance.ifeng.com/rss/headnews.xml` | Top headlines feed |
| **财联社** | Via RSSHub: `https://rsshub.app/cls/` | May need self-hosted RSSHub |
| **FT中文网** | `https://www.ftchinese.com/rss/feed` | Partial paywall, headlines + intro |
| **日经中文网** | `https://www.cn.nikkei.com/rss.html` | 10 latest headlines |

⚠️ Pitfall: The `finance.ifeng.com/rss/` page is a directory listing of RSS options, NOT the actual XML feed. Use `finance.ifeng.com/rss/headnews.xml` directly.

## Agentic Prompt Template

Each cron job prompt is self-contained (no session context available). Key requirements discovered through testing:

1. **Provide explicit tool alternatives** — `web_search` may 401 if Hermes hasn't been restarted since API key change. Include `tavily_search.sh` via terminal as fallback.
2. **Embed source URLs directly** — Don't say "go search for news", give specific URLs the agent can `web_extract()`.
3. **Include INCLUDE/EXCLUDE lists** — The agent needs clear curation criteria, especially what to skip.
4. **Remind "你不能提问"** — Cron runs autonomously, no user feedback possible.
5. **Quality-over-quantity directive** — "宁可推送1篇高质量也不要5篇水的" prevents the agent from padding output.

Template structure:

```markdown
You are Hermes AI Agent. Task: {task_description}

## Your tools
- `web_extract(urls=[...])` — Read web pages/full articles (works well for Chinese sites)
- `terminal(command, timeout=30)` — For running `tavily_search.sh` or feedparser
- `browser_navigate` + `browser_vision` — For JS-heavy sites

## Curation task
{detailed task}

### Content criteria
- MUST include: {include criteria}
- SKIP: {exclude criteria}

### Execution steps
1. **Collect** — Use these tools on these sources:
   a. {Source A} → `web_extract("{url}")`
   b. {Source B} → `terminal("bash ~/.hermes/scripts/tavily_search.sh '{query}'")`
   c. {Source C (RSS)} → `terminal("python3 -c \"import feedparser; f=feedparser.parse('{rss_url}'); [print(e.title+'|'+e.link) for e in f.entries[:10]]\"")`
2. **Curate** — Select top {N} items. Prioritize {priority criteria}.
3. **Explain** — For technical articles, explain all jargon concepts.
4. **Format** — Professional, directly pushable to WeChat.

### Quality standards
- Professional depth, not surface level
- AI Agent PM perspective preferred
- Always include concept explanations for technical jargon
- {task-specific criteria}

Note: You CANNOT ask questions during this cron run. Make decisions autonomously.
```

## Cron Jobs (Fully Agentic — No Scripts)

All 3 cron jobs run as autonomous AI agents. No Python scripts attached — the agent decides which sources to check and which tools to use.

| Name | Schedule | Approach | Status | Next Run |
|------|----------|----------|--------|----------|
| 每日科技早报 | 0 8 * * * | Agent: web_extract + tavily search → curate news → format morning briefing | ✅ Active | Daily 08:00 |
| 深度技术文章推送 | 30 12 * * * | Agent: tavily search + RSS parse + web_extract → full article reading → concept explanations | ✅ Active | Daily 12:30 |
| AI产品经理/商业分析职位推送 | 0 10 */3 * * | Agent: LinkedIn + job platforms + tavily search → filter → personalized job push | ✅ Active | Every 3 days 10:00 |

### Tool Usage Pattern

Each cron agent uses the following tools autonomously:

1. **`web_search(query)`** — Built-in web search (Tavily API), returns structured JSON results
2. **`web_extract(urls=[...])`** — Read full article/website content
3. **`terminal("python3 -c \"import feedparser; ...\"")`** — Parse RSS feeds
4. **`browser_navigate` + `browser_vision`** — For JS-heavy / dynamic sites (e.g., 阿里云开发者)

### Known Limitations

- **WeChat-only sources** (远传科技评论, 远川研究所, Careerfore, 来我青年, Youngs Blood): These accounts have no public website. Currently skipped. Future solution: `wewe-rss` 3rd-party RSS gateway.
- **Paywalled sites** (FT中文网 etc.): Agent gets headlines + intro only.

## Common Issues & Fixes

### "Timeout context manager should be used inside a task" (WeChat delivery)

**Symptom**: Cron jobs run successfully (`last_status: ok`) but delivery fails with:
```
Weixin send failed: Timeout context manager should be used inside a task
```

**Root cause**: aiohttp 3.13.x's `TimerContext.__enter__()` (in `helpers.py:678`) calls `asyncio.current_task(loop=self._loop)`. The live WeChat adapter's `ClientSession` was created in the gateway's main event loop. When `send_weixin_direct()` is called from cron delivery's ThreadPoolExecutor fallback path (different event loop), it tries to reuse the live adapter's session, but `TimerContext` can't find the current task in the session's original loop → `RuntimeError`.

Note: `adapter.send()` catches the RuntimeError internally (returns `SendResult(success=False, error=...)`), so wrapping the call in `try/except RuntimeError` won't work. The detection must happen BEFORE calling the adapter.

**Fix** (applied in `gateway/platforms/weixin.py`, `send_weixin_direct()`, line ~1983):

```python
live_adapter = _LIVE_ADAPTERS.get(resolved_token)
send_session = getattr(live_adapter, '_send_session', None)
if live_adapter is not None and send_session is not None and not send_session.closed:
    try:
        current_loop = asyncio.get_running_loop()
        if current_loop is not send_session._loop:
            raise RuntimeError("loop_mismatch")
    except (RuntimeError, AttributeError):
        # Loop mismatch — skip live adapter, create fresh session below
        logger.debug("[weixin] live adapter loop mismatch for %s, creating fresh session", ...)
    else:
        # Use the live adapter (same event loop)
        ...
        return {...}
    # Falls through to: async with aiohttp.ClientSession(trust_env=True, connector=...) as session:
```

### iLink `ret=-2` sendmessage error

**Symptom**: After fixing the loop mismatch, delivery returns:
```
iLink sendmessage error: ret=-2 errcode=None errmsg=unknown error
```

**Status**: Unresolved. The HTTP request to `ilinkai.weixin.qq.com` succeeds (HTTP 200) but the iLink application-level response returns `ret=-2`. Probable causes:
- Context token expired / invalid
- WeChat account not actively logged in on phone
- Rate limiting or message content rejection
- Fresh `ClientSession` (not the live polling session) may lack the proper iLink auth context

Try: Ensuring the WeChat account is actively logged in on the phone, or running a test message through the live adapter (non-cron) to confirm the iLink token is still valid.

## Legacy Scripts (no longer used by cron jobs)

- `~/.hermes/scripts/daily_fetch.py` — old headlines-only fetcher
- `~/.hermes/scripts/daily_fetch_news.py` — full-article news scraper (812 lines)
- `~/.hermes/scripts/daily_fetch_tech.py` — full-article tech scraper (931 lines)
- `~/.hermes/scripts/tavily_search.sh` — Tavily API curl wrapper (backup, no longer needed since web_search works after restart)
### web_extract() — Chinese Site Compatibility

| Site | URL | web_extract Result | Notes |
|------|-----|-------------------|-------|
| 财联社 | `cls.cn` | ✅ Excellent — returns structured content with headlines, market data, investment calendar, categories | Best single source for news |
| 科创板日报 | `chinastarmarket.cn` | ✅ Good — returns full article text | Works well |
| 腾讯云开发者 | `cloud.tencent.com/developer/article/{id}` | ✅ Good — returns full article text | Direct article URLs work |
| 日经中文网 | `cn.nikkei.com` | ⚠️ Partial — returns navigation + article previews | Need exact article URL |
| 凤凰网科技 | `tech.ifeng.com` | ⚠️ Partial — returns navigation + some article text | Better to use RSS XML |
| 凤凰RSS目录 | `finance.ifeng.com/rss/` | ❌ Just navigation — this is a directory page, not the XML feed | Must use `finance.ifeng.com/rss/headnews.xml` directly |
| 阿里云开发者 | `developer.aliyun.com` | ❌ JS-heavy — returns minimal content, most articles loaded dynamically | Needs browser tools |
| 腾讯科技 | WeChat mirror sites | ❌ Unreliable — mirror sites frequently change | Use RSS feed instead |

### terminal() — Working Commands

```bash
# Tavily search (replacement for broken web_search)
bash ~/.hermes/scripts/tavily_search.sh 'Chinese search query here'

# RSS feed parsing
python3 -c "import feedparser; f=feedparser.parse('https://feedmaker.kindle4rss.com/feeds/qqtech.weixin.xml'); [print(e.title+'|'+e.link) for e in f.entries[:10]]"

# RSS for 凤凰网科技 headlines
python3 -c "import feedparser; f=feedparser.parse('http://finance.ifeng.com/rss/headnews.xml'); [print(e.title+'|'+e.link) for e in f.entries[:10]]"
```

### Why Script-Based Approach Was Replaced

The old `daily_fetch.py --section tech` approach failed because:
1. Google Developers Blog RSS returned nothing (feed unavailable)
2. 阿里云开发者 is JS-heavy, regex-based scraping couldn't extract articles
3. AIGC开放社区 regex matching extracted garbage links (CSS files, author pages)
4. 36kr pages returned inconsistent HTML structure

The agentic approach solves this by letting the agent dynamically choose the right tool per source and adapt when one approach fails.

## Tavily API

- Free tier: 1,000 credits/month
- ✅ Built-in `web_search` works normally after Hermes restart
- `~/.hermes/scripts/tavily_search.sh` kept as backup
