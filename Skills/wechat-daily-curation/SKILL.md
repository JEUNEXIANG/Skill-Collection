---
name: wechat-daily-curation
description: Run and maintain an autonomous WeChat daily curation workflow for tech, AI, macro/finance, and AI PM/BA jobs using search, RSS, extraction, browser automation, and WeChat delivery.
---

# WeChat Daily Content Curation

Use this skill when asked to inspect, run, debug, or improve the Hermes WeChat daily curation pipeline.

The pipeline has three autonomous cron jobs and one draft research-paper push:

| Job | Schedule | Prompt |
| --- | --- | --- |
| 每日科技早报 | Daily 08:00 | `references/news_agent_prompt.txt` |
| 深度技术文章推送 | Daily 12:30 | `references/tech_agent_prompt.txt` |
| AI产品经理/商业分析职位推送 | Every 3 days 10:00 | `references/jobs_agent_prompt.txt` |
| Agent Research Papers 推送 | Draft; suggested weekly | `references/research_papers_agent_prompt.txt` |

## Operating Principles

- Cron jobs run without interactive user feedback. Prompts must tell the agent to decide autonomously and never ask follow-up questions.
- Prefer quality over volume. Send fewer strong items instead of padding with weak content.
- Every pushed item should include source, link when available, date or recency signal, and a short explanation of why it matters to the user.
- For AI/technical content, explain jargon in plain Chinese before making product or career implications.
- Treat source freshness as part of quality. Prefer items published within the last 24 hours for news, within the last 7 days for deep tech, and currently open roles for jobs.
- Deduplicate repeated stories across sources. Cluster similar reports and keep the clearest or most authoritative source.
- Keep WeChat output mobile-readable: concise sections, short paragraphs, no wide tables unless necessary.

## Source Registry

Use `references/source_registry.md` as the source map. It records websites, RSS feeds, access methods, and known weak spots.

Fastest reliable paths:

| Source | Preferred Access |
| --- | --- |
| 腾讯科技 | RSS: `https://feedmaker.kindle4rss.com/feeds/qqtech.weixin.xml` |
| 凤凰新闻/科技 | RSS: `http://finance.ifeng.com/rss/headnews.xml` |
| 财联社 | `web_extract(urls=["https://www.cls.cn"])`; RSSHub only if available |
| FT中文网 | RSS: `https://www.ftchinese.com/rss/feed` |
| 日经中文网 | RSS landing page: `https://www.cn.nikkei.com/rss.html`; exact article URLs work better |
| 阿里云开发者 | Browser tools for JS-heavy pages |

Do not use `finance.ifeng.com/rss/` as a feed URL. It is a directory page; use `http://finance.ifeng.com/rss/headnews.xml`.

## Tool Ladder

For each source, try the cheapest reliable method first:

1. RSS feed when a working feed is listed.
2. `web_extract(urls=[...])` for exact article URLs or static pages.
3. `web_search(query)` with source-specific queries such as `site:domain.com keyword`.
4. `browser_navigate` + `browser_snapshot`/`browser_vision` for JS-heavy pages.
5. Skip the source and mention the reason only in logs or maintenance notes, not in the final push.

If built-in `web_search` is unavailable (Tavily key not configured or process not restarted), use:

```bash
bash ~/.hermes/scripts/tavily_search.sh 'Chinese search query here'
```

As of 2026-05-12, `web_search` is confirmed working after a Hermes gateway restart.

For RSS parsing:

```bash
python3 -c "import feedparser; f=feedparser.parse('RSS_URL'); [print(e.title+'|'+e.link) for e in f.entries[:10]]"
```

## Prompt Maintenance Checklist

When editing a cron prompt:

- Keep the prompt self-contained; cron runs do not have session context.
- Embed concrete source URLs and fallback search queries.
- Include explicit include/exclude criteria.
- Use dynamic date language rather than hardcoding a year unless the task truly requires it.
- Require final output to be directly pushable to WeChat.
- Preserve the "cannot ask questions" instruction.
- Add or update source-specific notes in `references/source_registry.md` instead of burying them in prompts.

## Job Curation — Specific Sources & Findings (May 2026)

When running the job curation cron (AI PM / BA positions), use this multi-source strategy:

### Reliable sources (high yield)
| Source | Method | Notes |
|--------|--------|-------|
| 字节跳动校招 | `web_search("site:jobs.bytedance.com 2026 校招 AI产品经理")` | 2026春招窗口 ~04/23~06/23 |
| 阿里巴巴校招 | `campus-talent.alibaba.com` | AI Agent产品经理在淘天集团有专门岗位 |
| Shopee Careers | `careers.shopee.sg` | Regional AI Strategy PM + AI Agent Intern 两个岗位很对口 |
| 鼠鼠求职 | `web_search("site:shushuqiuzhi.com AI产品经理 2026")` | 经常能找到小公司+大厂的校招AI PM岗位 |
| BOSS直聘 | `web_search("site:zhipin.com AI产品经理 2026 校招")` | 量大但需要筛选 |
| 美团招聘 | `zhaopin.meituan.com` | 2026春招有AI产品经理提前批 |
| 小红书校招 | `job.xiaohongshu.com/campus` | 产品经理培训生项目 |

### Low-yield / dead sources (skip unless desperate)
- `site:jiancareer.com` — consistently returns no results
- `site:casemock.com` — consistently returns no results
- `site:cdc.sem.tsinghua.edu.cn` — campus CDC site rarely has recruitable job listings
- 罗兰贝格官网 `rolandberger.com/zh/Join/` — LinkedIn showed position as closed

### Critical: verify deadline freshness
Many job listings stay indexed long after deadlines. Always cross-reference:
- The actual company careers page (not aggregator)
- LinkedIn posting dates
- Recent update timestamps on aggregator sites (shushuqiuzhi.com shows "updated" dates)

### MBB Consulting deadlines for 2026 cycle
| Firm | Role | Deadline | Status |
|------|------|----------|--------|
| Bain | Associate Consultant (UG/MS) | July 19, 2026 | Upcoming |
| McKinsey | Business Analyst (UG/MS) | August 11, 2026 | Upcoming |
| BCG | Associate (UG/MS) | Expected Summer 2026 | TBD |
| L.E.K. | Analyst (Asia) | Dec 31, 2026 | Open |

Search queries that work well:
- `AI Agent Product Manager job 2026 hiring`
- `AI产品经理 招聘 2026 应届`
- `Business Analyst entry level 2026`
- `腾讯 2026 AI产品经理培训生`
- `阿里巴巴 2026 校招 AI Agent产品经理`

### Quality Rubric

Score candidate content before selection:

| Criterion | What Good Looks Like |
| --- | --- |
| Relevance | Strong connection to AI Agents, AI PM work, macro/finance context, or the user's job targets |
| Freshness | Clearly recent and not recycled |
| Authority | Primary source, reputable media, official job page, or credible technical author |
| Depth | Adds analysis, mechanism, data, or practical insight |
| Actionability | Helps the user understand what to watch, learn, apply to, or prepare for |

## Evaluation Workflow — Scoring Cron Output with eval_criteria.md

Use this when asked to evaluate the quality of a curation push, or to score output against `references/eval_criteria.md`.

### Step-by-step

1. **Trigger the cron job** (if it hasn't run yet):
   ```
   cronjob(action='list')          # find the job_id
   cronjob(action='run', job_id=...)  # run it manually
   ```

2. **Wait for completion** — poll `cronjob(action='list')` until `last_run_at` updates and `last_status` is `ok`.

3. **Locate the session file** — the actual push output is in the cron agent's session JSON:
   ```
   ~/.hermes/sessions/session_cron_<job_id>_<timestamp>.json
   ```

4. **Extract the final response** — the push content is the last `"role": "assistant"` message with `"finish_reason": "stop"` and non-empty `"content"`. This is what was delivered to WeChat.

5. **Run Gate checks** — programmatic (no LLM):
   - Link resolves: `curl -s -o /dev/null -w '%{http_code}' <url>` — accept 2xx, 3xx
   - Freshness window: news ≤24h · deep tech ≤7d · jobs = currently open · papers ≤7d
   - No duplicates: count unique URLs
   - Required fields: each item needs title + source + link + why-it-matters (in the push structure)
   - Delivered: check `last_delivery_error` is null on the cron job
   - Non-empty: at least 1 item
   - Gate = 0 if **any** check fails → R_push = 0

6. **Score each item** on the 5 rubric dimensions (simulate LLM reward model):
   - Relevance (0.30) — how directly on AI Agents / AI-PM
   - Actionability (0.25) — can the user act on it
   - Depth (0.20) — mechanism, data, analysis, connections
   - Authority (0.15) — primary source vs aggregator vs blog
   - Freshness (0.10) — graded bonus above gate (decay curve based on age)

7. **Aggregate**:
   ```
   top_k = sorted(scores, reverse=True)[:k]   # k=3 for news/tech, k=2 for jobs/papers
   quality = mean(top_k) - 0.5 * sum(max(0, 0.45 - s) for s in scores)  # clamp to [0,1]
   r_push = int(gate_pass) * quality
   ```

### Known Tensions / Edge Cases

- **Freshness gate vs deep analysis content**: The 7d window (`eval_criteria.md` §2) blocks long-form analysis articles that have lasting value. The cron agent naturally picks the best content regardless of age. This is a known tension in the eval design — deep tech analysis has longer relevance than the window captures. If this produces too many false-positive gate failures, consider relaxing to 30d for deep tech or adding a "depth override" exception.
- **Always use real output**: Fabricating test items and scoring them is worthless — it evaluates your own curation, not the agent's. Always extract from the cron session file.
- **Link status**: 3xx redirects count as "resolves" (the link works). Only 4xx/5xx are gate failures.

### Implementation Reference

A complete Python implementation of the eval_criteria.md reward function exists in the June 16, 2026 session (`cron_fe501b823c3c_20260616_225320`). The eval_criteria.md §9 calls for `scripts/score_push.py` — that script should:
- Take push items (parsed from cron session final response) as input
- Implement `Gate()` — programmatic checks returning 0/1 with reasons
- Implement `score_items()` — LLM call or simulation returning cᵢ per item
- Implement `aggregate()` — top-k minus penalty → Quality
- Output R_push, per-item breakdown, gate failures

## References

- `references/source_registry.md` — source map and access methods.
- `references/news_agent_prompt.txt` — morning news cron prompt.
- `references/tech_agent_prompt.txt` — deep tech cron prompt.
- `references/jobs_agent_prompt.txt` — job curation cron prompt.
- `references/research_papers_agent_prompt.txt` — research-paper curation prompt for Agent development, RAG, skill/tool use, harness engineering, and system prompts.
- `references/eval_criteria.md` — output scoring system (reward function) for grading push quality via LLM + human pairwise labels.
- `references/source_compatibility.md` — known `web_extract` and browser behavior by site.
- `references/troubleshooting.md` — WeChat delivery, iLink, and legacy script notes.
