# Source Compatibility Notes

## `web_extract()` behavior

| Site | URL | Result | Notes |
| --- | --- | --- | --- |
| 财联社 | `cls.cn` | Excellent | Returns structured content with headlines, market data, investment calendar, and categories. Strong single source for daily news. |
| 科创板日报 | `chinastarmarket.cn` | Good | Full article text usually works. |
| 腾讯云开发者 | `cloud.tencent.com/developer/article/{id}` | Good | Direct article URLs work better than index pages. |
| 日经中文网 | `cn.nikkei.com` | Partial | Index pages return navigation and previews. Prefer exact article URLs. |
| 凤凰网科技 | `tech.ifeng.com` | Partial | Better to use RSS XML first. |
| 凤凰 RSS directory | `finance.ifeng.com/rss/` | Bad | Directory page, not an XML feed. Use `http://finance.ifeng.com/rss/headnews.xml`. |
| 阿里云开发者 | `developer.aliyun.com` | Weak | JS-heavy; use browser tools when extraction is thin. |
| 腾讯科技 | WeChat mirror sites | Unreliable | Mirror sites change often. Use the Kindle4RSS feed first. |

## Working terminal commands (backup only — web_search works after restart)

```bash
# Tavily fallback search (only if web_search returns 401)
bash ~/.hermes/scripts/tavily_search.sh 'Chinese search query here'

# 腾讯科技 RSS
python3 -c "import feedparser; f=feedparser.parse('https://feedmaker.kindle4rss.com/feeds/qqtech.weixin.xml'); [print(e.title+'|'+e.link) for e in f.entries[:10]]"

# 凤凰网科技/新闻 RSS
python3 -c "import feedparser; f=feedparser.parse('http://finance.ifeng.com/rss/headnews.xml'); [print(e.title+'|'+e.link) for e in f.entries[:10]]"
```
