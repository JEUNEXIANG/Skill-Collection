# Troubleshooting Notes

## WeChat delivery: `Timeout context manager should be used inside a task`

Symptom:

```text
Weixin send failed: Timeout context manager should be used inside a task
```

Observed behavior: cron jobs can finish with `last_status: ok` while delivery fails.

Root cause: `aiohttp` 3.13.x requires `TimerContext` to run in the same event loop as the `ClientSession`. The live WeChat adapter's session is created in the gateway main event loop. Cron delivery may run through a `ThreadPoolExecutor` fallback path with a different event loop, then reuse the live adapter's session and fail.

Important detail: `adapter.send()` catches the `RuntimeError` internally and returns `SendResult(success=False, error=...)`, so loop mismatch detection must happen before calling the adapter.

Applied fix location and logic:

```text
gateway/platforms/weixin.py — send_weixin_direct()
```

Fix pattern — detect loop mismatch BEFORE calling the live adapter:

```python
try:
    current_loop = asyncio.get_running_loop()
    if current_loop is not send_session._loop:
        raise RuntimeError("loop_mismatch")
except (RuntimeError, AttributeError):
    # Live adapter session belongs to a different loop.
    # Fall through to create a fresh session below.
    pass
else:
    # Use the live adapter (correct loop).
    last_result = await live_adapter.send(chat_id, cleaned)
```

Note: wrapping `live_adapter.send()` in `try/except RuntimeError` does NOT work because `adapter.send()` catches RuntimeError internally and returns `SendResult(success=False, error=...)`, so the error never propagates to the caller. The check must be done BEFORE calling send().

## Manual delivery workaround

When cron delivery fails (regardless of error), content can still be pushed through the current WeChat conversation. The real-time messaging channel (via the gateway's long-poll adapter) works independently of the cron delivery path.

Steps:
1. Run the agent manually with the same prompt as the cron job.
2. The agent's final response is delivered through the current chat session normally.
3. This bypasses the cron scheduler's `_deliver_result()` and `send_weixin_direct()` entirely.

## iLink `ret=-2` sendmessage error

Symptom:

```text
iLink sendmessage error: ret=-2 errcode=None errmsg=unknown error
```

Status: **NOT FIXED by context_token stripping** — commit `d63a8c7e` was incorrect.

### Real root cause: expired bot credentials

Investigation 2026-05-14 found that `ret=-2` is NOT a stale-context issue. Even with no context_token in the request, iLink returns `ret=-2`. The real cause is expired bot credentials (QR login token). A fresh QR re-authentication resolves it.

### Resolution: re-authenticate via QR login

```bash
cd ~/.hermes/hermes-agent && source venv/bin/activate
hermes gateway setup
# Select option 14 (Weixin/WeChat), scan QR code with phone
```

Then update the env vars and fully restart all Hermes processes.

### Env var caching note

The `send_message` tool reads env at process startup time, not the current `.env` file. A full process restart (not just gateway) is required after credential changes.

## iLink `errcode=-14 (session timeout)` after fresh QR login

Symptom: fresh context_token saved from inbound message, but sends still fail with `ret=None errcode=-14`.

### Workaround

`send_weixin_direct()` from a standalone Python process (fresh session) succeeds. A full gateway restart after the token is stored resolves the live adapter issue.

Convenience script: `~/.hermes/scripts/test_weixin_send.py`.

## Gateway restart after code changes
```
# Fix triggered (token stripped, retried — still failed):
12:26:27 WARNING ... send context failed ret=-2 ... retrying without context_token
12:26:28 WARNING ... send chunk failed ... iLink sendmessage error: ret=-2  (attempt 2/3)
12:26:30 ERROR   ... send failed ... iLink sendmessage error: ret=-2

# Gateway restart → new gateway (no token at all — still ret=-2):
00:10:02 WARNING ... send chunk failed ... iLink sendmessage error: ret=-2 (attempt 1/3)
00:10:03 WARNING ... send chunk failed (2/3)
00:10:05 ERROR   ... send failed
```

### Adjacent problem: long-poll DNS failures

Since May 11, the long-poll has been intermittently failing with:
```text
Cannot connect to host ilinkai.weixin.qq.com:443 ssl:default
[nodename nor servname provided, or not known]
```
These come in waves — hours of DNS failures followed by recovery. This may not be directly related to `ret=-2` but it means the bot has missed inbound messages for extended periods and may have missed token-refresh signals.

### Likely root cause: expired bot credentials

Given that:
- Every send returns `ret=-2` regardless of context_token
- The iLink connection was unstable for days
- No successful poll has received inbound messages since May 11
- Token store is empty (no cached context_token)

...the most likely explanation is that the **bot token (QR-code login credential) has expired or been revoked by iLink**. A fresh QR-code re-authentication via `hermes gateway setup` is needed.

### Resolution path

1. Run `hermes gateway setup` and select Weixin/WeChat (option 14).
2. Scan the QR code with the WeChat mobile app.
3. After login, test with `send_message(target='weixin', message='test')`.
4. If successful, the new token will be persisted in the account config and the cron jobs should recover automatically.

## Gateway restart after code changes

When `gateway/platforms/weixin.py` or any gateway code changes at runtime, Python modules are already cached in memory. The gateway must be restarted to load the new code.

**⚠️ Critical timing check:** Compare the gateway process start time against the git commit time. If the gateway started BEFORE the commit, it's running old code:
```bash
# Check gateway start time
ps -o lstart= -p $(pgrep -f "gateway run" | head -1)

# Check commit time
git log --format="%ci %s" -1 <commit_hash>

# If gateway started before commit → restart needed
```

There is a launchd job (`ai.hermes.gateway`) that auto-restarts the gateway. To stop it:...

```bash
launchctl bootout gui/$(id -u)/ai.hermes.gateway
```

Then kill all existing gateway processes:

```bash
pkill -9 -f "gateway run"
```

Then start fresh:

```bash
cd ~/.hermes/hermes-agent
source venv/bin/activate
python -m hermes_cli.main gateway run --replace
```

(Use terminal() with background=true for the gateway process so Hermes can track its lifecycle.)

After starting, verify WeChat connection in the logs:

```bash
grep "weixin\|Weixin" ~/.hermes/logs/agent.log | tail -5
# Expected: "✓ weixin connected"
```

## Debugging checklist for `ret=-2` failures

When sends fail with `ret=-2`, go through this checklist before changing code:

1. **Check token store** — empty store means no cached context_token; ret=-2 is NOT a session expiry:
   ```bash
   cat ~/.hermes/weixin/accounts/*.context-tokens.json
   # Empty {} means no token exists at all
   ```

2. **Test DNS / connectivity** — iLink may be unreachable:
   ```bash
   ping -c 2 -t 3 ilinkai.weixin.qq.com
   ```

3. **Check long-poll health** — the poll must be running to receive inbound messages (which refresh context_token):
   ```bash
   grep "poll error\|getUpdates" ~/.hermes/logs/agent.log | tail -10
   # Expected: no recent errors, or only poll-timeout (harmless) lines
   ```

4. **Check gateway vs commit timing** — code changes don't take effect until restart:
   ```bash
   ps -o lstart= -p $(pgrep -f "gateway run" | head -1)
   git log --format="%ci %s" -1 -- gateway/platforms/weixin.py
   ```

5. **Test with fresh credentials** — if all else fails, re-authenticate:
   ```bash
   hermes gateway setup
   # Select option 14 (Weixin/WeChat)
   # Scan QR code with WeChat mobile app
   ```

## Legacy scripts

The cron jobs are currently agentic and do not depend on these old scripts:

- `~/.hermes/scripts/daily_fetch.py`
- `~/.hermes/scripts/daily_fetch_news.py`
- `~/.hermes/scripts/daily_fetch_tech.py`
- `~/.hermes/scripts/tavily_search.sh`

`tavily_search.sh` is still useful as a fallback if built-in `web_search` is unavailable.

## Known issues

### 深度技术文章推送 regularly times out

The deep tech article curation job (`fe501b823c3c`) has a history of idling out (>600s) or stalling. It's the heaviest of the three cron jobs — it searches for deep tech articles, extracts full content, generates concept explanations, and performs deep analysis. The deepseek-reasoner model also takes long reasoning chains.

Common failure pattern:
1. web_search completes (fast)
2. web_extract on article pages stalls (some sites timeout or return large content)
3. The 600s inactivity limit is hit

If this job fails after delivery is fixed, consider:
- Reducing the number of articles curated per run
- Using `web_extract` with fewer URLs per batch
- Increasing the cron job's `max_iterations` or the inactivity timeout in scheduler config
- Running the job manually via `cronjob(action='run', job_id='fe501b823c3c')` for a fresh attempt

## Gateway stuck process: `hermes gateway restart` hangs

Symptom: Running `hermes gateway restart --all` returns "✓ Service stopped" and "✓ Service started" but no new gateway process appears. Old gateway PID still shows in `ps`.

Root cause: The old gateway process received SIGTERM/SIGINT for graceful shutdown, but the Python process didn't fully exit (remains in S state). The `hermes gateway restart` wrapper spawned a bash loop `while kill -0 <old_pid>; do sleep 0.2; done` that blocks indefinitely waiting for the old PID to die.

Detection:
```bash
# Find the old gateway process
ps aux | grep "gateway run" | grep -v grep
# PID 12095 ... python -m hermes_cli.main gateway run --replace

# Check if a restart script is stuck waiting
ps aux | grep "kill -0" | grep -v grep
# bash -lc while kill -0 12095 2>/dev/null; do sleep 0.2; done;

# Check gateway recent log for shutdown signal
grep "SIGTERM\|SIGINT\|shutdown\|Stopping gateway" ~/.hermes/logs/agent.log | tail -5
```

Resolution — force-kill the stuck process to unblock the restart:
```bash
kill -9 <old_gateway_pid>
# The restart script detects the process is gone and proceeds to start a new gateway
```

After killing, wait 5-10s and verify:
```bash
ps aux | grep "gateway run" | grep -v grep
# Should show a new PID with recent start time
```

## Force-resetting a WeChat session via sessions.json

Use case: The user asks to restart their WeChat session, but you're in CLI (not on WeChat) and can't send `/new` or `/reset` commands through the gateway.

This can happen when the gateway is just starting, or when the user is on a different platform than the gateway session you need to reset.

Steps:

1. **Read `~/.hermes/sessions/sessions.json`** to find the WeChat session entry. The session key pattern is `agent:main:weixin:dm:<user_id>@im.wechat`.

2. **Edit sessions.json** to:
   - Change `session_id` to a new unique value (format: `YYYYMMDD_reset_XXXXXXXX` is fine)
   - Set `created_at` and `updated_at` to current time
   - Set `suspended: true` — this tells the gateway to auto-reset the session on next message
   - Zero out `input_tokens`, `output_tokens`, `last_prompt_tokens`, `total_tokens`
   - Set `memory_flushed: false`

3. **Restart the gateway** — the in-memory `_entries` dict in the running gateway won't see the file change. A gateway restart is required:
   ```bash
   cd ~/.hermes/hermes-agent && source venv/bin/activate
   hermes gateway restart --all
   # Or if restart hangs, force-kill and start manually:
   # pkill -9 -f "gateway run"; python -m hermes_cli.main gateway run --replace
   ```

4. **Verify** the new gateway loaded the updated sessions:
   ```bash
   grep "weixin connected" ~/.hermes/logs/agent.log | tail -1
   # Expected: ✓ weixin connected
   ```

5. **User sends a message on WeChat** — gateway finds `suspended: true`, auto-resets to a properly formatted session ID, and starts a fresh conversation.

Caveat: If the session key already exists in the gateway's in-memory cache (which it will if the gateway has been running with that session), editing the file alone does nothing until restart.

## Why the script-based approach was replaced

The old `daily_fetch.py --section tech` approach was brittle:

1. Google Developers Blog RSS returned no usable feed.
2. 阿里云开发者 is JS-heavy, so regex scraping missed articles.
3. AIGC开放社区 regex matching captured unrelated links such as CSS files and author pages.
4. 36kr pages returned inconsistent HTML.

The agentic approach lets the model choose RSS, extraction, search, or browser tools per source and recover when one path fails.
