# Platform SOB Agent — Harness Specification

This file documents the supporting infrastructure for the agent defined in `agent.md`.
It is a reference for developers and maintainers — the agent runtime reads `agent.md` and `references.md` for execution.

---

## 工作留痕 — Audit Trail
*(Implemented in `harness/audit_trail.py`)*

After every run, append a new entry to `audit_log.md` (stored in the same directory as this agent). Each entry documents one week's execution.

**Format per entry:**

```markdown
## Run: 2026-05-14 (Thursday) — Asia/Shanghai

- **Status:** ✅ Success / ❌ Failed / ⚠️ Partial
- **Steps Completed:** preflight, 1, 2, 3, 4, 5, 6, PC2-1
- **Steps Failed:** (none)
- **Halt Reason:** (none)
- **Tabs Created:**
  - `[Reg CNLS copy]`: `26/05/14 ADG ADO`
  - `[Reg CNLS copy]`: `26/05/14 By cluster`
  - `[Archive]`: `SOB-260514`
  - `[Archive]`: `PC2-260514`
- **Validation Results:**
  - Step 4 Zero Check: ✅ Pass
  - Step 5 Value Match: ✅ Pass (exact, whole number)
- **Duration:** 42 seconds
- **Notes:** (any manual observations)
```

---

## 操作许可 — Operation Permission Gate
*(Implemented in `harness/permission_gate.py`)*

**All steps require human confirmation before execution.**

For each step:
1. Agent prepares the action plan (what will be read/written, which cells, what values).
2. Agent sends a WeChat message with the summary and waits for user reply.
3. Only after receiving a clear "yes" / approval reply does the agent execute the write.
4. If no reply within a reasonable window, halt and log as "approval timeout".

**WeChat delivery** is configured in `wechat_config.json`. Three supported methods:

| Method | Use case |
|---|---|
| `terminal` | Local testing — prints to terminal, no WeChat |
| `wxpusher` | Personal WeChat via WxPusher subscription account |
| `wecom_webhook` | 企业微信 (WeCom) group bot webhook |

---

## 沙箱操作 — Sandbox Mode
*(Implemented in `harness/sandbox.py`)*

When running in sandbox mode (activated by flag `--sandbox`):
- All write operations are intercepted and logged as `[DRY_RUN]` — no actual changes to any sheet.
- A full intended-action log is printed and saved to `sandbox_run_log.md`.
- Use this mode for the first run after any changes to `agent.md`.

---

## 不确定要 Check — Uncertainty Escalation

If the agent cannot resolve an ambiguity after 1 retry:
- **Halt** the current step.
- Log the uncertainty with full context in the audit trail.
- Send a WeChat message describing exactly what is unclear and what manual action is needed.
- Wait for user instruction before resuming.

---

## Outcome Evaluation
*(Implemented in `harness/outcome_evaluator.py`)*

A run is **fully successful** if:
- All steps (preflight, 1–6, PC2-1) completed with user confirmation at each step.
- Step 4 zero-check passed.
- Step 5 value-match (whole number exact) passed.
- All new tabs created with correct naming and in correct date order.
- Audit log entry appended to `audit_log.md`.

**Outcome report is sent via WeChat after every run** (success or failure), including:
- Which steps passed / failed.
- Which tabs were created.
- Any validation warnings.
- Path to the audit log entry.

---

## Setup Checklist (One-Time)

- [ ] Configure `wechat_config.json` with your preferred delivery method
- [ ] Fill in the 4 Google Sheet IDs in `references.md` § Sheet Registry
- [ ] Run `python agent_runner.py --sandbox` and review `sandbox_run_log.md`
- [ ] Run first live supervised execution on a Thursday at 13:00 Asia/Shanghai
