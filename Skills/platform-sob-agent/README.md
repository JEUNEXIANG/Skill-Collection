# Platform SOB Agent

A two-agent system that runs and checks the weekly Platform SOB spreadsheet workflow.
The **executor** performs each step; the **verifier** audits it. Both are the same LLM
wearing a different system prompt.

There are **two ways to run the system, and keeping them separate is the core design rule:**

| Mode | Entry point | What it does | Verifier timing |
|------|-------------|--------------|-----------------|
| **Eval** | `eval/eval_runner.py` | Measures the executor's *unsupervised* end-to-end quality | Runs **once, after** the executor finishes the whole workflow (batch) |
| **Production** | `agent_runner.py` | Runs the workflow for real, gated step-by-step | Runs **per step, at T=0**; executor only advances on PASS |

Gating during eval would contaminate the measurement (you'd score the *gated* system,
not the executor), so the two never share an orchestrator.

---

## The architecture in four layers

Everything below the manager layer is the **substrate** — shared, identical in both modes.
Only the top layer changes.

```
   ┌──────────────────────────────────────────────────────────┐
   │  MANAGER   who calls run_agent, in what order             │  ← ONLY this differs
   │    eval: run_session   |   production: agent_runner.run    │     between modes
   ├──────────────────────────────────────────────────────────┤
   │  RHYTHM    run_agent()            ReAct loop               │ ┐
   │  HANDS     TOOLS + execute_tool   sheet read/write+sandbox │ │  SUBSTRATE
   │  ROLE      EXECUTOR_SPEC / VERIFIER_SPEC   (system prompt) │ │  (shared)
   │  BRAIN     the LLM (deepseek-chat)                         │ ┘
   └──────────────────────────────────────────────────────────┘
```

- **Brain** — a generic LLM. Knows nothing about SOB until given a role.
- **Role** — `SKILL.md` makes it the executor; `VERIFIER.md` makes it the verifier.
  Swapping the system prompt is the *entire* difference between the two agents.
- **Hands** — `TOOLS` + `execute_tool()` let it read/write Google Sheets. The `sandbox`
  flag disconnects the write hand (logs instead of writing) while reads stay live.
- **Rhythm** — `run_agent()` is the ReAct loop: **think → act (call a tool) → observe the
  result → think → … → answer.** It repeats as long as the model asks for tools, and stops
  the moment the model returns plain text instead of a tool request (cap: `max_turns=40`).

All four substrate parts live in `eval/eval_runner.py` and are imported by `agent_runner.py`.

---

## How a step is identified

An agent has no hidden "current step" counter. The step is **named in the user message**
for that one `run_agent` call, and `VERIFIER.md` defines what that step checks.

- **Production:** each step gets its *own* `run_agent` call with a fresh prompt naming
  only that step (`steps_to_verify: ["step4"]`). The call has no memory of other steps —
  so cross-step contamination is structurally impossible.
- **Eval:** one verifier call receives the full list and reports every step in one pass.

---

## Production runner — `agent_runner.py`

Per-step gate loop. For each step: (optional human approval on live writes) → executor does
that step → verifier audits that step at **T=0** → branch on the verdict.

```bash
python agent_runner.py              # SANDBOX (no writes) — safe to watch gate decisions
python agent_runner.py --live       # real writes + human approval gates
python agent_runner.py --gate-all   # treat every step as a hard gate (halt on any FAIL)
```

**Verdict branching:** `PASS` → next step · `FAIL` → halt (hard) or log (report) · `SKIP` →
skip (e.g. Step 5 archive is skipped unless Step 4 zero-check passed).

**Gate policy** — only steps that were *stable* in the latest eval (T=0 and T=0.7 agree) halt
on failure today. Unstable specs run in report-only mode so a flaky verdict can't wrongly
stop a real run:

| Step | Gate | Why |
|------|------|-----|
| Pre-flight, Step 2, Step 6 | **hard** | stable across temperatures |
| Step 3 | report | Type C divergence — add explicit Pass Condition first |
| Step 4 | report | SKIP-vs-FAIL boundary not yet exact |
| Step 5 | report | depends on Step 4 |
| PC2 Step 1 | report | check pending redesign |

To promote a step to `hard`: stabilize its spec in `VERIFIER.md`, confirm via the verifier
eval, then change its `gate` field in `STEPS` to `"hard"`. Stabilize the *spec* — don't lean
on the gate; a deterministic call to an ambiguous spec is consistently wrong, not safe.

**Harness wiring** — at the end of a run, `agent_runner` calls into `harness/`:
`RunRecord` (audit_trail) tracks run state, `permission_gate` handles human approval and
WeChat alerts, `outcome_evaluator.finalize_run` writes the `audit_log.md` entry and sends
the outcome report.

---

## Evaluation — see [`eval/README.md`](eval/README.md)

Two eval scripts, both **batch** (never gate per step):

- `eval/eval_checks_llm.py` — verifier consistency (T=0 baseline vs T=0.7 variance)
- `eval/eval_runner.py` (`run_session`) — executor quality (executor runs all steps → verifier judges once)

Run verifier eval first; a stable verifier is a prerequisite for trusting executor eval —
and for promoting any production gate to `hard`.

---

## File map

| File | Role |
|------|------|
| `SKILL.md` | Executor prompt spec — workflow instructions |
| `VERIFIER.md` | Verifier prompt spec — gate checks per step |
| `agent_runner.py` | **Production** orchestrator — per-step gate loop |
| `eval/eval_runner.py` | Shared substrate (`run_agent`, specs, tools) + **executor eval** (`run_session`) |
| `eval/eval_checks_llm.py` | **Verifier eval** — consistency across temperatures |
| `harness/audit_trail.py` | `RunRecord` — per-run state + `audit_log.md` writer |
| `harness/permission_gate.py` | Human approval + WeChat alerts |
| `harness/outcome_evaluator.py` | Success criteria + `finalize_run` |
| `harness/sandbox.py` | Dry-run guard (legacy direct-call path; LLM writes are sandboxed in `execute_tool`) |
| `references.md` | Sheet Registry — aliases → Google Sheet IDs |
| `gsheets_util.py` | Google Sheets API utility (OAuth, SSL, helpers) |
