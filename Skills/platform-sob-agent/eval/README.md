# Platform SOB — Evaluation Suite

> **Eval is batch, by design.** Both scripts run the verifier as a single post-hoc pass —
> the executor finishes the whole workflow, then the verifier judges. They never gate
> step-by-step; per-step gating is production-only (`../agent_runner.py`). Gating during
> eval would score the *gated* system instead of the executor's raw capability.
> For the overall architecture and the eval/production split, see [`../README.md`](../README.md).

Two scripts, two evaluation targets:

| Script | Evaluates | Question |
|--------|-----------|---------|
| `eval_checks_llm.py` | **Verifier** | Is the verifier's judgment consistent and reliable? |
| `eval_runner.py` | **Executor** | Does the executor correctly complete each workflow step? |

Run verifier eval first — a stable verifier is a prerequisite for meaningful executor eval,
and for promoting any production gate from report-only to hard.

> `eval_runner.py` also hosts the shared substrate (`run_agent`, specs, `TOOLS`) that
> `../agent_runner.py` imports. `run_agent` takes an optional `temperature` (default `None`
> = provider default, used by eval); production passes `temperature=0` for a deterministic
> gate. Eval behavior is unchanged.

---

## Prerequisites

```bash
pip install openai
export DEEPSEEK_API_KEY=your_key_here
```

Confirm Google Sheets OAuth token is valid before running either script:
```bash
python3 -c "import gsheets_util as gsu; print(gsu.get_auth_headers())"
```
If it raises an error, refresh the token before proceeding.

---

## eval_checks_llm.py — Verifier Evaluation

### How it works
```
Phase 1 (once):   Collect live sheet data from Google Sheets via gsheets_util.py
Phase 2 (×10):    Feed same data snapshot to DeepSeek verifier LLM 10 times
                  Run 1 at T=0 (deterministic baseline), runs 2-10 at T=0.7 (variance)
Phase 3 (once):   Report per-step PASS/FAIL rates, drift from T=0 baseline, token costs
```

### Run
```bash
cd ~/.hermes/skills/productivity/platform-sob-agent/eval
python3 eval_checks_llm.py
```

### Output
- Console: per-step consistency table, cost summary
- File: `eval_llm_report_YYMMDD_HHMM.json`

### Evaluation Procedure

#### Phase A — Stability

**Goal:** Confirm the verifier gives the same verdict on the same data every time.

**Steps:**
1. Run `eval_checks_llm.py`
2. In the console output, find the `CONSISTENCY VARIANCE BY STEP` table
3. For each step, check the `Drift from baseline` column:
   - `stable` → verdict never changed across 10 runs ✓
   - `drifted N/9` → verdict flipped N times → needs fixing
4. For any drifted step, open the JSON report and read the `raw_output` of the FAIL runs to find what the LLM flagged
5. Fix the corresponding gate in `../VERIFIER.md`
6. Re-run — repeat until all steps show `stable`

**Stability threshold:** `drifted 0/9` on all steps = Phase A complete.

**Report back:**
- The full console output
- List of any drifted steps and the specific check that caused the drift (from raw_output)
- The `eval_llm_report_*.json` filename

**Current status:** Fixes applied after 2 eval runs (260608). Re-run pending to confirm current VERIFIER.md is stable.

---

#### Phase B — Correctness

**Goal:** Confirm the verifier catches real errors and doesn't flag correct states.

Requires two ground truth fixtures — do this after Phase A is stable:

| Fixture | How to set up | Expected result | If wrong |
|---------|--------------|-----------------|----------|
| **Known-good** | Run the real workflow manually, confirm all steps are correct | ALL PASS | Any FAIL = false positive → fix that gate in VERIFIER.md |
| **Known-bad** | Introduce a deliberate error (e.g. paste wrong value into archive tab, then restore after) | FAIL on that step only | PASS = verifier missed the error → tighten that gate |

After any VERIFIER.md fix from Phase B, re-run Phase A to confirm stability is preserved.

**Current status:** Not yet built — no ground truth fixtures exist.

---

## eval_runner.py — Executor Evaluation

> Run this only after Phase A is complete.

### How it works
```
Per session:  Executor LLM runs all workflow steps (sandbox: writes simulated)
              → Verifier LLM reads real sheets and judges each step independently
              → PASS/FAIL is the executor's quality signal for that session
```

### Run
```bash
python3 eval_runner.py                 # 10 sessions, sandbox mode
python3 eval_runner.py --sessions 5   # custom count
python3 eval_runner.py --live         # real writes (use with caution)
```

### Output
- Console: per-session table (tool calls, tokens, cost, pass/fail by step)
- File: `eval_runner_report_YYMMDD.json`

### Report back
- Full console output
- Sessions where all checks passed vs failed
- Which steps failed most often across sessions
- Average tool call count and cost per session

---

## Key Files

| File | Role |
|------|------|
| `../VERIFIER.md` | Verifier prompt spec — 40+ gate checks across 8 steps |
| `../SKILL.md` | Executor prompt spec — workflow instructions |
| `../references.md` | Sheet Registry — aliases → Google Sheet IDs |
| `../gsheets_util.py` | Google Sheets API utility (OAuth, SSL, helpers) |
