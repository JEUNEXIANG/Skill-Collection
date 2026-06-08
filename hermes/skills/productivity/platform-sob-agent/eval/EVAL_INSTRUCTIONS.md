# Platform SOB Verifier — Evaluation Instructions (LLM-based)

You are running an evaluation of the Platform SOB verifier agent using the DeepSeek LLM API.
Your job: collect live sheet data once, run the LLM verifier **10 times** on the same data,
then report token costs and consistency variance across runs.

---

## Context

This is a weekly workflow that copies, validates, and archives Platform Share-of-Business (SOB)
data across multiple Google Sheets every Thursday. The verifier audits whether each step was
executed correctly. This evaluation measures:

1. **LLM consistency variance** — does the verifier give the same verdict every time on the
   same data, or does it flip between PASS/FAIL across runs?
2. **Token cost** — how much does each verifier run cost using DeepSeek pricing?

---

## Prerequisites

```bash
pip install openai            # DeepSeek uses OpenAI-compatible API
export DEEPSEEK_API_KEY=your_key_here
```

DeepSeek API key: https://platform.deepseek.com → API Keys

---

## Files to Load Before Starting

| File | Purpose |
|------|---------|
| `SKILL.md` | Executor workflow spec — what each step does |
| `VERIFIER.md` | Verifier spec — the checks the LLM must run |
| `references.md` | Sheet Registry — aliases → Google Sheet IDs and tab names |
| `gsheets_util.py` | Google Sheets API utility — used for all data collection |

---

## How It Works

```
Phase 1 (Python, once):   Read all relevant sheet data via gsheets_util.py
                           → Serialize into structured text context

Phase 2 (DeepSeek, ×10):  Call LLM with VERIFIER.md as system prompt
                           + sheet data as user context
                           → Parse structured PASS/FAIL output per step
                           → Record token usage

Phase 3 (Python, once):   Aggregate results
                           → Variance per check (how often did verdict flip?)
                           → Cost per run and total
                           → Consistency score
```

---

## Script — write to disk as `eval_checks_llm.py`

```python
"""
Platform SOB Verifier — LLM-based evaluation runner.
Phase 1: collect sheet data via gsheets_util.py (once).
Phase 2: call DeepSeek verifier LLM 10 times on the same data.
Phase 3: report consistency variance and token costs.

Requirements:
    pip install openai
    export DEEPSEEK_API_KEY=your_key_here
"""

import os
import sys
import json
import re
from pathlib import Path
from datetime import datetime
from openai import OpenAI

sys.path.insert(0, str(Path(__file__).parent))
import gsheets_util as gsu

# ── Config ─────────────────────────────────────────────────────────────────────
NUM_RUNS    = 10
MODEL       = "deepseek-chat"          # DeepSeek-V3; use "deepseek-reasoner" for R1
TEMPERATURE = 0.7                      # >0 to expose variance across runs
MAX_TOKENS  = 8192

# DeepSeek-V3 pricing per million tokens
PRICE_INPUT_CACHE_MISS = 0.27
PRICE_INPUT_CACHE_HIT  = 0.07
PRICE_OUTPUT           = 1.10

SCRIPT_DIR    = Path(__file__).parent
VERIFIER_SPEC = (SCRIPT_DIR / "VERIFIER.md").read_text()

# ── Sheet IDs ──────────────────────────────────────────────────────────────────
SID_WLV  = gsu.SHEETS["Weekly Live View"]
SID_RCT  = gsu.SHEETS["Reg Commercial Team"]
SID_CNLS = gsu.SHEETS["Reg CNLS copy"]
SID_ARCH = gsu.SHEETS["Archive"]
SID_PC2  = gsu.SHEETS["Platform PC2"]


# ── Helpers ────────────────────────────────────────────────────────────────────
import urllib.request, urllib.parse

def get_tabs(sid):
    data = gsu.api_get(sid, "sheets.properties")
    return [s["properties"]["title"] for s in data.get("sheets", [])]

def read_range(sid, tab, cell_range, max_rows=40):
    raw = gsu.values_get(sid, f"'{tab}'!{cell_range}")
    rows = raw.get("values", [])[:max_rows]
    return rows

def read_range_unformatted(sid, tab, cell_range, max_rows=40):
    """Read with UNFORMATTED_VALUE so number-format suffixes (e.g. '6.26x') are stripped."""
    headers = gsu.get_auth_headers()
    range_enc = urllib.parse.quote(f"'{tab}'!{cell_range}", safe='!')
    url = (f"https://sheets.googleapis.com/v4/spreadsheets/{sid}/values/{range_enc}"
           f"?valueRenderOption=UNFORMATTED_VALUE")
    req = urllib.request.Request(url, headers=headers)
    with urllib.request.urlopen(req, timeout=30, context=gsu.SSL_CTX) as resp:
        raw = json.loads(resp.read())
    return raw.get("values", [])[:max_rows]

def rows_to_text(rows, label):
    lines = [f"  {label}:"]
    for i, row in enumerate(rows):
        lines.append(f"    row {i}: {row}")
    return "\n".join(lines)

def find_latest_tab(tabs, keyword):
    matches = [t for t in tabs if keyword.lower() in t.lower()
               and re.search(r'\d{2}[/\-]\d{2}', t)]
    return matches[0] if matches else None


# ── Phase 1: Collect sheet data ────────────────────────────────────────────────
def collect_data():
    print("Phase 1: Collecting sheet data from Google Sheets...")
    d = {}

    # Tab lists
    d["wlv_tabs"]  = get_tabs(SID_WLV)
    d["rct_tabs"]  = get_tabs(SID_RCT)
    d["cnls_tabs"] = get_tabs(SID_CNLS)
    d["arch_tabs"] = get_tabs(SID_ARCH)
    d["pc2_tabs"]  = get_tabs(SID_PC2)

    # Latest tabs
    d["latest_adg_tab"]     = find_latest_tab(d["cnls_tabs"], "ADG ADO")
    d["latest_cluster_tab"] = find_latest_tab(d["cnls_tabs"], "cluster")
    d["latest_sob_tab"]     = next((t for t in d["arch_tabs"] if t.startswith("SOB-")), None)
    d["latest_pc2_tab"]     = next((t for t in d["arch_tabs"] if t.startswith("PC2-")), None)
    d["adg_src_tab"]        = next((t for t in d["wlv_tabs"]  if "ADG ADO" in t), None)
    d["cf_tab"]             = next((t for t in d["rct_tabs"]  if "Final" in t and "CF" in t), None)
    d["cluster_src_tab"]    = next((t for t in d["wlv_tabs"]  if "Cluster" in t), None)
    d["pc2_src_tab"]        = d["pc2_tabs"][0] if d["pc2_tabs"] else None

    print(f"  Latest ADG ADO tab  : {d['latest_adg_tab']}")
    print(f"  Latest cluster tab  : {d['latest_cluster_tab']}")
    print(f"  Latest SOB archive  : {d['latest_sob_tab']}")
    print(f"  Latest PC2 archive  : {d['latest_pc2_tab']}")

    # Step 1: source structure
    if d["adg_src_tab"]:
        d["step1_src_headers"] = read_range(SID_WLV, d["adg_src_tab"], "A1:Z5", max_rows=5)

    # Step 2: source data (for value match) + destination sample
    if d["adg_src_tab"]:
        d["step2_src_data"]    = read_range(SID_WLV,  d["adg_src_tab"],    "A1:Z80", max_rows=80)
    if d["latest_adg_tab"]:
        d["step2_dest_sample"] = read_range(SID_CNLS, d["latest_adg_tab"], "A1:Z80", max_rows=80)

    # Step 3: calculation section (paste source), results, cross-check, CF reference
    if d["latest_adg_tab"]:
        d["step3_calc"]       = read_range(SID_CNLS, d["latest_adg_tab"], "A98:Z130")
        d["step3_results"]    = read_range(SID_CNLS, d["latest_adg_tab"], "A137:Z169")
        d["step3_crosscheck"] = read_range(SID_CNLS, d["latest_adg_tab"], "A172:Z196")
    if d["cf_tab"]:
        d["step3_cf_src"]     = read_range(SID_RCT, d["cf_tab"], "A1:Z30")

    # Step 4: difference table
    if d["latest_adg_tab"]:
        d["step4_diff"] = read_range(SID_CNLS, d["latest_adg_tab"], "A199:Z223")

    # Step 5: archive (formatted + unformatted for x-suffix check) vs calculation section source
    if d["latest_sob_tab"]:
        d["step5_archive"]     = read_range(SID_ARCH,             d["latest_sob_tab"], "A1:Z40")
        d["step5_archive_raw"] = read_range_unformatted(SID_ARCH, d["latest_sob_tab"], "A1:Z40")
    if d["latest_adg_tab"]:
        d["step5_src"]         = read_range(SID_CNLS, d["latest_adg_tab"], "A98:Z130")  # calculation section

    # Step 6: cluster source vs destination
    if d["cluster_src_tab"]:
        d["step6_src"]  = read_range(SID_WLV,  d["cluster_src_tab"],    "A1:S10", max_rows=10)
    if d["latest_cluster_tab"]:
        d["step6_dest"] = read_range(SID_CNLS, d["latest_cluster_tab"], "A1:S10", max_rows=10)

    # PC2 Step 1: source and archive structure (formula error check)
    if d["pc2_src_tab"]:
        d["pc2_src"]  = read_range(SID_PC2,  d["pc2_src_tab"],    "A1:Z10", max_rows=10)
    if d["latest_pc2_tab"]:
        d["pc2_arch"] = read_range(SID_ARCH, d["latest_pc2_tab"], "A1:Z10", max_rows=10)

    print("  Data collection complete.\n")
    return d


# ── Format data as LLM context ─────────────────────────────────────────────────
def format_context(d):
    lines = ["## Live Sheet Data (collected at runtime)\n"]

    lines.append("### Tab Inventory")
    lines.append(f"[Weekly Live View] tabs     : {d.get('wlv_tabs', [])}")
    lines.append(f"[Reg Commercial Team] tabs  : {d.get('rct_tabs', [])}")
    lines.append(f"[Reg CNLS copy] tabs (latest 5): {d.get('cnls_tabs', [])[:5]}")
    lines.append(f"[Archive] tabs (latest 5)   : {d.get('arch_tabs', [])[:5]}")
    lines.append(f"[Platform PC2] tabs         : {d.get('pc2_tabs', [])}")
    lines.append(f"\nLatest ADG ADO tab   : {d.get('latest_adg_tab')}")
    lines.append(f"Latest cluster tab   : {d.get('latest_cluster_tab')}")
    lines.append(f"Latest SOB archive   : {d.get('latest_sob_tab')}")
    lines.append(f"Latest PC2 archive   : {d.get('latest_pc2_tab')}")

    if "step1_src_headers" in d:
        lines.append(rows_to_text(d["step1_src_headers"], "Step 1 — [Weekly Live View] source header rows (A1:Z5)"))
    if "step2_src_data" in d:
        lines.append(rows_to_text(d["step2_src_data"][:20], "Step 2 — [Weekly Live View] ADG ADO source data rows 1:20"))
    if "step2_dest_sample" in d:
        lines.append(rows_to_text(d["step2_dest_sample"][:20], "Step 2 — [Reg CNLS copy] destination tab rows 1:20"))
    if "step3_calc" in d:
        lines.append(rows_to_text(d["step3_calc"][:10], "Step 3 — calculation section rows 98:108 (paste source, first 10)"))
    if "step3_results" in d:
        lines.append(rows_to_text(d["step3_results"][:5], "Step 3 — results section rows 137:141 (first 5)"))
    if "step3_crosscheck" in d:
        lines.append(rows_to_text(d["step3_crosscheck"][:5], "Step 3 — cross-check section rows 172:176 (first 5)"))
    if "step3_cf_src" in d:
        lines.append(rows_to_text(d["step3_cf_src"][:5], "Step 3 — [Reg Commercial Team] CF source rows 1:5"))
    if "step4_diff" in d:
        lines.append(rows_to_text(d["step4_diff"][:10], "Step 4 — difference table rows 199:208 (first 10)"))
    if "step5_archive" in d:
        lines.append(rows_to_text(d["step5_archive"][:5], "Step 5 — SOB archive tab (formatted) rows 1:5"))
    if "step5_archive_raw" in d:
        lines.append(rows_to_text(d["step5_archive_raw"][:5], "Step 5 — SOB archive tab (UNFORMATTED_VALUE) rows 1:5"))
    if "step5_src" in d:
        lines.append(rows_to_text(d["step5_src"][:5], "Step 5 — [Reg CNLS copy] calculation section rows 98:103 (archive source, first 5)"))
    if "step6_src" in d:
        lines.append(rows_to_text(d["step6_src"], "Step 6 — [Weekly Live View] clusters source rows 1:10"))
    if "step6_dest" in d:
        lines.append(rows_to_text(d["step6_dest"], "Step 6 — [Reg CNLS copy] cluster destination rows 1:10"))
    if "pc2_src" in d:
        lines.append(rows_to_text(d["pc2_src"], "PC2 Step 1 — [Platform PC2] source rows 1:10"))
    if "pc2_arch" in d:
        lines.append(rows_to_text(d["pc2_arch"], "PC2 Step 1 — [Archive] PC2 archive rows 1:10"))

    return "\n".join(lines)


def build_user_task(d):
    return f"""
Audit the workflow using the live sheet data above.
The most recent tabs to verify are:
  today_tab_adg        : {d.get('latest_adg_tab', 'not found')}
  today_tab_cluster    : {d.get('latest_cluster_tab', 'not found')}
  today_tab_sob_archive: {d.get('latest_sob_tab', 'not found')}
  today_tab_pc2_archive: {d.get('latest_pc2_tab', 'not found')}
steps_to_verify: ["preflight", "step1", "step2", "step3", "step4", "step5", "step6", "pc2_step1"]

For each step output EXACTLY this format:

STEP: Pre-flight
STATUS: PASS
CHECKS:
  [✓] All sheet IDs resolve
  ...

STEP: Step 1 — Source Data Structure
STATUS: FAIL
CHECKS:
  [✓] ADG section keywords present
  [✗] Month sequence is continuous: gap detected between Feb'26 and Apr'26
  ...

Continue for all steps. End with: OVERALL: PASS or OVERALL: FAIL
"""


# ── Phase 2: Run LLM verifier ×N ───────────────────────────────────────────────
def run_llm_verifier(client, context_text, user_task, run_num, temperature):
    label = "baseline(T=0)" if temperature == 0 else f"T={temperature}"
    print(f"  Run {run_num:>2}/{NUM_RUNS} [{label}]...", end=" ", flush=True)

    response = client.chat.completions.create(
        model=MODEL,
        temperature=temperature,
        max_tokens=MAX_TOKENS,
        messages=[
            {"role": "system", "content": VERIFIER_SPEC},
            {"role": "user",   "content": context_text + "\n\n" + user_task}
        ]
    )

    text = response.choices[0].message.content
    usage = response.usage

    # Token costs
    cache_hit  = getattr(usage, "prompt_cache_hit_tokens", 0)
    cache_miss = getattr(usage, "prompt_cache_miss_tokens", usage.prompt_tokens)
    out_tokens = usage.completion_tokens
    cost = (cache_hit  / 1e6 * PRICE_INPUT_CACHE_HIT
          + cache_miss / 1e6 * PRICE_INPUT_CACHE_MISS
          + out_tokens / 1e6 * PRICE_OUTPUT)

    # Parse PASS/FAIL per step
    step_results = {}
    overall = None
    current_step = None
    for line in text.split("\n"):
        stripped = line.strip()
        if stripped.startswith("STEP:"):
            current_step = stripped[5:].strip()
        elif stripped.startswith("STATUS:") and current_step:
            step_results[current_step] = stripped[7:].strip()
            current_step = None
        elif stripped.startswith("OVERALL:"):
            overall = stripped[8:].strip()

    fail_count = sum(1 for v in step_results.values() if v == "FAIL")
    print(f"{fail_count} fail(s) | {usage.prompt_tokens + out_tokens:,} tokens | ${cost:.5f}")

    return {
        "run": run_num,
        "temperature": temperature,
        "is_baseline": temperature == 0,
        "step_results": step_results,
        "overall": overall,
        "fail_count": fail_count,
        "all_passed": fail_count == 0,
        "tokens": {
            "prompt": usage.prompt_tokens,
            "completion": out_tokens,
            "cache_hit": cache_hit,
            "cache_miss": cache_miss,
            "total": usage.prompt_tokens + out_tokens,
        },
        "cost_usd": cost,
        "raw_output": text,
    }


# ── Phase 3: Report ────────────────────────────────────────────────────────────
def generate_report(results):
    n = len(results)
    full_passes   = sum(1 for r in results if r["all_passed"])
    total_cost    = sum(r["cost_usd"] for r in results)
    avg_tokens    = sum(r["tokens"]["total"] for r in results) / n
    avg_cost      = total_cost / n

    print(f"\n{'='*70}")
    print(f"  EVALUATION REPORT — Platform SOB Verifier  ({n} runs, model={MODEL})")
    print(f"{'='*70}")

    # Per-run table
    print(f"\n{'Run':>4}  {'Fails':>5}  {'Prompt':>8}  {'Completion':>10}  {'CacheHit':>9}  {'Cost':>8}  {'Pass?'}")
    print("─" * 70)
    for r in results:
        t = r["tokens"]
        print(f"  {r['run']:>2}  {r['fail_count']:>5}  {t['prompt']:>8,}  "
              f"{t['completion']:>10,}  {t['cache_hit']:>9,}  "
              f"${r['cost_usd']:>6.5f}  {'  ✓' if r['all_passed'] else '  ✗'}")
    print("─" * 70)
    print(f"{'AVG':>4}  {sum(r['fail_count'] for r in results)/n:>5.1f}  "
          f"{avg_tokens:>8,.0f}  {'':>10}  {'':>9}  ${avg_cost:>6.5f}")
    print(f"\n  Runs with ALL checks passed : {full_passes}/{n}")
    print(f"  Total cost                  : ${total_cost:.5f}")

    # Consistency variance vs baseline (run 1, T=0)
    baseline = next((r for r in results if r["is_baseline"]), None)
    variance_runs = [r for r in results if not r["is_baseline"]]
    nv = len(variance_runs)

    all_steps = sorted({step for r in results for step in r["step_results"]})
    print(f"\n  Consistency variance by step (baseline=run1 T=0, variance runs={nv} T={TEMPERATURE}):")
    print(f"  {'Step':<42} {'Base':>5}  {'PASS':>5}  {'FAIL':>5}  {'Drift from baseline'}")
    print("  " + "─" * 68)
    for step in all_steps:
        base_verdict = baseline["step_results"].get(step, "—") if baseline else "—"
        verdicts = [r["step_results"].get(step) for r in variance_runs if step in r["step_results"]]
        passes = verdicts.count("PASS")
        fails  = verdicts.count("FAIL")
        drifted = sum(1 for v in verdicts if v and v != base_verdict)
        drift_str = f"drifted {drifted}/{nv}" if drifted > 0 else "stable"
        print(f"  {step:<42} {base_verdict:>5}  {passes:>5}  {fails:>5}  {drift_str}")

    unstable = [s for s in all_steps if baseline and
                any(r["step_results"].get(s) != baseline["step_results"].get(s)
                    for r in variance_runs if s in r["step_results"])]
    if unstable:
        print(f"\n  ⚠ Steps that drifted from baseline:")
        for s in unstable:
            drifted = sum(1 for r in variance_runs
                         if r["step_results"].get(s) != baseline["step_results"].get(s))
            print(f"    {s}: drifted in {drifted}/{nv} runs (baseline={baseline['step_results'].get(s)})")
    else:
        print(f"\n  All steps matched baseline across all {nv} variance runs.")

    # Save full JSON
    report_path = SCRIPT_DIR / f"eval_llm_report_{datetime.now().strftime('%y%m%d_%H%M')}.json"
    with open(report_path, "w") as f:
        json.dump(results, f, indent=2, default=str)
    print(f"\n  Full report saved → {report_path.name}")
    print(f"{'='*70}\n")


# ── Main ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        print("ERROR: DEEPSEEK_API_KEY not set. Run: export DEEPSEEK_API_KEY=your_key")
        sys.exit(1)

    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

    # Phase 1
    data = collect_data()
    context_text = format_context(data)
    user_task = build_user_task(data)
    print(f"Context size: {len(context_text):,} chars")
    print(f"Tabs under evaluation:")
    print(f"  ADG ADO  : {data.get('latest_adg_tab')}")
    print(f"  Cluster  : {data.get('latest_cluster_tab')}")
    print(f"  SOB arch : {data.get('latest_sob_tab')}")
    print(f"  PC2 arch : {data.get('latest_pc2_tab')}\n")

    # Phase 2: run 1 at T=0 (baseline), runs 2-N at T=TEMPERATURE
    print(f"Phase 2: Run 1 = baseline (T=0), runs 2-{NUM_RUNS} = variance (T={TEMPERATURE}), model={MODEL}")
    results = []
    for i in range(1, NUM_RUNS + 1):
        temp = 0 if i == 1 else TEMPERATURE
        result = run_llm_verifier(client, context_text, user_task, i, temp)
        results.append(result)

    # Phase 3
    generate_report(results)
```

---

## Run Instructions

```bash
# 1. Install dependency
pip install openai

# 2. Set API key
export DEEPSEEK_API_KEY=your_key_here

# 3. Write and run
python3 eval_checks_llm.py
```

---

## What the Report Shows

| Metric | What it tells you |
|--------|------------------|
| **Fails per run** | How many steps the verifier flagged in each run |
| **Prompt / Completion tokens** | Token breakdown per run |
| **Cache hit tokens** | DeepSeek prompt caching — reduces cost on repeated runs |
| **Cost per run** | USD cost at DeepSeek-V3 pricing |
| **Consistency variance** | PASS/FAIL counts per step across 10 runs — `FLIPPED Nx` = unstable |
| **Unstable checks** | Checks where the LLM gave different verdicts in different runs |

**Variance interpretation:**
- `stable` → LLM is consistent on this check (same verdict every run)
- `FLIPPED 1x` → minor instability, likely borderline data
- `FLIPPED 3x+` → the check is ambiguous — the verifier prompt or data context needs sharpening

**To use DeepSeek-R1 instead of V3:** change `MODEL = "deepseek-reasoner"` and update pricing:
- Input cache miss: $0.55/M, cache hit: $0.14/M, output: $2.19/M
