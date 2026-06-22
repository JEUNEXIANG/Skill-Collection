"""
Platform SOB Verifier — LLM-based evaluation runner.
Phase 1: collect sheet data via gsheets_util.py (once).
Phase 2: call DeepSeek verifier LLM 10 times on the same data.
Phase 3: report consistency variance and token costs.
"""

import os
import sys
import json
import re
from pathlib import Path
from datetime import datetime
from openai import OpenAI

AGENT_DIR = Path(__file__).parent.parent   # platform-sob-agent/
sys.path.insert(0, str(AGENT_DIR))
import gsheets_util as gsu

# ── Config ─────────────────────────────────────────────────────────────────────
NUM_RUNS    = 10
MODEL       = "deepseek-chat"
TEMPERATURE = 0.7
MAX_TOKENS  = 8192

PRICE_INPUT_CACHE_MISS = 0.27
PRICE_INPUT_CACHE_HIT  = 0.07
PRICE_OUTPUT           = 1.10

SCRIPT_DIR    = Path(__file__).parent      # eval/
VERIFIER_SPEC = (AGENT_DIR / "VERIFIER.md").read_text()

# ── Sheet IDs ──────────────────────────────────────────────────────────────────
SID_WLV  = gsu.SHEETS["Weekly Live View"]
SID_RCT  = gsu.SHEETS["Reg Commercial Team"]
SID_CNLS = gsu.SHEETS["Reg CNLS copy"]
SID_ARCH = gsu.SHEETS["Archive"]
SID_PC2  = gsu.SHEETS["Platform PC2"]


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

    d["wlv_tabs"]  = get_tabs(SID_WLV)
    d["rct_tabs"]  = get_tabs(SID_RCT)
    d["cnls_tabs"] = get_tabs(SID_CNLS)
    d["arch_tabs"] = get_tabs(SID_ARCH)
    d["pc2_tabs"]  = get_tabs(SID_PC2)

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

    # Step 2: source data (for value match) + destination sample
    if d["adg_src_tab"]:
        d["step2_src_data"]    = read_range(SID_WLV,  d["adg_src_tab"],     "A1:Z80", max_rows=80)
    if d["latest_adg_tab"]:
        d["step2_dest_sample"] = read_range(SID_CNLS, d["latest_adg_tab"],  "A1:Z80", max_rows=80)

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
        d["step5_archive"]     = read_range(SID_ARCH,            d["latest_sob_tab"], "A1:Z40")
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
steps_to_verify: ["preflight", "step2", "step3", "step4", "step5", "step6", "pc2_step1"]

For each step output EXACTLY this format (no deviation):

STEP: Pre-flight
STATUS: PASS
CHECKS:
  [✓] All sheet IDs resolve
  ...

Continue for all steps: Pre-flight, Step 2, Step 3, Step 4, Step 5, Step 6, PC2 Step 1.
End with a single line: OVERALL: PASS or OVERALL: FAIL
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

    cache_hit  = getattr(usage, "prompt_cache_hit_tokens", 0)
    cache_miss = getattr(usage, "prompt_cache_miss_tokens", usage.prompt_tokens)
    out_tokens = usage.completion_tokens
    cost = (cache_hit  / 1e6 * PRICE_INPUT_CACHE_HIT
          + cache_miss / 1e6 * PRICE_INPUT_CACHE_MISS
          + out_tokens / 1e6 * PRICE_OUTPUT)

    # ── Parse structured output ─────────────────────────────────────────────
    step_results = {}
    step_checks = {}   # step_name -> list of check strings
    overall = None
    current_step = None
    in_checks = False

    for line in text.split("\n"):
        stripped = line.strip()
        if stripped.startswith("STEP:"):
            current_step = stripped[5:].strip()
            step_checks[current_step] = []
            in_checks = False
        elif stripped.startswith("STATUS:") and current_step:
            step_results[current_step] = stripped[7:].strip()
            in_checks = False
        elif stripped.upper().startswith("CHECKS") or stripped.upper().startswith("CHECKS:"):
            in_checks = True
        elif in_checks and current_step:
            # Capture check lines - [✓] or [✗]
            if stripped.startswith("[") and ("]" in stripped):
                step_checks[current_step].append(stripped)
            # Also capture CONTEXT lines
            elif stripped.upper().startswith("CONTEXT:"):
                step_checks.setdefault(f"{current_step}_context", []).append(stripped)
        elif stripped.upper().startswith("OVERALL:"):
            overall = stripped[8:].strip()
        elif stripped.upper().startswith("SECOND ATTEMPT") or stripped.upper().startswith("HUMAN INTERVENTION"):
            overall = "FAIL_HUMAN_INTERVENTION"

    fail_count = sum(1 for v in step_results.values() if v == "FAIL")
    print(f"{fail_count} fail(s) | {usage.prompt_tokens + out_tokens:,} tokens | ${cost:.5f}")

    return {
        "run": run_num,
        "temperature": temperature,
        "is_baseline": temperature == 0,
        "step_results": step_results,
        "step_checks": step_checks,
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

    print(f"\n{'='*75}")
    print(f"  EVALUATION REPORT — Platform SOB Verifier  ({n} runs, model={MODEL})")
    print(f"{'='*75}")

    # ── Per-run table ─────────────────────────────────────────────────────
    print(f"\n{'Run':>4}  {'Fails':>5}  {'Prompt':>8}  {'Completion':>10}  {'CacheHit':>9}  {'Cost':>8}  {'Pass?'}")
    print("\u2500" * 75)
    for r in results:
        t = r["tokens"]
        print(f"  {r['run']:>2}  {r['fail_count']:>5}  {t['prompt']:>8,}  "
              f"{t['completion']:>10,}  {t['cache_hit']:>9,}  "
              f"${r['cost_usd']:>6.5f}  {'  \u2713' if r['all_passed'] else '  \u2717'}")
    print("\u2500" * 75)
    print(f"{'AVG':>4}  {sum(r['fail_count'] for r in results)/n:>5.1f}  "
          f"{avg_tokens:>8,.0f}  {'':>10}  {'':>9}  ${avg_cost:>6.5f}")
    print(f"\n  Runs with ALL checks passed : {full_passes}/{n}")
    print(f"  Total cost                  : ${total_cost:.5f}")

    # ── Consistency variance ──────────────────────────────────────────────
    all_steps = sorted({step for r in results for step in r["step_results"]})
    baseline      = next((r for r in results if r["is_baseline"]), None)
    variance_runs = [r for r in results if not r["is_baseline"]]
    nv            = len(variance_runs)

    print(f"\n  {'='*70}")
    print(f"  CONSISTENCY VARIANCE BY STEP (baseline=run1 T=0, variance={nv} runs T={TEMPERATURE})")
    print(f"  {'='*70}")
    print(f"  {'Step':<42} {'Base':>5}  {'PASS':>5}  {'FAIL':>5}  {'Drift from baseline'}")
    print("  " + "\u2500" * 68)
    for step in all_steps:
        base_verdict = baseline["step_results"].get(step, "\u2014") if baseline else "\u2014"
        verdicts  = [r["step_results"].get(step) for r in variance_runs if step in r["step_results"]]
        passes    = verdicts.count("PASS")
        fails     = verdicts.count("FAIL")
        drifted   = sum(1 for v in verdicts if v and v != base_verdict)
        drift_str = f"drifted {drifted}/{nv}" if drifted > 0 else "stable"
        print(f"  {step:<42} {base_verdict:>5}  {passes:>5}  {fails:>5}  {drift_str}")

    # ── Detailed failure criteria across all runs ─────────────────────────┐
    print(f"\n  {'='*70}")
    print(f"  FAILURE CRITERIA — All failed checks across all runs")
    print(f"  {'='*70}")

    # Collect all unique check texts that appeared when a step was FAIL
    all_fail_checks = {}  # step -> set of check descriptions
    all_pass_checks = {}  # step -> set of check descriptions (for comparison)
    for r in results:
        for step, checks in r["step_checks"].items():
            if "_context" in step:
                continue
            status = r["step_results"].get(step, "UNKNOWN")
            if status == "FAIL":
                all_fail_checks.setdefault(step, set()).update(checks)
            elif status == "PASS":
                all_pass_checks.setdefault(step, set()).update(checks)

    for step in all_steps:
        fail_checks = all_fail_checks.get(step, set())
        pass_checks = all_pass_checks.get(step, set())

        if not fail_checks:
            continue

        # Only failures that never appeared in any PASS run (or vice versa)
        print(f"\n  --- {step} ---")
        for c in sorted(fail_checks):
            marker = "✗" if "[✗]" in c else "  "
            times_failed = sum(1 for r in results
                               if step in r["step_checks"]
                               and c in r["step_checks"][step])
            # Sort: failing checks first
            if "[✗]" in c:
                print(f"    [FAIL REASON] {c[4:].strip()}  (seen in {times_failed}/{n} FAIL runs)")
            else:
                print(f"    [PASS CHECK]  {c[4:].strip()}  (seen in {times_failed}/{n} FAIL runs)")

        # Show checks that are in pass but not in fail (the ones that flip)
        only_pass_checks = pass_checks - fail_checks
        if only_pass_checks:
            for c in sorted(only_pass_checks):
                times_passed = sum(1 for r in results
                                   if step in r["step_checks"]
                                   and c in r["step_checks"][step])
                if "[✗]" in c:
                    print(f"    [FAIL (only in PASS runs)] {c[4:].strip()}  (seen in {times_passed}/{n} PASS runs)")

    # ── Aggregated check-level consistency ────────────────────────────────
    print(f"\n  {'='*70}")
    print(f"  CHECK-LEVEL CONSISTENCY — Which criteria flip most?")
    print(f"  {'='*70}")

    # Find the most common FAIL checks across all runs
    from collections import Counter
    fail_check_counter = Counter()
    check_step_map = {}
    for r in results:
        for step, checks in r["step_checks"].items():
            if "_context" in step:
                continue
            status = r["step_results"].get(step, "UNKNOWN")
            if status == "FAIL":
                for c in checks:
                    if "[✗]" in c:
                        # Normalize: remove cell numbers, keep the pattern
                        normalized = re.sub(r'\bcell [A-Z]+\d+\b', 'cell X', c)
                        normalized = re.sub(r'\brow \d+\b', 'row X', normalized)
                        normalized = re.sub(r'\be\.g\..*', 'e.g. ...', normalized)
                        fail_check_counter[normalized] += 1
                        check_step_map[normalized] = step

    for check_text, count in fail_check_counter.most_common(20):
        step_name = check_step_map.get(check_text, "?")
        stability = "UNSTABLE" if 0 < count < n else "stable" if count == n else ""
        print(f"  {stability:<10} {count:>2}/{n}  [{step_name}] {check_text[4:].strip()[:90]}")

    # ── Summary ───────────────────────────────────────────────────────────
    unstable = [s for s in all_steps
                if 0 < sum(1 for r in results if r["step_results"].get(s) == "FAIL") < n]
    if unstable:
        print(f"\n  {'='*70}")
        print(f"  STABILITY SUMMARY — Steps with flipped verdicts:")
        for s in unstable:
            fails = sum(1 for r in results if r["step_results"].get(s) == "FAIL")
            print(f"    {s:<48} FAIL in {fails}/{n} runs")
    else:
        print(f"\n  All checks were stable across {n} runs.")

    # ── Save ──────────────────────────────────────────────────────────────
    report_path = SCRIPT_DIR / f"eval_llm_report_{datetime.now().strftime('%y%m%d_%H%M')}.json"
    with open(report_path, "w") as f:
        json.dump(results, f, indent=2, default=str)
    print(f"\n  Full report saved \u2192 {report_path.name}")
    print(f"{'='*75}\n")


# ── Main ───────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # Try env var first, then fallback to Hermes config
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        hermes_cfg = Path.home() / ".hermes" / "config.yaml"
        if hermes_cfg.exists():
            for line in hermes_cfg.read_text().splitlines():
                if line.strip().startswith("api_key:"):
                    api_key = line.split(":", 1)[1].strip().strip('"').strip("'")
                    break
    if not api_key:
        print("ERROR: DEEPSEEK_API_KEY not set. Set it via env var or add to ~/.hermes/config.yaml")
        sys.exit(1)

    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")

    # Phase 1
    data         = collect_data()
    context_text = format_context(data)
    user_task    = build_user_task(data)
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
        temp   = 0 if i == 1 else TEMPERATURE
        result = run_llm_verifier(client, context_text, user_task, i, temp)
        results.append(result)

    # Phase 3
    generate_report(results)
