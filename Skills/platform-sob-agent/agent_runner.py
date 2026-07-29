"""
agent_runner.py — Platform SOB production orchestrator (per-step verifier gate).

THIS IS PRODUCTION ONLY. It is deliberately separate from the eval path.
  - Eval (eval/eval_runner.py::run_session) stays BATCH: the executor runs the
    whole workflow autonomously, then the verifier intervenes ONCE at the end.
    That measures the executor's raw end-to-end capability.
  - This runner GATES per step: executor does step N -> verifier checks step N
    at T=0 -> executor only advances to step N+1 on PASS. Never use this for eval;
    gating contaminates the measurement (you'd be scoring the gated system).

Loop per step:
    1. (write steps, live only) request human approval   -> permission_gate
    2. executor performs ONLY this step                  -> run_agent(EXECUTOR_SPEC)
    3. verifier audits ONLY this step at T=0             -> run_agent(VERIFIER_SPEC)
    4. branch on PASS / FAIL / SKIP
    5. record into RunRecord                             -> audit_trail
At the end: finalize_run -> audit log entry + WeChat outcome report.

Gate policy (gate="hard" halts on FAIL; gate="report" logs FAIL but continues):
    Only steps that were STABLE in the latest eval (T=0 and T=0.7 agree) are
    hard-gated today. Unstable steps (3/4/5) and the not-yet-redesigned PC2 check
    run in report-only mode so a flaky FAIL cannot wrongly halt a real run.
    Promote a step to "hard" once its spec is stabilized.

Usage:
    python agent_runner.py                # SANDBOX (no writes), full run
    python agent_runner.py --live         # LIVE: real writes, human approval gates
    python agent_runner.py --gate-all     # treat every step as a hard gate
"""

import os
import sys
import argparse
from datetime import datetime
from pathlib import Path

from openai import OpenAI

AGENT_DIR = Path(__file__).parent          # platform-sob-agent/
sys.path.insert(0, str(AGENT_DIR / "eval"))
sys.path.insert(0, str(AGENT_DIR / "harness"))

# Reuse the agentic loop + specs from the eval runner (single source of truth).
# Note: importing eval_runner instantiates a FastAPI app at module load, so
# fastapi/uvicorn must be installed in the environment.
try:
    from eval_runner import run_agent, parse_verifier_output, EXECUTOR_SPEC, VERIFIER_SPEC
except Exception as e:                      # pragma: no cover
    print(f"ERROR: could not import eval_runner ({e}).")
    print("       agent_runner reuses run_agent/specs from eval/eval_runner.py.")
    print("       Ensure fastapi + uvicorn are installed: pip install fastapi uvicorn")
    sys.exit(1)

# Harness components (imported by bare name — harness/ is on sys.path)
from audit_trail import RunRecord
from permission_gate import request_approval, alert_failure
from outcome_evaluator import finalize_run

# Deterministic gate temperature. The verdict here is control flow, not a
# diagnostic — randomness would be a bug. (T=0.7 was a *spec* stress test only.)
GATE_TEMPERATURE = 0


# ── Step plan ────────────────────────────────────────────────────────────────
# label MUST match the "STEP:" line the verifier emits, so verdicts parse back.
# gate:  "hard"   -> FAIL halts the run
#        "report" -> FAIL is logged but the run continues (spec not yet stable)
STEPS = [
    {"key": "preflight", "label": "Pre-flight",
     "writes": False, "gate": "report"},  # TEMP: demoted from hard — verifier hallucinates on healthy sheets
    {"key": "step2",     "label": "Step 2 — Copy SOB Values to [Reg CNLS copy]",
     "writes": True,  "gate": "hard"},
    {"key": "step3",     "label": "Step 3 — Paste Within [Reg CNLS copy]",
     "writes": True,  "gate": "report"},   # Type C divergence — add Pass Condition first
    {"key": "step4",     "label": "Step 4 — Zero Check",
     "writes": False, "gate": "report"},   # SKIP/FAIL distinction not yet exact
    {"key": "step5",     "label": "Step 5 — Archive SOB",
     "writes": True,  "gate": "report", "depends_on": "step4"},  # archive only if step4 PASS
    {"key": "step6",     "label": "Step 6 — Clusters Data",
     "writes": True,  "gate": "hard"},
    {"key": "pc2_step1", "label": "PC2 Step 1 — Archive PC2",
     "writes": True,  "gate": "report"},   # PC2 check pending redesign
]


# ── Verdict extraction (single-step) ─────────────────────────────────────────
def extract_verdict(verifier_text: str) -> str:
    """Collapse a single-step verifier output to PASS / FAIL / SKIP / UNKNOWN."""
    step_results, _, _ = parse_verifier_output(verifier_text)
    verdicts = [v.upper() for v in step_results.values()]
    if any(v == "FAIL" for v in verdicts):
        return "FAIL"
    if any(v.startswith("SKIP") for v in verdicts):
        return "SKIP"
    if any(v == "PASS" for v in verdicts):
        return "PASS"
    return "UNKNOWN"


# ── Prompts (scoped to ONE step) ─────────────────────────────────────────────
def executor_prompt(step, tabs, today, sandbox):
    mode = ("SANDBOX MODE: reads are live; all writes and tab duplications are "
            "simulated — make no actual changes.\n" if sandbox else "")
    return f"""Run ONLY this step of the Platform SOB workflow for today ({today}): {step['label']}.
Tabs in play: '{tabs['adg']}', '{tabs['cluster']}', '{tabs['sob']}', '{tabs['pc2']}'.
{mode}Do not perform any other step. Describe what you read and the action you took."""


def verifier_prompt(step, tabs, today_yymmdd):
    return f"""Verify ONLY this step of the Platform SOB workflow: {step['label']}.
today_date            : {today_yymmdd}
today_tab_adg         : {tabs['adg']}
today_tab_cluster     : {tabs['cluster']}
today_tab_sob_archive : {tabs['sob']}
today_tab_pc2_archive : {tabs['pc2']}
steps_to_verify       : ["{step['key']}"]

Read the relevant sheets directly and audit ONLY this step.
Output the structured block for this one step using your spec format,
ending with its STATUS line (PASS / FAIL / SKIP)."""


# ── Orchestrator loop ────────────────────────────────────────────────────────
def run(client, sandbox: bool, gate_all: bool):
    today        = datetime.now().strftime("%y/%m/%d")
    today_yymmdd = datetime.now().strftime("%y%m%d")
    tabs = {
        "adg":     f"{today} ADG ADO",
        "cluster": f"{today} By cluster",
        "sob":     f"SOB-{today_yymmdd}",
        "pc2":     f"PC2-{today_yymmdd}",
    }

    record   = RunRecord()
    verdicts = {}   # step key -> PASS/FAIL/SKIP

    print(f"\n{'='*68}")
    print(f"  Platform SOB — PRODUCTION run  ({'SANDBOX' if sandbox else 'LIVE'})")
    print(f"  Date {today}  |  gate {'ALL-HARD' if gate_all else 'per-step policy'}  |  run {record.run_id}")
    print(f"{'='*68}")

    for step in STEPS:
        key, label = step["key"], step["label"]
        gate = "hard" if gate_all else step["gate"]
        print(f"\n── {label} ──  (gate: {gate})")

        # 1. Dependency: skip a step whose prerequisite did not PASS.
        dep = step.get("depends_on")
        if dep and verdicts.get(dep) != "PASS":
            print(f"   SKIP — prerequisite {dep} did not PASS ({verdicts.get(dep)}).")
            record.add_note(f"{key}: skipped (prerequisite {dep} = {verdicts.get(dep)})")
            verdicts[key] = "SKIP"
            continue

        # 2. Human approval before any live write step.
        if step["writes"] and not sandbox:
            plan = f"About to execute {label}.\nTabs: {tabs}"
            if not request_approval(label, plan, sandbox_mode=False):
                record.fail(key, "User rejected or approval timed out")
                print("   HALT — approval not granted.")
                break

        # 3. Executor performs this step only.
        exec_text, _, _, _, exec_cost = run_agent(
            client, EXECUTOR_SPEC, executor_prompt(step, tabs, today, sandbox),
            sandbox, label=f"exec:{key}")
        print(f"   executor done (${exec_cost:.5f})")

        # 4. Verifier gate (deterministic, single step).
        verif_text, _, _, _, verif_cost = run_agent(
            client, VERIFIER_SPEC, verifier_prompt(step, tabs, today_yymmdd),
            sandbox=False, label=f"verify:{key}", temperature=GATE_TEMPERATURE)
        verdict = extract_verdict(verif_text)
        verdicts[key] = verdict
        print(f"   verifier verdict: {verdict}  (${verif_cost:.5f})")

        # Record step-4/5 validation outcomes for the outcome evaluator.
        if key == "step4":
            record.validation_results["step4_zero_check"] = (verdict == "PASS")
        if key == "step5":
            record.validation_results["step5_value_match"] = (verdict == "PASS")

        # 5. Branch.
        if verdict == "PASS":
            record.complete(key)
            for k, v in (("adg", tabs["adg"]), ) if key == "step2" else ():
                record.add_tab("Reg CNLS copy", v)
        elif verdict == "SKIP":
            record.add_note(f"{key}: verifier returned SKIP")
        else:  # FAIL or UNKNOWN
            if gate == "hard":
                record.fail(key, f"Verifier {verdict} on {label}")
                alert_failure(label, f"Verifier returned {verdict}. Run halted before next step.")
                print(f"   HALT — hard gate failed on {label}.")
                break
            else:
                record.add_note(f"{key}: verifier {verdict} (report-only gate — not halted)")
                record.complete(key)   # executed; spec not yet trusted to gate
                print(f"   report-only — logged {verdict}, continuing.")

    # 6. Close out: audit log entry + outcome report.
    finalize_run(record, sandbox=sandbox)


# ── Entry point ──────────────────────────────────────────────────────────────
def _load_api_key():
    key = os.environ.get("DEEPSEEK_API_KEY")
    if key:
        return key
    cfg = Path.home() / ".hermes" / "config.yaml"
    if cfg.exists():
        for line in cfg.read_text().splitlines():
            if line.strip().startswith("api_key:"):
                return line.split(":", 1)[1].strip().strip('"').strip("'")
    return None


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Platform SOB production orchestrator (per-step gate)")
    parser.add_argument("--live", action="store_true",
                        help="Real writes + human approval gates (default: sandbox, no writes)")
    parser.add_argument("--gate-all", action="store_true",
                        help="Treat every step as a hard gate (halt on any FAIL)")
    args = parser.parse_args()

    api_key = _load_api_key()
    if not api_key:
        print("ERROR: DEEPSEEK_API_KEY not set (env var or ~/.hermes/config.yaml).")
        sys.exit(1)

    client = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
    run(client, sandbox=not args.live, gate_all=args.gate_all)
