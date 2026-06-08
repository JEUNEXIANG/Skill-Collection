"""
Evaluation runner for Platform SOB multi-agent workflow.
Runs executor + verifier 10 times in sandbox mode and reports metrics.

Usage:
    python3 eval_runner.py                  # 10 sessions, sandbox mode
    python3 eval_runner.py --sessions 5     # custom session count
"""

import anthropic
import json
import sys
import argparse
from datetime import datetime
from pathlib import Path

AGENT_DIR = Path(__file__).parent.parent   # platform-sob-agent/
sys.path.insert(0, str(AGENT_DIR))
import gsheets_util as gsu

# ── Config ────────────────────────────────────────────────────────────────────
SCRIPT_DIR = Path(__file__).parent        # eval/
MODEL = "claude-sonnet-4-6"

# claude-sonnet-4-6 pricing (per million tokens)
PRICE_INPUT = 3.00
PRICE_OUTPUT = 15.00

# ── Load specs ────────────────────────────────────────────────────────────────
EXECUTOR_SPEC = (AGENT_DIR / "SKILL.md").read_text()
VERIFIER_SPEC = (AGENT_DIR / "VERIFIER.md").read_text()

# ── Tool definitions ──────────────────────────────────────────────────────────
TOOLS = [
    {
        "name": "sheets_get_metadata",
        "description": "Read sheet metadata: title, tab list, sheet properties.",
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string", "description": "Google Sheet ID"},
                "fields": {"type": "string", "description": "Fields to return, e.g. 'properties.title,sheets.properties'"}
            },
            "required": ["sheet_id", "fields"]
        }
    },
    {
        "name": "sheets_read",
        "description": "Read cell values from a Google Sheet range.",
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string"},
                "range": {"type": "string", "description": "A1 notation, e.g. \"'Tab Name'!A1:Z100\""}
            },
            "required": ["sheet_id", "range"]
        }
    },
    {
        "name": "sheets_write",
        "description": "Write values to a Google Sheet range (paste as values only).",
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string"},
                "range": {"type": "string"},
                "values": {"type": "array", "items": {"type": "array"}, "description": "2D array of values"}
            },
            "required": ["sheet_id", "range", "values"]
        }
    },
    {
        "name": "sheets_batch_update",
        "description": "Execute Sheets API batchUpdate requests (tab duplication, formatting, insertions).",
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet_id": {"type": "string"},
                "requests": {"type": "array", "description": "List of batchUpdate request objects"}
            },
            "required": ["sheet_id", "requests"]
        }
    }
]


# ── Tool executor ─────────────────────────────────────────────────────────────
def execute_tool(name, inp, sandbox):
    try:
        if name == "sheets_get_metadata":
            return gsu.api_get(inp["sheet_id"], inp["fields"])

        elif name == "sheets_read":
            return gsu.values_get(inp["sheet_id"], inp["range"])

        elif name == "sheets_write":
            if sandbox:
                return {"sandbox": True, "skipped": f"Would write {len(inp['values'])} rows to {inp['range']}"}
            return gsu.values_update(inp["sheet_id"], inp["range"], inp["values"])

        elif name == "sheets_batch_update":
            if sandbox:
                req_types = [list(r.keys())[0] for r in inp["requests"]]
                return {"sandbox": True, "skipped": f"Would execute: {req_types}"}
            return gsu.api_batch_update(inp["sheet_id"], inp["requests"])

        return {"error": f"Unknown tool: {name}"}

    except Exception as e:
        return {"error": str(e)}


# ── Agentic loop ──────────────────────────────────────────────────────────────
def run_agent(client, system_prompt, user_message, sandbox, label="agent", max_turns=40):
    """
    Run a single-agent ReAct loop.
    Returns: (final_text, tool_call_log, input_tokens, output_tokens)
    """
    messages = [{"role": "user", "content": user_message}]
    tool_call_log = []
    total_in = 0
    total_out = 0

    for turn in range(max_turns):
        response = client.messages.create(
            model=MODEL,
            max_tokens=4096,
            system=system_prompt,
            tools=TOOLS,
            messages=messages
        )
        total_in += response.usage.input_tokens
        total_out += response.usage.output_tokens

        tool_uses = [b for b in response.content if b.type == "tool_use"]

        if response.stop_reason == "end_turn" or not tool_uses:
            texts = [b.text for b in response.content if hasattr(b, "text")]
            return "\n".join(texts), tool_call_log, total_in, total_out

        # Execute tools
        tool_results = []
        for tu in tool_uses:
            result = execute_tool(tu.name, tu.input, sandbox)
            tool_call_log.append({"tool": tu.name, "input_keys": list(tu.input.keys())})
            tool_results.append({
                "type": "tool_result",
                "tool_use_id": tu.id,
                "content": json.dumps(result, default=str)
            })

        messages.append({"role": "assistant", "content": response.content})
        messages.append({"role": "user", "content": tool_results})

    texts = [b.text for b in response.content if hasattr(b, "text")]
    return "\n".join(texts), tool_call_log, total_in, total_out


# ── Parse verifier output ─────────────────────────────────────────────────────
def parse_verifier_output(text):
    """
    Parse structured PASS/FAIL blocks from verifier output.
    Returns: (step_results dict, fail_count, pass_count)
    """
    step_results = {}
    current_step = None
    fail_count = 0
    pass_count = 0

    for line in text.split("\n"):
        upper = line.upper().strip()
        if upper.startswith("STEP:"):
            current_step = line.strip()
        elif "STATUS: FAIL" in upper and current_step:
            step_results[current_step] = "FAIL"
            fail_count += 1
            current_step = None
        elif "STATUS: PASS" in upper and current_step:
            step_results[current_step] = "PASS"
            pass_count += 1
            current_step = None

    return step_results, fail_count, pass_count


# ── Session runner ─────────────────────────────────────────────────────────────
def run_session(client, session_num, total_sessions, sandbox, today, today_yymmdd):
    print(f"\n── Session {session_num}/{total_sessions} {'[SANDBOX]' if sandbox else '[LIVE]'} ──")

    today_tab_adg     = f"{today} ADG ADO"
    today_tab_cluster = f"{today} By cluster"
    today_tab_sob     = f"SOB-{today_yymmdd}"
    today_tab_pc2     = f"PC2-{today_yymmdd}"

    # Executor
    print("  → Executor...")
    executor_prompt = f"""Run the Platform SOB workflow for today ({today}).
Tabs to create: '{today_tab_adg}', '{today_tab_cluster}', '{today_tab_sob}', '{today_tab_pc2}'.
{"SANDBOX MODE: reads are live, all writes and tab duplications are simulated — no actual changes to sheets." if sandbox else ""}
Execute all steps: Pre-flight, Step 1, Step 2, Step 3a, Step 3b, Step 4, Step 5, Step 6, PC2 Step 1.
For each step, describe what you read and what action you would take."""

    exec_text, exec_tools, exec_in, exec_out = run_agent(
        client, EXECUTOR_SPEC, executor_prompt, sandbox, label="executor"
    )
    print(f"     {len(exec_tools)} tool calls | {exec_in + exec_out:,} tokens")

    # Verifier
    print("  → Verifier...")
    verifier_prompt = f"""Verify the Platform SOB workflow execution for today.
today_date: {today_yymmdd}
today_tab_adg: {today_tab_adg}
today_tab_cluster: {today_tab_cluster}
today_tab_sob_archive: {today_tab_sob}
today_tab_pc2_archive: {today_tab_pc2}
steps_to_verify: ["preflight", "step1", "step2", "step3a", "step3b", "step4", "step5", "step6", "pc2_step1"]

Read all relevant sheets directly and audit each step independently.
Output a structured report using the format defined in your spec."""

    verif_text, verif_tools, verif_in, verif_out = run_agent(
        client, VERIFIER_SPEC, verifier_prompt, sandbox=False, label="verifier"
    )
    print(f"     {len(verif_tools)} tool calls | {verif_in + verif_out:,} tokens")

    step_results, fail_count, pass_count = parse_verifier_output(verif_text)
    total_tokens = exec_in + exec_out + verif_in + verif_out
    cost_usd = (exec_in + verif_in) / 1e6 * PRICE_INPUT + (exec_out + verif_out) / 1e6 * PRICE_OUTPUT

    print(f"     Verifier: {fail_count} fail(s), {pass_count} pass(es) | ${cost_usd:.4f}")

    return {
        "session": session_num,
        "executor": {
            "tool_calls": exec_tools,
            "tool_counts": _count_tools(exec_tools),
            "input_tokens": exec_in,
            "output_tokens": exec_out,
        },
        "verifier": {
            "tool_calls": verif_tools,
            "tool_counts": _count_tools(verif_tools),
            "input_tokens": verif_in,
            "output_tokens": verif_out,
            "step_results": step_results,
            "fail_count": fail_count,
            "pass_count": pass_count,
            "output": verif_text,
        },
        "all_passed": fail_count == 0,
        "total_tokens": total_tokens,
        "cost_usd": cost_usd,
    }


def _count_tools(tool_log):
    counts = {}
    for tc in tool_log:
        counts[tc["tool"]] = counts.get(tc["tool"], 0) + 1
    return counts


# ── Report ─────────────────────────────────────────────────────────────────────
def generate_report(results, today_yymmdd):
    n = len(results)
    full_passes = sum(1 for r in results if r["all_passed"])
    avg_fails   = sum(r["verifier"]["fail_count"] for r in results) / n
    avg_tokens  = sum(r["total_tokens"] for r in results) / n
    total_cost  = sum(r["cost_usd"] for r in results)

    print(f"\n{'='*72}")
    print(f"  EVALUATION REPORT — Platform SOB Agent  ({n} sessions, SANDBOX)")
    print(f"{'='*72}")

    # Per-session table
    header = f"{'Sess':>5}  {'Verif Fails':>11}  {'Exec Tools':>10}  {'Verif Tools':>11}  {'Tokens':>8}  {'Cost':>7}  {'Pass?':>5}"
    print(f"\n{header}")
    print("─" * 72)
    for r in results:
        exec_total  = sum(r["executor"]["tool_counts"].values())
        verif_total = sum(r["verifier"]["tool_counts"].values())
        print(
            f"  {r['session']:>3}  "
            f"{r['verifier']['fail_count']:>11}  "
            f"{exec_total:>10}  "
            f"{verif_total:>11}  "
            f"{r['total_tokens']:>8,}  "
            f"${r['cost_usd']:>5.4f}  "
            f"{'  ✓' if r['all_passed'] else '  ✗'}"
        )
    print("─" * 72)
    print(f"{'AVG':>5}  {avg_fails:>11.1f}  {'':>10}  {'':>11}  {avg_tokens:>8,.0f}  ${total_cost/n:>5.4f}")
    print(f"\n  Sessions with ALL checks passed : {full_passes}/{n}")
    print(f"  Total cost across all sessions  : ${total_cost:.4f}")

    # Tool usage breakdown
    print(f"\n  Tool usage (executor, across all sessions):")
    exec_agg = {}
    for r in results:
        for tool, cnt in r["executor"]["tool_counts"].items():
            exec_agg[tool] = exec_agg.get(tool, 0) + cnt
    for tool, cnt in sorted(exec_agg.items(), key=lambda x: -x[1]):
        print(f"    {tool:<30} {cnt:>4}x  (avg {cnt/n:.1f}/session)")

    print(f"\n  Tool usage (verifier, across all sessions):")
    verif_agg = {}
    for r in results:
        for tool, cnt in r["verifier"]["tool_counts"].items():
            verif_agg[tool] = verif_agg.get(tool, 0) + cnt
    for tool, cnt in sorted(verif_agg.items(), key=lambda x: -x[1]):
        print(f"    {tool:<30} {cnt:>4}x  (avg {cnt/n:.1f}/session)")

    # Failure breakdown by step
    fail_tally = {}
    for r in results:
        for step, status in r["verifier"]["step_results"].items():
            if status == "FAIL":
                fail_tally[step] = fail_tally.get(step, 0) + 1

    if fail_tally:
        print(f"\n  Verifier failure breakdown by step:")
        for step, cnt in sorted(fail_tally.items(), key=lambda x: -x[1]):
            print(f"    {step:<40} failed {cnt}/{n}")
    else:
        print(f"\n  No verifier failures recorded across all sessions.")

    # Save JSON report
    report_path = SCRIPT_DIR / f"eval_report_{today_yymmdd}.json"
    with open(report_path, "w") as f:
        json.dump(results, f, indent=2, default=str)
    print(f"\n  Full report saved → {report_path.name}")
    print(f"{'='*72}\n")


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Evaluate Platform SOB multi-agent workflow")
    parser.add_argument("--sessions", type=int, default=10, help="Number of sessions to run (default: 10)")
    parser.add_argument("--live", action="store_true", help="Run in live mode (default: sandbox)")
    args = parser.parse_args()

    sandbox = not args.live
    today = datetime.now().strftime("%y/%m/%d")
    today_yymmdd = datetime.now().strftime("%y%m%d")

    print(f"Platform SOB Evaluation")
    print(f"  Sessions : {args.sessions}")
    print(f"  Mode     : {'SANDBOX (no writes)' if sandbox else 'LIVE (real writes)'}")
    print(f"  Date     : {today}")
    print(f"  Model    : {MODEL}")

    client = anthropic.Anthropic()
    results = []

    for i in range(1, args.sessions + 1):
        result = run_session(client, i, args.sessions, sandbox, today, today_yymmdd)
        results.append(result)

    generate_report(results, today_yymmdd)


if __name__ == "__main__":
    main()
