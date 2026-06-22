"""
Evaluation runner for Platform SOB multi-agent workflow.
Now also serves as a FastAPI server when run directly.

Usage:
    python3 eval_runner.py                  # Start API server (default, port 8888)
    python3 eval_runner.py --cli            # Old CLI mode: run N eval sessions
    python3 eval_runner.py --cli --sessions 5 --live

API endpoints (when running as server):
    GET  /                         -> API info
    GET  /health                  -> Health check
    POST /agent/run               -> Run a single agent (executor/verifier)
    POST /evaluate                -> Run N-session evaluation
    GET  /results                 -> Latest evaluation results

Requires: DEEPSEEK_API_KEY environment variable
"""

import os
import sys
import json
import argparse
import urllib.request, urllib.parse
from datetime import datetime
from pathlib import Path
from openai import OpenAI

# --- FastAPI ------------------------------------------------------------------
try:
    from fastapi import FastAPI, HTTPException
    from fastapi.responses import JSONResponse
    import uvicorn
    HAS_FASTAPI = True
except ImportError:
    HAS_FASTAPI = False

AGENT_DIR = Path(__file__).parent.parent   # platform-sob-agent/
sys.path.insert(0, str(AGENT_DIR))
import gsheets_util as gsu

# --- Config -------------------------------------------------------------------
SCRIPT_DIR  = Path(__file__).parent        # eval/
MODEL       = "deepseek-chat"              # DeepSeek-V3; use "deepseek-reasoner" for R1
MAX_TOKENS  = 8192

# DeepSeek-V3 pricing per million tokens
PRICE_INPUT_CACHE_MISS = 0.27
PRICE_INPUT_CACHE_HIT  = 0.07
PRICE_OUTPUT           = 1.10

# --- API Server config --------------------------------------------------------
API_HOST = "0.0.0.0"
API_PORT = 8888

# --- Load specs ---------------------------------------------------------------
EXECUTOR_SPEC = (AGENT_DIR / "SKILL.md").read_text()
VERIFIER_SPEC = (AGENT_DIR / "VERIFIER.md").read_text()

# --- In-memory results store --------------------------------------------------
_latest_results = []

# --- Tool definitions (OpenAI function-calling format) ------------------------
TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "sheets_get_metadata",
            "description": "Read sheet metadata: title, tab list, sheet properties.",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_id": {"type": "string", "description": "Google Sheet ID"},
                    "fields":   {"type": "string", "description": "Fields to return, e.g. 'properties.title,sheets.properties'"}
                },
                "required": ["sheet_id", "fields"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "sheets_read",
            "description": "Read cell values from a Google Sheet range.",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_id": {"type": "string"},
                    "range":    {"type": "string", "description": "A1 notation, e.g. \"'Tab Name'!A1:Z100\""}
                },
                "required": ["sheet_id", "range"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "sheets_read_unformatted",
            "description": "Read raw (unformatted) cell values. Use to distinguish a number-format x suffix (displays as '6.26x' but stores 6.26) from a literal string ending in x.",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_id": {"type": "string"},
                    "range":    {"type": "string"}
                },
                "required": ["sheet_id", "range"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "sheets_write",
            "description": "Write values to a Google Sheet range (paste as values only).",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_id": {"type": "string"},
                    "range":    {"type": "string"},
                    "values":   {"type": "array", "items": {"type": "array"}, "description": "2D array of values"}
                },
                "required": ["sheet_id", "range", "values"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "sheets_batch_update",
            "description": "Execute Sheets API batchUpdate requests (tab duplication, formatting, insertions).",
            "parameters": {
                "type": "object",
                "properties": {
                    "sheet_id": {"type": "string"},
                    "requests": {"type": "array", "description": "List of batchUpdate request objects"}
                },
                "required": ["sheet_id", "requests"]
            }
        }
    }
]


# --- Tool executor ------------------------------------------------------------
def execute_tool(name, inp, sandbox):
    try:
        if name == "sheets_get_metadata":
            return gsu.api_get(inp["sheet_id"], inp["fields"])

        elif name == "sheets_read":
            return gsu.values_get(inp["sheet_id"], inp["range"])

        elif name == "sheets_read_unformatted":
            headers   = gsu.get_auth_headers()
            range_enc = urllib.parse.quote(inp["range"], safe='!')
            url = (f"https://sheets.googleapis.com/v4/spreadsheets/{inp['sheet_id']}"
                   f"/values/{range_enc}?valueRenderOption=UNFORMATTED_VALUE")
            req = urllib.request.Request(url, headers=headers)
            with urllib.request.urlopen(req, timeout=30, context=gsu.SSL_CTX) as resp:
                return json.loads(resp.read())

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


# --- Token cost helper --------------------------------------------------------
def calc_cost(usage):
    cache_hit  = getattr(usage, "prompt_cache_hit_tokens", 0)
    cache_miss = getattr(usage, "prompt_cache_miss_tokens", usage.prompt_tokens)
    out        = usage.completion_tokens
    cost = (cache_hit  / 1e6 * PRICE_INPUT_CACHE_HIT
          + cache_miss / 1e6 * PRICE_INPUT_CACHE_MISS
          + out        / 1e6 * PRICE_OUTPUT)
    return cache_hit, cache_miss, out, cost


# --- Agentic loop -------------------------------------------------------------
def run_agent(client, system_prompt, user_message, sandbox, label="agent", max_turns=40, temperature=None):
    """
    Run a single-agent ReAct loop via DeepSeek (OpenAI-compatible).
    Returns: (final_text, tool_call_log, prompt_tokens, completion_tokens, total_cost)

    temperature: if provided, pins sampling (e.g. 0 for a deterministic production gate).
                 If None, uses the provider default — preserves prior eval behavior.
    """
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user",   "content": user_message},
    ]
    tool_call_log  = []
    total_prompt   = 0
    total_complete = 0
    total_cost     = 0.0

    for _ in range(max_turns):
        create_kwargs = dict(
            model=MODEL,
            max_tokens=MAX_TOKENS,
            tools=TOOLS,
            tool_choice="auto",
            messages=messages,
        )
        if temperature is not None:
            create_kwargs["temperature"] = temperature
        response = client.chat.completions.create(**create_kwargs)

        _, _, _, cost = calc_cost(response.usage)
        total_prompt   += response.usage.prompt_tokens
        total_complete += response.usage.completion_tokens
        total_cost     += cost

        msg = response.choices[0].message
        messages.append(msg)

        if not msg.tool_calls:
            return msg.content or "", tool_call_log, total_prompt, total_complete, total_cost

        for tc in msg.tool_calls:
            args   = json.loads(tc.function.arguments)
            result = execute_tool(tc.function.name, args, sandbox)
            tool_call_log.append({
                "tool":       tc.function.name,
                "input_keys": list(args.keys()),
            })
            messages.append({
                "role":         "tool",
                "tool_call_id": tc.id,
                "content":      json.dumps(result, default=str),
            })

    return msg.content or "", tool_call_log, total_prompt, total_complete, total_cost


# --- Parse verifier output ----------------------------------------------------
def parse_verifier_output(text):
    step_results = {}
    current_step = None
    fail_count   = 0
    pass_count   = 0

    for line in text.split("\n"):
        stripped = line.strip()
        upper    = stripped.upper()
        if upper.startswith("STEP:"):
            current_step = stripped[5:].strip()
        elif upper.startswith("STATUS:") and current_step:
            verdict = stripped[7:].strip().upper()
            step_results[current_step] = verdict
            if verdict == "FAIL":
                fail_count += 1
            elif verdict == "PASS":
                pass_count += 1
            current_step = None

    return step_results, fail_count, pass_count


# --- Session runner -----------------------------------------------------------
def run_session(client, session_num, total_sessions, sandbox, today, today_yymmdd):
    print(f"\n--- Session {session_num}/{total_sessions} {'[SANDBOX]' if sandbox else '[LIVE]'} ---")

    today_tab_adg     = f"{today} ADG ADO"
    today_tab_cluster = f"{today} By cluster"
    today_tab_sob     = f"SOB-{today_yymmdd}"
    today_tab_pc2     = f"PC2-{today_yymmdd}"

    # --- Executor ---
    print("  -> Executor...")
    executor_prompt = f"""Run the Platform SOB workflow for today ({today}).
Tabs to create: '{today_tab_adg}', '{today_tab_cluster}', '{today_tab_sob}', '{today_tab_pc2}'.
{"SANDBOX MODE: reads are live, all writes and tab duplications are simulated - no actual changes to sheets." if sandbox else ""}
Execute all steps: Pre-flight, Step 2, Step 3, Step 4, Step 5, Step 6, PC2 Step 1.
For each step, describe what you read and what action you would take."""

    exec_text, exec_tools, exec_in, exec_out, exec_cost = run_agent(
        client, EXECUTOR_SPEC, executor_prompt, sandbox, label="executor"
    )
    print(f"     {len(exec_tools)} tool calls | {exec_in + exec_out:,} tokens | ${exec_cost:.5f}")

    # --- Verifier ---
    print("  -> Verifier...")
    verifier_prompt = f"""Verify the Platform SOB workflow execution for today.
today_date            : {today_yymmdd}
today_tab_adg         : {today_tab_adg}
today_tab_cluster     : {today_tab_cluster}
today_tab_sob_archive : {today_tab_sob}
today_tab_pc2_archive : {today_tab_pc2}
steps_to_verify       : ["preflight", "step2", "step3", "step4", "step5", "step6", "pc2_step1"]

Read all relevant sheets directly and audit each step independently.
Output a structured report using the format defined in your spec."""

    verif_text, verif_tools, verif_in, verif_out, verif_cost = run_agent(
        client, VERIFIER_SPEC, verifier_prompt, sandbox=False, label="verifier"
    )
    print(f"     {len(verif_tools)} tool calls | {verif_in + verif_out:,} tokens | ${verif_cost:.5f}")

    step_results, fail_count, pass_count = parse_verifier_output(verif_text)
    total_tokens = exec_in + exec_out + verif_in + verif_out
    total_cost   = exec_cost + verif_cost
    print(f"     Verifier: {fail_count} fail(s), {pass_count} pass(es) | session ${total_cost:.5f}")

    return {
        "session": session_num,
        "executor": {
            "tool_calls":        exec_tools,
            "tool_counts":       _count_tools(exec_tools),
            "prompt_tokens":     exec_in,
            "completion_tokens": exec_out,
            "cost_usd":          exec_cost,
        },
        "verifier": {
            "tool_calls":        verif_tools,
            "tool_counts":       _count_tools(verif_tools),
            "prompt_tokens":     verif_in,
            "completion_tokens": verif_out,
            "cost_usd":          verif_cost,
            "step_results":      step_results,
            "fail_count":        fail_count,
            "pass_count":        pass_count,
            "output":            verif_text,
        },
        "all_passed":   fail_count == 0,
        "total_tokens": total_tokens,
        "cost_usd":     total_cost,
    }


def _count_tools(tool_log):
    counts = {}
    for tc in tool_log:
        counts[tc["tool"]] = counts.get(tc["tool"], 0) + 1
    return counts


# --- Report -------------------------------------------------------------------
def generate_report(results, today_yymmdd):
    n           = len(results)
    full_passes = sum(1 for r in results if r["all_passed"])
    avg_fails   = sum(r["verifier"]["fail_count"] for r in results) / n
    avg_tokens  = sum(r["total_tokens"] for r in results) / n
    total_cost  = sum(r["cost_usd"] for r in results)

    print(f"\n{'='*72}")
    print(f"  EVALUATION REPORT - Platform SOB Agent  ({n} sessions, model={MODEL})")
    print(f"{'='*72}")

    header = (f"{'Sess':>5}  {'Verif Fails':>11}  {'Exec Tools':>10}  "
              f"{'Verif Tools':>11}  {'Tokens':>8}  {'Cost':>8}  {'Pass?'}")
    print(f"\n{header}")
    print("-" * 72)
    for r in results:
        exec_total  = sum(r["executor"]["tool_counts"].values())
        verif_total = sum(r["verifier"]["tool_counts"].values())
        print(
            f"  {r['session']:>3}  "
            f"{r['verifier']['fail_count']:>11}  "
            f"{exec_total:>10}  "
            f"{verif_total:>11}  "
            f"{r['total_tokens']:>8,}  "
            f"${r['cost_usd']:>6.5f}  "
            f"{'  OK' if r['all_passed'] else '  XX'}"
        )
    print("-" * 72)
    print(f"{'AVG':>5}  {avg_fails:>11.1f}  {'':>10}  {'':>11}  "
          f"{avg_tokens:>8,.0f}  ${total_cost/n:>6.5f}")
    print(f"\n  Sessions with ALL checks passed : {full_passes}/{n}")
    print(f"  Total cost across all sessions  : ${total_cost:.5f}")

    # Tool usage breakdown
    print(f"\n  Tool usage - executor (across {n} sessions):")
    exec_agg = {}
    for r in results:
        for tool, cnt in r["executor"]["tool_counts"].items():
            exec_agg[tool] = exec_agg.get(tool, 0) + cnt
    for tool, cnt in sorted(exec_agg.items(), key=lambda x: -x[1]):
        print(f"    {tool:<35} {cnt:>4}x  (avg {cnt/n:.1f}/session)")

    print(f"\n  Tool usage - verifier (across {n} sessions):")
    verif_agg = {}
    for r in results:
        for tool, cnt in r["verifier"]["tool_counts"].items():
            verif_agg[tool] = verif_agg.get(tool, 0) + cnt
    for tool, cnt in sorted(verif_agg.items(), key=lambda x: -x[1]):
        print(f"    {tool:<35} {cnt:>4}x  (avg {cnt/n:.1f}/session)")

    # Failure breakdown by step
    fail_tally = {}
    for r in results:
        for step, status in r["verifier"]["step_results"].items():
            if status == "FAIL":
                fail_tally[step] = fail_tally.get(step, 0) + 1

    if fail_tally:
        print(f"\n  Verifier failure breakdown by step (out of {n} sessions):")
        for step, cnt in sorted(fail_tally.items(), key=lambda x: -x[1]):
            print(f"    {step:<42} {cnt}/{n}  ({cnt/n*100:.0f}%)")
    else:
        print(f"\n  No verifier failures across all {n} sessions.")

    report_path = SCRIPT_DIR / f"eval_runner_report_{today_yymmdd}.json"
    with open(report_path, "w") as f:
        json.dumps(results, f, indent=2, default=str)
    print(f"\n  Full report saved -> {report_path.name}")
    print(f"{'='*72}\n")


# ==============================================================================
#  FastAPI Server
# ==============================================================================

def _get_client():
    """Get authenticated OpenAI client (DeepSeek)."""
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        raise HTTPException(status_code=500, detail="DEEPSEEK_API_KEY not set")
    return OpenAI(api_key=api_key, base_url="https://api.deepseek.com")


app = FastAPI(
    title="Platform SOB Agent API",
    description="DeepSeek-powered multi-agent evaluation server for Platform SOB workflow",
    version="1.0.0",
)


@app.get("/")
async def root():
    return {
        "service": "Platform SOB Agent API",
        "model": MODEL,
        "docs": "/docs",
        "endpoints": {
            "GET  /health":    "Health check",
            "POST /agent/run": "Run a single agent (custom prompt)",
            "POST /evaluate":  "Run N-session evaluation",
            "GET  /results":   "Latest evaluation results",
        }
    }


@app.get("/health")
async def health():
    api_key = os.environ.get("DEEPSEEK_API_KEY")
    return {
        "status": "ok",
        "model": MODEL,
        "deepseek_api_key_set": api_key is not None,
        "uvicorn_version": getattr(uvicorn, "__version__", "unknown") if HAS_FASTAPI else "n/a",
        "timestamp": datetime.now().isoformat(),
    }


@app.post("/agent/run")
async def run_agent_endpoint(
    prompt: str,
    role: str = "executor",
    sandbox: bool = True,
    max_turns: int = 40,
):
    """
    Run a single agent with a custom prompt.

    - role: "executor" or "verifier" - determines which system spec to load
    - sandbox: if True, all writes are simulated (default: True)
    - max_turns: max tool-calling iterations (default: 40)
    """
    client = _get_client()
    system_prompt = EXECUTOR_SPEC if role == "executor" else VERIFIER_SPEC

    text, tool_log, prompt_tok, comp_tok, cost = run_agent(
        client, system_prompt, prompt, sandbox, label=role, max_turns=max_turns
    )

    return {
        "role": role,
        "sandbox": sandbox,
        "output": text,
        "tool_calls": len(tool_log),
        "tool_details": tool_log,
        "prompt_tokens": prompt_tok,
        "completion_tokens": comp_tok,
        "cost_usd": round(cost, 6),
    }


@app.post("/evaluate")
async def evaluate_endpoint(
    sessions: int = 10,
    live: bool = False,
):
    """
    Run the full N-session evaluation (executor + verifier per session).

    - sessions: number of evaluation sessions (default: 10)
    - live: if True, performs real sheet writes (default: False = sandbox)
    """
    client = _get_client()
    sandbox = not live
    today        = datetime.now().strftime("%y/%m/%d")
    today_yymmdd = datetime.now().strftime("%y%m%d")

    results = []
    for i in range(1, sessions + 1):
        result = run_session(client, i, sessions, sandbox, today, today_yymmdd)
        results.append(result)

    # Store in memory
    global _latest_results
    _latest_results = results

    n = len(results)
    full_passes = sum(1 for r in results if r["all_passed"])
    total_cost  = sum(r["cost_usd"] for r in results)

    return {
        "sessions_run": sessions,
        "mode": "sandbox" if sandbox else "live",
        "all_passed": full_passes == sessions,
        "sessions_passed": full_passes,
        "sessions_failed": sessions - full_passes,
        "total_cost_usd": round(total_cost, 6),
        "avg_cost_usd_per_session": round(total_cost / n, 6),
        "results": results,
    }


@app.get("/results")
async def get_results():
    """Get the latest evaluation results."""
    if not _latest_results:
        return {"results": [], "message": "No evaluation results yet. POST /evaluate first."}
    return {"results": _latest_results}


# ==============================================================================
#  CLI Main
# ==============================================================================

def cli_main():
    parser = argparse.ArgumentParser(description="Evaluate Platform SOB multi-agent workflow")
    parser.add_argument("--sessions", type=int, default=10, help="Number of sessions (default: 10)")
    parser.add_argument("--live",     action="store_true",  help="Live mode - real writes (default: sandbox)")
    args = parser.parse_args()

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

    sandbox      = not args.live
    today        = datetime.now().strftime("%y/%m/%d")
    today_yymmdd = datetime.now().strftime("%y%m%d")

    print("Platform SOB Evaluation")
    print(f"  Sessions : {args.sessions}")
    print(f"  Mode     : {'SANDBOX (no writes)' if sandbox else 'LIVE (real writes)'}")
    print(f"  Date     : {today}")
    print(f"  Model    : {MODEL}")

    client  = OpenAI(api_key=api_key, base_url="https://api.deepseek.com")
    results = []

    for i in range(1, args.sessions + 1):
        result = run_session(client, i, args.sessions, sandbox, today, today_yymmdd)
        results.append(result)

    generate_report(results, today_yymmdd)


# ==============================================================================
#  Entry point
# ==============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Platform SOB Agent - API server & evaluation runner"
    )
    parser.add_argument("--cli", action="store_true",
                        help="Run in CLI evaluation mode instead of API server")
    parser.add_argument("--port", type=int, default=API_PORT,
                        help=f"API server port (default: {API_PORT})")
    parser.add_argument("--host", type=str, default=API_HOST,
                        help=f"API server host (default: {API_HOST})")
    parser.add_argument("--sessions", type=int, default=10,
                        help="Number of sessions (only used with --cli)")
    parser.add_argument("--live", action="store_true",
                        help="Live mode for CLI (only used with --cli)")
    args, _ = parser.parse_known_args()

    if args.cli:
        # Delegate to CLI evaluation mode - parse remaining args properly
        sys.argv = [sys.argv[0], "--sessions", str(args.sessions)]
        if args.live:
            sys.argv.append("--live")
        return cli_main()

    # --- API Server mode (default) ---
    if not HAS_FASTAPI:
        print("ERROR: fastapi and uvicorn are required for API server mode.")
        print("       Install: pip3 install fastapi uvicorn")
        print("       Or run with --cli for the old CLI evaluation mode.")
        sys.exit(1)

    api_key = os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        print("WARNING: DEEPSEEK_API_KEY not set. API calls will fail until it's set.")

    print(f"Platform SOB Agent API server")
    print(f"  Host : http://{args.host}:{args.port}")
    print(f"  Model: {MODEL}")
    print(f"  Docs : http://{args.host}:{args.port}/docs")
    print(f"  CLI  : python3 {Path(__file__).name} --cli [--sessions 5] [--live]")
    print()

    uvicorn.run(app, host=args.host, port=args.port)


if __name__ == "__main__":
    main()
