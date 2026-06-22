"""
sandbox.py — 沙箱操作
Intercepts all sheet write operations when running in sandbox (dry-run) mode.
No changes are made to any spreadsheet. All intended actions are logged to
sandbox_run_log.md for review.

Usage:
    python agent_runner.py --sandbox

All write calls should route through SandboxGuard before executing.
"""

from datetime import datetime, timezone, timedelta
from pathlib import Path

# ── Config ──────────────────────────────────────────────────────────────────
AGENT_DIR = Path(__file__).parent.parent
SANDBOX_LOG = AGENT_DIR / "sandbox_run_log.md"
SHANGHAI_TZ = timezone(timedelta(hours=8))


class SandboxGuard:
    """
    Wraps every sheet write operation.
    - In sandbox mode: logs the intended action, returns a mock success.
    - In live mode:    executes the real function and logs the result.

    Usage:
        guard = SandboxGuard(sandbox=args.sandbox)

        # Instead of calling the API directly:
        guard.write(
            description="Paste SOB values rows 98:130 → 137:169",
            fn=lambda: gsi_update_range(spreadsheet_id, range_, values)
        )
    """

    def __init__(self, sandbox: bool = False):
        self.sandbox = sandbox
        self.actions: list[dict] = []
        self._run_time = datetime.now(SHANGHAI_TZ)

        if sandbox:
            print("[sandbox] 🔒 SANDBOX MODE ACTIVE — no writes will be made to any sheet.")
            self._init_log()

    def _init_log(self):
        ts = self._run_time.strftime("%Y-%m-%d %H:%M")
        header = (
            f"\n## Sandbox Run: {ts} (Asia/Shanghai)\n\n"
            "All actions below are **DRY RUN** — no spreadsheet was modified.\n\n"
            "| # | Step | Action | Status |\n"
            "|---|---|---|---|\n"
        )
        if not SANDBOX_LOG.exists():
            SANDBOX_LOG.write_text(
                "# Platform SOB Agent — Sandbox Run Log\n\n"
                "Each section is one dry-run execution.\n",
                encoding="utf-8"
            )
        with open(SANDBOX_LOG, "a", encoding="utf-8") as f:
            f.write(header)

    def write(self, step: str, description: str, fn=None, *args, **kwargs):
        """
        Execute a write operation (or mock it in sandbox mode).

        Args:
            step:        Step name, e.g. "Step 2" or "PC2-1"
            description: Human-readable description of what will be written.
            fn:          The actual write function to call in live mode.
                         Must accept no additional args (use a lambda).
        Returns:
            Result of fn() in live mode, or {"dry_run": True} in sandbox mode.
        """
        action_num = len(self.actions) + 1
        action = {"num": action_num, "step": step, "description": description}

        if self.sandbox:
            status = "⏭️ DRY RUN"
            self.actions.append({**action, "status": status})
            self._append_log_row(action_num, step, description, status)
            print(f"[sandbox] [{status}] {step}: {description}")
            return {"dry_run": True}
        else:
            try:
                result = fn() if fn else None
                status = "✅ Done"
                self.actions.append({**action, "status": status})
                print(f"[sandbox] [{status}] {step}: {description}")
                return result
            except Exception as e:
                status = f"❌ Error: {e}"
                self.actions.append({**action, "status": status})
                print(f"[sandbox] [{status}] {step}: {description}")
                raise

    def duplicate_tab(self, step: str, workbook: str, source_tab: str, new_tab: str, fn=None):
        """Convenience wrapper specifically for tab duplication actions."""
        description = (
            f"Duplicate tab `{source_tab}` in `[{workbook}]` → new tab `{new_tab}`"
        )
        return self.write(step=step, description=description, fn=fn)

    def paste_values(self, step: str, source: str, destination: str, fn=None):
        """Convenience wrapper for paste-as-values actions."""
        description = f"Paste values from `{source}` → `{destination}` (values only)"
        return self.write(step=step, description=description, fn=fn)

    def _append_log_row(self, num: int, step: str, description: str, status: str):
        row = f"| {num} | {step} | {description} | {status} |\n"
        with open(SANDBOX_LOG, "a", encoding="utf-8") as f:
            f.write(row)

    def finalize(self):
        """Write summary footer to sandbox log."""
        if not self.sandbox:
            return
        total = len(self.actions)
        dry_runs = sum(1 for a in self.actions if "DRY RUN" in a["status"])
        errors = sum(1 for a in self.actions if "Error" in a["status"])
        footer = (
            f"\n**Summary:** {total} actions logged — "
            f"{dry_runs} dry-run, {errors} errors.\n\n---\n"
        )
        with open(SANDBOX_LOG, "a", encoding="utf-8") as f:
            f.write(footer)
        print(f"[sandbox] Sandbox run complete. Log: {SANDBOX_LOG}")


# ── CLI test ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    guard = SandboxGuard(sandbox=True)

    guard.duplicate_tab(
        step="Step 2",
        workbook="Reg CNLS copy",
        source_tab="26/05/07 ADG ADO",
        new_tab="26/05/14 ADG ADO"
    )
    guard.paste_values(
        step="Step 2",
        source="[Weekly Live View] SHP/TTS ADG ADO B5:Z62 (red section)",
        destination="[Reg CNLS copy] '26/05/14 ADG ADO' B5:Z62"
    )
    guard.write(
        step="Step 3",
        description="Paste rows 98:130 → 137:169 in [Reg CNLS copy] '26/05/14 ADG ADO' (values only)"
    )
    guard.write(
        step="Step 4",
        description="Read rows 199:223 in [Reg CNLS copy] — validate all = 0"
    )
    guard.finalize()
    print("Sandbox test complete. Check sandbox_run_log.md.")
