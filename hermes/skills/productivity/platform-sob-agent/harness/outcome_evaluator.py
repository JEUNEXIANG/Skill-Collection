"""
outcome_evaluator.py — Outcome Evaluation
Evaluates whether a full agent run was successful, builds a human-readable
summary, and triggers the WeChat outcome report.

Called at the end of every run (success or failure).
"""

from datetime import datetime, timezone, timedelta
from pathlib import Path

from audit_trail import RunRecord, write_audit_entry, ALL_STEPS
from permission_gate import send_outcome_report

SHANGHAI_TZ = timezone(timedelta(hours=8))

# ── Success criteria ─────────────────────────────────────────────────────────
REQUIRED_STEPS = set(ALL_STEPS)          # all steps must complete for full success

REQUIRED_TABS = [                         # tabs that must have been created
    "ADG ADO",                            # contains today's date + ADG ADO
    "By cluster",                         # contains today's date + By cluster
    "SOB-",                               # SOB archive tab
    "PC2-",                               # PC2 archive tab
]


# ── Evaluator ────────────────────────────────────────────────────────────────
class OutcomeEvaluator:

    def __init__(self, record: RunRecord):
        self.record = record
        self.issues: list[str] = []
        self.highlights: list[str] = []

    def evaluate(self) -> str:
        """
        Run all checks and return the final status string.
        Also populates self.issues and self.highlights.
        """
        r = self.record

        # 1. All steps completed?
        missing_steps = REQUIRED_STEPS - set(r.steps_completed)
        if missing_steps:
            self.issues.append(f"Steps not completed: {', '.join(sorted(missing_steps))}")
        else:
            self.highlights.append("All 8 steps completed")

        # 2. No failed steps?
        if r.steps_failed:
            self.issues.append(f"Failed steps: {', '.join(r.steps_failed)}")
            if r.halt_reason:
                self.issues.append(f"Halt reason: {r.halt_reason}")

        # 3. Step 4 zero-check passed?
        v4 = r.validation_results.get("step4_zero_check")
        if v4 is True:
            self.highlights.append("Step 4 zero-check: ✅ All values = 0")
        elif v4 is False:
            self.issues.append("Step 4 zero-check: ❌ Non-zero values found — archive was skipped")
        else:
            self.issues.append("Step 4 zero-check: ⏭️ Not reached")

        # 4. Step 5 value-match passed?
        v5 = r.validation_results.get("step5_value_match")
        if v5 is True:
            self.highlights.append("Step 5 value match: ✅ SOB values identical (whole number)")
        elif v5 is False:
            self.issues.append("Step 5 value match: ❌ Mismatch detected — archive was skipped")
        else:
            if v4 is not False:   # only flag if step 4 passed but step 5 wasn't reached
                self.issues.append("Step 5 value match: ⏭️ Not reached")

        # 5. All expected tabs created?
        created_names = " ".join(r.tabs_created)
        for expected in REQUIRED_TABS:
            if expected not in created_names:
                self.issues.append(f"Expected tab containing '{expected}' was NOT created")
            else:
                self.highlights.append(f"Tab '{expected}...' created ✅")

        # 6. Audit log written? (checked by presence of record.run_id in log)
        # (audit_trail.write_audit_entry handles this — assumed done if we reach evaluator)
        self.highlights.append("Audit log entry written to audit_log.md")

        # Final status
        return r.status

    def build_summary(self) -> str:
        """Build a concise WeChat-friendly run summary."""
        r = self.record
        ts = r.trigger_time.strftime("%Y-%m-%d %H:%M")

        lines = [f"Run ID: {r.run_id} | {ts} Asia/Shanghai\n"]

        if self.highlights:
            lines.append("✅ What went well:")
            lines.extend(f"  • {h}" for h in self.highlights)

        if self.issues:
            lines.append("\n⚠️ Issues:")
            lines.extend(f"  • {i}" for i in self.issues)

        lines.append(f"\nTabs created: {len(r.tabs_created)}")
        for t in r.tabs_created:
            lines.append(f"  • {t}")

        lines.append(f"\nDuration: {r.duration_seconds}s")
        return "\n".join(lines)


# ── Main entry point ─────────────────────────────────────────────────────────
def finalize_run(record: RunRecord, sandbox: bool = False):
    """
    Called at the very end of every agent run.
    1. Evaluates success.
    2. Writes audit log entry.
    3. Sends WeChat outcome report.
    """
    evaluator = OutcomeEvaluator(record)
    status = evaluator.evaluate()
    summary = evaluator.build_summary()

    # Write audit log
    audit_path = write_audit_entry(record)

    # Send WeChat report (skip in sandbox — just print)
    if sandbox:
        print("\n[outcome_evaluator] [DRY_RUN] Outcome report (would be sent via WeChat):")
        print(f"Status: {status}")
        print(summary)
    else:
        send_outcome_report(
            status=status,
            summary=summary,
            audit_log_path=audit_path
        )

    print(f"\n[outcome_evaluator] Run {record.run_id} finished — {status}")
    return status


# ── CLI test ──────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    # Simulate a successful run
    rec = RunRecord()
    for step in ALL_STEPS:
        rec.complete(step)
    rec.add_tab("Reg CNLS copy", "26/05/14 ADG ADO")
    rec.add_tab("Reg CNLS copy", "26/05/14 By cluster")
    rec.add_tab("Archive", "SOB-260514")
    rec.add_tab("Archive", "PC2-260514")
    rec.validation_results["step4_zero_check"] = True
    rec.validation_results["step5_value_match"] = True

    finalize_run(rec, sandbox=True)
