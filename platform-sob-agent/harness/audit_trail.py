"""
audit_trail.py — 工作留痕
Appends a structured run entry to audit_log.md after every agent execution.
"""

import json
import uuid
from datetime import datetime, timezone, timedelta
from pathlib import Path

# ── Config ──────────────────────────────────────────────────────────────────
AGENT_DIR = Path(__file__).parent.parent          # platform-sob-agent/
AUDIT_LOG = AGENT_DIR / "audit_log.md"
SHANGHAI_TZ = timezone(timedelta(hours=8))

ALL_STEPS = ["preflight", "1", "2", "3", "4", "5", "6", "PC2-1"]


# ── Run Record ───────────────────────────────────────────────────────────────
class RunRecord:
    """Holds the state of a single agent run. Pass this object between steps."""

    def __init__(self):
        self.run_id: str = str(uuid.uuid4())[:8]
        self.trigger_time: datetime = datetime.now(SHANGHAI_TZ)
        self.steps_completed: list[str] = []
        self.steps_failed: list[str] = []
        self.tabs_created: list[str] = []
        self.cells_written: list[str] = []
        self.validation_results: dict = {
            "step4_zero_check": None,       # True / False / None (not reached)
            "step5_value_match": None,
        }
        self.halt_reason: str | None = None
        self.notes: list[str] = []
        self._start_time = datetime.now(SHANGHAI_TZ)

    # ── Step tracking ─────────────────────────────────────────────────────
    def complete(self, step: str):
        self.steps_completed.append(step)

    def fail(self, step: str, reason: str):
        self.steps_failed.append(step)
        self.halt_reason = reason

    def add_tab(self, workbook: str, tab_name: str):
        self.tabs_created.append(f"`[{workbook}]`: `{tab_name}`")

    def add_cells(self, ref: str):
        self.cells_written.append(ref)

    def add_note(self, note: str):
        self.notes.append(note)

    # ── Status ────────────────────────────────────────────────────────────
    @property
    def status(self) -> str:
        if self.steps_failed:
            if self.steps_completed:
                return "⚠️ Partial"
            return "❌ Failed"
        if set(ALL_STEPS).issubset(set(self.steps_completed)):
            return "✅ Success"
        return "⚠️ Partial"

    @property
    def duration_seconds(self) -> int:
        return int((datetime.now(SHANGHAI_TZ) - self._start_time).total_seconds())


# ── Append to audit_log.md ───────────────────────────────────────────────────
def write_audit_entry(record: RunRecord):
    """Append one markdown run entry to audit_log.md."""
    ts = record.trigger_time
    date_str = ts.strftime("%Y-%m-%d (%A)")
    run_date = ts.strftime("%Y-%m-%d")

    # Build tab list
    tabs_md = "\n".join(f"  - {t}" for t in record.tabs_created) or "  - (none)"

    # Build validation section
    def val(v):
        if v is True:   return "✅ Pass"
        if v is False:  return "❌ Fail"
        return "⏭️ Not reached"

    step4 = val(record.validation_results["step4_zero_check"])
    step5 = val(record.validation_results["step5_value_match"])

    # Build notes
    notes_md = "\n".join(f"  - {n}" for n in record.notes) if record.notes else "  - (none)"

    entry = f"""
## Run: {date_str} — Asia/Shanghai  *(ID: {record.run_id})*

- **Status:** {record.status}
- **Steps Completed:** {", ".join(record.steps_completed) or "(none)"}
- **Steps Failed:** {", ".join(record.steps_failed) or "(none)"}
- **Halt Reason:** {record.halt_reason or "(none)"}
- **Tabs Created:**
{tabs_md}
- **Validation Results:**
  - Step 4 Zero Check: {step4}
  - Step 5 Value Match: {step5}
- **Duration:** {record.duration_seconds} seconds
- **Notes:**
{notes_md}

---
"""

    # Create file with header if it doesn't exist
    if not AUDIT_LOG.exists():
        AUDIT_LOG.write_text(
            "# Platform SOB Agent — Audit Log\n\n"
            "One entry is appended after every Thursday run.\n\n---\n",
            encoding="utf-8"
        )

    with open(AUDIT_LOG, "a", encoding="utf-8") as f:
        f.write(entry)

    print(f"[audit_trail] Entry written to {AUDIT_LOG} (run {record.run_id})")
    return str(AUDIT_LOG)


# ── CLI (for manual inspection) ───────────────────────────────────────────────
if __name__ == "__main__":
    # Demo: create a dummy record and write it
    rec = RunRecord()
    rec.complete("preflight")
    rec.complete("1")
    rec.complete("2")
    rec.add_tab("Reg CNLS copy", "26/05/14 ADG ADO")
    rec.fail("3", "User rejected the paste plan at Step 3 approval prompt")
    rec.validation_results["step4_zero_check"] = None
    rec.add_note("Test run — not a real Thursday execution")
    write_audit_entry(rec)
    print("Demo audit entry written.")
