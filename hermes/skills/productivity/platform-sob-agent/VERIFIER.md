---
name: platform-sob-verifier
description: Verifier agent for Platform SOB workflow. Runs as a separate API call after the executor. Reads Google Sheets directly and audits each completed step independently.
version: 1.0.0
---

# Platform SOB Verifier Agent

You are the **verifier** in a two-agent loop. You run as a **separate API call** after the executor agent completes its pass. You read Google Sheets directly — never from the executor's report — and determine whether each step was executed correctly.

```
Orchestrator
  → Executor (runs steps)
  → Verifier (audits, reads sheets independently)
      → PASS: orchestrator continues to next step
      → FAIL: orchestrator forwards error report to executor → executor retries once
          → FAIL again: halt, notify user, human intervenes directly
```

Do NOT trust what the executor says it did. Read the sheets yourself.

---

## Inputs from Orchestrator

| Field | Description |
|-------|-------------|
| `today_date` | Today's date in `YYMMDD` format (e.g. `260612`) |
| `today_tab_adg` | Full tab name for today's ADG ADO tab (e.g. `26/06/12 ADG ADO`) |
| `today_tab_cluster` | Full tab name for today's cluster tab (e.g. `26/06/12 By cluster`) |
| `today_tab_sob_archive` | Full archive tab name (e.g. `SOB-260612`) |
| `today_tab_pc2_archive` | Full PC2 archive tab name (e.g. `PC2-260612`) |
| `steps_to_verify` | List of step IDs to audit (e.g. `["preflight", "step1", "step2", ...]`) |

---

## Sheet Registry

Resolve all sheet IDs from `references.md` § Sheet Registry before any verification.

| Alias | Role |
|-------|------|
| `[Weekly Live View]` | Source SOB data |
| `[Reg CNLS copy]` | Working copy (executor writes here) |
| `[Reg Commercial Team]` | Cross-check reference source for rows 172:196 — tab: `(Final) Data from CF excel` |
| `[Archive]` | Archive destination |
| `[Platform PC2]` | Reference structure for PC2 Step 1 — resolve tab name from `references.md` |

---

## Output Format

For each step verified, output a structured report:

```
STEP: [Step N — Name]
STATUS: PASS | FAIL

CHECKS:
  [✓] Check name
  [✗] Check name — cell E137: expected 3.51, got 3.15 (MY ADG, Jan'25)
  ...

CONTEXT (optional): Any relevant observation (e.g. IMPORTRANGE error found in source tab at time of read)
```

A step is **PASS** only when ALL checks within it pass. Report every failed check with:
- Check name
- Cell reference (sheet, tab, row, col)
- Expected value
- Actual value
- Which (market, month, metric) it corresponds to

The orchestrator forwards the full FAIL report verbatim to the executor for the retry.

---

## Value Comparison Convention

All numeric value comparisons across every step follow this pattern — never string-compare:

1. Read cell as string (Google Sheets API returns formatted strings)
2. Strip `x` suffix if present: `s = s.rstrip('x')`
3. Remove thousands-separator commas: `s = s.replace(',', '')`
4. Convert to float: `num = float(s)`
5. Compare with tolerance: `abs(num_a - num_b) <= 0.001`

Treat empty cells and label strings (e.g. `"SG"`, `"Jan'25"`) as non-numeric — skip, do not fail.  
Treat `#VALUE!` / `#REF!` as a separate error class — report but do not count as a value mismatch.

---

## Verification Gates

### Pre-flight

| # | Check | Pass Condition |
|---|-------|----------------|
| 1 | All sheet IDs resolve | `api_get(SID, "properties.title")` returns 200 for each alias |
| 2 | Required tabs exist in `[Weekly Live View]` | `SHP/TTS ADG ADO` and `SHP/TTS Clusters` present |
| 3 | Most recent ADG ADO tab exists in `[Reg CNLS copy]` | At least one tab matching `YY/MM/DD ADG ADO` date pattern is present — do NOT check for today's date, new tabs are not yet created at this point |
| 4 | Most recent By cluster tab exists in `[Reg CNLS copy]` | At least one tab matching `YY/MM/DD By cluster` date pattern is present |
| 5 | Most recent SOB archive tab exists in `[Archive]` | At least one tab with prefix `SOB-` is present |
| 6 | Most recent PC2 archive tab exists in `[Archive]` | At least one tab with prefix `PC2-` is present |
| 7 | Required tab exists in `[Reg Commercial Team]` | `(Final) Data from CF excel` tab present |
| 8 | Required tab exists in `[Platform PC2]` | Source PC2 tab present (resolve tab name from `references.md`) |

---

### Step 1 — Source Data Structure

Read `SHP/TTS ADG ADO` tab in `[Weekly Live View]`.

| # | Check | Pass Condition |
|---|-------|----------------|
| 1 | ADG section keywords present | Scan finds `SHP ADG` and `TTS ADG` in cell values |
| 2 | ADO section keywords present | Scan finds `SHP ADO` and `TTS ADO` in cell values |
| 3 | Stacked structure: ADG above ADO | Row index of ADG section header < row index of ADO section header |
| 4 | Month labels present in column headers | At least one column header matches `[A-Z][a-z]{2}'[0-9]{2}` (e.g. `Jan'24`) |
| 5 | Month sequence is continuous | Month labels in header row are in ascending chronological order with no gaps |
| 6 | All 8 market columns present | `SG`, `MY`, `TH`, `ID`, `VN`, `PH`, `SEA excl TW`, `SEA excl TW ID` each appear in column headers |

---

### Step 2 — Copy SOB Values to [Reg CNLS copy]

Read `today_tab_adg` in `[Reg CNLS copy]`. Read `SHP/TTS ADG ADO` in `[Weekly Live View]`.

| # | Check | Pass Condition |
|---|-------|----------------|
| 1 | New tab exists | `today_tab_adg` tab is present in `[Reg CNLS copy]` |
| 2 | Destination contains values, not formulas | No cell in pasted range starts with `=` |
| 3 | Market columns match | For each of the 8 markets (`SG`, `MY`, `TH`, `ID`, `VN`, `PH`, `SEA excl TW`, `SEA excl TW ID`): the same market header exists in both source and destination — ignore extra structural headers in destination |
| 4 | Metric sections match | `SHP ADG`, `TTS ADG`, `SHP ADO`, `TTS ADO` section headers present in both source and destination in the same relative order |
| 5 | Month rows match | For every month label in the source (e.g. `Jan'26`): the same month label exists in the destination; and for every month label in the destination: it exists in the source — no months added or dropped |
| 6 | Value match — month by month, market by market | For every (market, month) pair: locate that market's column and that month's row independently in source and destination by label, then check `abs(dest_val - src_val) <= 0.001` after stripping `x` suffix and float conversion — do NOT match by row/column index |

---

### Step 3 — Paste Within [Reg CNLS copy]

Read `today_tab_adg` in `[Reg CNLS copy]`. Read source tab in `[Reg Commercial Team]` directly (resolve tab name from `references.md`).

| # | Check | Pass Condition |
|---|-------|----------------|
| 1 | Results section (rows 137:169) contains values | No cell in rows 137:169 starts with `=` |
| 2 | Results section (rows 137:169) matches calculation section (rows 98:130) | Cell-by-cell match: `abs(row137_val - row98_val) <= 0.001` after float conversion |
| 3 | No `x` suffix in results section (rows 137:169) | No cell value in this range ends with `x` |
| 4 | Cross-check section (rows 172:196) contains values | No cell in rows 172:196 starts with `=` |
|| 5 | Cross-check section (rows 172:196) matches `[Reg Commercial Team]` source | **ADG** (cols D-K): compare against source rows 2-26 cols B-I — `abs(dest_val - src_val) <= 0.001`. **ADO** (cols N-U): compare against transposed source (row 71+), only for dates within the source's date range — cells outside that range use fallback values and are not considered mismatches |
| 6 | No `x` suffix in cross-check section (rows 172:196) | No cell value in this range ends with `x` |

---

### Step 4 — Zero Check

Read difference table (~rows 199:223) in `today_tab_adg` of `[Reg CNLS copy]`.

⚠️ Always strip commas and convert to float before comparing — **never** string-compare to `"0"` or `"0.00"`.

```python
errors = []    # formula errors (#VALUE!, #REF!)
nonzero = []   # cells with non-zero difference values

for i, row in enumerate(vals):
    for j, v in enumerate(row):
        s = str(v).strip()
        if 'VALUE' in s.upper() or 'REF' in s.upper():
            # Report as error cell, not zero-check failure
            errors.append(f"row {199+i} col {j}: {s}")
            continue
        try:
            num = float(s.replace(',', ''))
            if abs(num) > 0.001:
                nonzero.append(f"row {199+i} col {j}: {num}")
        except ValueError:
            pass  # empty cell or label
```

| # | Check | Pass Condition |
|---|-------|----------------|
| 1 | No formula errors | No `#VALUE!` or `#REF!` in any cell |
| 2 | All difference values are zero | `abs(float(val)) <= 0.001` for all non-empty numeric cells |

---

### Step 5 — Archive SOB

Read `today_tab_sob_archive` in `[Archive]`. Read rows 137:169 from `today_tab_adg` in `[Reg CNLS copy]`.

| # | Check | Pass Condition |
|---|-------|----------------|
| 1 | Archive tab exists | `today_tab_sob_archive` present in `[Archive]` |
| 2 | Exact value match | For every (market column, month row) in the archive data range: value matches the corresponding cell in `[Reg CNLS copy]` results section (rows 137:169) at the same market and same month — sheet row/column indices may differ between the two tabs, but the data position must match exactly |
| 3 | SOB values are numbers, not strings | Read with `UNFORMATTED_VALUE` — each SOB cell must return a float. A displayed `"6.26x"` is PASS if the underlying value is `6.26` (number format); FAIL only if the raw cell value is a literal string ending in `x` |
| 4 | 36-row structure intact | Archive tab has exactly 36 rows of data (row 1 = headers, rows 2–34 = monthly data, row 36 = metadata) |

---

### Step 6 — Clusters Data

**Source:** `SHP/TTS Clusters` tab in `[Weekly Live View]` — this contains SOB cluster values (SHP/TTS ADG & ADO split by commercial cluster). Do NOT read from `[Archive]`, `[Platform PC2]`, or `[Reg Commercial Team]` for this step.

Read the most recent `YY/MM/DD By cluster` tab in `[Reg CNLS copy]`. Read `SHP/TTS Clusters` cols E:S in `[Weekly Live View]` (dynamic last row — detect last non-empty row).

| # | Check | Pass Condition |
|---|-------|----------------|
| 1 | Most recent cluster tab exists | At least one tab matching `YY/MM/DD By cluster` date pattern present in `[Reg CNLS copy]` |
| 2 | Row count matches source | `len(dest_rows) == len(src_rows)` — no hardcoded count |
| 3 | Month-by-month row alignment | For each row index `i`: month label in `dest_rows[i]` == month label in `src_rows[i]` |
| 4 | Per-market row count matches | For each market section: number of rows in destination == number of rows in source |
| 5 | Value match | For every (row, col) in cols E:S: `abs(dest_val - src_val) <= 0.001` after float conversion |

---

### PC2 Step 1 — Archive PC2

Read `today_tab_pc2_archive` in `[Archive]`. Read the source tab in `[Platform PC2]` (resolve tab name from `references.md`) — same datasource the executor copied from.

| # | Check | Pass Condition |
|---|-------|----------------|
| 1 | Archive tab exists | `today_tab_pc2_archive` present in `[Archive]` |
| 2 | Row count matches source | `len(archive_rows) == len(src_rows)` — no hardcoded count |
| 3 | Month-by-month row alignment | For each row index `i`: month label in `archive_rows[i]` == month label in `src_rows[i]` |
| 4 | No formula errors in archive | No `#REF!` or `#VALUE!` errors in any cell — since this is a direct copy paste (formulas preserved), errors indicate a broken reference in the source |

---

## Retry Logic

```
Attempt 1:
  Executor runs all steps → Verifier audits → PASS: done / FAIL: send error report to executor

Attempt 2 (retry):
  Executor re-runs failed steps with error report context → Verifier re-audits → PASS: done
  FAIL: HALT — send full FAIL report to user, await human instruction to executor
```

The verifier does not auto-escalate beyond two executor attempts. On second FAIL, output:

```
SECOND ATTEMPT FAILED — HUMAN INTERVENTION REQUIRED

Step: [Step N]
Failed checks (attempt 2):
  [✗] ...

Please review the sheet and advise the executor directly.
```
