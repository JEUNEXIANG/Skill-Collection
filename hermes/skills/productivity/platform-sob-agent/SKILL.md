---
name: platform-sob-agent
description: Weekly Platform SOB & PC2 update workflow — reads SOB data from Google Sheets, cross-checks, validates, and archives. Runs every Thursday 13:00 Asia/Shanghai.
version: 1.2.0
author: Hermes Agent
tags: [google-sheets, sob, pc2, weekly, archive]
---

# Platform SOB Agent — Executor

Weekly workflow to update Platform Share-of-Business (SOB) data and PC2 data across multiple Google Sheets, validate cross-source consistency, and archive the results.

**Schedule:** Thursday 13:00 Asia/Shanghai (weekly)

**Prerequisite:** `g do you aoogle-sheets-intelligence` skill (loaded automatically). Google Sheets OAuth must be set up.

**Multi-agent context:** This agent runs as the **executor** in a two-agent loop. After each pass, the orchestrator invokes the verifier (`VERIFIER.md`) as a separate API call. If the verifier returns FAIL, the orchestrator sends the error report back here for one retry. If the second attempt also fails, execution halts and the user intervenes directly.

---

## Before Starting

**ALL steps require user approval before execution.** The permission model:
1. Read data and prepare intended action
2. Present clear summary to user (cells, ranges, tab names, values)
3. Wait for user confirmation before executing any write

If uncertainty cannot be resolved after 1 retry: halt, log the issue, and notify the user.

---

## Sheet Registry

The 4 Google Sheet IDs **must be filled in** before this agent can run. They are in:
`~/.hermes/skills/productivity/platform-sob-agent/references.md` § Sheet Registry

| Alias | Contents |
|-------|----------|
| `[Weekly Live View]` | Source SOB data — tabs: `SHP/TTS ADG ADO`, `SHP/TTS Clusters` |
| `[Reg Commercial Team]` | Reference SOB cross-check data — tab: `(Final) Data from CF excel` |
| `[Platform PC2]` | PC2 data — tab: `SHP & TTS PC2` |
| `[Reg CNLS copy]` | Working copy — tabs: `YY/MM/DD ADG ADO`, `YY/MM/DD By cluster` |
| `[Archive]` | Archive — tabs: `SOB-YYMMDD`, `PC2-YYMMDD` |

Set a shorthand for GSI:
```bash
GSI="python ~/.hermes/skills/productivity/google-sheets-intelligence/scripts/sheets_intelligence.py"
```

---

## Workflow Steps

### Pre-flight: Validate Source Tabs

1. Load `references.md` to resolve all sheet aliases → IDs
2. Confirm all required tabs exist and are accessible:
   - `$GSI structure SPREADSHEET_ID --sheet "TAB_NAME"`
3. Run `$GSI scan-sections` on each source tab to detect stacked table sections
4. Use the detected section metadata to understand what each table is about, including headers, labels, surrounding context, and likely business meaning
5. Match the user's table description or instruction against the detected section descriptions before selecting a source table
6. Present section map and intended table match to user for confirmation

### Step 1 — Read & Validate Source Data Structure

Read table structure of the source tabs. Use `$GSI scan-sections` and present the detected structure to the user. Validation has two purposes:
- Verify that the table layout is structurally usable before any copy/paste or archive action.
- Help the agent understand what each detected table represents, then choose the correct table based on the explicit matching guidance below.

**Matching Guidance for `[Weekly Live View]` Tables:**

1. **Tab: `SHP/TTS ADG ADO`** (Invariant identity / Raw ingredient table)
   - **Identity**: Contains platform-level (all-seller) ADG and ADO **absolute values** for SHP and TTS. It does NOT contain precomputed SOB ratios.
   - **How to recognize dynamically**:
     - **Stacked vertical sections**: An ADG section (always above) and an ADO section (always below), separated by a blank row.
     - **Section headers**: Keywords `"SHP ADG"`, `"TTS ADG"`, then later `"SHP ADO"`, `"TTS ADO"`.
     - **Column headers**: `"Period"` + market codes (`SG`, `MY`, `TH`, `ID`, `VN`, `PH`, `SEA excl TW`, `SEA excl TW ID`).
     - **Data rows**: Monthly periods from `Jan'24` onward.
     - **Sub-tables**: 3 side-by-side sub-tables per section: SHP absolutes (left), TTS absolutes (center), TTS Multiple (right).
   - **Selection**: Match when the user needs raw, platform-level ADG/ADO absolute values to compute Platform SOB. Do NOT match if you need pre-computed ratios, cluster-level breakdowns, or PC2 metrics.

2. **Tab: `SHP/TTS Clusters`**
   - **Identity**: Contains cluster/category-level SHP/TTS ADG & ADO absolute data split by market and commercial cluster (e.g., Overall, EL, LS, FMCG excl HB, Fashion, HB).
   - **Selection**: Match when the user asks for cluster, category, segment, or SOB by cluster data. Ensure detected headers/labels include cluster names and ADG/ADO values.

**Matching Guidance for Other Tables:**
- **`(Final) Data from CF excel`** in `[Reg Commercial Team]`: Match when looking for the final cross-check reference values.

### Step 2 — Copy SOB Values (ADG ADO) to [Reg CNLS copy]

1. From the source data tab (`[Weekly Live View] SHP/TTS ADG ADO` or fallback `Sheet6`/`Sheet5`), read the relevant ADG & ADO values.
2. In `[Reg CNLS copy]`, find the most recent tab (format: `YY/MM/DD ADG ADO`). See ⚠️ Date parsing below.
3. Duplicate that tab, name it `{YY}/{MM}/{DD} ADG ADO` with today's date.
4. In the new tab, use the **red-bordered section** (`effectiveFormat.borders`) as a positional reminder for the paste target range. If no red borders found, paste data from Jan 2026 onwards for all markets, matched by month label and market column name.
5. Paste source values **as values only**, aligned **month by month and market by market**: match each source row by its month label and each source column by its market name to the corresponding row/column in the destination. If the red section dimensions do not match the source layout, fall back to label-based matching — do not assume positional identity. Present the mapping to the user and wait for confirmation before pasting.

#### ⚠️ Step 2 Implementation Details

**The destination has structural gaps between data blocks.** After Dec'25, there are non-data rows (headers like "1st Copy", "2nd Copy") before Jan'26 starts. Do NOT assume contiguous rows:

```python
# WRONG — assumes continuous data
for i in range(4, 40):
    dst_row = dest_rows[i]  # fails at Dec'25/Jun'26 boundary

# RIGHT — skip non-data rows, match by month label
for row in dest_rows:
    label = str(row[1]).strip()
    if not label or any(kw in label for kw in ['Copy','Q1','Q2','Q3','Q4']):
        continue
    if label.startswith('Q'): continue
    if label in src_map:
        update_row(src_map[label])
```

**Use batchUpdate for non-contiguous rows.** Since the destination may have gaps, write individual rows via the Sheets API `batchUpdate` endpoint:

```python
from urllib.request import Request, urlopen
import json

batch_data = []
for row_idx, vals in matched_updates:
    batch_data.append({
        'range': f"'{tab}'!D{row_idx+1}:K{row_idx+1}",
        'majorDimension': 'ROWS',
        'values': [vals]
    })

url = f"https://sheets.googleapis.com/v4/spreadsheets/{SID}/values:batchUpdate"
body = {'valueInputOption': 'USER_ENTERED', 'data': batch_data}
req = Request(url, data=json.dumps(body).encode(), headers=headers)
req.add_header('Content-Type', 'application/json')
with urlopen(req, timeout=60, context=SSL_CTX) as resp:
    json.loads(resp.read())
```

**Source has SHP and TTS side by side; destination has them stacked.** The Weekly Live View source places SHP ADG (cols D-K) and TTS ADG (cols N-U) in the same rows. The destination separates them into two vertically stacked sections (SHP ADG at rows ~4-39, TTS ADG at rows ~44-79). When reading source values for the TTS section, read from the same source rows but different columns:

```python
# SHP ADG: source cols D-K (indices 3-10)
shp_adg = build_map(src_rows, 3)
paste_by_label(dest_shp_section, shp_adg)

# TTS ADG: source cols N-U (indices 13-20) — SAME rows, different cols
tts_adg = build_map(src_rows, 13)
paste_by_label(dest_tts_section, tts_adg)
```

**ADO sections use dates (col C) not month labels (col B).** The ADO sections in the destination have date strings like "2024-01-01" in col C instead of month labels in col B. Match by converting dates to month labels or match by date string directly.


### Step 3 — Paste SOB Values Within [Reg CNLS copy] Tab

In the newly duplicated tab:
1. Copy rows ~98:130 → rows ~137:169 (SOB calculation section → results section) as values only
2. Copy values from `[Reg Commercial Team]` tab → rows ~172:196 of the new tab (cross-check section)

Both require user approval before execution.

⚠️ **After pasting, strip "x" suffix from all values in rows 137:169 and 172:196** so the difference formulas can compute correctly. Write clean numbers (e.g. `3.51` not `"3.51x"`).

### Step 4 — Validate Zero Check

Read the difference table (~rows 199:223).

**IMPORTANT — Handle FORMATTED_VALUE correctly:**
`gsu.values_get()` returns **formatted string values** (e.g. `"0.00"`, `"-31.56"`, `"#VALUE!"`), not raw numbers. Always convert to float before comparing:

```python
r = gsu.values_get(SID_CNLS, "'26/06/03 ADG ADO'!D199:U223")
vals = r.get('values', [])

for i, row in enumerate(vals):
    for j, v in enumerate(row):
        s = str(v).strip()
        # Skip #VALUE! errors — flag separately
        if 'VALUE' in s.upper() or 'REF' in s.upper():
            print(f'  ERROR at row {199+i}, col {chr(68+j)}: {s}')
            continue
        # Convert formatted number string to float
        try:
            num = float(s.replace(',', ''))
            if abs(num) > 0.001:
                print(f'  NON-ZERO at row {199+i}, col {chr(68+j)}: {num}')
        except ValueError:
            pass  # empty cell or label
```

- ✅ All values = 0: notify user, proceed to Step 5
- ❌ Any non-zero value: **halt**, alert user with details of which cells differ
- ❌ Any `#VALUE!` or `#REF!` error: flag to user — likely caused by "n/a" in cross-check source or format mismatch

### Step 5 — Archive SOB to [Archive]

Conditions (both must be true):
- Step 4 zero-check passed

Action: duplicate most recent `SOB-YYMMDD` tab in `[Archive]`, name it `SOB-{YYMMDD}`. Copy the SOB results section (rows ~137:169) from the `{YY}/{MM}/{DD} ADG ADO` tab in `[Reg CNLS copy]` and paste as values into `SOB-{YYMMDD}`.

### Step 6 — Copy Clusters Data to [Reg CNLS copy]

Copy columns E:S (dynamic last row) from `SHP/TTS Clusters` in `[Weekly Live View]`. Duplicate most recent `By cluster` tab in `[Reg CNLS copy]`, then paste the copied values into the new tab as values only.

### PC2 Step 1 — Archive PC2 to [Archive]

Duplicate most recent `PC2-YYMMDD` tab in `[Archive]`, rename to `PC2-{YYMMDD}`. Copy all data from `SHP & TTS PC2` in `[Platform PC2]` and paste directly into the new tab (preserve formulas and formatting — not paste-as-values).

---

## Pitfalls & Learnings

### ⚠️ Tab Ordering: Newer = Lower Index (Counter-Intuitive)

In Google Sheets, when tabs are sorted with the newest first (chronologically descending), the **newest tab has the lowest index**. Index 0 is often a non-data tab like "Direction". The SOB tabs in the Archive sheet follow this pattern:
- `idx=1: SOB- As of 260518` (newest)
- `idx=2: SOB-260514`
- `idx=3: SOB-260506`

When duplicating a tab and inserting it in chronological order:
```python
# Find the newest tab (lowest index)
sob_tabs = [(t, i, idx) for t, i, idx in tabs if t.startswith("SOB")]
sob_tabs.sort(key=lambda x: x[2])  # ascending = newest first
newest_name, newest_id, newest_idx = sob_tabs[0]

# Insert right after it = same position (Google shifts existing tabs down)
insert_idx = newest_idx + 1
```

**⚠️ Do NOT use sheetId for ordering.** Google Sheets sheetIds are arbitrary numbers — they are NOT monotonic with tab position. Always sort by the tab's `index` field from `api_get()`, NOT by `sheetId`.

### ⚠️ Archive Tab Structure (SOB-YYMMDD)

The archive SOB tabs have a specific structure with 36 rows:
- **Row 1:** Headers: A=period label, B-J=market names (SG, MY, TH, ID, VN, PH, SEA excl TW, SEA excl TW ID, SEA excl TW SG)
- **Rows 2-34:** Monthly data (Dec'24 → Q4'26 Target), with period labels in column A and ADG SOB values in columns B-J
- **Row 36:** Metadata row with commercial update time and archive reference

When pasting updated values into the archive:
```python
# Read ADG values from Reg CNLS copy (cols D-L = 9 market columns)
adg_values = gsu.values_get(SID_CNLS, "'26/05/18 ADG ADO'!D137:L169")
# Paste into archive (cols B-J = same 9 markets, rows 2-34)
gsu.values_update(SID_ARCH, "'SOB-260518'!B2:J34", cleaned_values)
```

### ⚠️ Cross-Check Source Mapping

The cross-check section (rows 172:196) takes data from `(Final) Data from CF excel` in `[Reg Commercial Team]`:

| Source (RC) | Destination (Cross-check) | Content |
|---|---|---|
| Row 2-26, Col A | Col C (serial date) | Date |
| Row 2-26, Col B | Col D | SG ADG SOB value |
| Row 2-26, Col C | Col E | MY ADG SOB value |

The ADO part of the cross-check (cols N-U) comes from the same source but is transposed — see ⚠️ Cross-Check ADO Section below for how to handle it.

Convert date strings to serial numbers:
```python
import datetime
epoch = datetime.date(1899, 12, 30)
dt = datetime.datetime.strptime("2024-12-01", "%Y-%m-%d").date()
serial = (dt - epoch).days  # 45627
```

### ⚠️ Date Parsing: YY/MM/DD vs MM/DD

The tab names use two formats:
- **Pre-2026:** `MM/DD` format, e.g. `12/25 ADG ADO`, `9/19 By cluster`
- **2026+:** `YY/MM/DD` format, e.g. `26/05/14 ADG ADO`, `26/05/14 By cluster`

When finding the most recent tab, distinguish between:
- 3-part date `YY/MM/DD`: group1=year, group2=month, group3=day
- 2-part date `MM/DD` (no year): group1=month, group2=day

```python
import re
def parse_date(tab_name):
    m = re.search(r'(\d{1,2})/(\d{1,2})(?:/(\d{2,4}))? ADG ADO', tab_name)
    if m:
        if m.group(3):  # YY/MM/DD format
            y = "20" + m.group(1)
            return f"{y}{m.group(2).zfill(2)}{m.group(3).zfill(2)}"
        else:  # MM/DD format (pre-2026)
            return f"2025{m.group(1).zfill(2)}{m.group(2).zfill(2)}"
    return None
```

### ⚠️ IMPORTRANGE Tabs Are Often Inaccessible

The `[Weekly Live View] SHP/TTS ADG ADO` and `SHP/TTS Clusters` tabs often contain IMPORTRANGE formulas that haven't been authorized. They'll show `#REF!` instead of data.

**Fallback approach:**
- For ADG ADO data: Use `Sheet6` or `Sheet5` tab in the same spreadsheet (raw absolute values)
- For Clusters data: The latest By cluster tab in `[Reg CNLS copy]` already has the data from a previous run

### ⚠️ "Red Section" = Destination (Paste Target), Not Source

The "red section" (cells with **red borders** via `effectiveFormat.borders`) marks the **destination** where data gets **pasted into** in the `[Reg CNLS copy]` ADG ADO tab, NOT the source range in `[Weekly Live View]`.

- **DO NOT** search for red borders in the source tab to find what data to copy
- **DO** search for red borders in the destination tab to find the paste target range
- If no red borders found, ask the user what range to paste into
- Common fallback: rows ~137:169 in the ADG ADO tab of `[Reg CNLS copy]`

When in doubt, read both source and destination formats and ask the user which they mean.

### ⚠️ Tab Names May Have Trailing Spaces

The `[Weekly Live View] SHP/TTS Clusters` tab has a trailing space in its name: `'[Weekly Live View] SHP/TTS Clusters '`. Always URL-encode carefully:

```python
quoted = urllib.parse.quote("'[Weekly Live View] SHP/TTS Clusters '!A1", safe="'!")
```

### ⚠️ "x" Suffix on SOB Values — Cell Format vs Stored Value

SOB values often show with an "x" suffix (e.g. `"6.26x"`). This can come from two sources:

**Source 1: Text strings (actual "x" in the cell value)**
The value stored is literally the string `"6.26x"`. This is common when pasting from certain source systems.

```python
if isinstance(v, str) and v.endswith("x"):
    new_row.append(float(v[:-1]))
```

**Source 2: Custom number format (`0.00"x"`)**
The cell has a number format that appends "x" visually. The underlying value IS a number (6.26), but `FORMATTED_VALUE` render returns `"6.26x"`. Writing the number as `6.26` with `USER_ENTERED` still displays as `"6.26x"` because the format persists. To actually remove the "x" from display, you'd need to clear the cell format via `repeatCell` with `fields: "userEnteredFormat.numberFormat"`.

**Always clean both the results section and the Archive tab after pasting:**

```python
# Read as FORMATTED_VALUE (catches both text and format-driven x's)
# Strip x and store as number
# The Archive tab will still show x from format — that's OK for archive purposes
```

### ⚠️ "SOB by cluster" → "By cluster" Tab Mapping

The `SOB by cluster` tab in `[Reg Commercial Team]` has a different structure than the `By cluster` tab in `[Reg CNLS copy]`:

| Aspect | Source (SOB by cluster) | Target (YY/MM/DD By cluster) |
|--------|------------------------|------------------------------|
| ADG values | D-I (6 cols: Overall, EL, LS, FMCG excl HB, Fashion, HB) | F-K (same 6 categories) |
| ADO/2nd values | N-S (6 cols: same categories) | N-S (6 cols: same categories) |
| Data type | SHP/TTS ratio (decimal ~0.5-9.99) | Absolute values (large numbers) |
| Rows per market | 25 rows (months Dec'24-Dec'26) | 48 rows (months + quarters + projections) |
| Row structure | A=market, B=ADG/ADO, C=date, D-I=values, K-L=repeat, N-S=values | A=market, B=date, C=ADG, D=Actuals/Target, E=period, F-K=values, N-S=values |

**Critical: Match by row position, not by absolute sheet rows. This took 3 iterations to get right — trial and error proved that blind range pasting destroys the target structure.**

Source data rows (0-indexed) — within the source array from `SOB by cluster!A1:S200`:
- SG: rows 3-27 (A1=row4-row28)
- MY: rows 31-55 (row32-row56)  
- TH: rows 59-83 (row60-row84)
- VN: rows 87-111 (row88-row112)

Target data rows (1-indexed, for direct API range reference):
- SG: rows 6-30 (0-indexed 5-29)
- MY: rows 57-81
- TH: rows 108-132
- VN: rows 159-183

**Pasting as clean numbers:**
```python
def clean_val(v):
    """Strip x suffix, convert to float. Returns number or original value."""
    if v:
        s = str(v)
        if s.endswith("x"):
            try: return float(s[:-1])
            except: pass
        try: return float(s)
        except: pass
    return v

# For each market, paste ADG (D-I) → F-K and ADO (N-S) → N-S
market_info = {
    "SG": {"src_start": 3, "tgt_start": 6},   # 1-indexed target row
    "MY": {"src_start": 31, "tgt_start": 57},
    "TH": {"src_start": 59, "tgt_start": 108},
    "VN": {"src_start": 87, "tgt_start": 159},
}

for mkt, info in market_info.items():
    adg_rows = []
    ado_rows = []
    for i in range(info["src_start"], info["src_start"] + 25):
        r = src_vals[i]
        adg = [clean_val(r[j]) for j in range(3, 9)]   # D-I
        ado = [clean_val(r[j]) for j in range(13, 19)]  # N-S
        adg_rows.append(adg)
        ado_rows.append(ado)
    
    tgt = info["tgt_start"]
    write_range(sid, f"'Tab'!F{tgt}:K{tgt+24}", adg_rows)
    write_range(sid, f"'Tab'!N{tgt}:S{tgt+24}", ado_rows)
```

**Copy approach:** Match market by market name, then copy by row position (first 25 source rows → first 25 data rows of each market section in target). Source col D-I → target F-K, source col N-S → target N-S.

### ⚠️ `As of...` Tab Names Are Not Standard

The Reg CNLS copy may have tabs named `As of 26/05/18 ADG ADO` which use a different naming format than the standard `YY/MM/DD ADG ADO`. When checking if a tab for today's date already exists, only check the standard format — `As of...` tabs are a different context and should not block creation of the standard-format tab.

### ⚠️ Non-Zero Difference Override

If Step 4 validation fails (non-zero differences found), the spec says to halt. However, the user may explicitly override this and proceed to Step 5 anyway. In that case, re-paste the current results values into the Archive tab (which was already created in Step 5 but now needs updated data).

## Supporting Files

- `agent.md` — Full workflow specification (original source)
- `agent.md.backup-20260518` — Previous version of agent.md
- `references.md` — Business concepts, sheet registry, tab/row references
- `harness.md` — Infrastructure documentation (audit trail, permission gate, sandbox)
- `harness/` — Python modules:
  - `audit_trail.py` — RunRecord class + audit_log.md writer
  - `permission_gate.py` — WeChat notification + approval gate
  - `sandbox.py` — SandboxGuard for dry-run mode
  - `outcome_evaluator.py` — Final run evaluation + outcome report
- `wechat_config.json` — WeChat delivery configuration (currently: terminal mode)
- `audit_log.md` — Run history (one entry per Thursday)

## Key API Patterns

### ⚡ Preferred: Use `gsheets_util.py` (handles all macOS/SSL/auth issues)

This skill directory includes `gsheets_util.py` — a reusable utility that handles:
- **macOS SSL cert fix** (uses `certifi` — required for Python 3.x on macOS)
- **Auto token refresh** (direct `urllib` call, bypasses hanging `google-auth` library)
- **Clean wrappers** for all common Google Sheets API operations

Import and use it for ALL API calls in this workflow:

```python
import sys; sys.path.insert(0, "/Users/apple/.hermes/skills/productivity/platform-sob-agent")
import gsheets_util as gsu

# Read sheet metadata
data = gsu.api_get(SPREADSHEET_ID, "properties.title,sheets.properties")

# Get all values
rows = gsu.values_get(SID, "'Tab Name'!A1:Z100")

# Write values
gsu.values_update(SID, "'Tab Name'!A1:C3", [[1,2,3],[4,5,6]])

# Batch update (for tab duplication, formatting changes)
gsu.api_batch_update(SID, [{"duplicateSheet": {...}}])
```

The utility auto-refreshes the OAuth token when it's within 2 minutes of expiry.

### Reading red-bordered section (Step 2 — locating the paste target range)

The "red section" refers to cells with **red borders** in the **destination tab** (`[Reg CNLS copy]` ADG ADO tab), not in the source. The red borders mark where the pasted values should go. Use `gsheets_util.py` to query `effectiveFormat.borders` in the destination sheet:

```python
import sys; sys.path.insert(0, "/Users/apple/.hermes/skills/productivity/platform-sob-agent")
import gsheets_util as gsu

# Query the DESTINATION tab, not the source
SID = gsu.SHEETS["Reg CNLS copy"]
tab_name = "26/05/18 ADG ADO"  # Replace with actual target tab name

# Use the direct API with includeGridData
headers = gsu.get_auth_headers()
import urllib.request, urllib.parse, json
quoted = urllib.parse.quote(f"'{tab_name}'", safe="'")
url = f"https://sheets.googleapis.com/v4/spreadsheets/{SID}?ranges={quoted}&fields=sheets.data.rowData.values(effectiveFormat.borders,effectiveValue,formattedValue)"
req = urllib.request.Request(url, headers=headers)
with urllib.request.urlopen(req, timeout=30, context=gsu.SSL_CTX) as resp:
    data = json.loads(resp.read())
```

**Finding the red border range:** For each cell, check if any border side has a red-ish color. A common red threshold is `red > 0.8` and `green < 0.3` and `blue < 0.3`:

```python
def is_red_border(borders):
    if not borders:
        return False
    for side in ('top', 'bottom', 'left', 'right'):
        side_data = borders.get(side, {})
        if not side_data:
            continue
        color = side_data.get('color', side_data.get('colorStyle', {}).get('rgbColor', {}))
        if color.get('red', 0) > 0.8 and color.get('green', 0) < 0.3 and color.get('blue', 0) < 0.3:
            return True
    return False
```

Then scan all rows and build the contiguous range of rows/columns with red borders.

### Duplicating a tab

```python
import sys; sys.path.insert(0, "/Users/apple/.hermes/skills/productivity/platform-sob-agent")
import gsheets_util as gsu

SID = gsu.SHEETS["Reg CNLS copy"]

# Get all tabs to find source ID and insert position
spreadsheet = gsu.api_get(SID, "sheets.properties")

tabs = [(s["properties"]["title"], s["properties"]["sheetId"]) for s in spreadsheet["sheets"]]

# Find source tab by name
source_name = "26/05/14 ADG ADO"  # most recent tab to duplicate
source_id = None
for t, sid_ in tabs:
    if t == source_name:
        source_id = sid_
        break

# Find insert index (after the source tab, for chronological order)
insert_index = next(i+1 for i, (t,_) in enumerate(tabs) if t == source_name)

# Duplicate
gsu.api_batch_update(SID, [{
    "duplicateSheet": {
        "sourceSheetId": source_id,
        "insertSheetIndex": insert_index,
        "newSheetName": "26/05/18 ADG ADO"
    }
}])
```

### Writing values as values-only (paste-as-value)

```python
import sys; sys.path.insert(0, "/Users/apple/.hermes/skills/productivity/platform-sob-agent")
import gsheets_util as gsu

gsu.values_update(
    SID,
    "'Tab Name'!A1:Z100",
    [[1, 2, 3], [4, 5, 6]]
)
```

Note: `valueInputOption="USER_ENTERED"` by default — handles percentages, dates, and formulas (if value starts with `=`).

### Reading values (including formulas)

```python
import sys; sys.path.insert(0, "/Users/apple/.hermes/skills/productivity/platform-sob-agent")
import gsheets_util as gsu

# Values only (default):
result = gsu.values_get(SID, "'Tab Name'!A1:Z100")
rows = result.get("values", [])

# With formulas — use raw API:
import urllib.request, urllib.parse, json
headers = gsu.get_auth_headers()
quoted = urllib.parse.quote("'Tab Name'!A1:Z100", safe="'!")
url = f"https://sheets.googleapis.com/v4/spreadsheets/{SID}/values/{quoted}?valueRenderOption=FORMULA"
req = urllib.request.Request(url, headers=headers)
with urllib.request.urlopen(req, timeout=30, context=gsu.SSL_CTX) as resp:
    data = json.loads(resp.read())
formula_rows = data.get("values", [])

---

### ⚠️ Table Description Format: Crystallized, Not Coordinate-Based

Table descriptions in `references.md` should follow the **crystallized format** — organized around what doesn't change, not fixed row/column numbers.

**Bad (coordinate-based):** "ADG in rows 2-39 cols C-K, ADO in rows 41-77 cols C-K"

**Good (crystallized):**
- "Invariant identity" — what the table IS (e.g. raw ADG/ADO absolutes, not precomputed SOB)
- "How to recognize it" — keywords to scan for, structural patterns (stacked? horizontal?), column sequences that repeat
- What to score/match against and what to exclude
- "Structure detected dynamically — never assume fixed row/column numbers"

This format was developed iteratively across multiple corrections from the user:
1. First attempt used row/column numbers → user pointed out I missed the ADO section below
2. Added more coordinates → user asked: "make it survive layout changes"
3. Final format focuses on **semantic invariants** (SHP ADG keyword → 8-market sequence repeats → 2 stacked vertical sections → ADG above ADO)

Apply this format to ALL table descriptions. Key invariants to identify for any table:
- **Data type:** absolute values, SOB ratios, percentages, dollar amounts
- **Platform pair:** SHP/TTS, SHP/LZD, TTS-only, SHP-only
- **Metric family:** ADG, ADO, PC2, TR%, CIR%, multiple
- **Layout pattern:** stacked vertically (sections one below another), stacked horizontally (side by side with same row range), or transposed (markets in rows not columns)
- **Value format:** raw numbers, "x" suffix, % format
- **Date range** and direction (oldest-first vs newest-first)

### ⚠️ Verify Sheet IDs by Reading, Not Assuming

When the user provides a URL, always read the workbook title with the Sheets API (`api_get(SID, "properties.title")`) and compare against the expected name, even if the ID looks different from the registry. The actual working copy may have been re-shared or re-created with a new ID.

### ⚠️ Cross-Check Source: Only SHP/TTS Section Matters

The `(Final) Data from CF excel` tab has multiple stacked sections (SHP/TTS, SHP/LZD, transposed views). For the workflow's cross-check (rows 172:196 in the ADG ADO tab), **only the SHP/TTS ADG Multiple section** (first section, ~rows 1-26) is relevant. The section starts with a cell containing `"SHP/TTS ADG Multiple"` — scan for this keyword to find it dynamically.

### ⚠️ PC2 Tab is SHP-Only; TTS Data Elsewhere

The `SHP & TTS PC2` tab contains only Shopee (SHP) metrics despite its name. TTS PC2 data may be in a separate tab or workbook. When the workflow says "archive PC2 to [Archive]", it archives only SHP-side PC2 data.

### ⚠️ Weekly Live View: SHP and TTS Values Are Side-by-Side, Not Stacked

The source tab `[Weekly Live View] SHP/TTS ADG ADO` stores SHP and TTS values **in the same rows, different column ranges**:

| Col Range | Content |
|-----------|---------|
| D-K | SHP ADG absolute values |
| N-U | TTS ADG absolute values |

Both SHP and TTS ADG live in rows 4-39 of the same section. The ADO section (rows 41+) follows the same pattern.

However, the destination tab `[Reg CNLS copy]` has them **stacked vertically** — SHP ADG in rows 4-43, TTS ADG in a separate section below.

**When pasting Step 2 values:** Read SHP ADG from source cols D-K → dest SHP ADG section. Read TTS ADG from source cols N-U (same rows) → dest TTS ADG section. Do NOT look for TTS section headers in the source — there aren't any.

### ⚠️ Structural Gap in Destination: Dec'25 ↔ Jan'26

The destination tab's SHP ADG raw data is **not continuous**. Between Dec'25 (row 27) and Jan'26 (row 32), there are structural header rows (e.g., "1st Copy") that break positional mapping.

When pasting fresh source values:
1. **Do NOT assume rows are continuous** — the destination has a 4-row gap between the Dec'25 data block and the Jan'26+ data block
2. **Match by month label (col B)** across the entire section — scan all rows, skip non-data labels ("1st Copy", "Q1", "Q2", etc.)
3. For non-contiguous rows, use `values:batchUpdate` API instead of a single range update
4. The source has 36 months of data; the destination expects data at specific row positions regardless of the gap

### ⚠️ Cross-Check ADO Section: Transposed and Incomplete

The Reg Commercial Team source tab `(Final) Data from CF excel` has ADG and ADO in different layouts:

**ADG section** (rows 1-26): Standard row-based format — dates in rows, markets in columns. Always use this as the source for ADG in the cross-check (rows 172:196, cols D-K).

**ADO section** (row 71+): **Transposed** — markets in rows, dates in columns. Only covers 12 months (2025-01 to 2025-12) and 7 markets (SG, SEA excl TW, MY, TH, VN, PH, ID — missing "SEA excl TW ID").

When building the cross-check ADO section (rows 172:196, cols N-U):
1. For dates within 2025: read from the transposed ADO source, mapping source market names to destination market columns
2. For dates outside 2025: **fall back to existing ADO values** from the duplicated tab — the Reg Commercial Team has no reference data for these periods
3. Market "SEA excl TW ID" has no source data in the ADO section — preserve from duplication or leave empty

```python
# Transposed ADO parsing pattern
ado_src = {}  # {date: {market: value}}
ado_start = index_of_row_containing("SHP/TTS ADO Multiple", all_rows)
dates = all_rows[ado_start][1:]  # column headers = dates
for r in range(ado_start + 1, ado_start + 8):
    market = all_rows[r][0]
    for di, dt in enumerate(dates):
        ado_src.setdefault(dt, {})[market] = all_rows[r][di + 1]
```

### ⚠️ Verifier Must Check Full Cross-Check (ADG + ADO)

When verifying Step 3 (S3-5), the verifier must check **both** parts of the cross-check section against the Reg Commercial Team source:

- **ADG** (cols D-K): Compare against source rows 2-26, cols B-I — 200 cells (25 rows × 8 markets)
- **ADO** (cols N-U): Compare against transposed source — only for 2025 dates where source data exists (84 cells: 12 dates × 7 markets)

A common bug is only checking ADG and tacitly passing the ADO part. The verifier must explicitly compare ADO as well, recognizing that cells outside the source's date range are expected to contain fallback values (not mismatches).

### ⚠️ Verifier Cascade: SKIP vs PASS on Missing Tabs

When the executor hasn't run (e.g., non-Thursday), today's tab doesn't exist. The verifier must distinguish between:

- **PASS**: Check ran and all conditions were met
- **FAIL**: Check ran and conditions were not met  
- **SKIP**: Check could not run because a prerequisite is missing

Steps 3, 4, 5, and PC2-1 depend on Step 2 creating the today's tab. If the tab doesn't exist, these steps should report `[SKIP (no tab)]`, not `[PASS]` or `[FAIL]`. A flat `PASS` on a step that never ran is misleading — it appears to confirm correct execution when no execution occurred.

---

## Adding New Month Rows to ADG/ADO SOB Tabs

A recurring maintenance task: insert a new projection month (e.g. Jun'26) below May in all tables, rename the previous projection month to actual, and apply correct formulas.

### Critical: Write Formulas, Not Static Values

The SOB sheets have two layers of interconnected formulas — do NOT copy static values. Always replicate the formula pattern:

**Reporting Tables (SOB ratios):** Each cell is a division formula referencing Raw Data rows:
- Table 1 (CNCB & CNLS): `=C{Shopee_Jun}/C{TTS_Jun}` — divides Shopee by TTS for the same market/column
- Table 2 (CNCB): `=C{Shopee_CNCB_Jun}/C{TTS_CNCB_Jun}`
- Table 3 (CNLS): `=C{Shopee_CNLS_Jun}/C{TTS_CNLS_Jun}`

**Raw Data Tables (absolute values):**
- **CNCB+CNLS:** Platform columns = static values (from external source). Sub-group columns = `=CNCB + CNLS` (e.g. `=D88+D101`). Regional = SUM formulas.
- **CNCB:** Platform columns = `=CNCB+CNLS` (reference formula). Sub-group columns = static values.
- **CNLS:** Platform columns = empty (CNLS has no SG/Platform). Sub-group columns = static values.

**Static values** (external source data): copy from the previous month's row as placeholders — the user replaces with actual projection data.

### Adding a New Month — Step by Step

**1. Rename labels:** Update `May'26 (Proj.)` to `May'26` in all 9 tables (column B).

**2. Insert rows bottom-to-top:** Use `insertDimension` in a single `batchUpdate`, processing from the lowest table upward so insertions don't cascade:

```python
# 0-indexed insertion positions for ADO SOB tab (after May row):
insert_positions = [139, 127, 114, 101, 88, 75, 56, 37, 16]  # TTS CNLS → Table 1
```

Set `inheritFromBefore: False` to avoid inheriting formatting.

**3. Write formulas per table pattern:**

| Table | May | Jun | Formula Pattern |
|-------|-----|-----|----------------|
| T1 CNCB&CNLS SOB | 15 | 16 | `=C75/C114` (Shopee/TTS CNCB+CNLS) |
| T2 CNCB SOB | 36 | 37 | `=C88/C127` (Shopee/TTS CNCB) |
| T3 CNLS SOB | 55 | 56 | `=F101/F140` (Shopee/TTS CNLS) |
| Shopee CNCB+CNLS | 74 | 75 | Sub-group: `=D88+D101` |
| TTS CNCB+CNLS | 113 | 114 | Sub-group: `=D127+D140` |

**4. Copy formatting:** Use `copyPaste` with `pasteType: 'PASTE_FORMAT'` from May row to Jun row. This preserves formulas while applying background colors, bold, and number formats.

### Column Layout

23 columns (A-W): A=Date, B=Month, C-E=SG, F-H=MY, I-K=TH, L-N=VN, O-Q=PH, R-T=ID (CNLS, not CNCB+CNLS), U-W=Regional. Each market: Platform, sub-group, vs Platform.

### Verification

Check FORMATTED_VALUE renders correctly (e.g. `3.96 x` not `=C75/C114`). Verify sub-group formulas sum correctly. Flag any pre-existing #REF! errors to the user.
```
