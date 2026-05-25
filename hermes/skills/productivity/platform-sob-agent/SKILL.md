---
name: platform-sob-agent
description: Weekly Platform SOB & PC2 update workflow — reads SOB data from Google Sheets, cross-checks, validates, and archives. Runs every Thursday 13:00 Asia/Shanghai.
version: 1.1.0
author: Hermes Agent
tags: [google-sheets, sob, pc2, weekly, archive]
---

# Platform SOB Agent

Weekly workflow to update Platform Share-of-Business (SOB) data and PC2 data across multiple Google Sheets, validate cross-source consistency, and archive the results.

**Schedule:** Thursday 13:00 Asia/Shanghai (weekly)

**Prerequisite:** `google-sheets-intelligence` skill (loaded automatically). Google Sheets OAuth must be set up.

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
| `[Reg Commercial Team]` | Reference PC2 data — tab: `(Final) Data from CF excel` |
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
4. Present section map to user for confirmation

### Step 1 — Read & Validate Source Data Structure

Read table structure of:
- `[Weekly Live View] SHP/TTS ADG ADO` in `[Weekly Live View]`
- `[Weekly Live View] SHP/TTS Clusters` in `[Weekly Live View]`
- `(Final) Data from CF excel` in `[Reg Commercial Team]`

Use `$GSI scan-sections` and present detected structure to user.

### Step 2 — Copy SOB Values (ADG ADO) to [Reg CNLS copy]

1. In the source data tab, identify the **red section** (cells with **red borders** via `effectiveFormat.borders`). Using `effectiveFormat.backgroundColor` is WRONG — the "red section" means red borders, not red background fill.
   - If no red borders found, ask the user what range to use as the source data.
   - Common fallback: `Sheet6` or `Sheet5` tab in `[Weekly Live View]`, range `D4:V27`.
2. In `[Reg CNLS copy]`, find the most recent tab (format: `YY/MM/DD ADG ADO`). See ⚠️ Date parsing below.
3. Duplicate that tab, name it `{YY}/{MM}/{DD} ADG ADO` with today's date.
4. Paste the source values as values only (after user approval).

### Step 3 — Paste SOB Values Within [Reg CNLS copy] Tab

In the newly duplicated tab:
1. Copy rows ~98:130 → rows ~137:169 (SOB calculation section → results section) as values only
2. Copy values from `[Reg Commercial Team]` tab → rows ~172:196 of the new tab (cross-check section)

Both require user approval before execution.

### Step 4 — Validate Zero Check

Read the difference table (~rows 199:223):
- ✅ All values = 0: notify user, proceed to Step 5
- ❌ Any value ≠ 0: **halt**, alert user with details of non-zero cells

### Step 5 — Archive SOB to [Archive]

Conditions (both must be true):
- Step 4 zero-check passed
- SOB values (rows ~137:169) match exactly with `[Reg Commercial Team]` values

Action: duplicate most recent `SOB-YYMMDD` tab in `[Archive]`, name it `SOB-{YYMMDD}`, paste as values.

### Step 6 — Copy Clusters Data to [Reg CNLS copy]

From `SHP/TTS Clusters` columns E:S (dynamic last row), duplicate most recent `By cluster` tab in `[Reg CNLS copy]`, paste as values.

### PC2 Step 1 — Archive PC2 to [Archive]

From `(Final) Data from CF excel` (dynamic range), duplicate most recent `PC2-YYMMDD` tab in `[Archive]`, paste as values.

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
| Row 2-26, Col A | Col C (serial date) | Date as serial number (days since 1899-12-30) |
| Row 2-26, Col B | Col D | SG ADG SOB value |
| Row 2-26, Col C | Col E | MY ADG SOB value |

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

### ⚠️ "Red Section" = Red Borders, Not Background

The agent spec originally described the red section using `backgroundColor`, but the actual visual marker is **red borders** (`effectiveFormat.borders`). These are completely different API fields:

```python
# ❌ WRONG — queries background fill color
fields="...effectiveFormat.backgroundColor..."

# ✅ CORRECT — queries border colors
fields="...effectiveFormat.borders..."
```

When in doubt, read both and ask the user which they mean.

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

### Reading red-bordered section (Step 2 — "red section" detection)

The "red section" refers to cells with **red borders**, not red background fill. Use `gsheets_util.py` to query `effectiveFormat.borders`:

```python
import sys; sys.path.insert(0, "/Users/apple/.hermes/skills/productivity/platform-sob-agent")
import gsheets_util as gsu

SID = gsu.SHEETS["Weekly Live View"]
tab_name = "[Weekly Live View] SHP/TTS ADG ADO"

# Use the direct API with includeGridData
headers = gsu.get_auth_headers()
import urllib.request, json
url = f"https://sheets.googleapis.com/v4/spreadsheets/{SID}?ranges={urllib.parse.quote(f\"'{tab_name}'\")}&fields=sheets.data.rowData.values(effectiveFormat.borders,effectiveValue,formattedValue)"
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
```
