---
name: platform-sob-agent
description: Weekly Platform SOB & PC2 update workflow — reads SOB data from Google Sheets, cross-checks, validates, and archives. Runs every Thursday 13:00 Asia/Shanghai.
version: 1.7.0
author: Hermes Agent
tags: [google-sheets, sob, pc2, weekly, archive]
---

# Platform SOB Agent — Executor

Weekly workflow to update Platform Share-of-Business (SOB) data and PC2 data across multiple Google Sheets, validate cross-source consistency, and archive the results.

**Schedule:** Thursday 13:00 Asia/Shanghai (weekly)

**Prerequisite:** `google-sheets-intelligence` skill (loaded automatically). Google Sheets OAuth must be set up.

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
7. If the user provides a sheet URL, verify the workbook title via `gsu.api_get(SID, "properties.title")` and compare against the expected name — the working copy may have been re-shared or re-created with a new ID

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

### Step 2 — Copy ADO & ADG Values to [Reg CNLS copy]

1. From the source data tab (`[Weekly Live View] SHP/TTS ADG ADO`), read the relevant ADG & ADO values.
2. In `[Reg CNLS copy]`, find the most recent tab (format: `YY/MM/DD ADG ADO`). Tab names use two formats: pre-2026 `MM/DD` and 2026+ `YY/MM/DD` — use `gsu.parse_tab_date(tab_name)` to get a sortable `YYYYMMDD` string. Only match standard-format tabs — `As of...` tabs (e.g. `As of 26/05/18 ADG ADO`) are a different context and must not block creation of today's tab.
3. Duplicate that tab, name it `{YY}/{MM}/{DD} ADG ADO` with today's date.
4. In the new tab, use the **red-bordered section** (`effectiveFormat.borders`) in the **destination tab** as a positional reminder for the paste target range. If no red borders found, fall back to rows ~137:169 and match by month label and market column name.
5. Paste ADG & ADO values of both SHP and TTS **as values only**, aligned **month by month and market by market**: match each source row by its month label and each source column by its market name to the corresponding row/column in the destination. If the red section dimensions do not match the source layout, fall back to label-based matching — do not assume positional identity. Present the mapping to the user and wait for confirmation before pasting.

### ⚠️ Step 2 Implementation Details — CRITICAL LEARNINGS

#### API Batch Size Limits (HTTP 400 on large writes)

`values:batchUpdate` with more than ~15 data entries causes **HTTP 400 Bad Request** (Google's undocumented batch size limit). Always batch writes in chunks of **10 or fewer** entries. Chunks of 5-10 are reliable; 20+ fails even with valid data.

**Pattern that works:**
```python
batch_size = 10
for start in range(0, len(all_writes), batch_size):
    chunk = all_writes[start:start+batch_size]
    resp = values_batch_write(SID, {"valueInputOption": "USER_ENTERED", "data": chunk})
```

**Write full rows, not separate column ranges:** When writing ADG values to the destination, write each row as a single range (e.g. `'Tab'!D{dest_row}:U{dest_row}` = 18 cols) rather than separate D:K and N:U ranges. This halves the API calls and reduces batch-induced errors. The destination has SHP and TTS side by side in the same rows, so D:U covers SHP (D-K) + blanks (L-M) + TTS (N-U).

#### Tab Name Quoting (HTTP 400 Errors)

Source tab names with brackets (e.g. `[Weekly Live View] SHP/TTS ADG ADO`) cause HTTP 400 errors with `gsu.values_get()` and `urllib.parse.quote()` on the full range string. **Fix:** Use `values:batchGet` endpoint with the `ranges` parameter and `doseq=True` for reading:

```python
# ✅ WORKS — batchGet with ranges param
url = f"https://sheets.googleapis.com/v4/spreadsheets/{SID}/values:batchGet"
params = urllib.parse.urlencode({"ranges": [range_string]}, quote_via=urllib.parse.quote, doseq=True)
```

For **writing**, use `values:batchUpdate` (POST) — ranges go in the JSON body, not the URL:

```python
url = f"https://sheets.googleapis.com/v4/spreadsheets/{SID}/values:batchUpdate"
body = json.dumps({"valueInputOption": "USER_ENTERED", "data": data_updates}).encode()
```

This avoids all URL quote-related 404/400 errors. Also add 503 retry logic (3 attempts, 2s delay).

#### Red-Bordered Section for Paste Target

The red-bordered section defines the **raw ADG/ADO paste target** (rows 4-39, cols D-K for SHP, N-U for TTS). Use `includeGridData=True` on a **single specific tab range** to avoid timeout. Full grid data on the entire workbook times out at 120s.

**Scope to one tab:**
```python
url = f"https://sheets.googleapis.com/v4/spreadsheets/{SID}?ranges={urllib.parse.quote('TabName!A1:V250')}&fields=sheets.data.rowData.values(effectiveFormat.borders,formattedValue)&includeGridData=true"
```

Red border threshold: `red > 0.8 AND green < 0.3 AND blue < 0.3`.

#### Cross-Check Column Structure (MUST match exactly)

The cross-check section (rows 172-196) has a specific column layout expected by diff formulas:

| Col | Content | Source Reg Commercial Team |
|-----|---------|---------------------------|
| C | Date | Source A |
| D-K | ADG SOB (8 markets) | Source B-I |
| L | Blank | — |
| M | Blank | — |
| N | Date (ADO section) | Same as C |
| O-V | ADO SOB (8 markets) | Source K-R (7 markets available — leave 8th empty) |

⚠️ **Diff formulas use `=O172-O137` for ADO (col O, not N).** Col N contains a date in both sections. Source has 7 ADO markets (no SEAexTWID ADO) — dest expects 8.

**Write pattern:**
```python
write_row = [date_val] + adg_vals + ['', '', date_val] + ado_vals
# Range: f"'{TabName}'!C{row}:V{row}" (20 cols)
```

#### "x" Suffix Stripping — Order Matters

Both **results** (rows 137-169) and **cross-check** (rows 172-196) have "x" suffixes. Strip AFTER writing both sections or diff formulas produce `#VALUE!`:

```python
cleaned = [v.replace('x', '').strip() if v and isinstance(v, str) and v.endswith('x') else (v or '') for v in row]
```

Do NOT strip from date columns (C, N) or blanks (L, M).

#### ADO Section — Different from ADG

- **No month labels in col B** — dates are in col C as serial numbers (e.g. `45292`)
- **No "Period" header row** — data starts right after section header
- Match by date (col C), not month label (col B)

#### Structural Gaps

After Dec'25, non-data rows ("1st Copy", Q1-Q4 averages) appear before Jan'26 data. Skip rows where col B contains "Copy" or starts with "Q". Match by month label, not row index.

#### API Write Method Preference

| Operation | Method | Notes |
|-----------|--------|-------|
| Read values | `values:batchGet` with `ranges` param | Handles special chars, `doseq=True` |
| Write values | `values:batchUpdate` (POST) | Ranges in JSON body, no URL issues |
| Tab duplication | `sheets().batchUpdate` with `duplicateSheet` | Structural change |

Never use `values().update()` (PUT) — fails on URL-encoded single quotes in tab names.

**The destination has structural gaps between data blocks.** After Dec'25, there are non-data rows (headers like "1st Copy", "2nd Copy") before Jan'26 starts. Do NOT assume contiguous rows — skip rows where col B contains "Copy", "Q1"–"Q4", or is empty, and match each row by its month label, not by row index.

**Use `gsu.values_batch_write(sheet_id, batch_data)` for non-contiguous rows.** Each entry: `{'range': "'Tab'!D5:K5", 'majorDimension': 'ROWS', 'values': [[...]]}`. Note: this calls `values:batchUpdate` (multi-range value writes) — distinct from `gsu.api_batch_update` (structural changes like tab duplication).

**Source has SHP and TTS side by side; destination has them stacked.** The source places SHP ADG (cols D-K) and TTS ADG (cols N-U) in the same rows. The destination separates them into two vertically stacked sections (SHP ADG at rows ~4-39, TTS ADG at rows ~44-79). When reading TTS values, use the same source rows but shift to cols N-V.

**ADO sections use col C as month labels** — col B is empty. The ADO section has two data blocks separated by Q-summary formula rows:
  - **2024-2025 block**: rows ~52-75 (col C = month labels)
  - **Q-summary block (SKIP)**: rows ~76-79 (col B="2nd Copy", col C=Q1-Q4 labels, formula rows in D-K and N-V)
  - **2026 block**: rows ~80-91 (col C = month labels)
  When iterating `A50:U95`, use `r = 50 + offset` (NOT `+1` — offset is 0-based within the range). Match each data row by month label in col C directly. Skip formula rows by checking if any cell in the formula row starts with "=".

**⚠️ Critical: after writing ADO data, the SOB calc formulas (rows ~98-130) recalculate automatically.** Must re-run Step 2 (copy calc→results) to pick up the updated values. Re-run the zero check after.

**Cross-check ADO note:** The `(Final) Data from CF excel` source has only **7 ADO markets** (cols K-Q: SG, MY, TH, ID, VN, PH, SEA excl TW), NOT 8. Col R is empty. When building the cross-check write row, use 8 ADO slots (leave the 8th empty):

```python
# Source indices: K=10=SG, L=11=MY, M=12=TH, N=13=ID, O=14=VN, P=15=PH, Q=16=SEAexTW, R=17=(empty)
ado_vals = [str(r[i]) for i in range(10, 18)]  # 8 values, last may be empty
write_row = [date_val] + adg_vals + ['', '', date_val] + ado_vals
# Range: C{row}:V{row} = 20 cols
```

**⚠️ Diff formula structure — ADO uses cols O-U, NOT N-U:**
The diff formulas are:
- ADG diff: `=D172-D137` through `=K172-K137` (cols D-K)
- ADO diff: `=O172-O137` through `=U172-U137` (cols O-U, skipping N)
Col N contains a date in BOTH sections (results row 137 and cross-check row 172), so `=N172-N137` would compare date serial numbers and give 0 — but that's NOT the ADO check. The real ADO check starts at col O.

If you misalign the cross-check data by even 1 column (putting SG ADO at col N instead of col O), the ADO diff will show huge numbers from date-serial subtraction.

**Red-bordered section:** Use `gsu.get_cell_formats(sheet_id, tab_name)` → `rowData` list. Call `gsu.is_red_border(cell["effectiveFormat"]["borders"])` on each cell and scan all rows to build the contiguous paste target range. Query the destination tab, not the source.

### Step 3 — Paste SOB Values Within [Reg CNLS copy] Tab

In the newly duplicated tab:
1. Copy rows ~98:130 → rows ~137:169 (SOB calculation section → results section) as values only, all columns through V (including SEA excl TW SG ADO SOB in col V)
2. Copy SHP/TTS ADG SoB & ADO SoB values from the `[Reg Commercial Team]` tab → rows ~172:196 of the new tab (cross-check section), matching the same markets and same month range as the corresponding table already in the cross-check section of the duplicated tab. Column mapping from source `(Final) Data from CF excel` rows 2-26:

   | Source col | Destination col | Content |
   |---|---|---|
   | A | C | Month label — paste directly |
   | B | D | SG ADG SOB |
   | C | E | MY ADG SOB |
   | D | F | TH ADG SOB |
   | E | G | ID ADG SOB |
   | F | H | VN ADG SOB |
   | G | I | PH ADG SOB |
   | H | J | SEA excl TW ADG SOB |
   | I | K | SEA excl TW ID ADG SOB |
   | K | N | SG ADO SOB |
   | L | O | MY ADO SOB |
   | M | P | TH ADO SOB |
   | N | Q | ID ADO SOB |
   | O | R | VN ADO SOB |
   | P | S | PH ADO SOB |
   | Q | T | SEA excl TW ADO SOB |
   | R | U | SEA excl TW ID ADO SOB |

   ⚠️ Source column letters above are indicative — always locate columns dynamically by scanning for market header labels. Do not hardcode column indices.

   The tab has multiple stacked sections — **only the SHP/TTS ADG Multiple section** (first section, ~rows 1-26) is relevant. Scan for a cell containing `"SHP/TTS ADG Multiple"` to locate it dynamically.

Both require user approval before execution.

⚠️ **After pasting, strip "x" suffix from all values in rows 137:169 and 172:196** so the difference formulas can compute correctly. Write clean numbers (e.g. `3.51` not `"3.51x"`).

### Step 4 — Validate Zero Check

Read the difference table (~rows 199:223).

**IMPORTANT — `values_get()` returns formatted strings, not raw numbers** (e.g. `"0.00"`, `"-31.56"`, `"#VALUE!"`). Convert to float before comparing; treat `#VALUE!`/`#REF!` as errors, not zeros. Use `gsu.scan_nonzero(sheet_id, range_)` which returns `(errors, nonzero)` lists.

- ✅ All values = 0: notify user, proceed to Step 5
- ❌ Any non-zero value: **halt**, alert user with details of which cells differ. If the user explicitly overrides and instructs to proceed, continue to Step 5 and re-paste the updated results values into the Archive tab after it is created.
- ❌ Any `#VALUE!` or `#REF!` error: flag to user — likely caused by "n/a" in cross-check source or format mismatch

### Step 5 — Archive SOB to [Archive]

Conditions (both must be true):
- Step 4 zero-check passed

Action: duplicate most recent `SOB-YYMMDD` tab in `[Archive]`, name it `SOB-{YYMMDD}`. Read today's ADG ADO tab `D137:L169` (9 market cols D-L) from `[Reg CNLS copy]` → paste as values into the new archive tab `B2:J34` (cols B-J, rows 2-34). Apply `gsu.clean_val(v)` to strip "x" suffix before writing.

Archive tab structure (36 rows):
- **Row 1:** Headers — A=period label, B-J=market names (SG, MY, TH, ID, VN, PH, SEA excl TW, SEA excl TW ID, SEA excl TW SG)
- **Rows 2-34:** Monthly data (Dec'24 → Q4'26 Target)
- **Row 36:** Metadata row with commercial update time and archive reference

### Step 6 — Copy Clusters ADG & ADO Data to [Reg CNLS copy]

Read `E4:K1226` from `SHP/TTS Clusters` in `[Weekly Live View]` (note: tab name has a trailing space — `'SHP/TTS Clusters '`; `gsu.values_get` URL-encodes it automatically). Duplicate most recent `By cluster` tab in `[Reg CNLS copy]`, then paste the copied values into the new tab at the same position `E4:K1226` as values only. Source and destination use the same row and column — do not offset.

What is pasted is ADG and ADO **values** (not SOB ratios). Apply `gsu.clean_val(v)` on each value before writing.

### PC2 Step 1 — Archive PC2 to [Archive] (⚠️ Formatting Critical)

Must copy **both values AND formatting** from `SHP & TTS PC2` in `[Platform PC2]`. Since source and destination are different spreadsheets, `copyPaste` alone won't work — use this 7-step workflow:

1. **Find source sheetId** — `gsu.api_get(PC2_SID, "sheets.properties")` → locate `SHP & TTS PC2`, note its `sheetId`

2. **Delete any existing `PC2-{YYMMDD}`** tab in `[Archive]` — `gsu.api_batch_update(ARCHIVE_SID, [{"deleteSheet": {"sheetId": existing_id}}])`

3. **`copyTo` — bring source into Archive** (the only cross-spreadsheet way to carry formatting) — `gsu.sheet_copy_to(PC2_SID, src_sheetId, ARCHIVE_SID)` → note the returned `sheetId` as `temp_sheet_id`

4. **Create final tab at correct position** — use `gsu.find_newest_tab(ARCHIVE_SID, "PC2-")` to get `(name, newest_pc2_id, newest_pc2_index)`, then `gsu.api_batch_update(ARCHIVE_SID, [{"duplicateSheet": {"sourceSheetId": newest_pc2_id, "insertSheetIndex": newest_pc2_index + 1, "newSheetName": "PC2-{YYMMDD}"}}])` → note `dest_sheet_id` from `replies[0]["duplicateSheet"]["properties"]["sheetId"]`

5. **Copy formatting** — `gsu.api_batch_update(ARCHIVE_SID, [{"copyPaste": {"source": {"sheetId": temp_sheet_id, "startRowIndex": 2, "endRowIndex": 505, "startColumnIndex": 0, "endColumnIndex": 75}, "destination": {"sheetId": dest_sheet_id, "startRowIndex": 0, "endRowIndex": 503, "startColumnIndex": 0, "endColumnIndex": 75}, "pasteType": "PASTE_FORMAT", "pasteOrientation": "NORMAL"}}])`. ⚠️ `PASTE_FORMAT` only — never `PASTE_VALUES`, it copies `#REF!` errors from broken cross-sheet references.

6. **Write values** — read `A1:BW2` (headers) and `A3:BW505` (data) from source via `gsu.values_get` (returns `FORMATTED_VALUE` by default). Write both to destination via `gsu.values_batch_write(ARCHIVE_SID, [...], input_option="USER_ENTERED")` — `USER_ENTERED` lets Google Sheets parse `"8.98%"` as a number with format intact, and does not overwrite formatting already applied by step 5.

7. **Delete temp copy** — `gsu.api_batch_update(ARCHIVE_SID, [{"deleteSheet": {"sheetId": temp_sheet_id}}])`

**Key pitfalls:**
- `copyPaste` only works within the same spreadsheet — `copyTo` (step 3) must come first to bring the source into Archive
- `moveSheet` fails with 400 on `copyTo`-created sheets — always position via `duplicateSheet` with `insertSheetIndex` (step 4)
- Write values (step 6) after `PASTE_FORMAT` (step 5) — order matters
- Read FORMATTED_VALUE from the ORIGINAL source spreadsheet, not from the temp copy (which already has broken references).

---



## Supporting Files

- `VERIFIER.md` — Verifier agent spec (audits executor output independently)
- `references.md` — Business concepts, sheet registry, tab/row references, table description format
- `gsheets_util.py` — Shared Google Sheets auth + API helpers (use for all API calls)
- `harness.md` — Infrastructure documentation (audit trail, permission gate, sandbox)
- `harness/` — Python modules:
  - `audit_trail.py` — RunRecord class + audit_log.md writer
  - `permission_gate.py` — WeChat notification + approval gate
  - `sandbox.py` — SandboxGuard for dry-run mode
  - `outcome_evaluator.py` — Final run evaluation + outcome report
- `eval/` — Evaluation scripts for verifier consistency and end-to-end system testing
- `verifier-implementation-guide.md` — Implementation notes for the verifier
- `wechat_config.json` — WeChat delivery configuration (currently: terminal mode)
- `audit_log.md` — Run history (one entry per Thursday)

## Key API Patterns

### ⚡ Raw API via `urllib` (preferred — faster, no timeout issues)

`gsheets_util.py` can **hang/timeout (60s+)** on large sheets or slow connections. The raw `urllib` approach below completes in **1-3s** and handles bracket tab names correctly. Use it for all read/write operations in this workflow.

**Token refresh (one-time at start):**

```python
import json, os, ssl, urllib.request, urllib.parse
import certifi
ctx = ssl.create_default_context(cafile=certifi.where())

token_path = os.path.expanduser("~/.hermes/google_token.json")
with open(token_path) as f: creds = json.load(f)

if creds.get("expired", True):
    refresh_data = urllib.parse.urlencode({
        "client_id": creds["client_id"], "client_secret": creds["client_secret"],
        "refresh_token": creds["refresh_token"], "grant_type": "refresh_token",
    }).encode()
    req = urllib.request.Request(creds["token_uri"], data=refresh_data,
        headers={"Content-Type": "application/x-www-form-urlencoded"})
    with urllib.request.urlopen(req, timeout=30, context=ctx) as resp:
        token_resp = json.loads(resp.read())
        creds["token"] = token_resp["access_token"]
        with open(token_path, "w") as f: json.dump(creds, f, indent=2)
```

**Reading values (use `values:batchGet` with `doseq=True`):**

```python
def batch_get(sid, ranges, render="FORMATTED_VALUE"):
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sid}/values:batchGet"
    range_params = "&ranges=".join(urllib.parse.quote(r, safe='') for r in ranges)
    full_url = f"{url}?ranges={range_params}&valueRenderOption={render}"
    req = urllib.request.Request(full_url, headers={"Authorization": f"Bearer {creds['token']}"})
    with urllib.request.urlopen(req, timeout=60, context=ctx) as resp:
        return json.loads(resp.read())
```

**Writing values (use `values:batchUpdate` POST — NOT `values().update()` PUT):**

PUT requests (like `gsu.values_update`) fail with **404** when tab names have single quotes. Always use the POST `batchUpdate` endpoint which puts ranges in the JSON body:

```python
def values_batch_write(sid, data_updates):
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sid}/values:batchUpdate"
    body = json.dumps({
        "valueInputOption": "USER_ENTERED",
        "data": data_updates  # [{"range": "'Tab'!A1:C3", "majorDimension": "ROWS", "values": [[1,2,3]]}, ...]
    }).encode()
    req = urllib.request.Request(url, data=body, headers={
        "Authorization": f"Bearer {creds['token']}", "Content-Type": "application/json"
    })
    with urllib.request.urlopen(req, timeout=60, context=ctx) as resp:
        return json.loads(resp.read())
```

**Tab duplication & structural changes (use `batchUpdate` POST):**

```python
def api_batch_update(sid, requests):
    url = f"https://sheets.googleapis.com/v4/spreadsheets/{sid}:batchUpdate"
    body = json.dumps({"requests": requests}).encode()
    req = urllib.request.Request(url, data=body, headers={
        "Authorization": f"Bearer {creds['token']}", "Content-Type": "application/json"
    })
    with urllib.request.urlopen(req, timeout=60, context=ctx) as resp:
        return json.loads(resp.read())
```

**Cross-spreadcraft sheet copy (carries formatting):**

```python
url = f"https://sheets.googleapis.com/v4/spreadsheets/{SOURCE_SID}/sheets/{src_sheetId}:copyTo"
body = {"destinationSpreadsheetId": DEST_SID}
req = urllib.request.Request(url, data=json.dumps(body).encode(), headers={
    "Authorization": f"Bearer {creds['token']}", "Content-Type": "application/json"
})
with urllib.request.urlopen(req, timeout=60, context=ctx) as resp:
    result = json.loads(resp.read())
```

**⚠️ 503 Retry:** Add 3-attempt retry with 2s delay for transient Google API errors (503 Service Unavailable happens frequently on busy sheets):

```python
for attempt in range(3):
    try:
        with urllib.request.urlopen(req, timeout=60, context=ctx) as resp:
            return json.loads(resp.read())
    except urllib.error.HTTPError as e:
        if e.code == 503 and attempt < 2: time.sleep(2); continue
        raise
```

### Using `gsheets_util.py` (fallback — may time out)

If the raw API setup is inconvenient, `gsheets_util.py` handles auth and SSL automatically but may **time out (60s+)** on large sheets:

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

### Duplicating a tab

Use `gsu.find_newest_tab(sheet_id, prefix)` → returns `(name, sheetId, index)` of the most recent tab. Pass `sheetId` and `index + 1` to `gsu.api_batch_update` with a `duplicateSheet` request. Never hardcode tab names or sheet IDs.

**Tab ordering:** when tabs are sorted newest-first, the newest tab has the **lowest index** — not the highest. Do NOT sort or filter by `sheetId`; sheetIds are arbitrary numbers, not positional. Always use the `index` field from `api_get`.

### Writing and reading values

- Write: `gsu.values_update(sheet_id, range_, values)` — `valueInputOption="USER_ENTERED"` by default, handles percentages, dates, and formulas
- Read (values): `gsu.values_get(sheet_id, range_)` → `result.get("values", [])`
- Read (formulas): `gsu.values_get_formula(sheet_id, range_)` → returns raw formula strings

---

### Verification

Check FORMATTED_VALUE renders correctly (e.g. `3.96 x` not `=C75/C114`). Verify sub-group formulas sum correctly. Flag any pre-existing #REF! errors to the user.
```
